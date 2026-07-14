#Requires -Version 7
$ErrorActionPreference = 'Stop'

$script:ScriptPath = Join-Path $PSScriptRoot '..' 'Export-ExchangeObjectDataWorkbook.ps1'
$script:Failures = [System.Collections.Generic.List[string]]::new()
$script:Passes = 0

function Assert-True {
    param(
        [Parameter(Mandatory = $true)] [bool]$Condition,
        [Parameter(Mandatory = $true)] [string]$Name
    )
    if ($Condition) {
        $script:Passes++
        Write-Host "PASS $Name" -ForegroundColor Green
    }
    else {
        $script:Failures.Add($Name) | Out-Null
        Write-Host "FAIL $Name" -ForegroundColor Red
    }
}

if (-not (Get-Module -ListAvailable ImportExcel)) {
    Write-Host 'SKIP ImportExcel module not installed; workbook checks require it.' -ForegroundColor Yellow
    exit 0
}
Import-Module ImportExcel -ErrorAction Stop

# ---------------------------------------------------------------------------
# Fixture: a CSV folder shaped like the export jobs' output.
# FullAccess.csv and the membership CSVs are intentionally missing.
# ---------------------------------------------------------------------------
$fixtureFolder = Join-Path ([System.IO.Path]::GetTempPath()) "workbook-checks-$([guid]::NewGuid())"
New-Item -Path $fixtureFolder -ItemType Directory | Out-Null

@(
    [pscustomobject]@{ Identity = 'Jane Doe'; DisplayName = 'Jane Doe'; RecipientType = 'UserMailbox'; RecipientTypeDetails = 'UserMailbox'; MailboxType = 'UserMailbox'; PrimarySmtpAddress = 'jane@contoso.com'; Alias = 'jane'; ExternalDirectoryObjectId = 'aaaa'; ExchangeObjectId = 'bbbb'; DistinguishedName = 'CN=Jane' },
    [pscustomobject]@{ Identity = 'Sales DL'; DisplayName = 'Sales DL'; RecipientType = 'MailUniversalDistributionGroup'; RecipientTypeDetails = 'MailUniversalDistributionGroup'; MailboxType = ''; PrimarySmtpAddress = 'sales@contoso.com'; Alias = 'sales'; ExternalDirectoryObjectId = 'cccc'; ExchangeObjectId = 'dddd'; DistinguishedName = 'CN=Sales' }
) | Export-Csv -Path (Join-Path $fixtureFolder 'Objects.csv') -NoTypeInformation

@(
    [pscustomobject]@{ Identity = 'Jane Doe'; PrimarySmtpAddress = 'jane@contoso.com'; AddressType = 'SMTP'; ProxyAddress = 'jane@contoso.com'; RawProxyAddress = 'SMTP:jane@contoso.com'; IsPrimarySmtp = 'True' },
    [pscustomobject]@{ Identity = 'Jane Doe'; PrimarySmtpAddress = 'jane@contoso.com'; AddressType = 'smtp'; ProxyAddress = 'jane2@contoso.com'; RawProxyAddress = 'smtp:jane2@contoso.com'; IsPrimarySmtp = 'False' },
    [pscustomobject]@{ Identity = 'Sales DL'; PrimarySmtpAddress = 'sales@contoso.com'; AddressType = 'SMTP'; ProxyAddress = 'sales@contoso.com'; RawProxyAddress = 'SMTP:sales@contoso.com'; IsPrimarySmtp = 'True' }
) | Export-Csv -Path (Join-Path $fixtureFolder 'ProxyAddresses.csv') -NoTypeInformation

@(
    [pscustomobject]@{ Identity = 'Jane Doe'; PrimarySmtpAddress = 'jane@contoso.com'; Trustee = 'bob@contoso.com'; AccessRights = 'SendAs'; IsInherited = 'False' }
) | Export-Csv -Path (Join-Path $fixtureFolder 'SendAs.csv') -NoTypeInformation

@(
    [pscustomobject]@{ Identity = 'Sales DL'; Stage = 'SendAs'; Operation = 'Get-EXORecipientPermission'; Message = 'boom' }
) | Export-Csv -Path (Join-Path $fixtureFolder 'Errors-SendAs.csv') -NoTypeInformation

@(
    [pscustomobject]@{ Identity = ''; Stage = 'RecipientLookup'; Operation = 'Get-EXORecipient'; Message = 'not found' }
) | Export-Csv -Path (Join-Path $fixtureFolder 'Errors-Inventory.csv') -NoTypeInformation

try {
    # ---------------------------------------------------------------------------
    # End-to-end run
    # ---------------------------------------------------------------------------
    $workbookPath = Join-Path $fixtureFolder 'ExchangeObjectData.xlsx'
    & $script:ScriptPath -InputFolder $fixtureFolder 3>$null | Out-Null
    Assert-True -Condition (Test-Path -Path $workbookPath) -Name 'Workbook created at default path'

    $sheetNames = @(Get-ExcelSheetInfo -Path $workbookPath | ForEach-Object { $_.Name })
    Assert-True -Condition ($sheetNames[0] -eq 'Summary' -and $sheetNames[1] -eq 'Consolidated') -Name 'Summary and Consolidated are the first sheets'
    Assert-True -Condition ('Objects' -in $sheetNames -and 'ProxyAddresses' -in $sheetNames -and 'SendAs' -in $sheetNames -and 'Errors' -in $sheetNames) -Name 'Workbook has one sheet per non-empty dataset'
    Assert-True -Condition ('FullAccess' -notin $sheetNames -and 'ExchangeGroupMemberships' -notin $sheetNames -and 'EntraGroupMemberships' -notin $sheetNames) -Name 'Workbook skips empty datasets'

    $consolidated = @(Import-Excel -Path $workbookPath -WorksheetName 'Consolidated')
    $jane = $consolidated | Where-Object { $_.Identity -eq 'Jane Doe' }
    $sales = $consolidated | Where-Object { $_.Identity -eq 'Sales DL' }
    Assert-True -Condition ($consolidated.Count -eq 2) -Name 'Consolidated sheet has one row per object'
    Assert-True -Condition ($jane.ProxyAddresses -eq 'SMTP:jane@contoso.com; smtp:jane2@contoso.com') -Name 'Consolidated joins proxy addresses'
    Assert-True -Condition ($jane.SendAsTrustees -eq 'bob@contoso.com') -Name 'Consolidated joins SendAs trustees'
    Assert-True -Condition ([string]$jane.HasErrors -eq 'False' -and [string]$sales.HasErrors -eq 'True') -Name 'Consolidated flags per-object errors from merged error CSVs'

    $errorRows = @(Import-Excel -Path $workbookPath -WorksheetName 'Errors')
    Assert-True -Condition ($errorRows.Count -eq 2) -Name 'Errors sheet merges every Errors*.csv'

    $summary = @(Import-Excel -Path $workbookPath -WorksheetName 'Summary')
    $consolidatedSummary = $summary | Where-Object { $_.Dataset -eq 'Consolidated' }
    $fullAccessSummary = $summary | Where-Object { $_.Dataset -eq 'FullAccess' }
    Assert-True -Condition ([int]$consolidatedSummary.Rows -eq 2 -and [int]$fullAccessSummary.Rows -eq 0) -Name 'Summary sheet reports row counts including empty datasets'

    # ---------------------------------------------------------------------------
    # Overwrite behavior
    # ---------------------------------------------------------------------------
    $threw = $false
    try { & $script:ScriptPath -InputFolder $fixtureFolder 3>$null | Out-Null }
    catch { $threw = $_.Exception.Message -match 'Force' }
    Assert-True -Condition $threw -Name 'Existing workbook without -Force throws'

    & $script:ScriptPath -InputFolder $fixtureFolder -Force 3>$null | Out-Null
    $sheetNames = @(Get-ExcelSheetInfo -Path $workbookPath | ForEach-Object { $_.Name })
    Assert-True -Condition ('Consolidated' -in $sheetNames -and 'FullAccess' -notin $sheetNames) -Name 'Force overwrites the workbook without stale sheets'

    # ---------------------------------------------------------------------------
    # Missing Objects.csv fails fast
    # ---------------------------------------------------------------------------
    $emptyFolder = Join-Path ([System.IO.Path]::GetTempPath()) "workbook-checks-empty-$([guid]::NewGuid())"
    New-Item -Path $emptyFolder -ItemType Directory | Out-Null
    try {
        $threw = $false
        try { & $script:ScriptPath -InputFolder $emptyFolder 3>$null | Out-Null }
        catch { $threw = $_.Exception.Message -match 'Objects\.csv' }
        Assert-True -Condition $threw -Name 'Missing Objects.csv throws a clear error'
    }
    finally {
        Remove-Item -Path $emptyFolder -Recurse -Force -ErrorAction SilentlyContinue
    }
}
finally {
    Remove-Item -Path $fixtureFolder -Recurse -Force -ErrorAction SilentlyContinue
}

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
if ($script:Failures.Count -gt 0) {
    Write-Host "$($script:Failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host "All $($script:Passes) checks passed." -ForegroundColor Green
exit 0
