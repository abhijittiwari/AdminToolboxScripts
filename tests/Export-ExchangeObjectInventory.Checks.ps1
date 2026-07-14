#Requires -Version 7
$ErrorActionPreference = 'Stop'

$script:ScriptPath = Join-Path $PSScriptRoot '..' 'Export-ExchangeObjectInventory.ps1'
$script:Failures = [System.Collections.Generic.List[string]]::new()
$script:Passes = 0

function Get-ScriptAst {
    $tokens = $null
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:ScriptPath, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count -gt 0) {
        throw "Parse errors in $($script:ScriptPath): $($parseErrors[0].Message)"
    }
    return $ast
}

function Import-ScriptFunctions {
    $ast = Get-ScriptAst
    $functions = $ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)
    foreach ($functionAst in $functions) {
        . ([scriptblock]::Create($functionAst.Extent.Text))
    }
}

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

# Error sink used by Add-ExportError inside the script under test.
$script:Errors = [System.Collections.Generic.List[object]]::new()

. Import-ScriptFunctions

# ---------------------------------------------------------------------------
# Checks
# ---------------------------------------------------------------------------

function Test-ConvertToProxyRows {
    $recipient = [pscustomobject]@{
        Identity           = 'Jane Doe'
        PrimarySmtpAddress = 'jane@contoso.com'
        EmailAddresses     = @('SMTP:jane@contoso.com', 'smtp:jane.doe@contoso.com', 'SPO:SPO_abc123')
    }
    $rows = @(ConvertTo-ProxyRows -Recipient $recipient)
    Assert-True -Condition ($rows.Count -eq 3) -Name 'ConvertTo-ProxyRows returns one row per address'
    Assert-True -Condition ($rows[0].AddressType -eq 'SMTP' -and $rows[0].ProxyAddress -eq 'jane@contoso.com' -and $rows[0].IsPrimarySmtp -eq $true) -Name 'ConvertTo-ProxyRows parses primary SMTP'
    Assert-True -Condition ($rows[1].IsPrimarySmtp -eq $false -and $rows[1].AddressType -eq 'smtp') -Name 'ConvertTo-ProxyRows keeps secondary smtp non-primary'
    Assert-True -Condition ($rows[2].AddressType -eq 'SPO' -and $rows[2].ProxyAddress -eq 'SPO_abc123') -Name 'ConvertTo-ProxyRows preserves non-SMTP prefixes'

    $emptyRecipient = [pscustomobject]@{ Identity = 'X'; PrimarySmtpAddress = ''; EmailAddresses = $null }
    Assert-True -Condition (@(ConvertTo-ProxyRows -Recipient $emptyRecipient).Count -eq 0) -Name 'ConvertTo-ProxyRows returns empty for null addresses'
}
Test-ConvertToProxyRows

function Test-RecipientTypeMatching {
    $userMailbox = [pscustomobject]@{ RecipientTypeDetails = 'UserMailbox'; RecipientType = 'UserMailbox' }
    $groupMailbox = [pscustomobject]@{ RecipientTypeDetails = 'GroupMailbox'; RecipientType = 'MailUniversalDistributionGroup' }
    $distributionGroup = [pscustomobject]@{ RecipientTypeDetails = 'MailUniversalDistributionGroup'; RecipientType = 'MailUniversalDistributionGroup' }
    $mailUser = [pscustomobject]@{ RecipientTypeDetails = 'MailUser'; RecipientType = 'MailUser' }
    $mailContact = [pscustomobject]@{ RecipientTypeDetails = 'MailContact'; RecipientType = 'MailContact' }

    Assert-True -Condition (Test-RecipientMatchesType -Recipient $userMailbox -SelectedRecipientType 'Mailbox') -Name 'Test-RecipientMatchesType accepts UserMailbox as Mailbox'
    Assert-True -Condition (-not (Test-RecipientMatchesType -Recipient $groupMailbox -SelectedRecipientType 'Mailbox')) -Name 'Test-RecipientMatchesType rejects GroupMailbox as Mailbox'
    Assert-True -Condition (Test-RecipientMatchesType -Recipient $groupMailbox -SelectedRecipientType 'Group') -Name 'Test-RecipientMatchesType accepts GroupMailbox as Group'
    Assert-True -Condition (Test-RecipientMatchesType -Recipient $distributionGroup -SelectedRecipientType 'Group') -Name 'Test-RecipientMatchesType accepts distribution group as Group'
    Assert-True -Condition (Test-RecipientMatchesType -Recipient $mailUser -SelectedRecipientType 'MailUser') -Name 'Test-RecipientMatchesType accepts MailUser'
    Assert-True -Condition (Test-RecipientMatchesType -Recipient $mailContact -SelectedRecipientType 'MailContact') -Name 'Test-RecipientMatchesType accepts MailContact'
    Assert-True -Condition (Test-RecipientMatchesType -Recipient $mailContact -SelectedRecipientType 'All') -Name 'Test-RecipientMatchesType accepts anything for All'
}
Test-RecipientTypeMatching

function Test-ConvertToExportObject {
    $recipient = [pscustomobject]@{
        Identity                  = 'Jane Doe'
        DisplayName               = 'Jane Doe'
        RecipientType             = 'UserMailbox'
        RecipientTypeDetails      = 'UserMailbox'
        PrimarySmtpAddress        = 'jane@contoso.com'
        Alias                     = 'jane'
        ExternalDirectoryObjectId = 'aaaa'
        Guid                      = [guid]'22222222-2222-2222-2222-222222222222'
        DistinguishedName         = 'CN=Jane'
    }
    $object = ConvertTo-ExportObject -Recipient $recipient
    Assert-True -Condition ($object.ExchangeObjectId -eq '22222222-2222-2222-2222-222222222222') -Name 'ConvertTo-ExportObject maps Guid to ExchangeObjectId'
    Assert-True -Condition ($object.MailboxType -eq 'UserMailbox') -Name 'ConvertTo-ExportObject derives MailboxType'
    Assert-True -Condition ($object.PrimarySmtpAddress -eq 'jane@contoso.com' -and $object.DistinguishedName -eq 'CN=Jane') -Name 'ConvertTo-ExportObject maps identity columns'
}
Test-ConvertToExportObject

function Test-ScriptStaticContract {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw
    Assert-True -Condition ($scriptText -match 'Errors-Inventory\.csv') -Name 'Inventory writes job-specific errors CSV'
    Assert-True -Condition ($scriptText -match 'Get-ConnectionInformation') -Name 'Inventory reuses an existing Exchange Online connection'
    Assert-True -Condition ($scriptText -notmatch 'Disconnect-ExchangeOnline') -Name 'Inventory leaves the Exchange Online session open for chained jobs'
    Assert-True -Condition ($scriptText -match 'Objects\.csv' -and $scriptText -match 'ProxyAddresses\.csv') -Name 'Inventory writes Objects.csv and ProxyAddresses.csv'
}
Test-ScriptStaticContract

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
if ($script:Failures.Count -gt 0) {
    Write-Host "$($script:Failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host "All $($script:Passes) checks passed." -ForegroundColor Green
exit 0
