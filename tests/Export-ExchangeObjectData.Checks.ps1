#Requires -Version 7
$ErrorActionPreference = 'Stop'

$script:ScriptPath = Join-Path $PSScriptRoot '..' 'Export-ExchangeObjectData.ps1'
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

# Functions used by Add-ExportError inside the script under test.
$script:Errors = [System.Collections.Generic.List[object]]::new()

# Load every function from the script under test into this scope.
. Import-ScriptFunctions

# ---------------------------------------------------------------------------
# Checks (appended by implementation tasks)
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

function Test-ProxyLookupLayerRemoved {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw
    Assert-True -Condition ($scriptText -notmatch 'Get-ProxyAddressRows') -Name 'Old Get-ProxyAddressRows removed'
    Assert-True -Condition ($scriptText -match 'Properties\s+[^|]*EmailAddresses') -Name 'Inventory requests EmailAddresses property'
}
Test-ProxyLookupLayerRemoved

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
if ($script:Failures.Count -gt 0) {
    Write-Host "$($script:Failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host "All $($script:Passes) checks passed." -ForegroundColor Green
exit 0
