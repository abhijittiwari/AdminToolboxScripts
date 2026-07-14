#Requires -Version 7
$ErrorActionPreference = 'Stop'

$script:ScriptPath = Join-Path $PSScriptRoot '..' 'Export-ExchangeSendAsPermissions.ps1'
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

function Test-SendAsRows {
    $recipients = @(
        [pscustomobject]@{ Identity = 'Jane Doe'; PrimarySmtpAddress = 'jane@contoso.com'; RecipientTypeDetails = 'UserMailbox'; RecipientType = 'UserMailbox' },
        [pscustomobject]@{ Identity = 'No SendAs'; PrimarySmtpAddress = 'nosendas@contoso.com'; RecipientTypeDetails = 'UserMailbox'; RecipientType = 'UserMailbox' }
    )

    # Multi-object run with the EXO cmdlet available: per-recipient SMTP lookups, SELF and
    # inherited entries dropped.
    $script:ExoSendAsLookups = [System.Collections.Generic.List[string]]::new()
    function Get-EXORecipientPermission {
        param($Identity, $ResultSize, $ErrorAction)
        if ([string]::IsNullOrWhiteSpace([string]$Identity)) { throw 'No-identity EXO SendAs query should not be used.' }
        $script:ExoSendAsLookups.Add([string]$Identity) | Out-Null
        if ([string]$Identity -ne 'jane@contoso.com') { return @() }
        return @(
            [pscustomobject]@{ Identity = 'Jane Doe'; Trustee = 'bob@contoso.com'; AccessRights = @('SendAs'); IsInherited = $false },
            [pscustomobject]@{ Identity = 'Jane Doe'; Trustee = 'NT AUTHORITY\SELF'; AccessRights = @('SendAs'); IsInherited = $false },
            [pscustomobject]@{ Identity = 'Jane Doe'; Trustee = 'eve@contoso.com'; AccessRights = @('SendAs'); IsInherited = $true }
        )
    }
    function Get-RecipientPermission {
        param($Identity, $ResultSize, $ErrorAction)
        throw 'Legacy Get-RecipientPermission should not be used when Get-EXORecipientPermission exists.'
    }
    $rows = @(Get-SendAsRows -Recipients $recipients -UseOrgWideQuery $true)
    Assert-True -Condition ($script:ExoSendAsLookups.Count -eq 2 -and $script:ExoSendAsLookups[0] -eq 'jane@contoso.com' -and $script:ExoSendAsLookups[1] -eq 'nosendas@contoso.com') -Name 'Get-SendAsRows uses per-recipient EXO queries for multi-object runs'
    Assert-True -Condition ($rows.Count -eq 1 -and $rows[0].Trustee -eq 'bob@contoso.com') -Name 'Get-SendAsRows drops SELF and inherited entries'

    # Single-object mode also uses the SMTP lookup.
    $script:SendAsLookups = [System.Collections.Generic.List[string]]::new()
    function Get-EXORecipientPermission {
        param($Identity, $ResultSize, $ErrorAction)
        $script:SendAsLookups.Add([string]$Identity) | Out-Null
        @([pscustomobject]@{ Identity = 'Jane Doe'; Trustee = 'bob@contoso.com'; AccessRights = @('SendAs'); IsInherited = $false })
    }
    $rows = @(Get-SendAsRows -Recipients @($recipients[0]) -UseOrgWideQuery $false)
    Assert-True -Condition ($script:SendAsLookups[0] -eq 'jane@contoso.com') -Name 'Get-SendAsRows per-identity uses primary SMTP address'
    Assert-True -Condition ($rows.Count -eq 1 -and $rows[0].PrimarySmtpAddress -eq 'jane@contoso.com') -Name 'Get-SendAsRows maps row columns'
}
Test-SendAsRows

function Test-SelfTrustee {
    Assert-True -Condition (Test-IsSelfTrustee -Trustee 'NT AUTHORITY\SELF') -Name 'Test-IsSelfTrustee matches NT AUTHORITY\SELF'
    Assert-True -Condition (Test-IsSelfTrustee -Trustee 'S-1-5-10') -Name 'Test-IsSelfTrustee matches the SELF SID'
    Assert-True -Condition (-not (Test-IsSelfTrustee -Trustee 'bob@contoso.com')) -Name 'Test-IsSelfTrustee keeps real trustees'
}
Test-SelfTrustee

function Test-ScriptStaticContract {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw
    Assert-True -Condition ($scriptText -match 'Errors-SendAs\.csv') -Name 'SendAs job writes job-specific errors CSV'
    Assert-True -Condition ($scriptText -match 'Get-ConnectionInformation') -Name 'SendAs job reuses an existing Exchange Online connection'
    Assert-True -Condition ($scriptText -notmatch 'Disconnect-ExchangeOnline') -Name 'SendAs job leaves the Exchange Online session open for chained jobs'
    Assert-True -Condition ($scriptText -match 'SendAs\.csv') -Name 'SendAs job writes SendAs.csv'
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
