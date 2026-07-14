#Requires -Version 7
$ErrorActionPreference = 'Stop'

$script:ScriptPath = Join-Path $PSScriptRoot '..' 'Export-ExchangeFullAccessPermissions.ps1'
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

function Test-FullAccessRows {
    $mailbox = [pscustomobject]@{ Identity = 'Jane Doe'; PrimarySmtpAddress = 'jane@contoso.com'; RecipientTypeDetails = 'UserMailbox'; RecipientType = 'UserMailbox' }
    $contact = [pscustomobject]@{ Identity = 'Ext Contact'; PrimarySmtpAddress = 'ext@fabrikam.com'; RecipientTypeDetails = 'MailContact'; RecipientType = 'MailContact' }

    $script:FullAccessLookups = [System.Collections.Generic.List[string]]::new()
    function Get-EXOMailboxPermission {
        param($Identity, $ErrorAction)
        $script:FullAccessLookups.Add([string]$Identity) | Out-Null
        @(
            [pscustomobject]@{ User = 'bob@contoso.com'; AccessRights = @('FullAccess'); IsInherited = $false; Deny = $false },
            [pscustomobject]@{ User = 'NT AUTHORITY\SELF'; AccessRights = @('FullAccess'); IsInherited = $false; Deny = $false },
            [pscustomobject]@{ User = 'inh@contoso.com'; AccessRights = @('FullAccess'); IsInherited = $true; Deny = $false },
            [pscustomobject]@{ User = 'ro@contoso.com'; AccessRights = @('ReadPermission'); IsInherited = $false; Deny = $false }
        )
    }

    $rows = @(Get-FullAccessRows -Recipient $mailbox)
    Assert-True -Condition ($script:FullAccessLookups[0] -eq 'jane@contoso.com') -Name 'Get-FullAccessRows uses primary SMTP address'
    Assert-True -Condition ($rows.Count -eq 1 -and $rows[0].Trustee -eq 'bob@contoso.com') -Name 'Get-FullAccessRows keeps only non-inherited non-SELF FullAccess'

    $rows = @(Get-FullAccessRows -Recipient $contact)
    Assert-True -Condition ($rows.Count -eq 0 -and $script:FullAccessLookups.Count -eq 1) -Name 'Get-FullAccessRows skips non-mailbox recipients without querying'
}
Test-FullAccessRows

function Test-MailboxTypeMapping {
    Assert-True -Condition ((Get-MailboxType -Recipient ([pscustomobject]@{ RecipientTypeDetails = 'SharedMailbox'; RecipientType = 'UserMailbox' })) -eq 'SharedMailbox') -Name 'Get-MailboxType keeps mailbox type details'
    Assert-True -Condition ((Get-MailboxType -Recipient ([pscustomobject]@{ RecipientTypeDetails = 'MailContact'; RecipientType = 'MailContact' })) -eq '') -Name 'Get-MailboxType returns empty for non-mailboxes'
}
Test-MailboxTypeMapping

function Test-ScriptStaticContract {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw
    Assert-True -Condition ($scriptText -match 'Errors-FullAccess\.csv') -Name 'FullAccess job writes job-specific errors CSV'
    Assert-True -Condition ($scriptText -match 'Get-ConnectionInformation') -Name 'FullAccess job reuses an existing Exchange Online connection'
    Assert-True -Condition ($scriptText -notmatch 'Disconnect-ExchangeOnline') -Name 'FullAccess job leaves the Exchange Online session open for chained jobs'
    Assert-True -Condition ($scriptText -match 'FullAccess\.csv') -Name 'FullAccess job writes FullAccess.csv'
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
