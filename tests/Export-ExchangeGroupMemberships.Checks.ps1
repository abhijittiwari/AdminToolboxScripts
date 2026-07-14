#Requires -Version 7
$ErrorActionPreference = 'Stop'

$script:ScriptPath = Join-Path $PSScriptRoot '..' 'Export-ExchangeGroupMemberships.ps1'
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

function Test-ImportObjectsCsv {
    $tempCsv = Join-Path ([System.IO.Path]::GetTempPath()) "objects-checks-$([guid]::NewGuid()).csv"
    @(
        [pscustomobject]@{ Identity = 'Jane Doe'; PrimarySmtpAddress = 'jane@contoso.com'; ExternalDirectoryObjectId = 'aaaa'; RecipientTypeDetails = 'UserMailbox'; RecipientType = 'UserMailbox' }
    ) | Export-Csv -Path $tempCsv -NoTypeInformation

    try {
        $rows = @(Import-ObjectsCsv -Path $tempCsv -RequiredColumns @('Identity', 'PrimarySmtpAddress', 'ExternalDirectoryObjectId'))
        Assert-True -Condition ($rows.Count -eq 1 -and $rows[0].Identity -eq 'Jane Doe') -Name 'Import-ObjectsCsv returns rows'

        $threw = $false
        try { Import-ObjectsCsv -Path $tempCsv -RequiredColumns @('Identity', 'MissingColumn') | Out-Null }
        catch { $threw = $_.Exception.Message -match 'MissingColumn' }
        Assert-True -Condition $threw -Name 'Import-ObjectsCsv names the missing column when validation fails'
    }
    finally {
        Remove-Item -Path $tempCsv -ErrorAction SilentlyContinue
    }
}
Test-ImportObjectsCsv

function Test-InvokeGraphBatch {
    # 45 requests -> 3 batch calls (20 + 20 + 5), all succeed.
    $script:BatchCalls = [System.Collections.Generic.List[object]]::new()
    $mockOk = {
        param($Method, $Uri, $Body)
        $script:BatchCalls.Add($Body.requests.Count) | Out-Null
        return @{ responses = @($Body.requests | ForEach-Object { @{ id = $_.id; status = 200; body = @{ value = @($_.url) } } }) }
    }
    $requests = @(1..45 | ForEach-Object { @{ Id = "r$_"; Url = "/users/$_/memberOf" } })
    $results = Invoke-GraphBatch -Requests $requests -GraphRequest $mockOk
    Assert-True -Condition ($script:BatchCalls.Count -eq 3 -and $script:BatchCalls[0] -eq 20 -and $script:BatchCalls[2] -eq 5) -Name 'Invoke-GraphBatch chunks into 20s'
    Assert-True -Condition ($results.Count -eq 45 -and $results['r7'].value[0] -eq '/users/7/memberOf') -Name 'Invoke-GraphBatch maps responses to request ids'

    # Permanent 404 -> OnItemError once, not in results.
    $script:ItemErrors = [System.Collections.Generic.List[string]]::new()
    $mock404 = {
        param($Method, $Uri, $Body)
        return @{ responses = @($Body.requests | ForEach-Object { @{ id = $_.id; status = 404; body = @{ error = @{ message = 'not found' } } } }) }
    }
    $onError = { param($id, $message) $script:ItemErrors.Add("$id|$message") | Out-Null }
    $results = Invoke-GraphBatch -Requests @(@{ Id = 'r9'; Url = '/gone' }) -GraphRequest $mock404 -OnItemError $onError
    Assert-True -Condition ($results.Count -eq 0 -and $script:ItemErrors.Count -eq 1 -and $script:ItemErrors[0] -eq 'r9|not found') -Name 'Invoke-GraphBatch reports permanent item failures once'

    # Always-throttled item exhausts MaxAttempts then errors.
    $script:ItemErrors = [System.Collections.Generic.List[string]]::new()
    $mockAlways429 = {
        param($Method, $Uri, $Body)
        return @{ responses = @($Body.requests | ForEach-Object { @{ id = $_.id; status = 429; headers = @{ 'Retry-After' = '0' }; body = @{} } }) }
    }
    $results = Invoke-GraphBatch -Requests @(@{ Id = 'r1'; Url = '/x' }) -GraphRequest $mockAlways429 -MaxAttempts 2 -OnItemError $onError
    Assert-True -Condition ($results.Count -eq 0 -and $script:ItemErrors.Count -eq 1) -Name 'Invoke-GraphBatch gives up after MaxAttempts'
}
Test-InvokeGraphBatch

function Test-GetMembershipRows {
    $script:MembershipUrls = [System.Collections.Generic.List[string]]::new()
    $recipients = @(
        [pscustomobject]@{ Identity = 'Jane Doe'; PrimarySmtpAddress = 'jane@contoso.com'; ExternalDirectoryObjectId = 'aaaaaaaa-0000-0000-0000-000000000001'; RecipientTypeDetails = 'UserMailbox'; RecipientType = 'UserMailbox' },
        [pscustomobject]@{ Identity = 'Sales Team'; PrimarySmtpAddress = 'salesteam@contoso.com'; ExternalDirectoryObjectId = 'aaaaaaaa-0000-0000-0000-000000000002'; RecipientTypeDetails = 'GroupMailbox'; RecipientType = 'MailUniversalDistributionGroup' },
        [pscustomobject]@{ Identity = 'Ext Contact'; PrimarySmtpAddress = 'ext@fabrikam.com'; ExternalDirectoryObjectId = 'aaaaaaaa-0000-0000-0000-000000000003'; RecipientTypeDetails = 'MailContact'; RecipientType = 'MailContact' },
        [pscustomobject]@{ Identity = 'No Entra'; PrimarySmtpAddress = 'x@contoso.com'; ExternalDirectoryObjectId = '' }
    )
    $mock = {
        param($Method, $Uri, $Body)
        if ($Method -eq 'POST') {
            return @{ responses = @($Body.requests | ForEach-Object {
                $script:MembershipUrls.Add([string]$_.url) | Out-Null
                if ($_.id -ne 'aaaaaaaa-0000-0000-0000-000000000001') {
                    @{ id = $_.id; status = 200; body = @{ value = @() } }
                }
                else {
                    @{ id = $_.id; status = 200; body = @{
                        value = @(
                            @{ '@odata.type' = '#microsoft.graph.group'; id = 'g1'; displayName = 'Sales DL'; mail = 'sales@contoso.com'; mailEnabled = $true; securityEnabled = $false },
                            @{ '@odata.type' = '#microsoft.graph.group'; id = 'g2'; displayName = 'Sec Only'; mail = $null; mailEnabled = $false; securityEnabled = $true },
                            @{ '@odata.type' = '#microsoft.graph.directoryRole'; id = 'r1'; displayName = 'Role'; mail = $null; mailEnabled = $false; securityEnabled = $false }
                        )
                        '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/page2'
                    } }
                }
            }) }
        }
        # nextLink page fetch
        return @{ value = @(@{ '@odata.type' = '#microsoft.graph.group'; id = 'g3'; displayName = 'Second Page DL'; mail = 'p2@contoso.com'; mailEnabled = $true; securityEnabled = $false }) }
    }
    $result = Get-MembershipRows -Recipients $recipients -GraphRequest $mock 3>$null
    $entra = @($result.Entra)
    $exchange = @($result.Exchange)
    # directoryObject has no memberOf navigation in Graph v1.0 metadata; requests must address
    # the concrete resource type or every batch item fails with 400.
    Assert-True -Condition (@($script:MembershipUrls | Where-Object { $_ -match '/directoryObjects/' }).Count -eq 0) -Name 'Get-MembershipRows never uses the invalid directoryObjects memberOf path'
    Assert-True -Condition ($script:MembershipUrls[0] -match '^/users/aaaaaaaa-0000-0000-0000-000000000001/memberOf\?') -Name 'Get-MembershipRows queries mailboxes via /users memberOf'
    Assert-True -Condition ($script:MembershipUrls[1] -match '^/groups/aaaaaaaa-0000-0000-0000-000000000002/memberOf\?') -Name 'Get-MembershipRows queries group recipients via /groups memberOf'
    Assert-True -Condition ($script:MembershipUrls[2] -match '^/contacts/aaaaaaaa-0000-0000-0000-000000000003/memberOf\?') -Name 'Get-MembershipRows queries mail contacts via /contacts memberOf'
    Assert-True -Condition ($entra.Count -eq 3) -Name 'Get-MembershipRows returns all groups as Entra rows'
    Assert-True -Condition ($exchange.Count -eq 2) -Name 'Get-MembershipRows filters mailEnabled groups into Exchange rows'
    Assert-True -Condition ($exchange[0].GroupIdentity -eq 'Sales DL (sales@contoso.com)') -Name 'Get-MembershipRows formats GroupIdentity with mail'
    Assert-True -Condition (@($entra | Where-Object { $_.GroupDisplayName -eq 'Second Page DL' }).Count -eq 1) -Name 'Get-MembershipRows follows nextLink pages'
    Assert-True -Condition (@($entra | Where-Object { $_.Identity -eq 'No Entra' }).Count -eq 0) -Name 'Get-MembershipRows skips recipients without ExternalDirectoryObjectId'
}
Test-GetMembershipRows

function Test-ScriptStaticContract {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw
    Assert-True -Condition ($scriptText -match 'Errors-GroupMemberships\.csv') -Name 'Memberships job writes job-specific errors CSV'
    Assert-True -Condition ($scriptText -match 'Get-MgContext') -Name 'Memberships job reuses an existing Graph context'
    Assert-True -Condition ($scriptText -notmatch 'Connect-ExchangeOnline') -Name 'Memberships job never connects to Exchange Online'
    Assert-True -Condition ($scriptText -match 'EntraGroupMemberships\.csv' -and $scriptText -match 'ExchangeGroupMemberships\.csv') -Name 'Memberships job writes both membership CSVs'
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
