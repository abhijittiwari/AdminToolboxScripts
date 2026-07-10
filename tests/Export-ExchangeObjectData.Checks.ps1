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

function Test-InvokeGraphBatch {
    # 45 requests -> 3 batch calls (20 + 20 + 5), all succeed.
    $script:BatchCalls = [System.Collections.Generic.List[object]]::new()
    $mockOk = {
        param($Method, $Uri, $Body)
        $script:BatchCalls.Add($Body.requests.Count) | Out-Null
        return @{ responses = @($Body.requests | ForEach-Object { @{ id = $_.id; status = 200; body = @{ value = @($_.url) } } }) }
    }
    $requests = @(1..45 | ForEach-Object { @{ Id = "r$_"; Url = "/directoryObjects/$_/memberOf" } })
    $results = Invoke-GraphBatch -Requests $requests -GraphRequest $mockOk
    Assert-True -Condition ($script:BatchCalls.Count -eq 3 -and $script:BatchCalls[0] -eq 20 -and $script:BatchCalls[2] -eq 5) -Name 'Invoke-GraphBatch chunks into 20s'
    Assert-True -Condition ($results.Count -eq 45 -and $results['r7'].value[0] -eq '/directoryObjects/7/memberOf') -Name 'Invoke-GraphBatch maps responses to request ids'

    # One item throttled once (429 with Retry-After 0), succeeds on retry.
    $script:Attempt = 0
    $mockThrottle = {
        param($Method, $Uri, $Body)
        $script:Attempt++
        $responses = @($Body.requests | ForEach-Object {
            if ($_.id -eq 'r1' -and $script:Attempt -eq 1) {
                @{ id = $_.id; status = 429; headers = @{ 'Retry-After' = '0' }; body = @{ error = @{ message = 'throttled' } } }
            }
            else { @{ id = $_.id; status = 200; body = @{ value = @() } } }
        })
        return @{ responses = $responses }
    }
    $results = Invoke-GraphBatch -Requests @(@{ Id = 'r1'; Url = '/x' }, @{ Id = 'r2'; Url = '/y' }) -GraphRequest $mockThrottle
    Assert-True -Condition ($results.Count -eq 2 -and $script:Attempt -eq 2) -Name 'Invoke-GraphBatch retries throttled items'

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

    # Batch POST succeeds but returns no responses -> chunk requeued, then succeeds.
    $script:EmptyAttempt = 0
    $mockEmptyThenOk = {
        param($Method, $Uri, $Body)
        $script:EmptyAttempt++
        if ($script:EmptyAttempt -eq 1) { return @{ responses = @() } }
        return @{ responses = @($Body.requests | ForEach-Object { @{ id = $_.id; status = 200; body = @{ value = @() } } }) }
    }
    $results = Invoke-GraphBatch -Requests @(@{ Id = 'r1'; Url = '/x' }) -GraphRequest $mockEmptyThenOk
    Assert-True -Condition ($results.Count -eq 1 -and $script:EmptyAttempt -eq 2) -Name 'Invoke-GraphBatch requeues chunk when batch response is empty'
}
Test-InvokeGraphBatch

function Test-GetMembershipRows {
    $recipients = @(
        [pscustomobject]@{ Identity = 'Jane Doe'; PrimarySmtpAddress = 'jane@contoso.com'; ExternalDirectoryObjectId = 'aaaaaaaa-0000-0000-0000-000000000001' },
        [pscustomobject]@{ Identity = 'No Entra'; PrimarySmtpAddress = 'x@contoso.com'; ExternalDirectoryObjectId = '' }
    )
    $mock = {
        param($Method, $Uri, $Body)
        if ($Method -eq 'POST') {
            return @{ responses = @($Body.requests | ForEach-Object {
                @{ id = $_.id; status = 200; body = @{
                    value = @(
                        @{ id = 'g1'; displayName = 'Sales DL'; mail = 'sales@contoso.com'; mailEnabled = $true; securityEnabled = $false },
                        @{ id = 'g2'; displayName = 'Sec Only'; mail = $null; mailEnabled = $false; securityEnabled = $true }
                    )
                    '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/page2'
                } }
            }) }
        }
        # nextLink page fetch
        return @{ value = @(@{ id = 'g3'; displayName = 'Second Page DL'; mail = 'p2@contoso.com'; mailEnabled = $true; securityEnabled = $false }) }
    }
    $result = Get-MembershipRows -Recipients $recipients -GraphRequest $mock 3>$null
    $entra = @($result.Entra)
    $exchange = @($result.Exchange)
    Assert-True -Condition ($entra.Count -eq 3) -Name 'Get-MembershipRows returns all groups as Entra rows'
    Assert-True -Condition ($exchange.Count -eq 2) -Name 'Get-MembershipRows filters mailEnabled groups into Exchange rows'
    Assert-True -Condition ($exchange[0].GroupIdentity -eq 'Sales DL (sales@contoso.com)') -Name 'Get-MembershipRows formats GroupIdentity with mail'
    Assert-True -Condition ($entra[1].GroupDisplayName -eq 'Sec Only' -and $entra[1].GroupMail -eq '') -Name 'Get-MembershipRows maps Entra row columns'
    Assert-True -Condition (@($entra | Where-Object { $_.GroupDisplayName -eq 'Second Page DL' }).Count -eq 1) -Name 'Get-MembershipRows follows nextLink pages'
    Assert-True -Condition (@($entra | Where-Object { $_.Identity -eq 'No Entra' }).Count -eq 0) -Name 'Get-MembershipRows skips recipients without ExternalDirectoryObjectId'
}
Test-GetMembershipRows

function Test-OldMembershipFunctionsRemoved {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw
    Assert-True -Condition ($scriptText -notmatch 'Get-ExchangeGroupMembershipRows' -and $scriptText -notmatch 'Get-EntraGroupMembershipRows') -Name 'Old membership functions removed'
    Assert-True -Condition ($scriptText -notmatch 'Get-MgDirectoryObjectMemberOf' -and $scriptText -notmatch 'Microsoft\.Graph\.DirectoryObjects') -Name 'Graph SDK entity cmdlets removed'
}
Test-OldMembershipFunctionsRemoved

function Test-SendAsAndFullAccess {
    $recipients = @(
        [pscustomobject]@{ Identity = 'Jane Doe'; PrimarySmtpAddress = 'jane@contoso.com'; Guid = [guid]'22222222-2222-2222-2222-222222222222'; RecipientTypeDetails = 'UserMailbox'; RecipientType = 'UserMailbox' }
    )

    # Org-wide mode: mock returns rows for an in-scope recipient, an out-of-scope one, and a SELF trustee.
    function Get-RecipientPermission {
        param($Identity, $ResultSize, $ErrorAction)
        @(
            [pscustomobject]@{ Identity = 'Jane Doe'; Trustee = 'bob@contoso.com'; AccessRights = @('SendAs'); IsInherited = $false },
            [pscustomobject]@{ Identity = 'Jane Doe'; Trustee = 'NT AUTHORITY\SELF'; AccessRights = @('SendAs'); IsInherited = $false },
            [pscustomobject]@{ Identity = 'Someone Else'; Trustee = 'eve@contoso.com'; AccessRights = @('SendAs'); IsInherited = $false }
        )
    }
    $rows = @(Get-SendAsRows -Recipients $recipients -UseOrgWideQuery $true)
    Assert-True -Condition ($rows.Count -eq 1 -and $rows[0].Trustee -eq 'bob@contoso.com') -Name 'Get-SendAsRows org-wide filters to inventory and drops SELF'

    # Per-identity mode must pass the Guid, not the display name.
    $script:SendAsLookups = [System.Collections.Generic.List[string]]::new()
    function Get-RecipientPermission {
        param($Identity, $ResultSize, $ErrorAction)
        $script:SendAsLookups.Add([string]$Identity) | Out-Null
        @([pscustomobject]@{ Identity = 'Jane Doe'; Trustee = 'bob@contoso.com'; AccessRights = @('SendAs'); IsInherited = $false })
    }
    $rows = @(Get-SendAsRows -Recipients $recipients -UseOrgWideQuery $false)
    Assert-True -Condition ($script:SendAsLookups[0] -eq '22222222-2222-2222-2222-222222222222') -Name 'Get-SendAsRows per-identity uses Guid'

    # FullAccess must pass the Guid too.
    $script:FullAccessLookups = [System.Collections.Generic.List[string]]::new()
    function Get-EXOMailboxPermission {
        param($Identity, $ErrorAction)
        $script:FullAccessLookups.Add([string]$Identity) | Out-Null
        @([pscustomobject]@{ User = 'bob@contoso.com'; AccessRights = @('FullAccess'); IsInherited = $false; Deny = $false })
    }
    $rows = @(Get-FullAccessRows -Recipient $recipients[0])
    Assert-True -Condition ($script:FullAccessLookups[0] -eq '22222222-2222-2222-2222-222222222222') -Name 'Get-FullAccessRows uses Guid'
    Assert-True -Condition ($rows.Count -eq 1 -and $rows[0].Trustee -eq 'bob@contoso.com') -Name 'Get-FullAccessRows maps rows'
}
Test-SendAsAndFullAccess

function Test-LookupIdentityHelperRemoved {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw
    Assert-True -Condition ($scriptText -notmatch 'Get-ExchangeLookupIdentity') -Name 'Get-ExchangeLookupIdentity removed'
}
Test-LookupIdentityHelperRemoved

function Test-Consolidation {
    $objects = @([pscustomobject]@{
        Identity = 'Jane Doe'; DisplayName = 'Jane Doe'; RecipientType = 'UserMailbox'; RecipientTypeDetails = 'UserMailbox'
        MailboxType = 'UserMailbox'; PrimarySmtpAddress = 'jane@contoso.com'; Alias = 'jane'
        ExternalDirectoryObjectId = 'aaaa'; ExchangeObjectId = 'bbbb'; DistinguishedName = 'CN=Jane'
    })
    $proxy = @(
        [pscustomobject]@{ Identity = 'Jane Doe'; RawProxyAddress = 'SMTP:jane@contoso.com' },
        [pscustomobject]@{ Identity = 'Jane Doe'; RawProxyAddress = 'smtp:jane2@contoso.com' }
    )
    $sendAs = @([pscustomobject]@{ Identity = 'Jane Doe'; Trustee = 'bob@contoso.com' })
    $errors = @([pscustomobject]@{ Identity = 'Jane Doe'; Stage = 'X'; Operation = 'Y'; Message = 'Z' })

    $index = Group-RowsByIdentity -Rows $proxy
    Assert-True -Condition ($index['Jane Doe'].Count -eq 2) -Name 'Group-RowsByIdentity groups rows'
    Assert-True -Condition ((Group-RowsByIdentity -Rows @()).Count -eq 0) -Name 'Group-RowsByIdentity handles empty input'

    $consolidated = @(New-ConsolidatedRows -Objects $objects -ProxyRows $proxy -FullAccessRows @() -SendAsRows $sendAs -ExchangeGroupRows @() -EntraGroupRows @() -ErrorRows $errors)
    Assert-True -Condition ($consolidated.Count -eq 1) -Name 'New-ConsolidatedRows emits one row per object'
    Assert-True -Condition ($consolidated[0].ProxyAddresses -eq 'SMTP:jane@contoso.com; smtp:jane2@contoso.com') -Name 'New-ConsolidatedRows joins proxies sorted'
    Assert-True -Condition ($consolidated[0].SendAsTrustees -eq 'bob@contoso.com' -and $consolidated[0].FullAccessTrustees -eq '') -Name 'New-ConsolidatedRows joins trustees and handles empty sets'
    Assert-True -Condition ($consolidated[0].HasErrors -eq $true) -Name 'New-ConsolidatedRows flags errors'
}
Test-Consolidation

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
if ($script:Failures.Count -gt 0) {
    Write-Host "$($script:Failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host "All $($script:Passes) checks passed." -ForegroundColor Green
exit 0
