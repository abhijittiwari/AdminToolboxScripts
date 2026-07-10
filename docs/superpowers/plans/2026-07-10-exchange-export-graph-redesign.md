# Export-ExchangeObjectData.ps1 Graph-First Redesign Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Rebuild the internals of `Export-ExchangeObjectData.ps1` as a five-phase pipeline that uses one bulk Exchange Online inventory call plus Microsoft Graph `$batch` for per-recipient work, keeping the external surface (parameters, CSVs, dashboard) identical.

**Architecture:** One paged `Get-EXORecipient` call (now including `EmailAddresses`) is the spine. Proxy rows are parsed in memory. Group memberships come from Graph `/$batch` `memberOf` calls (20 per HTTP request) and feed both membership CSVs. SendAs uses one org-wide `Get-RecipientPermission` sweep on `-All` runs; FullAccess stays per-mailbox but always looks up by immutable `Guid`. Consolidation joins through hashtable indexes.

**Tech Stack:** PowerShell 7+, `ExchangeOnlineManagement`, `Microsoft.Graph.Authentication` (raw REST via `Invoke-MgGraphRequest` — no Graph SDK entity cmdlets).

**Spec:** `docs/superpowers/specs/2026-07-10-exchange-export-redesign-design.md`

## Global Constraints

- Only two module imports in the script: `ExchangeOnlineManagement`, `Microsoft.Graph.Authentication`.
- All Graph traffic via `Invoke-MgGraphRequest`; never `Get-MgUser`/`Get-MgGroup`/`Get-MgDirectoryObjectMemberOf`.
- Graph base URL `https://graph.microsoft.com/v1.0`; batch endpoint `https://graph.microsoft.com/v1.0/$batch`; 20 sub-requests per batch; maximum 5 attempts per item; honor `Retry-After`.
- Graph scopes exactly: `User.Read.All`, `Group.Read.All`, `Directory.Read.All`.
- External surface unchanged: parameter sets `All`/`Csv`/`Identity`; parameters `-All -InputCsv -Identity -IdentityColumn -RecipientType -OutputFolder -SkipGraph -NoDashboard`; CSV filenames and headers exactly as today (see Task 6 header lists); dashboard behavior unchanged.
- Accepted output delta (only one): `ExchangeGroupMemberships.csv` column `GroupIdentity` holds the Graph group display name, with `" (<mail>)"` appended when the group has a mail address.
- Exchange lookups that take an identity always use `[string]$Recipient.Guid` — never display name or alias.
- Every failure logs an `Errors.csv` row via `Add-ExportError` (`Identity, Stage, Operation, Message`) and never aborts the run.
- Tests: `tests/Export-ExchangeObjectData.Checks.ps1`, run with `pwsh -NoLogo -NoProfile -File tests/Export-ExchangeObjectData.Checks.ps1`; exit 0 = all pass, exit 1 = failures. No Pester dependency.
- Commit after every task with the message given in the task.

---

### Task 1: Retire the superseded workaround and create the test harness

The working tree has an uncommitted retry-loop change to `Export-ExchangeObjectData.ps1` that this redesign supersedes. Stash it (do not delete it), then build the check harness every later task uses.

**Files:**
- Modify: `Export-ExchangeObjectData.ps1` (restore to HEAD via stash)
- Create: `tests/Export-ExchangeObjectData.Checks.ps1`

**Interfaces:**
- Produces: `Import-ScriptFunctions` (dot-source every function from the main script into the check script's scope), `Assert-True -Condition <bool> -Name <string>`, `Get-ScriptAst` (returns the parsed AST of the main script). Later tasks append check functions and calls to this file.

- [ ] **Step 1: Stash the superseded working-tree change**

```bash
cd /Users/abhijittiwari/Documents/AdminToolboxScripts
git stash push -m "superseded proxy retry-loop fix (replaced by Graph redesign)" -- Export-ExchangeObjectData.ps1
git status --short   # Export-ExchangeObjectData.ps1 must no longer appear as modified
```

- [ ] **Step 2: Create the check harness**

Create `tests/Export-ExchangeObjectData.Checks.ps1` with exactly:

```powershell
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

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
if ($script:Failures.Count -gt 0) {
    Write-Host "$($script:Failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host "All $($script:Passes) checks passed." -ForegroundColor Green
exit 0
```

- [ ] **Step 3: Run the harness to verify it loads and passes with zero checks**

Run: `pwsh -NoLogo -NoProfile -File tests/Export-ExchangeObjectData.Checks.ps1`
Expected: `All 0 checks passed.` and exit code 0.

- [ ] **Step 4: Commit**

```bash
git add tests/Export-ExchangeObjectData.Checks.ps1
git commit -m "Add check harness for Exchange export redesign"
```

---

### Task 2: Replace the proxy-address lookup layer with in-memory parsing

Proxy addresses come from the Phase 1 inventory (`EmailAddresses` property), not from per-recipient lookups. This deletes `Get-ProxyAddressRows`, `Get-ExchangeLookupIdentities` (if present), and the per-recipient `Get-Recipient` proxy call.

**Files:**
- Modify: `Export-ExchangeObjectData.ps1` (function `Get-ProxyAddressRows` → new `ConvertTo-ProxyRows`; both `Get-EXORecipient` calls in `Get-TargetRecipients`; the main collection loop)
- Modify: `tests/Export-ExchangeObjectData.Checks.ps1` (append checks)

**Interfaces:**
- Consumes: recipient objects with `Identity`, `PrimarySmtpAddress`, `EmailAddresses` properties.
- Produces: `ConvertTo-ProxyRows -Recipient <object>` → array of `pscustomobject` rows with properties `Identity, PrimarySmtpAddress, AddressType, ProxyAddress, RawProxyAddress, IsPrimarySmtp`. Task 6 consumes these rows.

- [ ] **Step 1: Append the failing checks**

In `tests/Export-ExchangeObjectData.Checks.ps1`, insert into the Checks region:

```powershell
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
```

- [ ] **Step 2: Run checks to verify they fail**

Run: `pwsh -NoLogo -NoProfile -File tests/Export-ExchangeObjectData.Checks.ps1`
Expected: FAIL — `ConvertTo-ProxyRows` is not defined (CommandNotFoundException) or the assertions fail; exit code 1.

- [ ] **Step 3: Implement**

In `Export-ExchangeObjectData.ps1`:

3a. Delete the entire `Get-ProxyAddressRows` function (and `Get-ExchangeLookupIdentities`/`Test-IsMailboxRecipient` if the stash restore left them; at HEAD they do not exist). Replace with:

```powershell
function ConvertTo-ProxyRows {
    param([Parameter(Mandatory = $true)] $Recipient)

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($address in @($Recipient.EmailAddresses)) {
        if ($null -eq $address) { continue }
        $value = $address.ToString()
        $parts = $value -split ':', 2
        $prefix = if ($parts.Count -eq 2) { $parts[0] } else { '' }
        $addressValue = if ($parts.Count -eq 2) { $parts[1] } else { $value }

        $rows.Add([pscustomobject]@{
            Identity           = [string]$Recipient.Identity
            PrimarySmtpAddress = [string]$Recipient.PrimarySmtpAddress
            AddressType        = $prefix
            ProxyAddress       = $addressValue
            RawProxyAddress    = $value
            IsPrimarySmtp      = $value.StartsWith('SMTP:')
        }) | Out-Null
    }
    return @($rows)
}
```

3b. In `Get-TargetRecipients`, add `EmailAddresses` to the `-Properties` list of **both** `Get-EXORecipient` calls:

```powershell
-Properties RecipientTypeDetails,ExternalDirectoryObjectId,PrimarySmtpAddress,Alias,Guid,DistinguishedName,EmailAddresses
```

3c. In the main collection loop, replace the line
`foreach ($row in @(Get-ProxyAddressRows -Recipient $recipient)) { $proxyRows.Add($row) | Out-Null }`
with
`foreach ($row in @(ConvertTo-ProxyRows -Recipient $recipient)) { $proxyRows.Add($row) | Out-Null }`

- [ ] **Step 4: Run checks to verify they pass**

Run: `pwsh -NoLogo -NoProfile -File tests/Export-ExchangeObjectData.Checks.ps1`
Expected: all PASS, exit code 0.

- [ ] **Step 5: Commit**

```bash
git add Export-ExchangeObjectData.ps1 tests/Export-ExchangeObjectData.Checks.ps1
git commit -m "Parse proxy addresses from bulk inventory instead of per-recipient lookups"
```

---

### Task 3: Add the Graph batch engine

Pure addition: `New-GraphRequestExecutor` (the seam that lets checks mock Graph) and `Invoke-GraphBatch` (chunking, retry, per-item error reporting).

**Files:**
- Modify: `Export-ExchangeObjectData.ps1` (add two functions after `Connect-GraphIfNeeded`)
- Modify: `tests/Export-ExchangeObjectData.Checks.ps1` (append checks)

**Interfaces:**
- Produces:
  - `New-GraphRequestExecutor` → `[scriptblock]` taking `($Method, $Uri, $Body)` and calling `Invoke-MgGraphRequest`.
  - `Invoke-GraphBatch -Requests <object[]> [-GraphRequest <scriptblock>] [-MaxAttempts <int>] [-OnItemError <scriptblock>] [-Activity <string>]` where each request is `@{ Id = <string>; Url = <string> }` → returns `[hashtable]` mapping request Id → response body (hashtable with `value` / `@odata.nextLink` keys). `OnItemError` is invoked as `& $OnItemError $id $message`. Task 4 consumes both.

- [ ] **Step 1: Append the failing checks**

```powershell
function Test-InvokeGraphBatch {
    # 45 requests -> 3 batch calls (20 + 20 + 5), all succeed.
    $script:BatchCalls = New-Object System.Collections.Generic.List[object]
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
    $script:ItemErrors = New-Object System.Collections.Generic.List[string]
    $mock404 = {
        param($Method, $Uri, $Body)
        return @{ responses = @($Body.requests | ForEach-Object { @{ id = $_.id; status = 404; body = @{ error = @{ message = 'not found' } } } }) }
    }
    $onError = { param($id, $message) $script:ItemErrors.Add("$id|$message") | Out-Null }
    $results = Invoke-GraphBatch -Requests @(@{ Id = 'r9'; Url = '/gone' }) -GraphRequest $mock404 -OnItemError $onError
    Assert-True -Condition ($results.Count -eq 0 -and $script:ItemErrors.Count -eq 1 -and $script:ItemErrors[0] -eq 'r9|not found') -Name 'Invoke-GraphBatch reports permanent item failures once'

    # Always-throttled item exhausts MaxAttempts then errors.
    $script:ItemErrors = New-Object System.Collections.Generic.List[string]
    $mockAlways429 = {
        param($Method, $Uri, $Body)
        return @{ responses = @($Body.requests | ForEach-Object { @{ id = $_.id; status = 429; headers = @{ 'Retry-After' = '0' }; body = @{} } }) }
    }
    $results = Invoke-GraphBatch -Requests @(@{ Id = 'r1'; Url = '/x' }) -GraphRequest $mockAlways429 -MaxAttempts 2 -OnItemError $onError
    Assert-True -Condition ($results.Count -eq 0 -and $script:ItemErrors.Count -eq 1) -Name 'Invoke-GraphBatch gives up after MaxAttempts'
}
Test-InvokeGraphBatch
```

- [ ] **Step 2: Run checks to verify the new ones fail**

Run: `pwsh -NoLogo -NoProfile -File tests/Export-ExchangeObjectData.Checks.ps1`
Expected: prior checks PASS; `Invoke-GraphBatch` not defined → failure; exit code 1.

- [ ] **Step 3: Implement**

Add to `Export-ExchangeObjectData.ps1`, directly after `Connect-GraphIfNeeded`:

```powershell
function New-GraphRequestExecutor {
    return {
        param($Method, $Uri, $Body)
        if ($null -ne $Body) {
            Invoke-MgGraphRequest -Method $Method -Uri $Uri -Body ($Body | ConvertTo-Json -Depth 10) -ContentType 'application/json'
        }
        else {
            Invoke-MgGraphRequest -Method $Method -Uri $Uri
        }
    }
}

function Invoke-GraphBatch {
    param(
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] [object[]]$Requests,
        [Parameter(Mandatory = $false)] [scriptblock]$GraphRequest,
        [Parameter(Mandatory = $false)] [int]$MaxAttempts = 5,
        [Parameter(Mandatory = $false)] [scriptblock]$OnItemError,
        [Parameter(Mandatory = $false)] [string]$Activity = 'Calling Microsoft Graph'
    )

    if (-not $GraphRequest) { $GraphRequest = New-GraphRequestExecutor }

    $results = @{}
    $pending = New-Object System.Collections.Generic.List[object]
    foreach ($request in @($Requests)) { $pending.Add($request) | Out-Null }
    $totalCount = [Math]::Max($pending.Count, 1)
    $attempt = 1

    while ($pending.Count -gt 0 -and $attempt -le $MaxAttempts) {
        $retryQueue = New-Object System.Collections.Generic.List[object]
        $delaySeconds = [Math]::Pow(2, $attempt)

        for ($offset = 0; $offset -lt $pending.Count; $offset += 20) {
            $chunk = @($pending[$offset..([Math]::Min($offset + 19, $pending.Count - 1))])
            Write-Progress -Activity $Activity -Status "attempt $attempt, $([Math]::Min($offset + 20, $pending.Count)) of $($pending.Count)" -PercentComplete ((($totalCount - $pending.Count + $offset) / $totalCount) * 100)
            $body = @{ requests = @($chunk | ForEach-Object { @{ id = [string]$_.Id; method = 'GET'; url = [string]$_.Url } }) }

            try {
                $response = & $GraphRequest 'POST' 'https://graph.microsoft.com/v1.0/$batch' $body
            }
            catch {
                foreach ($item in $chunk) { $retryQueue.Add($item) | Out-Null }
                continue
            }

            foreach ($item in @($response.responses)) {
                $itemId = [string]$item.id
                $status = [int]$item.status
                if ($status -ge 200 -and $status -lt 300) {
                    $results[$itemId] = $item.body
                }
                elseif ($status -in @(429, 502, 503, 504)) {
                    $retryQueue.Add(($chunk | Where-Object { [string]$_.Id -eq $itemId } | Select-Object -First 1)) | Out-Null
                    if ($item.headers -and $item.headers.'Retry-After') {
                        $headerDelay = [double]$item.headers.'Retry-After'
                        if ($headerDelay -gt $delaySeconds) { $delaySeconds = $headerDelay }
                    }
                }
                else {
                    $message = if ($item.body -and $item.body.error -and $item.body.error.message) { [string]$item.body.error.message } else { "HTTP $status" }
                    if ($OnItemError) { & $OnItemError $itemId $message }
                }
            }
        }

        $pending = $retryQueue
        if ($pending.Count -gt 0 -and $attempt -lt $MaxAttempts) { Start-Sleep -Seconds $delaySeconds }
        $attempt++
    }

    Write-Progress -Activity $Activity -Completed
    foreach ($leftover in $pending) {
        if ($OnItemError) { & $OnItemError ([string]$leftover.Id) "Gave up after $MaxAttempts attempts (throttled)" }
    }

    return $results
}
```

- [ ] **Step 4: Run checks to verify they pass**

Run: `pwsh -NoLogo -NoProfile -File tests/Export-ExchangeObjectData.Checks.ps1`
Expected: all PASS, exit code 0. (The always-throttled check sleeps ~2s once; total runtime stays under ~10s.)

- [ ] **Step 5: Commit**

```bash
git add Export-ExchangeObjectData.ps1 tests/Export-ExchangeObjectData.Checks.ps1
git commit -m "Add Graph batch engine with Retry-After handling"
```

---

### Task 4: Collect group memberships through Graph batch

Replaces `Get-ExchangeGroupMembershipRows` (per-recipient EXO `Get-Recipient`) and `Get-EntraGroupMembershipRows` (Graph SDK cmdlet) with one `Get-MembershipRows` that feeds both CSVs. Also trims `Connect-GraphIfNeeded` to a single module import.

**Files:**
- Modify: `Export-ExchangeObjectData.ps1` (delete `Get-ExchangeGroupMembershipRows`, `Get-EntraGroupMembershipRows`; add `Get-MembershipRows`; edit `Connect-GraphIfNeeded`; rewire the main loop's membership collection)
- Modify: `tests/Export-ExchangeObjectData.Checks.ps1` (append checks)

**Interfaces:**
- Consumes: `Invoke-GraphBatch`, `New-GraphRequestExecutor`, `Add-ExportError` from earlier tasks.
- Produces: `Get-MembershipRows -Recipients <object[]> [-GraphRequest <scriptblock>]` → `[hashtable]` with keys `Entra` (rows: `Identity, PrimarySmtpAddress, GroupId, GroupDisplayName, GroupMail, GroupSecurityEnabled`) and `Exchange` (rows: `Identity, PrimarySmtpAddress, GroupIdentity`). Task 6 consumes both row sets.

- [ ] **Step 1: Append the failing checks**

```powershell
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
```

- [ ] **Step 2: Run checks to verify the new ones fail**

Run: `pwsh -NoLogo -NoProfile -File tests/Export-ExchangeObjectData.Checks.ps1`
Expected: earlier checks PASS; `Get-MembershipRows` undefined and removal checks FAIL; exit code 1.

- [ ] **Step 3: Implement**

3a. In `Connect-GraphIfNeeded`, replace the two `Import-Module` lines with one:

```powershell
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
```

(keep `Connect-MgGraph -Scopes @('User.Read.All', 'Group.Read.All', 'Directory.Read.All') -NoWelcome` unchanged).

3b. Delete `Get-ExchangeGroupMembershipRows` and `Get-EntraGroupMembershipRows` entirely. Add in their place:

```powershell
function Get-MembershipRows {
    param(
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] [object[]]$Recipients,
        [Parameter(Mandatory = $false)] [scriptblock]$GraphRequest
    )

    if (-not $GraphRequest) { $GraphRequest = New-GraphRequestExecutor }

    $entraRows = New-Object System.Collections.Generic.List[object]
    $exchangeRows = New-Object System.Collections.Generic.List[object]

    $recipientsById = @{}
    $requests = New-Object System.Collections.Generic.List[object]
    foreach ($recipient in @($Recipients)) {
        $externalId = [string]$recipient.ExternalDirectoryObjectId
        if ([string]::IsNullOrWhiteSpace($externalId)) { continue }
        if ($recipientsById.ContainsKey($externalId)) { continue }
        $recipientsById[$externalId] = $recipient
        $requests.Add(@{
            Id  = $externalId
            Url = "/directoryObjects/$externalId/memberOf/microsoft.graph.group?`$select=id,displayName,mail,mailEnabled,securityEnabled&`$top=999"
        }) | Out-Null
    }

    $onItemError = {
        param($id, $message)
        $identity = if ($recipientsById.ContainsKey($id)) { [string]$recipientsById[$id].Identity } else { [string]$id }
        Add-ExportError -Identity $identity -Stage 'GroupMemberships' -Operation 'Graph memberOf' -Message $message
    }

    $responses = Invoke-GraphBatch -Requests @($requests) -GraphRequest $GraphRequest -OnItemError $onItemError -Activity 'Collecting group memberships via Microsoft Graph'

    foreach ($externalId in @($responses.Keys)) {
        $recipient = $recipientsById[$externalId]
        $body = $responses[$externalId]

        $groups = New-Object System.Collections.Generic.List[object]
        while ($true) {
            foreach ($group in @($body.value)) {
                if ($null -ne $group) { $groups.Add($group) | Out-Null }
            }
            $nextLink = [string]$body.'@odata.nextLink'
            if ([string]::IsNullOrWhiteSpace($nextLink)) { break }
            try {
                $body = & $GraphRequest 'GET' $nextLink $null
            }
            catch {
                Add-ExportError -Identity ([string]$recipient.Identity) -Stage 'GroupMemberships' -Operation 'Graph memberOf paging' -Message $_.Exception.Message
                break
            }
        }

        foreach ($group in $groups) {
            $groupMail = [string]$group.mail
            $entraRows.Add([pscustomobject]@{
                Identity             = [string]$recipient.Identity
                PrimarySmtpAddress   = [string]$recipient.PrimarySmtpAddress
                GroupId              = [string]$group.id
                GroupDisplayName     = [string]$group.displayName
                GroupMail            = $groupMail
                GroupSecurityEnabled = [string]$group.securityEnabled
            }) | Out-Null

            if ([bool]$group.mailEnabled) {
                $groupIdentity = [string]$group.displayName
                if (-not [string]::IsNullOrWhiteSpace($groupMail)) { $groupIdentity = "$groupIdentity ($groupMail)" }
                $exchangeRows.Add([pscustomobject]@{
                    Identity           = [string]$recipient.Identity
                    PrimarySmtpAddress = [string]$recipient.PrimarySmtpAddress
                    GroupIdentity      = $groupIdentity
                }) | Out-Null
            }
        }
    }

    return @{ Entra = @($entraRows); Exchange = @($exchangeRows) }
}
```

3c. In the main collection loop, delete these two lines:

```powershell
    foreach ($row in @(Get-ExchangeGroupMembershipRows -Recipient $recipient)) { $exchangeGroupMembershipRows.Add($row) | Out-Null }
    if ($graphConnected) {
        foreach ($row in @(Get-EntraGroupMembershipRows -Recipient $recipient)) { $entraGroupMembershipRows.Add($row) | Out-Null }
    }
```

and after the loop (before `$consolidatedRows = ...`) add:

```powershell
if ($graphConnected) {
    Write-Host 'Collecting group memberships via Microsoft Graph...' -ForegroundColor Cyan
    $membershipRows = Get-MembershipRows -Recipients $targetRecipients
    foreach ($row in @($membershipRows.Exchange)) { $exchangeGroupMembershipRows.Add($row) | Out-Null }
    foreach ($row in @($membershipRows.Entra)) { $entraGroupMembershipRows.Add($row) | Out-Null }
}
else {
    Write-Host 'Microsoft Graph not connected: membership CSVs will contain headers only.' -ForegroundColor Yellow
}
```

- [ ] **Step 4: Run checks to verify they pass**

Run: `pwsh -NoLogo -NoProfile -File tests/Export-ExchangeObjectData.Checks.ps1`
Expected: all PASS, exit code 0.

- [ ] **Step 5: Commit**

```bash
git add Export-ExchangeObjectData.ps1 tests/Export-ExchangeObjectData.Checks.ps1
git commit -m "Collect both membership exports via Graph batch memberOf"
```

---

### Task 5: Bulk SendAs, Guid-keyed FullAccess, delete the lookup-identity helper

`-All` runs sweep SendAs org-wide in one paged call. FullAccess and per-identity SendAs look up by `Guid`. `Get-ExchangeLookupIdentity` (the source of the display-name ambiguity bugs) is deleted.

**Files:**
- Modify: `Export-ExchangeObjectData.ps1` (rewrite `Get-SendAsRows`, edit `Get-FullAccessRows`, delete `Get-ExchangeLookupIdentity`, rewire main loop)
- Modify: `tests/Export-ExchangeObjectData.Checks.ps1` (append checks)

**Interfaces:**
- Consumes: `Test-IsSelfTrustee`, `Get-MailboxType`, `Add-ExportError` (all unchanged).
- Produces:
  - `Get-SendAsRows -Recipients <object[]> -UseOrgWideQuery <bool>` → array of rows `Identity, PrimarySmtpAddress, Trustee, AccessRights, IsInherited`.
  - `Get-FullAccessRows -Recipient <object>` → array of rows `Identity, PrimarySmtpAddress, Trustee, AccessRights, Deny, IsInherited` (unchanged shape, now Guid lookup). Task 6 consumes both.

- [ ] **Step 1: Append the failing checks**

```powershell
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
    $script:SendAsLookups = New-Object System.Collections.Generic.List[string]
    function Get-RecipientPermission {
        param($Identity, $ResultSize, $ErrorAction)
        $script:SendAsLookups.Add([string]$Identity) | Out-Null
        @([pscustomobject]@{ Identity = 'Jane Doe'; Trustee = 'bob@contoso.com'; AccessRights = @('SendAs'); IsInherited = $false })
    }
    $rows = @(Get-SendAsRows -Recipients $recipients -UseOrgWideQuery $false)
    Assert-True -Condition ($script:SendAsLookups[0] -eq '22222222-2222-2222-2222-222222222222') -Name 'Get-SendAsRows per-identity uses Guid'

    # FullAccess must pass the Guid too.
    $script:FullAccessLookups = New-Object System.Collections.Generic.List[string]
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
```

- [ ] **Step 2: Run checks to verify the new ones fail**

Run: `pwsh -NoLogo -NoProfile -File tests/Export-ExchangeObjectData.Checks.ps1`
Expected: new checks FAIL (old `Get-SendAsRows` takes `-Recipient`, not `-Recipients`; helper still present); exit code 1.

- [ ] **Step 3: Implement**

3a. Delete the `Get-ExchangeLookupIdentity` function.

3b. Replace `Get-SendAsRows` entirely with:

```powershell
function Get-SendAsRows {
    param(
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] [object[]]$Recipients,
        [Parameter(Mandatory = $true)] [bool]$UseOrgWideQuery
    )

    $rows = New-Object System.Collections.Generic.List[object]

    if ($UseOrgWideQuery) {
        $recipientsByIdentity = @{}
        foreach ($recipient in @($Recipients)) { $recipientsByIdentity[[string]$recipient.Identity] = $recipient }
        try {
            foreach ($permission in @(Get-RecipientPermission -ResultSize Unlimited -ErrorAction Stop)) {
                $identity = [string]$permission.Identity
                if (-not $recipientsByIdentity.ContainsKey($identity)) { continue }
                if ($permission.IsInherited) { continue }
                if (Test-IsSelfTrustee -Trustee ([string]$permission.Trustee)) { continue }
                if ($permission.AccessRights -notcontains 'SendAs') { continue }
                $rows.Add([pscustomobject]@{
                    Identity           = $identity
                    PrimarySmtpAddress = [string]$recipientsByIdentity[$identity].PrimarySmtpAddress
                    Trustee            = [string]$permission.Trustee
                    AccessRights       = ($permission.AccessRights -join '; ')
                    IsInherited        = [bool]$permission.IsInherited
                }) | Out-Null
            }
        }
        catch {
            Add-ExportError -Identity '' -Stage 'SendAs' -Operation 'Get-RecipientPermission (org-wide)' -Message $_.Exception.Message
        }
        return @($rows)
    }

    foreach ($recipient in @($Recipients)) {
        try {
            foreach ($permission in @(Get-RecipientPermission -Identity ([string]$recipient.Guid) -ErrorAction Stop)) {
                if ($permission.IsInherited) { continue }
                if (Test-IsSelfTrustee -Trustee ([string]$permission.Trustee)) { continue }
                if ($permission.AccessRights -notcontains 'SendAs') { continue }
                $rows.Add([pscustomobject]@{
                    Identity           = [string]$recipient.Identity
                    PrimarySmtpAddress = [string]$recipient.PrimarySmtpAddress
                    Trustee            = [string]$permission.Trustee
                    AccessRights       = ($permission.AccessRights -join '; ')
                    IsInherited        = [bool]$permission.IsInherited
                }) | Out-Null
            }
        }
        catch {
            Add-ExportError -Identity ([string]$recipient.Identity) -Stage 'SendAs' -Operation 'Get-RecipientPermission' -Message $_.Exception.Message
        }
    }
    return @($rows)
}
```

3c. In `Get-FullAccessRows`, replace the two lines

```powershell
    $lookupIdentity = Get-ExchangeLookupIdentity -Recipient $Recipient
```
and
```powershell
            Get-EXOMailboxPermission -Identity $lookupIdentity -ErrorAction Stop |
```
with a direct Guid lookup:
```powershell
    $lookupIdentity = [string]$Recipient.Guid
```
(keep the rest, including the error message's `(lookup: $lookupIdentity)` suffix).

3d. In the main collection loop, delete the line
`foreach ($row in @(Get-SendAsRows -Recipient $recipient)) { $sendAsRows.Add($row) | Out-Null }`
and add after the membership block from Task 4:

```powershell
Write-Host 'Collecting SendAs permissions...' -ForegroundColor Cyan
foreach ($row in @(Get-SendAsRows -Recipients $targetRecipients -UseOrgWideQuery ($PSCmdlet.ParameterSetName -eq 'All'))) { $sendAsRows.Add($row) | Out-Null }
```

- [ ] **Step 4: Run checks to verify they pass**

Run: `pwsh -NoLogo -NoProfile -File tests/Export-ExchangeObjectData.Checks.ps1`
Expected: all PASS, exit code 0.

- [ ] **Step 5: Commit**

```bash
git add Export-ExchangeObjectData.ps1 tests/Export-ExchangeObjectData.Checks.ps1
git commit -m "Bulk SendAs sweep and Guid-keyed permission lookups"
```

---

### Task 6: Hashtable-indexed consolidation and final phase ordering

Kill the quadratic `Where-Object` joins and settle the main block into the five-phase order with progress reporting.

**Files:**
- Modify: `Export-ExchangeObjectData.ps1` (add `Group-RowsByIdentity`, rewrite `New-ConsolidatedRows` internals, restructure the main collection block)
- Modify: `tests/Export-ExchangeObjectData.Checks.ps1` (append checks)

**Interfaces:**
- Consumes: row shapes from Tasks 2, 4, 5.
- Produces: `Group-RowsByIdentity -Rows <object[]>` → `[hashtable]` Identity → `List[object]`. `New-ConsolidatedRows` signature and output columns unchanged: `Identity, DisplayName, RecipientType, RecipientTypeDetails, MailboxType, PrimarySmtpAddress, Alias, ExternalDirectoryObjectId, ExchangeObjectId, ProxyAddresses, FullAccessTrustees, SendAsTrustees, ExchangeGroupMemberships, EntraGroupMemberships, HasErrors`.

- [ ] **Step 1: Append the failing checks**

```powershell
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
```

- [ ] **Step 2: Run checks to verify the new ones fail**

Run: `pwsh -NoLogo -NoProfile -File tests/Export-ExchangeObjectData.Checks.ps1`
Expected: `Group-RowsByIdentity` undefined → FAIL; exit code 1.

- [ ] **Step 3: Implement**

3a. Add above `New-ConsolidatedRows`:

```powershell
function Group-RowsByIdentity {
    param([Parameter(Mandatory = $false)] [AllowEmptyCollection()] [object[]]$Rows)

    $index = @{}
    foreach ($row in @($Rows)) {
        if ($null -eq $row) { continue }
        $key = [string]$row.Identity
        if (-not $index.ContainsKey($key)) { $index[$key] = New-Object System.Collections.Generic.List[object] }
        $index[$key].Add($row) | Out-Null
    }
    return $index
}
```

3b. In `New-ConsolidatedRows`, keep the signature and replace the body with:

```powershell
    $proxyIndex = Group-RowsByIdentity -Rows $ProxyRows
    $fullAccessIndex = Group-RowsByIdentity -Rows $FullAccessRows
    $sendAsIndex = Group-RowsByIdentity -Rows $SendAsRows
    $exchangeGroupIndex = Group-RowsByIdentity -Rows $ExchangeGroupRows
    $entraGroupIndex = Group-RowsByIdentity -Rows $EntraGroupRows
    $errorIndex = Group-RowsByIdentity -Rows $ErrorRows

    foreach ($object in $Objects) {
        $identity = [string]$object.Identity
        [pscustomobject]@{
            Identity                  = $object.Identity
            DisplayName               = $object.DisplayName
            RecipientType             = $object.RecipientType
            RecipientTypeDetails      = $object.RecipientTypeDetails
            MailboxType               = $object.MailboxType
            PrimarySmtpAddress        = $object.PrimarySmtpAddress
            Alias                     = $object.Alias
            ExternalDirectoryObjectId = $object.ExternalDirectoryObjectId
            ExchangeObjectId          = $object.ExchangeObjectId
            ProxyAddresses            = Join-Values (@($proxyIndex[$identity]) | ForEach-Object { $_.RawProxyAddress })
            FullAccessTrustees        = Join-Values (@($fullAccessIndex[$identity]) | ForEach-Object { $_.Trustee })
            SendAsTrustees            = Join-Values (@($sendAsIndex[$identity]) | ForEach-Object { $_.Trustee })
            ExchangeGroupMemberships  = Join-Values (@($exchangeGroupIndex[$identity]) | ForEach-Object { $_.GroupIdentity })
            EntraGroupMemberships     = Join-Values (@($entraGroupIndex[$identity]) | ForEach-Object { $_.GroupDisplayName })
            HasErrors                 = [bool]($errorIndex.ContainsKey($identity) -and $errorIndex[$identity].Count -gt 0)
        }
    }
```

3c. Restructure the main collection block (between `$graphConnected = Connect-GraphIfNeeded` and `$consolidatedRows = @(...)`) to exactly:

```powershell
$proxyRows = New-Object System.Collections.Generic.List[object]
$fullAccessRows = New-Object System.Collections.Generic.List[object]
$sendAsRows = New-Object System.Collections.Generic.List[object]
$exchangeGroupMembershipRows = New-Object System.Collections.Generic.List[object]
$entraGroupMembershipRows = New-Object System.Collections.Generic.List[object]

# Phase 2: proxy addresses (in-memory, from the Phase 1 inventory)
Write-Host 'Deriving proxy address rows from inventory...' -ForegroundColor Cyan
foreach ($recipient in $targetRecipients) {
    foreach ($row in @(ConvertTo-ProxyRows -Recipient $recipient)) { $proxyRows.Add($row) | Out-Null }
}

# Phase 3: group memberships (Microsoft Graph $batch)
if ($graphConnected) {
    Write-Host 'Collecting group memberships via Microsoft Graph...' -ForegroundColor Cyan
    $membershipRows = Get-MembershipRows -Recipients $targetRecipients
    foreach ($row in @($membershipRows.Exchange)) { $exchangeGroupMembershipRows.Add($row) | Out-Null }
    foreach ($row in @($membershipRows.Entra)) { $entraGroupMembershipRows.Add($row) | Out-Null }
}
else {
    Write-Host 'Microsoft Graph not connected: membership CSVs will contain headers only.' -ForegroundColor Yellow
}

# Phase 4a: SendAs permissions (org-wide sweep on -All)
Write-Host 'Collecting SendAs permissions...' -ForegroundColor Cyan
foreach ($row in @(Get-SendAsRows -Recipients $targetRecipients -UseOrgWideQuery ($PSCmdlet.ParameterSetName -eq 'All'))) { $sendAsRows.Add($row) | Out-Null }

# Phase 4b: FullAccess permissions (per mailbox; no bulk API exists)
$mailboxRecipients = @($targetRecipients | Where-Object { -not [string]::IsNullOrWhiteSpace((Get-MailboxType -Recipient $_)) })
$index = 0
foreach ($recipient in $mailboxRecipients) {
    $index++
    Write-Progress -Activity 'Collecting FullAccess permissions' -Status "$index of $($mailboxRecipients.Count): $($recipient.Identity)" -PercentComplete (($index / [Math]::Max($mailboxRecipients.Count, 1)) * 100)
    foreach ($row in @(Get-FullAccessRows -Recipient $recipient)) { $fullAccessRows.Add($row) | Out-Null }
}
Write-Progress -Activity 'Collecting FullAccess permissions' -Completed
```

(This replaces the old single `foreach ($recipient in $targetRecipients)` loop and any membership/SendAs lines added by Tasks 4–5; the `Write-Host`/`Get-MembershipRows`/`Get-SendAsRows` snippets those tasks introduced move into this structure verbatim.)

- [ ] **Step 4: Run checks and parser validation**

Run: `pwsh -NoLogo -NoProfile -File tests/Export-ExchangeObjectData.Checks.ps1`
Expected: all PASS, exit code 0 (the harness's `Get-ScriptAst` already fails the run on any parse error).

- [ ] **Step 5: Commit**

```bash
git add Export-ExchangeObjectData.ps1 tests/Export-ExchangeObjectData.Checks.ps1
git commit -m "Index consolidation joins and settle five-phase pipeline"
```

---

### Task 7: Documentation, final verification, live-validation handoff

**Files:**
- Modify: `README.md` (Export-ExchangeObjectData section: module requirements, Graph-first description)
- Modify: `docs/script-reference.md` (same section: requirements, GroupIdentity value note)

**Interfaces:**
- Consumes: final script behavior from Tasks 1–6. Produces: none (docs only).

- [ ] **Step 1: Update docs**

In both files, in the Export-ExchangeObjectData.ps1 sections:
- State requirements as: `ExchangeOnlineManagement` and `Microsoft.Graph.Authentication` (previously the full Microsoft.Graph modules).
- Describe collection: bulk recipient inventory (including proxy addresses) from Exchange Online; group memberships via Microsoft Graph `$batch`; SendAs via one org-wide sweep on `-All` runs; FullAccess per mailbox by GUID.
- Add a note under the outputs table: "`ExchangeGroupMemberships.csv` `GroupIdentity` contains the group display name, with the group's mail address in parentheses when present (sourced from Microsoft Graph)."

- [ ] **Step 2: Run full verification**

```bash
pwsh -NoLogo -NoProfile -File tests/Export-ExchangeObjectData.Checks.ps1
git diff --check
```
Expected: all checks PASS; no whitespace errors.

- [ ] **Step 3: Commit**

```bash
git add README.md docs/script-reference.md
git commit -m "Document Graph-first Exchange export requirements and outputs"
```

- [ ] **Step 4: Hand off live validation to the user**

Offline checks cannot reach a tenant. Report to the user the staged live validation from the spec, to run in order:
1. `./Export-ExchangeObjectData.ps1 -Identity <one-user@domain>` — confirm all CSVs populate and `Errors.csv` is headers-only.
2. `./Export-ExchangeObjectData.ps1 -InputCsv <small.csv>` — confirm mixed types behave.
3. `./Export-ExchangeObjectData.ps1 -All` — compare row counts (`Objects`, `ProxyAddresses`, `SendAs`, `FullAccess`) against the old script's last output on the same tenant.
