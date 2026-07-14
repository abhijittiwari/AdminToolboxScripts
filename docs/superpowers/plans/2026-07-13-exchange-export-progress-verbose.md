# Exchange Export Progress and Verbose Output Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add visible phase progress and useful `-Verbose` diagnostics to `Export-ExchangeObjectData.ps1` so long PowerShell 7/macOS runs do not appear stuck.

**Architecture:** Keep the single-script structure. Add minimal helper functions for progress/status output, then call them from existing phases and long loops without changing export data or permission logic.

**Tech Stack:** PowerShell 7, ExchangeOnlineManagement, Microsoft.Graph.Authentication, existing AST-based PowerShell check script.

## Global Constraints

- Use `Write-Progress` for progress bars.
- Use `Write-Host` for concise default phase status and counts.
- Use `Write-Verbose` for per-object and diagnostic details.
- Do not change export schema, permission logic, Graph query shape, or dashboard behavior.
- Sync `/Users/abhijittiwari/Documents/Fortescue/Export-ExchangeObjectData.ps1` with `/Users/abhijittiwari/Documents/AdminToolboxScripts/Export-ExchangeObjectData.ps1` after implementation.
- Do not introduce new dependencies.

---

## File Structure

- Modify `Export-ExchangeObjectData.ps1`: add progress/status helpers and instrumentation around existing phases.
- Modify `tests/Export-ExchangeObjectData.Checks.ps1`: add static checks for new helper usage and previously silent phases.
- No new runtime files are required.

---

### Task 1: Add Observability Regression Checks

**Files:**
- Modify: `tests/Export-ExchangeObjectData.Checks.ps1`

**Interfaces:**
- Consumes: existing `$script:ScriptPath` and `Assert-True` helpers.
- Produces: `Test-ProgressAndVerboseInstrumentation`, a static check proving the script contains required progress and verbose instrumentation.

- [ ] **Step 1: Add failing static instrumentation checks**

Insert this function after `Test-ConnectionOrder`:

```powershell
function Test-ProgressAndVerboseInstrumentation {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw
    Assert-True -Condition ($scriptText -match 'function Write-PhaseStatus') -Name 'Progress helper Write-PhaseStatus exists'
    Assert-True -Condition ($scriptText -match 'function Write-PhaseProgress') -Name 'Progress helper Write-PhaseProgress exists'
    Assert-True -Condition ($scriptText -match 'Write-PhaseProgress -Activity ''Connecting to Microsoft Graph''') -Name 'Graph connection has progress'
    Assert-True -Condition ($scriptText -match 'Write-PhaseProgress -Activity ''Connecting to Exchange Online''') -Name 'Exchange connection has progress'
    Assert-True -Condition ($scriptText -match 'Write-PhaseProgress -Activity ''Deriving proxy address rows''') -Name 'Proxy derivation has progress'
    Assert-True -Condition ($scriptText -match 'Write-PhaseProgress -Activity ''Collecting SendAs permissions''') -Name 'SendAs collection has progress'
    Assert-True -Condition ($scriptText -match 'Write-PhaseProgress -Activity ''Creating consolidated rows''') -Name 'Consolidation has progress'
    Assert-True -Condition ($scriptText -match 'Write-PhaseProgress -Activity ''Exporting CSV reports''') -Name 'CSV export has progress'
    Assert-True -Condition ($scriptText -match 'Write-Verbose') -Name 'Verbose diagnostics are present'
}
Test-ProgressAndVerboseInstrumentation
```

- [ ] **Step 2: Run checks to verify they fail**

Run:

```bash
pwsh -NoProfile -File "tests/Export-ExchangeObjectData.Checks.ps1"
```

Expected: FAIL for the new instrumentation checks because helper functions and phase progress calls are not implemented yet.

---

### Task 2: Implement Progress and Verbose Output

**Files:**
- Modify: `Export-ExchangeObjectData.ps1`
- Modify: `/Users/abhijittiwari/Documents/Fortescue/Export-ExchangeObjectData.ps1`

**Interfaces:**
- Consumes: existing script phases and existing `Write-Progress` behavior in `Invoke-GraphBatch` and FullAccess loop.
- Produces: `Write-PhaseStatus -Message <string> [-Color <ConsoleColor>]` and `Write-PhaseProgress -Activity <string> -Status <string> -Current <int> -Total <int> [-Completed]`.

- [ ] **Step 1: Add minimal helper functions**

Add these functions after `Initialize-OutputFolder`:

```powershell
function Write-PhaseStatus {
    param(
        [Parameter(Mandatory = $true)] [string]$Message,
        [Parameter(Mandatory = $false)] [ConsoleColor]$Color = [ConsoleColor]::Cyan
    )

    Write-Host $Message -ForegroundColor $Color
}

function Write-PhaseProgress {
    param(
        [Parameter(Mandatory = $true)] [string]$Activity,
        [Parameter(Mandatory = $false)] [string]$Status = '',
        [Parameter(Mandatory = $false)] [int]$Current = 0,
        [Parameter(Mandatory = $false)] [int]$Total = 1,
        [Parameter(Mandatory = $false)] [switch]$Completed
    )

    if ($Completed) {
        Write-Progress -Activity $Activity -Completed
        return
    }

    $safeTotal = [Math]::Max($Total, 1)
    $safeCurrent = [Math]::Min([Math]::Max($Current, 0), $safeTotal)
    Write-Progress -Activity $Activity -Status $Status -PercentComplete (($safeCurrent / $safeTotal) * 100)
}
```

- [ ] **Step 2: Instrument connection and inventory phases**

Replace direct phase `Write-Host` calls near the main script start with `Write-PhaseStatus`, and add progress before/after Graph and Exchange connection:

```powershell
Write-PhaseStatus -Message "Output folder: $resolvedOutputFolder"
Write-PhaseStatus -Message "Recipient type: $RecipientType"

Write-PhaseProgress -Activity 'Connecting to Microsoft Graph' -Status 'Starting Graph authentication' -Current 0 -Total 1
$graphConnected = Connect-GraphIfNeeded
Write-PhaseProgress -Activity 'Connecting to Microsoft Graph' -Completed
Write-Verbose "Microsoft Graph connected: $graphConnected"

Write-PhaseProgress -Activity 'Connecting to Exchange Online' -Status 'Importing ExchangeOnlineManagement' -Current 0 -Total 2
Import-Module ExchangeOnlineManagement -ErrorAction Stop
Write-PhaseProgress -Activity 'Connecting to Exchange Online' -Status 'Connecting session' -Current 1 -Total 2
Connect-ExchangeOnline -ShowBanner:$false
Write-PhaseProgress -Activity 'Connecting to Exchange Online' -Completed

Write-PhaseProgress -Activity 'Loading Exchange recipients' -Status "Recipient type: $RecipientType" -Current 0 -Total 1
$targetRecipients = @(Get-TargetRecipients)
$objects = @($targetRecipients | ForEach-Object { ConvertTo-ExportObject -Recipient $_ })
Write-PhaseProgress -Activity 'Loading Exchange recipients' -Completed
Write-PhaseStatus -Message "Objects selected: $($objects.Count)" -Color Green
Write-Verbose "Loaded $($targetRecipients.Count) target recipients."
```

- [ ] **Step 3: Instrument proxy derivation loop**

Replace the proxy loop with:

```powershell
Write-PhaseStatus -Message 'Deriving proxy address rows from inventory...'
$proxyIndex = 0
foreach ($recipient in $targetRecipients) {
    $proxyIndex++
    Write-PhaseProgress -Activity 'Deriving proxy address rows' -Status "$proxyIndex of $($targetRecipients.Count): $($recipient.Identity)" -Current $proxyIndex -Total $targetRecipients.Count
    Write-Verbose "Deriving proxy rows for $($recipient.Identity)."
    foreach ($row in @(ConvertTo-ProxyRows -Recipient $recipient)) { $proxyRows.Add($row) | Out-Null }
}
Write-PhaseProgress -Activity 'Deriving proxy address rows' -Completed
Write-PhaseStatus -Message "Proxy rows derived: $($proxyRows.Count)" -Color Green
```

- [ ] **Step 4: Instrument SendAs collection**

Replace the SendAs collection line with:

```powershell
Write-PhaseStatus -Message 'Collecting SendAs permissions...'
Write-PhaseProgress -Activity 'Collecting SendAs permissions' -Status 'Querying recipient permissions' -Current 0 -Total 1
foreach ($row in @(Get-SendAsRows -Recipients $targetRecipients -UseOrgWideQuery ($PSCmdlet.ParameterSetName -eq 'All'))) { $sendAsRows.Add($row) | Out-Null }
Write-PhaseProgress -Activity 'Collecting SendAs permissions' -Completed
Write-PhaseStatus -Message "SendAs rows collected: $($sendAsRows.Count)" -Color Green
```

Inside `Get-SendAsRows`, add this before per-recipient lookup:

```powershell
Write-Verbose "Collecting SendAs for $($recipient.Identity) using $permissionCommand lookup '$lookupIdentity'."
```

- [ ] **Step 5: Instrument FullAccess, consolidation, CSV, dashboard, cleanup**

Add count status after FullAccess loop:

```powershell
Write-PhaseStatus -Message "FullAccess rows collected: $($fullAccessRows.Count)" -Color Green
```

Wrap consolidation:

```powershell
Write-PhaseProgress -Activity 'Creating consolidated rows' -Status 'Combining collected data' -Current 0 -Total 1
$consolidatedRows = @(
    New-ConsolidatedRows `
        -Objects $objects `
        -ProxyRows @($proxyRows) `
        -FullAccessRows @($fullAccessRows) `
        -SendAsRows @($sendAsRows) `
        -ExchangeGroupRows @($exchangeGroupMembershipRows) `
        -EntraGroupRows @($entraGroupMembershipRows) `
        -ErrorRows @($script:Errors)
)
Write-PhaseProgress -Activity 'Creating consolidated rows' -Completed
Write-PhaseStatus -Message "Consolidated rows created: $($consolidatedRows.Count)" -Color Green
```

Wrap CSV export:

```powershell
Write-PhaseProgress -Activity 'Exporting CSV reports' -Status 'Writing report CSV files' -Current 0 -Total 1
Export-ReportCsvs `
    -Folder $resolvedOutputFolder `
    -Objects $objects `
    -ConsolidatedRows $consolidatedRows `
    -ProxyRows @($proxyRows) `
    -FullAccessRows @($fullAccessRows) `
    -SendAsRows @($sendAsRows) `
    -ExchangeGroupRows @($exchangeGroupMembershipRows) `
    -EntraGroupRows @($entraGroupMembershipRows) `
    -ErrorRows @($script:Errors)
Write-PhaseProgress -Activity 'Exporting CSV reports' -Completed
Write-PhaseStatus -Message "CSV exports written to: $resolvedOutputFolder" -Color Green
```

Wrap dashboard generation with `Write-PhaseProgress -Activity 'Generating dashboard'`, and add cleanup progress before disconnect calls.

- [ ] **Step 6: Run checks to verify implementation**

Run:

```bash
pwsh -NoProfile -File "tests/Export-ExchangeObjectData.Checks.ps1"
```

Expected: PASS with all checks passing.

- [ ] **Step 7: Sync Fortescue copy**

Run:

```bash
cp "/Users/abhijittiwari/Documents/AdminToolboxScripts/Export-ExchangeObjectData.ps1" "/Users/abhijittiwari/Documents/Fortescue/Export-ExchangeObjectData.ps1"
```

- [ ] **Step 8: Verify both copies match**

Run:

```bash
cmp -s "/Users/abhijittiwari/Documents/AdminToolboxScripts/Export-ExchangeObjectData.ps1" "/Users/abhijittiwari/Documents/Fortescue/Export-ExchangeObjectData.ps1"; rc=$?; if [ $rc -eq 0 ]; then printf 'MATCH\n'; else printf 'DIFFER\n'; fi
```

Expected: `MATCH`.

---

## Self-Review

Spec coverage: The plan covers phase progress, default concise status, verbose diagnostics, no schema/logic changes, validation, and copy sync.

Placeholder scan: No placeholders remain.

Type consistency: Helper function names and parameters are consistent across tasks.
