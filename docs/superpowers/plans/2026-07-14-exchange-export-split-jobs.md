# Exchange Export Split Jobs and Excel Workbook Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Ship five standalone scripts — recipient inventory, group memberships, SendAs, FullAccess, and an Excel workbook collator — that together reproduce `Export-ExchangeObjectData.ps1`'s outputs job by job, leaving the monolith untouched.

**Architecture:** Each job script is self-contained (repo convention) and embeds verbatim copies of the monolith helpers it needs. The inventory script is the only one that enumerates recipients; the other jobs consume its `Objects.csv`. The workbook script is fully offline and reads the CSV folder.

**Tech Stack:** PowerShell 7, ExchangeOnlineManagement, Microsoft.Graph.Authentication, ImportExcel 7.x (installed).

## Global Constraints

- `Export-ExchangeObjectData.ps1` must not change (spec: "stays exactly as it is").
- CSV file names and column sets must match the monolith so the workbook script accepts either producer.
- Job scripts reuse an existing EXO/Graph session and do not disconnect (spec: Connection Behavior).
- Each job writes its own `Errors-<Job>.csv`; workbook merges `Errors*.csv`.
- Tests follow `tests/<Script>.Checks.ps1` AST-extraction pattern from `tests/Export-ExchangeObjectData.Checks.ps1`.
- Do not commit; the user commits when ready (session precedent: working tree already holds uncommitted reviewed work).

---

### Task 1: Export-ExchangeObjectInventory.ps1

**Files:**
- Create: `Export-ExchangeObjectInventory.ps1`
- Test: `tests/Export-ExchangeObjectInventory.Checks.ps1`

**Interfaces:**
- Produces: `Objects.csv` with headers `Identity, DisplayName, RecipientType, RecipientTypeDetails, MailboxType, PrimarySmtpAddress, Alias, ExternalDirectoryObjectId, ExchangeObjectId, DistinguishedName`; `ProxyAddresses.csv` with headers `Identity, PrimarySmtpAddress, AddressType, ProxyAddress, RawProxyAddress, IsPrimarySmtp`; `Errors-Inventory.csv` with headers `Identity, Stage, Operation, Message`. Tasks 2–4 consume `Objects.csv`.

- [ ] **Step 1: Write failing checks** — new checks file asserting: `ConvertTo-ProxyRows` parses primary/secondary/non-SMTP addresses (copy the four assertions from `Test-ConvertToProxyRows` in the monolith checks); `Test-RecipientMatchesType` accepts UserMailbox as Mailbox, GroupMailbox as Group not Mailbox, MailUser, MailContact, and All; `ConvertTo-ExportObject` maps `Guid` → `ExchangeObjectId`; static: script text contains `Errors-Inventory.csv`, contains `Get-ConnectionInformation`, does not contain `Disconnect-ExchangeOnline`.
- [ ] **Step 2: Run** `pwsh -NoProfile -File tests/Export-ExchangeObjectInventory.Checks.ps1` — expect parse failure (script missing).
- [ ] **Step 3: Implement.** Param block: parameter sets `All` (`-All`), `Csv` (`-InputCsv` + `-IdentityColumn`, default `Identity`), `Identity` (`-Identity`); shared `-RecipientType` (same ValidateSet), `-OutputFolder` default `.\ExchangeObjectInventory_$(Get-Date -Format 'yyyyMMdd_HHmmss')`. Copy verbatim from `Export-ExchangeObjectData.ps1`: `Add-ExportError` (+ `$script:Errors`), `Initialize-OutputFolder`, `Get-InputIdentityList`, `Test-RecipientMatchesType`, `Get-MailboxType`, `ConvertTo-ExportObject`, `ConvertTo-ProxyRows`, `Get-TargetRecipients`, `Export-CsvWithHeaders`. New: `Connect-ExchangeOnlineIfNeeded` = `Import-Module ExchangeOnlineManagement`; `if (-not (Get-ConnectionInformation)) { Connect-ExchangeOnline -ShowBanner:$false }`. Main body: connect → `Get-TargetRecipients` → build object + proxy rows with `Write-Progress` → write the three CSVs → print counts. No disconnect.
- [ ] **Step 4: Run checks — expect all PASS.**

### Task 2: Export-ExchangeGroupMemberships.ps1

**Files:**
- Create: `Export-ExchangeGroupMemberships.ps1`
- Test: `tests/Export-ExchangeGroupMemberships.Checks.ps1`

**Interfaces:**
- Consumes: `Objects.csv` (Task 1) via `-ObjectsCsv`.
- Produces: `EntraGroupMemberships.csv` (`Identity, PrimarySmtpAddress, GroupId, GroupDisplayName, GroupMail, GroupSecurityEnabled`), `ExchangeGroupMemberships.csv` (`Identity, PrimarySmtpAddress, GroupIdentity`), `Errors-GroupMemberships.csv`.
- Shared helper `Import-ObjectsCsv -Path -RequiredColumns` (returns rows array; throws naming the first missing column) — same shape reused in Tasks 3–4.

- [ ] **Step 1: Write failing checks** — port `Test-InvokeGraphBatch` and `Test-GetMembershipRows` (including the `/users/`, `/groups/`, `/contacts/` URL-segment and no-`/directoryObjects/` assertions) from the monolith checks; add `Import-ObjectsCsv` happy path + missing-column throw using a temp CSV; static: script contains `Errors-GroupMemberships.csv` and `Get-MgContext`, not `Connect-ExchangeOnline`.
- [ ] **Step 2: Run checks — expect FAIL.**
- [ ] **Step 3: Implement.** Params: `-ObjectsCsv` (mandatory, `Test-Path` leaf), `-OutputFolder` default `Split-Path -Parent` of resolved objects CSV. Copy verbatim: `Add-ExportError`, `Initialize-OutputFolder`, `New-GraphRequestExecutor`, `Invoke-GraphBatch`, `Get-GraphMemberOfResourceSegment`, `Get-MembershipRows`, `Export-CsvWithHeaders`. New: `Import-ObjectsCsv`; `Connect-GraphIfNeeded` variant = `Import-Module Microsoft.Graph.Authentication`; `if (-not (Get-MgContext)) { Connect-MgGraph -Scopes @('User.Read.All','Group.Read.All','Directory.Read.All') -NoWelcome }`. Main: import rows (required: `Identity, PrimarySmtpAddress, ExternalDirectoryObjectId, RecipientTypeDetails, RecipientType`) → connect → `Get-MembershipRows` → write CSVs.
- [ ] **Step 4: Run checks — expect PASS.**

### Task 3: Export-ExchangeSendAsPermissions.ps1

**Files:**
- Create: `Export-ExchangeSendAsPermissions.ps1`
- Test: `tests/Export-ExchangeSendAsPermissions.Checks.ps1`

**Interfaces:**
- Consumes: `Objects.csv` via `-ObjectsCsv` (required columns `Identity, PrimarySmtpAddress, RecipientTypeDetails, RecipientType`).
- Produces: `SendAs.csv` (`Identity, PrimarySmtpAddress, Trustee, AccessRights, IsInherited`), `Errors-SendAs.csv`.

- [ ] **Step 1: Write failing checks** — port `Test-SendAsAndFullAccess`'s SendAs sections (per-recipient EXO lookups by SMTP, SELF/inherited filtering, legacy-cmdlet guard); static: contains `Errors-SendAs.csv` and `Get-ConnectionInformation`, no `Disconnect-ExchangeOnline`.
- [ ] **Step 2: Run checks — expect FAIL.**
- [ ] **Step 3: Implement.** Copy verbatim: `Add-ExportError`, `Initialize-OutputFolder`, `Test-IsSelfTrustee`, `Get-SendAsPermissionCommandName`, `Get-SendAsRows`, `Export-CsvWithHeaders`; plus `Import-ObjectsCsv`, `Connect-ExchangeOnlineIfNeeded` (Task 1 shape). Main: import rows → connect → `Get-SendAsRows -Recipients $rows -UseOrgWideQuery ($rows.Count -gt 1)` → write CSVs.
- [ ] **Step 4: Run checks — expect PASS.**

### Task 4: Export-ExchangeFullAccessPermissions.ps1

**Files:**
- Create: `Export-ExchangeFullAccessPermissions.ps1`
- Test: `tests/Export-ExchangeFullAccessPermissions.Checks.ps1`

**Interfaces:**
- Consumes: `Objects.csv` via `-ObjectsCsv` (required columns as Task 3).
- Produces: `FullAccess.csv` (`Identity, PrimarySmtpAddress, Trustee, AccessRights, Deny, IsInherited`), `Errors-FullAccess.csv`.

- [ ] **Step 1: Write failing checks** — `Get-FullAccessRows` uses PrimarySmtpAddress lookup and maps rows (port from monolith checks); returns `@()` for a MailContact row (non-mailbox); static: contains `Errors-FullAccess.csv`.
- [ ] **Step 2: Run checks — expect FAIL.**
- [ ] **Step 3: Implement.** Copy verbatim: `Add-ExportError`, `Initialize-OutputFolder`, `Get-MailboxType`, `Get-FullAccessRows`, `Export-CsvWithHeaders`; plus `Import-ObjectsCsv`, `Connect-ExchangeOnlineIfNeeded`. Main: import rows → filter to mailbox types via `Get-MailboxType` → connect → per-mailbox loop with `Write-Progress` → write CSVs.
- [ ] **Step 4: Run checks — expect PASS.**

### Task 5: Export-ExchangeObjectDataWorkbook.ps1

**Files:**
- Create: `Export-ExchangeObjectDataWorkbook.ps1`
- Test: `tests/Export-ExchangeObjectDataWorkbook.Checks.ps1`

**Interfaces:**
- Consumes: CSV folder from Tasks 1–4 or from the monolith (`-InputFolder`); optional `-OutputPath`, `-Force`.
- Produces: one `.xlsx` with sheets `Summary`, `Consolidated`, then each non-empty dataset (`Objects`, `ProxyAddresses`, `FullAccess`, `SendAs`, `ExchangeGroupMemberships`, `EntraGroupMemberships`, `Errors`).

- [ ] **Step 1: Write failing checks** — end-to-end: build a temp folder with `Objects.csv` (2 objects), `ProxyAddresses.csv`, `SendAs.csv`, `Errors-SendAs.csv` + `Errors-Inventory.csv`, run the script, then `Import-Excel` and assert: Summary lists row counts; Consolidated has 2 rows with joined `ProxyAddresses`/`SendAsTrustees` and `HasErrors`; Errors sheet merges both error CSVs; no `FullAccess` sheet (empty dataset skipped). Negative: missing `Objects.csv` throws; existing output without `-Force` throws; with `-Force` overwrites (no stale sheets).
- [ ] **Step 2: Run checks — expect FAIL.**
- [ ] **Step 3: Implement.** Guard: `if (-not (Get-Module -ListAvailable ImportExcel)) { throw 'ImportExcel module required. Install with: Install-Module ImportExcel -Scope CurrentUser' }`. Copy verbatim: `Join-Values`, `Group-RowsByIdentity`, `New-ConsolidatedRows`. New: `Import-DataSet -Folder -FileName` (missing optional file → `Write-Warning` + `@()`), error-CSV glob merge, output handling (`Test-Path` + `-Force` → `Remove-Item` before export so stale sheets never survive), sheet writer loop:

```powershell
$rows | Export-Excel -Path $OutputPath -WorksheetName $name -AutoFilter -FreezeTopRow -BoldTopRow
```

- [ ] **Step 4: Run checks — expect PASS.**

### Task 6: Documentation and full verification

**Files:**
- Modify: `README.md` (new section after the Export-ExchangeObjectData docs)

- [ ] **Step 1:** Add README section "Modular Exchange export jobs" documenting the chained workflow (`inventory → memberships/SendAs/FullAccess → workbook`), each script's parameters and outputs, the shared-session behavior, and the ImportExcel prerequisite.
- [ ] **Step 2:** Run every checks file under `tests/` — all must pass, including the two pre-existing ones.
