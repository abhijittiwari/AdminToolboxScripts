# Exchange Export Split Jobs and Excel Workbook Design

## Goal

Provide each Export-ExchangeObjectData job as a standalone script that can run independently, plus a new script that collates the exported CSVs into a single Excel workbook. `Export-ExchangeObjectData.ps1` stays exactly as it is.

## Selected Approach

Five new standalone scripts at the repo root, matching the toolbox convention of self-contained single-file scripts (no shared module). Each job script embeds only the helper functions it needs, copied from the monolith so behavior stays identical.

The inventory script talks to Exchange Online and produces `Objects.csv`. Every other job script takes that `Objects.csv` as input (`-ObjectsCsv`), so recipients are resolved once and each job does exactly one thing. The workbook script is offline: it reads a folder of CSVs — produced either by the job scripts or by the monolith, which uses the same file names — and writes one `.xlsx`.

Alternatives considered:

- A shared `.psm1` module for common helpers. Rejected: breaks the copy-one-file portability of an admin toolbox and changes the repo's established layout.
- Job scripts that each re-query Exchange Online with `-All`/`-InputCsv`/`-Identity`. Rejected: repeats the expensive recipient inventory per job and duplicates the selection logic four times.

## Scripts

### Export-ExchangeObjectInventory.ps1

- Parameters: `-All` / `-InputCsv` + `-IdentityColumn` / `-Identity`, `-RecipientType`, `-OutputFolder` (default `.\ExchangeObjectInventory_<timestamp>`).
- Connects to Exchange Online (reuses an existing connection when `Get-ConnectionInformation` reports one).
- Writes `Objects.csv`, `ProxyAddresses.csv`, `Errors-Inventory.csv` with the same columns as the monolith.

### Export-ExchangeGroupMemberships.ps1

- Parameters: `-ObjectsCsv` (mandatory), `-OutputFolder` (default: folder containing the objects CSV).
- Connects to Microsoft Graph only (no Exchange Online session needed). Reuses the monolith's `$batch` + `memberOf` logic, including the per-recipient-type resource segment (`/users`, `/groups`, `/contacts`).
- Writes `EntraGroupMemberships.csv`, `ExchangeGroupMemberships.csv`, `Errors-GroupMemberships.csv`.

### Export-ExchangeSendAsPermissions.ps1

- Parameters: `-ObjectsCsv` (mandatory), `-OutputFolder` (default: folder containing the objects CSV).
- Connects to Exchange Online if needed. Same trustee filtering as the monolith (drops SELF, inherited, non-SendAs). Multi-object runs use the org-wide query only when just the legacy `Get-RecipientPermission` cmdlet is available, mirroring the monolith.
- Writes `SendAs.csv`, `Errors-SendAs.csv`.

### Export-ExchangeFullAccessPermissions.ps1

- Parameters: `-ObjectsCsv` (mandatory), `-OutputFolder` (default: folder containing the objects CSV).
- Connects to Exchange Online if needed. Mailbox-type recipients only, per-mailbox `Get-EXOMailboxPermission`, same filters as the monolith.
- Writes `FullAccess.csv`, `Errors-FullAccess.csv`.

### Export-ExchangeObjectDataWorkbook.ps1

- Parameters: `-InputFolder` (mandatory), `-OutputPath` (default `<InputFolder>\ExchangeObjectData.xlsx`), `-Force` to overwrite an existing workbook.
- Requires the ImportExcel module (works without Excel installed, cross-platform); fails fast with install instructions if missing.
- Reads `Objects.csv` (mandatory), `ProxyAddresses.csv`, `FullAccess.csv`, `SendAs.csv`, `ExchangeGroupMemberships.csv`, `EntraGroupMemberships.csv`, and all `Errors*.csv` (merged), tolerating missing optional files with a warning.
- Rebuilds the consolidated per-object rows with the same join logic as the monolith's `Objects-Consolidated.csv`.
- Writes worksheets: `Summary` (dataset, row count, source file), `Consolidated`, then one sheet per non-empty dataset, each with bold frozen header row and auto-filter.

## Connection Behavior

Job scripts connect when no session exists and leave the session open so chained runs (inventory → memberships → SendAs → FullAccess) authenticate once. This intentionally differs from the monolith, which owns its whole run and disconnects at the end.

## Error Handling

Same `Add-ExportError` pattern as the monolith: recoverable per-object failures are logged as warning + error row and the run continues. Each job writes its own uniquely named errors CSV so jobs sharing one output folder never clobber each other. The workbook script merges every `Errors*.csv` into one `Errors` sheet; the `Stage` column identifies the source job.

## Testing

One `tests/<Script>.Checks.ps1` per new script, following the existing AST-extraction pattern with mocked Exchange/Graph cmdlets. The workbook checks run end-to-end: build sample CSVs in a temp folder, produce a real `.xlsx`, read it back with `Import-Excel`, and assert sheet names and row contents.

## Out of Scope

- Any change to `Export-ExchangeObjectData.ps1`, its output schema, or its dashboard.
- A shared module or refactoring of the monolith to consume the new scripts.
- Workbook charts, pivot tables, or styling beyond header formatting and auto-filter.
