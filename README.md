# Admin Toolbox Scripts

Collection of administrative scripts for Microsoft 365, Entra ID, Active Directory, and mail DNS troubleshooting.

## Documentation

- [Quick Start](docs/quick-start.md)
- [Script Reference](docs/script-reference.md)
- [Operations and Security](docs/operations-and-security.md)

## Scripts

| Script | Platform | Purpose |
| --- | --- | --- |
| `Export-ExchangeObjectData.ps1` | PowerShell | Exports Exchange Online recipient proxies, permissions, Exchange/Entra group memberships, CSV reports, and a searchable HTML dashboard. |
| `Export-ExchangeObjectInventory.ps1` | PowerShell | Modular job: exports the recipient inventory (`Objects.csv`) and proxy addresses that feed the other export jobs. |
| `Export-ExchangeGroupMemberships.ps1` | PowerShell | Modular job: exports Entra ID and Exchange group memberships via Microsoft Graph for an `Objects.csv` inventory. |
| `Export-ExchangeSendAsPermissions.ps1` | PowerShell | Modular job: exports SendAs permissions for an `Objects.csv` inventory. |
| `Export-ExchangeFullAccessPermissions.ps1` | PowerShell | Modular job: exports FullAccess mailbox permissions for an `Objects.csv` inventory. |
| `Export-ExchangeObjectDataWorkbook.ps1` | PowerShell | Collates the export CSVs into a single Excel workbook (requires ImportExcel; no Excel needed). |
| `Invoke-ExchangeObjectDataExport.ps1` | PowerShell | Runs all five modular export jobs in sequence: inventory, memberships, SendAs, FullAccess, workbook. |
| `Get-EntraAdminAccounts.ps1` | PowerShell | Exports Entra ID admin users, active and PIM-eligible role assignments, MFA status, licenses, mailbox details, group membership, and sign-in activity. |
| `Get-CrossTenantAdminRoleAssignments.ps1` | PowerShell | Uses a WAE/Fortescue admin mapping CSV to export Entra ID active and PIM-eligible directory roles, immutable IDs, UPNs, display names, and licenses from both tenants. |
| `Find-ADDuplicateEmailProxyAddresses.ps1` | PowerShell | Finds duplicate mail-related values across on-prem Active Directory objects. |
| `Get-MailDnsRecords.ps1` | PowerShell | Checks MX, SPF, DMARC, and DKIM DNS records on Windows. |
| `get-mail-dns-records.sh` | Bash | Checks MX, SPF, DMARC, and DKIM DNS records on macOS/Linux using `dig`. |

## Export-ExchangeObjectData.ps1

Exports Exchange Online recipient data for mailboxes by default, or for selected recipient object types. The script can process all matching recipients, multiple identities from a CSV, or a single identity. It builds a bulk recipient inventory (including proxy addresses) from Exchange Online, collects Exchange and Entra ID group memberships via Microsoft Graph `$batch` requests, gathers SendAs permissions in one org-wide sweep on `-All` runs, and looks up FullAccess permissions per mailbox by GUID, then writes normalized CSV files, a consolidated CSV, and a searchable HTML dashboard for multi-object runs.

### Requirements

- PowerShell 7 recommended (Windows PowerShell 5.1 compatible).
- ExchangeOnlineManagement module: `Install-Module ExchangeOnlineManagement`.
- Microsoft.Graph.Authentication module: `Install-Module Microsoft.Graph.Authentication`.
- Exchange Online permission to read recipients and permissions.
- Microsoft Graph delegated scopes for group memberships:
  - `User.Read.All`
  - `Group.Read.All`
  - `Directory.Read.All`

Graph permissions may require admin consent depending on tenant policy.

### Connection Behavior

The script connects to Exchange Online on every run. Unless `-SkipGraph` is used, it also connects to Microsoft Graph to collect group memberships; both `ExchangeGroupMemberships.csv` and `EntraGroupMemberships.csv` are Graph-sourced, so with `-SkipGraph`, Graph connection is skipped and both files are written headers-only.

### Parameters

| Parameter | Type | Default | Description |
| --- | --- | --- | --- |
| `-All` | Switch | Off | Exports all recipients matching `-RecipientType`. |
| `-InputCsv` | String | None | CSV file containing object identities to export. |
| `-Identity` | String | None | Single Exchange-resolvable identity to export. |
| `-IdentityColumn` | String | `Identity` | CSV column containing identities. |
| `-RecipientType` | String | `Mailbox` | Object type filter: `Mailbox`, `Group`, `MailUser`, `MailContact`, or `All`. |
| `-OutputFolder` | String | Timestamped folder | Folder where CSVs and dashboard are written. |
| `-SkipGraph` | Switch | Off | Skips Microsoft Graph connection; `ExchangeGroupMemberships.csv` and `EntraGroupMemberships.csv` are written headers-only. |
| `-NoDashboard` | Switch | Off | Suppresses `Dashboard.html` generation for multi-object runs. |

### Usage

```powershell
./Export-ExchangeObjectData.ps1 -All
```

```powershell
./Export-ExchangeObjectData.ps1 -All -RecipientType All
```

```powershell
./Export-ExchangeObjectData.ps1 -InputCsv ./objects.csv -IdentityColumn Identity -RecipientType Group
```

```powershell
./Export-ExchangeObjectData.ps1 -Identity user@contoso.com
```

### Output Files

- `Objects-Consolidated.csv`
- `Objects.csv`
- `ProxyAddresses.csv`
- `FullAccess.csv`
- `SendAs.csv`
- `ExchangeGroupMemberships.csv`
- `EntraGroupMemberships.csv`
- `Errors.csv`
- `Dashboard.html` for `-All` and `-InputCsv` runs unless `-NoDashboard` is used

### Dashboard

The dashboard is a static local HTML file. It supports search, checkbox selection, expandable object details, and browser-side CSV export of selected objects and selected detail rows.

### Manual Validation

- Confirm the output folder contains the expected CSV files and, for `-All` or `-InputCsv` runs without `-NoDashboard`, `Dashboard.html`.
- For multi-object runs, open `Dashboard.html` and verify search, checkbox selection, row expansion, and selected-row CSV export; for single `-Identity` runs, confirm no dashboard is expected.
- Review `Errors.csv` after every run. An empty file or headers-only file indicates no logged lookup/export errors; any rows should be checked against console warnings and expected access limitations.
- Spot-check a small sample of objects against Exchange Online: primary SMTP and proxy addresses in `ProxyAddresses.csv`, mailbox type in `Objects.csv`, FullAccess/SendAs permissions, and Exchange/Entra group memberships when Graph was not skipped.

### Notes

- The default recipient type is `Mailbox`.
- Mailbox exports include `MailboxType`, such as `UserMailbox`, `SharedMailbox`, `RoomMailbox`, or `EquipmentMailbox`.
- The script is read-only.
- If a lookup fails for an object or permission type, the script continues and records the failure in `Errors.csv`.
- Export files and the dashboard may contain sensitive proxy addresses, mailbox permissions, group memberships, and object identifiers. Handle them as sensitive administrative data.

## Modular Exchange Export Jobs

Each stage of `Export-ExchangeObjectData.ps1` is also available as a standalone job script, plus a collation script that turns the CSV folder into a single Excel workbook. The monolith is unchanged; use whichever fits the run.

### Workflow

Run everything in one go with the orchestrator:

```powershell
./Invoke-ExchangeObjectDataExport.ps1 -All -RecipientType Mailbox -OutputFolder ./export
```

It takes the same selection parameters as the inventory script plus `-Force` (passed to the workbook step). If the inventory fails the run aborts; if a later job fails the remaining jobs still run, the workbook is built from whatever CSVs exist, and the script exits with code 1 naming the failed jobs.

Or run the jobs individually:

```powershell
# 1. Inventory (connects to Exchange Online, writes Objects.csv + ProxyAddresses.csv)
./Export-ExchangeObjectInventory.ps1 -All -RecipientType Mailbox -OutputFolder ./export

# 2. Independent jobs against the inventory (run any subset, any order)
./Export-ExchangeGroupMemberships.ps1 -ObjectsCsv ./export/Objects.csv      # Microsoft Graph only
./Export-ExchangeSendAsPermissions.ps1 -ObjectsCsv ./export/Objects.csv     # Exchange Online
./Export-ExchangeFullAccessPermissions.ps1 -ObjectsCsv ./export/Objects.csv # Exchange Online

# 3. Collate everything in the folder into one workbook (offline)
./Export-ExchangeObjectDataWorkbook.ps1 -InputFolder ./export
```

### Requirements

- `Export-ExchangeObjectInventory.ps1`, `Export-ExchangeSendAsPermissions.ps1`, `Export-ExchangeFullAccessPermissions.ps1`: ExchangeOnlineManagement module.
- `Export-ExchangeGroupMemberships.ps1`: Microsoft.Graph.Authentication module (no Exchange session needed).
- `Export-ExchangeObjectDataWorkbook.ps1`: ImportExcel module (`Install-Module ImportExcel -Scope CurrentUser`). Excel itself is not required and the workbook builds on macOS/Linux/Windows.

### Parameters

- `Export-ExchangeObjectInventory.ps1` takes the same selection parameters as the monolith: `-All`, `-InputCsv` + `-IdentityColumn`, or `-Identity`, plus `-RecipientType` and `-OutputFolder`.
- The membership/permission jobs take `-ObjectsCsv` (path to `Objects.csv`) and optional `-OutputFolder` (defaults to the folder containing the objects CSV).
- `Export-ExchangeObjectDataWorkbook.ps1` takes `-InputFolder`, optional `-OutputPath` (default `<InputFolder>/ExchangeObjectData.xlsx`), and `-Force` to overwrite an existing workbook.

### Connection Behavior

Job scripts reuse an existing Exchange Online or Microsoft Graph session when one exists and leave the session open, so a chained run authenticates once. Disconnect manually with `Disconnect-ExchangeOnline` / `Disconnect-MgGraph` when finished. (The monolith still manages and disconnects its own sessions.)

### Output Files

The jobs write the same CSV names and columns as the monolith (`Objects.csv`, `ProxyAddresses.csv`, `EntraGroupMemberships.csv`, `ExchangeGroupMemberships.csv`, `SendAs.csv`, `FullAccess.csv`), so the workbook script accepts a folder produced by either. Each job writes its own `Errors-<Job>.csv`; the workbook merges every `Errors*.csv` into one `Errors` sheet.

The workbook contains a `Summary` sheet (row counts per dataset), a `Consolidated` sheet (one row per object with joined proxies, trustees, memberships, and a `HasErrors` flag), and one sheet per non-empty dataset, each with a bold frozen header row and auto-filter.

### Notes

- All job scripts are read-only against Exchange and Entra.
- Recoverable per-object failures are logged to the job's errors CSV and the run continues, matching the monolith.
- The export CSVs and workbook may contain sensitive proxy addresses, mailbox permissions, group memberships, and object identifiers. Handle them as sensitive administrative data.

## Get-EntraAdminAccounts.ps1

Reports all Entra ID user accounts with directory admin access. The report includes active role assignments, PIM-eligible role assignments, role-assignable group membership expansion, license details, MFA registration, mailbox information, group memberships, account status, created date, and last sign-in activity.

### Requirements

- PowerShell.
- Microsoft Graph PowerShell SDK: `Install-Module Microsoft.Graph`.
- Recommended Entra role: Global Reader or equivalent.
- Microsoft Graph permissions:
  - `Directory.Read.All`
  - `RoleManagement.Read.Directory`
  - `AuditLog.Read.All`
  - `UserAuthenticationMethod.Read.All`
- Optional Exchange Online mailbox lookup requires the `ExchangeOnlineManagement` module.
- MFA registration fields are read from the Microsoft Graph authentication method user registration report.

### Parameters

| Parameter | Type | Description |
| --- | --- | --- |
| `-IncludeMailbox` | Switch | Connects to Exchange Online and reports mailbox type and primary SMTP address. If omitted, the script uses the user's Graph `Mail` attribute. |
| `-OutputPath` | String | CSV output path. Defaults to `./EntraAdminAccounts_<timestamp>.csv`. |

### Usage

```powershell
./Get-EntraAdminAccounts.ps1
```

```powershell
./Get-EntraAdminAccounts.ps1 -IncludeMailbox
```

```powershell
./Get-EntraAdminAccounts.ps1 -OutputPath "C:\Temp\EntraAdminAccounts.csv"
```

### Output

The script exports a CSV and prints a console summary table. CSV columns include:

- `DisplayName`
- `UPN`
- `ImmutableId`
- `AccountEnabled`
- `Roles`
- `Licenses`
- `MfaRegistered`
- `MfaMethods`
- `MfaDefaultMethod`
- `Mailbox`
- `GroupMembership`
- `LastSignIn`
- `LastInteractiveSignIn`
- `LastNonInteractiveSignIn`
- `CreatedDate`

### Notes

- Service principals are excluded from the user report.
- Group-assigned roles are expanded to transitive user members and marked as `via group` in the `Roles` column.
- PIM eligibility collection may warn or skip if the tenant lacks Entra ID P2 or the signed-in account lacks permission.
- If MFA registration details cannot be retrieved for a user, the script writes a warning and leaves `MfaRegistered` as `Unknown` with blank method fields.
- The script disconnects from Microsoft Graph and Exchange Online at the end of execution.

## Get-CrossTenantAdminRoleAssignments.ps1

Exports WAE and Fortescue admin account details from Microsoft Graph using a mapping CSV with `WAEUPN`, `Prefix`, and `FortescueUPN` columns. The report includes Entra ID active directory roles, PIM-eligible directory roles when readable, immutable ID, actual user principal name, display name, assigned license SKU part numbers, lookup status, and errors. If a Fortescue UPN is missing or cannot be found, the script strips `(Admin)` from the matching WAE display name and tries exact Fortescue `displayName` lookups for `Admin <name>` first, then `<name>`.

### Requirements

- PowerShell.
- Microsoft Graph PowerShell SDK: `Install-Module Microsoft.Graph`.
- Recommended Entra role: Global Reader or equivalent in each tenant.
- Microsoft Graph permissions:
  - `Directory.Read.All`
  - `RoleManagement.Read.Directory`

### Parameters

| Parameter | Type | Description |
| --- | --- | --- |
| `-InputCsv` | String | Admin mapping CSV path. Defaults to `./Admin.csv`. |
| `-OutputPath` | String | Combined CSV output path. Defaults to `./CrossTenantAdminRoles_yyyyMMdd_HHmmss.csv`. |

### Usage

```powershell
./Get-CrossTenantAdminRoleAssignments.ps1 -InputCsv ./Admin.csv
```

```powershell
./Get-CrossTenantAdminRoleAssignments.ps1 -InputCsv ./Admin.csv -OutputPath ./WAE-Fortescue-AdminRoles.csv
```

### Notes

- The script connects to Microsoft Graph once for WAE and once for Fortescue; sign into the requested tenant at each prompt.
- Exchange Online RBAC roles are not included.
- If PIM eligibility cannot be read, active roles are still exported and the PIM warning is included in the `Error` column.
- Fortescue rows resolved by display-name fallback use `LookupStatus=FoundByDisplayName`; ambiguous display-name matches are not guessed.
- Missing users are exported with `LookupStatus=NotFound`.

## Find-ADDuplicateEmailProxyAddresses.ps1

Scans on-prem Active Directory objects for duplicate mail-related values. It checks `proxyAddresses` and `mail` by default, with optional support for `targetAddress`, `userPrincipalName`, and custom attributes.

The script is read-only. It does not modify Active Directory.

### Requirements

- Windows PowerShell or PowerShell with RSAT Active Directory module available.
- RSAT Active Directory PowerShell module.
- Permission to read the selected AD search base and schema.

### Parameters

| Parameter | Type | Default | Description |
| --- | --- | --- | --- |
| `-SearchBase` | String | Domain default naming context | Distinguished name where the AD scan starts. |
| `-Server` | String | Current logon domain controller | Domain controller or AD server to query. |
| `-OutputCsv` | String | `./AD-Duplicate-Address-Pairs.csv` | Path for pair-based duplicate results. |
| `-OccurrenceCsv` | String | `./AD-Duplicate-Address-Occurrences.csv` | Path for duplicate occurrence details. |
| `-AdditionalAttributes` | String array | Empty | Extra AD attributes to scan. Attributes are validated against schema before scanning. |
| `-IncludeTargetAddress` | Switch | Off | Includes `targetAddress` in the scan. |
| `-IncludeUserPrincipalName` | Switch | Off | Includes `userPrincipalName` in the scan. |
| `-SmtpOnly` | Switch | Off | Limits `proxyAddresses` checks to SMTP values only. |
| `-NoOccurrenceCsv` | Switch | Off | Suppresses the occurrence-level CSV export. |

### Usage

```powershell
./Find-ADDuplicateEmailProxyAddresses.ps1
```

```powershell
./Find-ADDuplicateEmailProxyAddresses.ps1 `
    -SearchBase "DC=contoso,DC=com" `
    -Server "dc01.contoso.com"
```

```powershell
./Find-ADDuplicateEmailProxyAddresses.ps1 `
    -IncludeTargetAddress `
    -IncludeUserPrincipalName `
    -SmtpOnly
```

```powershell
./Find-ADDuplicateEmailProxyAddresses.ps1 `
    -OutputCsv "C:\Temp\AD-Duplicate-Address-Pairs.csv" `
    -OccurrenceCsv "C:\Temp\AD-Duplicate-Address-Occurrences.csv"
```

### Duplicate Matching Behavior

- SMTP `proxyAddresses` values are normalized to the bare lower-case email address so they can match `mail` values.
- Non-SMTP `proxyAddresses` values keep their address type prefix, such as `x500:` or `sip:`, to avoid false matches across unrelated address types.
- Empty values are ignored.
- A duplicate is only reported when the same normalized value appears on more than one AD object.

### Output

The script writes pair-based duplicate rows to `-OutputCsv` and also returns those rows to the PowerShell pipeline.

Pair CSV fields include:

- `DuplicateMatchValue`
- `DuplicateObjectCount`
- Object 1 attribute, object type, object GUID, SID, name, sAMAccountName, and distinguished name
- Object 2 attribute, object type, object GUID, SID, name, sAMAccountName, and distinguished name

Unless `-NoOccurrenceCsv` is set, the script also exports every duplicate occurrence to `-OccurrenceCsv`.

### Notes

- Parent folders for CSV output paths are created automatically when needed.
- The script validates requested attributes against the AD schema and skips missing attributes with a warning.
- Output is pair-based, so a duplicate value on three objects produces three pair rows.

## Get-MailDnsRecords.ps1

Fetches mail-related DNS records for a domain from Windows PowerShell using `Resolve-DnsName`. It checks MX, SPF, DMARC, and DKIM records.

### Requirements

- Windows with the `DnsClient` module available.
- `Resolve-DnsName` command.
- Network access to the configured DNS resolver.

### Parameters

| Parameter | Type | Default | Description |
| --- | --- | --- | --- |
| `-Domain` | String | Required | Domain to inspect. URLs are normalized to the domain name. |
| `-DkimSelectors` | String array | Common selectors | DKIM selectors to query. |
| `-Server` | String | System default resolver | DNS resolver to use, such as `8.8.8.8` or `1.1.1.1`. |
| `-Json` | Switch | Off | Outputs the result object as JSON instead of formatted console output. |

### Usage

```powershell
./Get-MailDnsRecords.ps1 -Domain example.com
```

```powershell
./Get-MailDnsRecords.ps1 -Domain example.com -DkimSelectors selector1,selector2,google
```

```powershell
./Get-MailDnsRecords.ps1 -Domain example.com -Server 8.8.8.8
```

```powershell
./Get-MailDnsRecords.ps1 -Domain example.com -Json
```

### What It Checks

- MX records at the domain root.
- SPF TXT records at the domain root beginning with `v=spf1`.
- DMARC TXT records at `_dmarc.<domain>` beginning with `v=DMARC1`.
- DKIM TXT records at `<selector>._domainkey.<domain>`.
- DKIM CNAME records, then follows the CNAME target for TXT lookup when present.

### Output

Console output is grouped into sections:

- `MX Records`
- `SPF Records`
- `DMARC Records`
- `DKIM Records`
- `Warnings`, when applicable

With `-Json`, the script emits a structured object containing:

- `Domain`
- `QueriedAt`
- `DnsServer`
- `MX`
- `SPF`
- `DMARC`
- `DKIM`
- `Warnings`

### Notes

- DKIM selectors are not reliably discoverable from DNS alone. Provide known selectors when possible.
- The script warns when multiple SPF or DMARC records are found.
- DKIM status can be `Found`, `Not found`, or `TXT found; verify manually`.

## get-mail-dns-records.sh

Bash version of the mail DNS record checker. It is intended for macOS or Linux systems with `dig` available.

### Requirements

- Bash.
- `dig` command.
- Network access to the configured DNS resolver.

### Options

| Option | Description |
| --- | --- |
| `<domain>` | Domain to inspect. Can be provided positionally. |
| `-d`, `--domain DOMAIN` | Domain to inspect. URLs are normalized to the domain name. |
| `-s`, `--server DNS_SERVER` | DNS resolver to use, such as `8.8.8.8` or `1.1.1.1`. |
| `-k`, `--dkim-selectors SELECTORS` | Comma-separated DKIM selectors to query. |
| `-h`, `--help` | Shows usage help. |

### Usage

```bash
./get-mail-dns-records.sh example.com
```

```bash
./get-mail-dns-records.sh -d example.com -s 8.8.8.8
```

```bash
./get-mail-dns-records.sh -d example.com -k selector1,selector2,google
```

### What It Checks

- MX records at the domain root.
- SPF TXT records at the domain root beginning with `v=spf1`.
- DMARC TXT records at `_dmarc.<domain>` beginning with `v=DMARC1`.
- DKIM TXT records at `<selector>._domainkey.<domain>`.
- DKIM CNAME records, then follows the CNAME target for TXT lookup when present.

### Output

Console output is grouped into sections:

- `MX Records`
- `SPF Records`
- `DMARC Records`
- `DKIM Records`
- `Warnings`

### Notes

- TXT record chunks returned by `dig` are concatenated so long SPF, DMARC, and DKIM records are easier to read.
- The script warns when multiple SPF or DMARC records are found.
- DKIM status can be `Found`, `Not found`, or `TXT found; verify manually`.
- The script exits with an error if `dig` is not available or the domain cannot be parsed.

## General Safety Notes

- `Get-EntraAdminAccounts.ps1` reads Microsoft Graph and optionally Exchange Online data, then exports a CSV.
- `Find-ADDuplicateEmailProxyAddresses.ps1` is read-only against Active Directory and exports CSV reports.
- DNS scripts perform read-only DNS lookups.
- Review exported CSVs before sharing because they may contain user account identifiers, role assignments, group memberships, and mailbox information.
