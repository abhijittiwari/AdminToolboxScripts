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
| `Export-ExchangeMigrationReadiness.ps1` | PowerShell | Identifies mail objects blocked, risky, or unsuitable for migration due to holds, retention, licensing, or compliance constraints. |
| `Export-DomainCutoverAssessment.ps1` | PowerShell | Assesses an accepted domain ahead of a tenant-to-tenant cutover: accepted-domain config, MX/SPF/DKIM DNS, and shared mailbox domain usage with hold status. |
| `Export-TenantDomainDependencyInventory.ps1` | PowerShell | Inventories groups, shared/resource mailboxes, routing, moderation, permissions, and address rewrite dependencies on a target SMTP domain. |
| `Get-EntraAdminAccounts.ps1` | PowerShell | Exports Entra ID admin users, active and PIM-eligible role assignments, MFA status, licenses, mailbox details, group membership, and sign-in activity. |
| `Get-EntraAdminMailboxInventory.ps1` | PowerShell | Exports Entra ID admin users, including active and PIM-eligible role assignments, with mailbox presence, Exchange GUID, license SKU names, and mailbox type. |
| `Get-EntraDirSyncProtectionSettings.ps1` | PowerShell | Checks Microsoft Graph directory synchronization protection flags for hard-match cloud object takeover and soft match. |
| `Get-CrossTenantAdminRoleAssignments.ps1` | PowerShell | Uses a source/target admin mapping CSV to export Entra ID active and PIM-eligible directory roles, immutable IDs, UPNs, display names, and licenses from both tenants. |
| `Get-CrossTenantUserMapping.ps1` | PowerShell | Exports all users from a source tenant, then matches each one to a target tenant account whose synced on-premises `extensionAttribute13` or `extensionAttribute14` contains the source UPN. |
| `Export-AdminCaPimPosture.ps1` | PowerShell | Maps Conditional Access policy applicability and PIM-eligible directory roles for in-scope admins into a combined Excel workbook with posture metrics. |
| `Invoke-AdSecurityHardeningAssessment.ps1` | PowerShell | Runs the Appendix B hardening assessment modules and merges their per-control results into one CSV with a consolidated summary. |
| `Export-DcHardeningAssessment.ps1` | PowerShell | Assesses Domain Controller hardening controls DC-001..DC-022 (supported OS, patching, Secure Boot/BitLocker, LDAP/SMB signing, NTLM, Netlogon, Print Spooler, audit policy, krbtgt age, backups). |
| `Export-DomainComputerHardeningAssessment.ps1` | PowerShell | Assesses domain-joined computer hardening controls DJ-001..DJ-020 (patching, LAPS, BitLocker, Secure Boot/TPM, Defender/EDR, ASR, legacy protocols, UAC, Credential Guard, stale accounts, machine account quota). |
| `Export-EntraIdHardeningAssessment.ps1` | PowerShell | Assesses Entra ID hardening controls AAD-001..AAD-015 via Microsoft Graph (MFA, phishing-resistant admin MFA, legacy auth, Conditional Access, PIM, break-glass, guest/consent restrictions, Identity Protection, access reviews). |
| `Export-DnsHardeningAssessment.ps1` | PowerShell | Assesses Windows DNS Server hardening controls DNS-001..DNS-010 (secure dynamic updates, zone transfers, scavenging, forwarders, recursion, DNSSEC, logging, global query block list, socket pool/cache locking). |
| `Export-AdcsHardeningAssessment.ps1` | PowerShell | Assesses AD CS hardening controls ADCS-001..ADCS-010 including the ESC1-ESC8 attack paths, CA audit logging, key protection, and CRL/OCSP management. |
| `Export-AdBackupRecoveryAssessment.ps1` | PowerShell | Assesses AD backup and forest-recovery controls BR-001..BR-010 (per-DC backup coverage from replication metadata, encryption/residency, immutable copies, runbook, restore testing, retention). |
| `Export-AdMonitoringAssessment.ps1` | PowerShell | Assesses monitoring and detection controls MON-001..MON-010 (SIEM forwarding, audit policy, privileged-group alerting, suspicious-auth detection, time sync, honeytokens, health monitoring). |
| `Find-ADDuplicateEmailProxyAddresses.ps1` | PowerShell | Finds duplicate mail-related values across on-prem Active Directory objects. |
| `Get-MailDnsRecords.ps1` | PowerShell | Checks MX, SPF, DMARC, and DKIM DNS records on Windows. |
| `get-mail-dns-records.sh` | Bash | Checks MX, SPF, DMARC, and DKIM DNS records on macOS/Linux using `dig`. |

## CSV Inputs at a Glance

Column names each script reads from its input CSV(s). A CSV produced by the referenced upstream script always qualifies; extra columns are ignored everywhere.

| Script | CSV parameter | Column names used |
| --- | --- | --- |
| `Export-ExchangeObjectData.ps1` | `-InputCsv` | One identity column, named by `-IdentityColumn` (default `Identity`), holding UPNs, primary SMTP addresses, aliases, or GUIDs. |
| `Export-ExchangeObjectInventory.ps1` | `-InputCsv` | Same as above. |
| `Invoke-ExchangeObjectDataExport.ps1` | `-InputCsv` | Same as above (passed through to the inventory job). |
| `Export-ExchangeGroupMemberships.ps1` | `-ObjectsCsv` | `Identity`, `PrimarySmtpAddress`, `ExternalDirectoryObjectId`, `RecipientTypeDetails`, `RecipientType` |
| `Export-ExchangeSendAsPermissions.ps1` | `-ObjectsCsv` | `Identity`, `PrimarySmtpAddress`, `RecipientTypeDetails`, `RecipientType` |
| `Export-ExchangeFullAccessPermissions.ps1` | `-ObjectsCsv` | `Identity`, `PrimarySmtpAddress`, `RecipientTypeDetails`, `RecipientType` |
| `Export-ExchangeObjectDataWorkbook.ps1` | `-InputFolder` | Consumes the export CSVs as written by the job scripts (`Objects.csv` mandatory; others optional). |
| `Export-ExchangeMigrationReadiness.ps1` | `-InputCsv` | One identity column, named by `-IdentityColumn` (default `Identity`). |
| `Export-ExchangeMigrationReadiness.ps1` | `-ObjectsCsv` | `Identity`, `DisplayName`, `RecipientType`; uses `PrimarySmtpAddress`, `ExternalDirectoryObjectId` when present. |
| `Export-DomainCutoverAssessment.ps1` | `-ObjectsCsv` | `Identity`, `DisplayName`, `RecipientTypeDetails`, `PrimarySmtpAddress` |
| `Export-DomainCutoverAssessment.ps1` | `-ProxyAddressesCsv` | `Identity`, `AddressType`, `ProxyAddress` |
| `Export-DomainCutoverAssessment.ps1` | `-MigrationReadinessCsv` | `Identity`, `LitigationHold`, `ComplianceHolds`, `RetentionPolicy`, `MigrationStatus`, `BlockingReasons`; uses `HasArchive`, `ArchiveState` when present. |
| `Export-DomainCutoverAssessment.ps1` | `-AcceptedDomainsCsv` | `DomainName`, `DomainType`, `Default`, `InitialDomain`, `MatchSubDomains`, `PendingRemoval`, `WhenCreated` |
| `Get-CrossTenantAdminRoleAssignments.ps1` | `-InputCsv` | `Prefix` plus the two UPN columns named by `-SourceUpnColumn` / `-TargetUpnColumn` (defaults `SourceUPN` / `TargetUPN`). |
| `Export-AdminCaPimPosture.ps1` | `-InputCsv` | `UserPrincipalName` |
| `Export-DomainComputerHardeningAssessment.ps1` | `-ComputersCsv` | `ComputerName` (falls back to `Name`); each value is a computer to assess remotely. |
| `Invoke-AdSecurityHardeningAssessment.ps1` | `-ComputersCsv` | Same as above (passed through to the domain-joined computer module). |

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
| `-InputCsv` | String | None | CSV file containing object identities to export. Needs one identity column (named by `-IdentityColumn`) with Exchange-resolvable values (UPN, primary SMTP, alias, or GUID); other columns are ignored. |
| `-Identity` | String | None | Single Exchange-resolvable identity to export. |
| `-IdentityColumn` | String | `Identity` | Name of the identity column in `-InputCsv`. |
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

- `Export-ExchangeObjectInventory.ps1` takes the same selection parameters as the monolith: `-All`, `-InputCsv` + `-IdentityColumn`, or `-Identity`, plus `-RecipientType` and `-OutputFolder`. The input CSV needs one identity column (named by `-IdentityColumn`, default `Identity`) with Exchange-resolvable values (UPN, primary SMTP, alias, or GUID); other columns are ignored.
- The membership/permission jobs take `-ObjectsCsv` (path to `Objects.csv`) and optional `-OutputFolder` (defaults to the folder containing the objects CSV). Required `Objects.csv` columns: `Identity`, `PrimarySmtpAddress`, `RecipientTypeDetails`, `RecipientType` — plus `ExternalDirectoryObjectId` for `Export-ExchangeGroupMemberships.ps1`. An `Objects.csv` produced by the inventory script or the monolith always qualifies.
- `Export-ExchangeObjectDataWorkbook.ps1` takes `-InputFolder`, optional `-OutputPath` (default `<InputFolder>/ExchangeObjectData.xlsx`), and `-Force` to overwrite an existing workbook. `Objects.csv` must exist in the folder; `ProxyAddresses.csv`, `FullAccess.csv`, `SendAs.csv`, `ExchangeGroupMemberships.csv`, `EntraGroupMemberships.csv`, and `Errors*.csv` are included when present.
- `Invoke-ExchangeObjectDataExport.ps1` takes the inventory selection parameters plus `-Force`; `-InputCsv` and `-IdentityColumn` are passed through to the inventory job unchanged.

### Connection Behavior

Job scripts reuse an existing Exchange Online or Microsoft Graph session when one exists and leave the session open, so a chained run authenticates once. Disconnect manually with `Disconnect-ExchangeOnline` / `Disconnect-MgGraph` when finished. (The monolith still manages and disconnects its own sessions.)

### Output Files

The jobs write the same CSV names and columns as the monolith (`Objects.csv`, `ProxyAddresses.csv`, `EntraGroupMemberships.csv`, `ExchangeGroupMemberships.csv`, `SendAs.csv`, `FullAccess.csv`), so the workbook script accepts a folder produced by either. Each job writes its own `Errors-<Job>.csv`; the workbook merges every `Errors*.csv` into one `Errors` sheet.

The workbook contains a `Summary` sheet (row counts per dataset), a `Consolidated` sheet (one row per object with joined proxies, trustees, memberships, and a `HasErrors` flag), and one sheet per non-empty dataset, each with a bold frozen header row and auto-filter.

### Notes

- All job scripts are read-only against Exchange and Entra.
- Recoverable per-object failures are logged to the job's errors CSV and the run continues, matching the monolith.
- The export CSVs and workbook may contain sensitive proxy addresses, mailbox permissions, group memberships, and object identifiers. Handle them as sensitive administrative data.

## Export-ExchangeMigrationReadiness.ps1

Assesses mailboxes for migration readiness and flags objects that are blocked, risky, or need review because of holds, retention, licensing, or mailbox type. The script is read-only.

### Migration Status Logic

| Status | Triggered By |
| --- | --- |
| `Blocked` | Litigation hold enabled, in-place/compliance holds present, or (with `-IncludeLicensing`) no license assigned to a mailbox type that requires one. |
| `Risky` | A retention policy is assigned. |
| `Review` | Uncommon mailbox type (`DiscoveryMailbox`, `LegacyMailbox`, `LinkedMailbox`, `LinkedRoomMailbox`) or an online archive is present (converting to a MailUser deletes the archive; `HasArchive`/`ArchiveState` columns identify these). |
| `Ready` | None of the above. |

### Requirements

- ExchangeOnlineManagement module.
- Microsoft.Graph.Authentication module only when `-IncludeLicensing` is used (delegated `User.Read.All`).

### Parameters

- Selection (pick one): `-All`, `-InputCsv` + `-IdentityColumn`, `-Identity`, or `-ObjectsCsv` (an `Objects.csv` inventory from `Export-ExchangeObjectInventory.ps1` or the monolith).
- `-InputCsv` needs one identity column (named by `-IdentityColumn`, default `Identity`) with Exchange-resolvable values (UPN, primary SMTP, alias, or GUID); other columns are ignored.
- `-ObjectsCsv` requires columns `Identity`, `DisplayName`, and `RecipientType`; `PrimarySmtpAddress` and `ExternalDirectoryObjectId` are used for mailbox and licensing lookups when present.
- `-OutputFolder` (optional): defaults to the `Objects.csv` folder when `-ObjectsCsv` is used, otherwise a timestamped `ExchangeMigrationReadiness_<timestamp>` folder.
- `-IncludeLicensing` (optional): checks license assignment via Microsoft Graph; unlicensed users are marked `Blocked` and assigned license SKU names (for example `SPE_E5`, `EXCHANGESTANDARD`) are written to the `Licenses` column.

### Usage

```powershell
# Assess every recipient in the tenant
./Export-ExchangeMigrationReadiness.ps1 -All

# Assess the objects from a previous modular export, including licensing
./Export-ExchangeMigrationReadiness.ps1 -ObjectsCsv ./export/Objects.csv -IncludeLicensing

# Assess a single mailbox
./Export-ExchangeMigrationReadiness.ps1 -Identity user@contoso.com
```

### Output Files

| File | Contents |
| --- | --- |
| `MigrationReadiness.csv` | One row per assessed object with hold, retention, licensing (including license SKU names), status, and blocking reasons. |
| `BlockedObjects.csv` | The subset with `MigrationStatus = Blocked`, including license SKU names. |
| `Errors-MigrationReadiness.csv` | Objects that could not be assessed. Headers-only means no errors. |

### Notes

- Mailbox lookups fall back through `PrimarySmtpAddress`, then `ExternalDirectoryObjectId`, then the raw `Identity`, so mixed-identifier CSVs resolve correctly.
- Non-mailbox recipients (mail contacts, mail users, distribution groups) cannot be assessed by `Get-EXOMailbox` and are logged to the errors CSV.
- The script reuses existing Exchange Online and Microsoft Graph sessions and leaves them open, matching the modular job scripts.

## Export-DomainCutoverAssessment.ps1

Assesses an accepted domain ahead of a tenant-to-tenant domain cutover. Confirms the accepted-domain configuration (live via `Get-AcceptedDomain` or from a previously exported CSV), resolves the domain's MX, SPF, Autodiscover, and DKIM records (via `Resolve-DnsName` on Windows or `dig` on macOS/Linux), and analyzes shared mailboxes from an `Objects.csv` inventory: which use the target domain as primary SMTP or alias, and which carry litigation or compliance holds. The script is read-only.

### Parameters

- `-Domain` (mandatory): the accepted domain to assess.
- `-ObjectsCsv` (optional): inventory from `Export-ExchangeObjectInventory.ps1`; `ProxyAddresses.csv` and `MigrationReadiness.csv` are auto-discovered from the same folder (or passed explicitly with `-ProxyAddressesCsv` / `-MigrationReadinessCsv`). Required columns: `Identity`, `DisplayName`, `RecipientTypeDetails`, `PrimarySmtpAddress`.
- `-ProxyAddressesCsv` (optional): required columns `Identity`, `AddressType`, `ProxyAddress`.
- `-MigrationReadinessCsv` (optional): columns used are `Identity`, `LitigationHold`, `ComplianceHolds`, `RetentionPolicy`, `MigrationStatus`, `BlockingReasons`, plus `HasArchive` and `ArchiveState` when present.
- `-AcceptedDomainsCsv` (optional): a previously exported `AcceptedDomains.csv` (from this script or a `Get-AcceptedDomain` export), enabling fully offline runs. Columns used: `DomainName`, `DomainType`, `Default`, `InitialDomain`, `MatchSubDomains`, `PendingRemoval`, `WhenCreated`.
- `-SkipExchange` (optional): skip the Exchange Online connection.
- `-OutputFolder` (optional): defaults to the `Objects.csv` folder or a timestamped folder.

### Usage

```powershell
# Live tenant assessment
./Export-DomainCutoverAssessment.ps1 -Domain contoso.com -ObjectsCsv ./export/Objects.csv

# Fully offline from previous exports
./Export-DomainCutoverAssessment.ps1 -Domain contoso.com -ObjectsCsv ./export/Objects.csv -AcceptedDomainsCsv ./AcceptedDomains.csv -SkipExchange
```

### Output Files

| File | Contents |
| --- | --- |
| `AcceptedDomains.csv` | All accepted domains with `IsTargetDomain` flag. |
| `DnsRecords.csv` | MX, SPF, Autodiscover, and DKIM records with per-record assessments. |
| `SharedMailboxDomainUsage.csv` | Per shared mailbox: primary/alias usage of the target domain joined with hold and readiness status. |
| `CutoverImpactSummary.csv` | Headline metrics (domain type, MX/EOP status, third-party SPF senders, hold counts). |
| `Errors-DomainCutover.csv` | Lookup failures. Headers-only means no errors. |

## Export-TenantDomainDependencyInventory.ps1

Inventories tenant resources that depend on a target SMTP domain. It scans distribution lists, mail-enabled security groups, dynamic distribution groups, Microsoft 365 groups, shared mailboxes, room/equipment mailboxes, mailbox permissions, moderation and sender restrictions, forwarding/routing properties, resource delegates, and address rewrite entries.

### Requirements

- PowerShell 7 recommended.
- ExchangeOnlineManagement module: `Install-Module ExchangeOnlineManagement`.
- Microsoft.Graph.Authentication module for Microsoft 365 group owner/member fallback when `Get-UnifiedGroupLinks` fails.
- Exchange Online read access for recipients, groups, mailboxes, permissions, calendar processing, and address rewrite entries.
- Microsoft Graph delegated scopes for fallback: `Group.Read.All`, `GroupMember.Read.All`, and `User.Read.All`.

### Parameters

| Parameter | Type | Description |
| --- | --- | --- |
| `-Domain` | String | Domain to inventory. Mandatory. |
| `-OutputFolder` | String | Destination folder. Defaults to `./TenantDomainDependencyInventory_<timestamp>`. |
| `-IncludeAll` | Switch | Include all scoped objects, not just objects with target-domain dependencies. |
| `-SkipMemberExpansion` | Switch | Skip group member expansion for very large tenants. |
| `-SkipAddressRewriteRules` | Switch | Skip `Get-AddressRewriteEntry`. |

### Usage

```powershell
./Export-TenantDomainDependencyInventory.ps1 -Domain contoso.com
```

```powershell
./Export-TenantDomainDependencyInventory.ps1 -Domain contoso.com -OutputFolder ./ContosoDependencyInventory -SkipMemberExpansion
```

### Output Files

| File | Contents |
| --- | --- |
| `DomainDependencyInventory.csv` | Impacted groups, shared/resource mailboxes, owners, aliases/proxies, member samples, sender restrictions, moderation, forwarding/routing, permissions, recommended disposition, and testing/comms needs. |
| `AddressRewriteRules.csv` | Address rewrite entries referencing the target domain when the cmdlet is available. |
| `Summary.csv` | Counts by action category. |
| `Errors-DomainDependencyInventory.csv` | Recoverable lookup failures. |

### Notes

- Default output includes only objects with a target-domain dependency; use `-IncludeAll` for a full scoped inventory.
- If `Get-UnifiedGroupLinks` fails with the Exchange Online `GetResponseHeader` error, the script falls back to Microsoft Graph for Microsoft 365 group owners/members.
- `RecommendedDisposition` flags `Migration`, `Reconfiguration`, `OwnerDecision`, or `ExclusionCandidate`.
- `TestingAndCommsNeeds` identifies downstream delivery, membership, moderation, booking, external sender, rewrite, and owner communication needs.
- The script is read-only and leaves the Exchange Online session open for follow-up checks.

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

## Get-EntraAdminMailboxInventory.ps1

Exports a focused CSV of Entra ID administrator users and mailbox status. It includes active directory role assignments, PIM-eligible directory role assignments, and users assigned through role-assignable groups. Service principals are excluded.

### Requirements

- PowerShell.
- Microsoft Graph PowerShell SDK: `Install-Module Microsoft.Graph`.
- ExchangeOnlineManagement module: `Install-Module ExchangeOnlineManagement`.
- Microsoft Graph permissions:
  - `Directory.Read.All`
  - `RoleManagement.Read.Directory`
  - `User.Read.All`
- Exchange Online read access to run `Get-EXOMailbox`.

### Parameters

| Parameter | Type | Description |
| --- | --- | --- |
| `-OutputPath` | String | CSV output path. Defaults to `./EntraAdminMailboxInventory_<timestamp>.csv`. |

### Usage

```powershell
./Get-EntraAdminMailboxInventory.ps1
```

```powershell
./Get-EntraAdminMailboxInventory.ps1 -OutputPath ./AdminMailboxInventory.csv
```

### Output

The CSV columns are:

- `DisplayName`
- `UserPrincipalName`
- `HasMailbox`
- `ExchangeGuid`
- `LicenseName`
- `MailboxType`

### Notes

- `LicenseName` contains semicolon-joined Microsoft Graph `SkuPartNumber` values.
- `HasMailbox=False` means `Get-EXOMailbox` did not find a mailbox for the admin UPN.
- PIM eligibility collection may warn or skip if the tenant lacks Entra ID P2 or the signed-in account lacks permission.

## Get-EntraDirSyncProtectionSettings.ps1

Checks the tenant-level Entra ID directory synchronization protection flags exposed by Microsoft Graph v1.0 at `/directory/onPremisesSynchronization`: `blockCloudObjectTakeoverThroughHardMatchEnabled` and `blockSoftMatchEnabled`. The script is read-only and does not use the MSOnline module.

### Requirements

- PowerShell.
- Microsoft Graph PowerShell SDK authentication module: `Install-Module Microsoft.Graph.Authentication`.
- Signed-in account must be Global Administrator for this Graph operation.
- Microsoft Graph permission:
  - `OnPremDirectorySynchronization.Read.All`

### Parameters

| Parameter | Type | Description |
| --- | --- | --- |
| `-OutputPath` | String | Optional CSV output path. If omitted, results are printed only. |

### Usage

```powershell
./Get-EntraDirSyncProtectionSettings.ps1
```

```powershell
./Get-EntraDirSyncProtectionSettings.ps1 -OutputPath ./DirSyncProtectionSettings.csv
```

### Notes

- `RecommendedState` is `Enabled` for both settings.
- `Status=Unknown` means Graph did not return the property in the `features` payload.

## Get-CrossTenantAdminRoleAssignments.ps1

Exports admin account details from two tenants (source and target) using a mapping CSV with `Prefix` plus a UPN column per tenant (`SourceUPN` and `TargetUPN` by default; override with `-SourceUpnColumn` / `-TargetUpnColumn`). The report includes Entra ID active directory roles, PIM-eligible directory roles when readable, immutable ID, actual user principal name, display name, assigned license SKU part numbers, lookup status, and errors. If a target UPN is missing or cannot be found, the script strips `(Admin)` from the matching source display name and tries exact target `displayName` lookups for `Admin <name>` first, then `<name>`.

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
| `-InputCsv` | String | Admin mapping CSV path. Defaults to `./Admin.csv`. Required columns: `Prefix` plus the two UPN columns named by `-SourceUpnColumn` / `-TargetUpnColumn`. |
| `-OutputPath` | String | Combined CSV output path. Defaults to `./CrossTenantAdminRoles_yyyyMMdd_HHmmss.csv`. |
| `-SourceUpnColumn` | String | CSV column holding source-tenant UPNs. Defaults to `SourceUPN`. |
| `-TargetUpnColumn` | String | CSV column holding target-tenant UPNs. Defaults to `TargetUPN`. |
| `-SourceLabel` | String | Friendly source-tenant name used in prompts and the `Environment` column. Defaults to `Source`. |
| `-TargetLabel` | String | Friendly target-tenant name. Defaults to `Target`. |
| `-SourceTenantId` | String | Optional tenant ID/domain passed to `Connect-MgGraph` for the source sign-in. |
| `-TargetTenantId` | String | Optional tenant ID/domain passed to `Connect-MgGraph` for the target sign-in. |

### Usage

```powershell
./Get-CrossTenantAdminRoleAssignments.ps1 -InputCsv ./Admin.csv
```

```powershell
./Get-CrossTenantAdminRoleAssignments.ps1 -InputCsv ./Admin.csv `
    -SourceUpnColumn ContosoUPN -TargetUpnColumn FabrikamUPN `
    -SourceLabel Contoso -TargetLabel Fabrikam `
    -SourceTenantId contoso.onmicrosoft.com -TargetTenantId fabrikam.onmicrosoft.com `
    -OutputPath ./Contoso-Fabrikam-AdminRoles.csv
```

### Notes

- The script connects to Microsoft Graph once for the source tenant and once for the target tenant; sign into the requested tenant at each prompt (or pass `-SourceTenantId` / `-TargetTenantId` to force the right tenant).
- Exchange Online RBAC roles are not included.
- If PIM eligibility cannot be read, active roles are still exported and the PIM warning is included in the `Error` column.
- Target rows resolved by display-name fallback use `LookupStatus=FoundByDisplayName`; ambiguous display-name matches are not guessed.
- Missing users are exported with `LookupStatus=NotFound`.

## Get-CrossTenantUserMapping.ps1

Exports all users from a source tenant, then connects to a target tenant and finds each source user's corresponding account by matching the source user principal name against the on-premises synced `extensionAttribute13` and `extensionAttribute14` values on target users. Matching is case-insensitive; an exact attribute value match is tried first, then a fallback scan for attribute values that contain the source UPN inside surrounding text. Multiple candidates are reported as `MatchStatus=Ambiguous` rather than guessed. The script is read-only.

### Requirements

- PowerShell.
- Microsoft Graph PowerShell SDK: `Install-Module Microsoft.Graph.Users` (and `Microsoft.Graph.Authentication`).
- Recommended Entra role: Global Reader or equivalent in each tenant.
- Microsoft Graph permission in both tenants:
  - `User.Read.All`

### Parameters

| Parameter | Type | Description |
| --- | --- | --- |
| `-SourceTenantId` | String | Optional source tenant ID or domain, forces `Connect-MgGraph` to the correct tenant. |
| `-TargetTenantId` | String | Optional target tenant ID or domain, forces `Connect-MgGraph` to the correct tenant. |
| `-IncludeGuests` | Switch | Include guest users on both sides. Defaults to `userType=Member` accounts only. |
| `-OutputPath` | String | CSV output path. Defaults to `./CrossTenantUserMapping_yyyyMMdd_HHmmss.csv`. |

### Usage

```powershell
./Get-CrossTenantUserMapping.ps1
```

```powershell
./Get-CrossTenantUserMapping.ps1 -SourceTenantId contoso.onmicrosoft.com -TargetTenantId fabrikam.onmicrosoft.com -OutputPath ./Contoso-Fabrikam-UserMapping.csv
```

### Notes

- The script connects to Microsoft Graph once for the source tenant and once for the target tenant; sign into the requested tenant at each prompt.
- Each source user row reports `MatchStatus` (`Matched`, `Ambiguous`, or `NotFound`), `MatchType` (`Exact` or `Contains`), the matched attribute(s), and the target account UPN, display name, object ID, enabled state, sync state, and both extension attribute values.
- A target user whose 13 and 14 attributes both hold the same UPN is collapsed into a single match listing both attributes.
- Ambiguous rows join all candidate target values with `; ` so they can be reviewed manually.

## Export-AdminCaPimPosture.ps1

Maps Conditional Access and PIM posture for in-scope admins. In-scope admins are discovered from active and PIM-eligible Entra ID directory role assignments (including assignments through role-assignable groups), or supplied explicitly via a CSV of UPNs. For each admin, every Conditional Access policy's user-scope conditions (included/excluded users, groups, and directory roles) are evaluated and reported as `Applies`, `Excluded`, `AppliesOnPimActivation`, or skipped when not in scope. The script is read-only.

The output is one Excel workbook with three worksheets:

- `Summary`: posture metrics such as admins with no enabled CA policy applying and admins without an MFA-requiring enabled CA policy.
- `AdminPosture`: one row per admin with active roles, PIM-eligible roles, applying/report-only/excluded CA policies, and an MFA coverage flag.
- `CaPolicyMap`: one row per admin per relevant CA policy with policy state, applicability status, grant controls, and included applications.

### Requirements

- PowerShell.
- Microsoft Graph PowerShell SDK: `Install-Module Microsoft.Graph`.
- ImportExcel module: `Install-Module ImportExcel -Scope CurrentUser` (Excel itself is not required).
- Recommended Entra role: Global Reader or Security Reader.
- Microsoft Graph permissions:
  - `Policy.Read.All`
  - `RoleManagement.Read.Directory`
  - `Directory.Read.All`

### Parameters

| Parameter | Type | Description |
| --- | --- | --- |
| `-InputCsv` | String | Optional CSV with a `UserPrincipalName` column defining the in-scope admins. When omitted, admins are discovered from role assignments. |
| `-OutputPath` | String | Excel workbook output path. Defaults to `./AdminCaPimPosture_yyyyMMdd_HHmmss.xlsx`. |

### Usage

```powershell
./Export-AdminCaPimPosture.ps1
```

```powershell
./Export-AdminCaPimPosture.ps1 -InputCsv ./InScopeAdmins.csv -OutputPath ./AdminCaPimPosture.xlsx
```

### Notes

- CA applicability is evaluated against user-scope conditions only; application, platform, location, client app, and risk conditions are reported for context but not evaluated.
- `AppliesOnPimActivation` marks role-scoped policies that only start applying once the admin activates a PIM-eligible role.
- Policies the admin is explicitly excluded from are listed separately so exclusion gaps are visible.
- If PIM eligibility cannot be read, active-role mapping still completes and a warning is shown.

## Appendix B Security Hardening Assessment

Seven read-only assessment scripts implement the Appendix B baseline hardening controls, plus an orchestrator that runs them and merges the results. Each script evaluates its controls and emits one row per control (per target where relevant) with a `Status`, `ControlId`, `Control`, `Target`, and `Evidence` column. Nothing is ever changed — every module only reads.

Every control resolves to one of five statuses:

| Status | Meaning |
| --- | --- |
| `Pass` | The control is met based on collected evidence. |
| `Fail` | The control is not met; the evidence explains why. |
| `Warning` | Partially met, or met via a weaker mechanism (e.g. Security Defaults instead of Conditional Access). |
| `Manual` | Not fully verifiable programmatically (SIEM onboarding, restore testing, ACL review); the evidence states exactly what to confirm. |
| `Error` | The fact could not be collected (host unreachable, insufficient rights). |

### Modules and control coverage

| Script | Controls | What it checks |
| --- | --- | --- |
| `Export-DcHardeningAssessment.ps1` | DC-001..DC-022 | Per-DC (via WinRM): OS support, patch currency, Secure Boot, BitLocker, firewall, anti-malware/EDR, unnecessary services, PowerShell restrictions, non-AD roles, LDAP signing/channel binding, SMB signing, NTLM/LM, Netlogon secure channel, hardened UNC paths, Print Spooler, audit policy + log collection, user rights. Domain-level: built-in Administrator, privileged group sizes, krbtgt age, per-DC system-state backups. |
| `Export-DomainComputerHardeningAssessment.ps1` | DJ-001..DJ-020 | Per-computer (sampled via WinRM): OS support/patching, BitLocker, Secure Boot + TPM 2.0, firewall, Defender/EDR, ASR rules, legacy protocols (SMBv1/LLMNR/NBT-NS), NTLM/SMB signing, local admin least privilege, UAC, Credential Guard + LSA protection, PowerShell logging, baseline GPOs, removable media/AutoRun, screen lock/banner, time hierarchy. Domain-level: Windows LAPS coverage, stale computer accounts, `ms-DS-MachineAccountQuota`. |
| `Export-EntraIdHardeningAssessment.ps1` | AAD-001..AAD-015 | Via Microsoft Graph: MFA for all users, phishing-resistant MFA for admins, legacy auth blocking, Conditional Access vs Security Defaults, per-user MFA retirement, PIM JIT for Global Admins, break-glass accounts + CA exclusions, GA minimization, guest invitation restrictions, user consent restrictions, log retention/SIEM, password protection, compliant-device for admin portals, Identity Protection risk policies, access reviews. |
| `Export-DnsHardeningAssessment.ps1` | DNS-001..DNS-010 | Per DNS server (via the DnsServer module): AD-integrated zones with secure dynamic updates, zone transfer restrictions, aging/scavenging, forwarders, recursion/root hints, DNSSEC, query/change logging, global query block list (WPAD/ISATAP), socket pool + cache locking, zone/record permissions. |
| `Export-AdcsHardeningAssessment.ps1` | ADCS-001..ADCS-010 | From the AD configuration partition + certutil: template enrollment least privilege, requester-supplied subject (ESC1), dangerous EKUs (ESC2/ESC3), template/CA ACLs (ESC4/ESC5), `EDITF_ATTRIBUTESUBJECTALTNAME2` (ESC6), CA admin permissions (ESC7), NTLM relay/web enrollment (ESC8), CA audit logging, key protection, CRL/OCSP. Emits a single not-applicable row when no CA is present. |
| `Export-AdBackupRecoveryAssessment.ps1` | BR-001..BR-010 | Per-DC backup coverage and recency from AD replication metadata (`dsaSignature`), plus procedural controls (encryption/residency, immutable copies, forest-recovery runbook, restore testing, retention, isolated recovery, alerting, access control, restore validation). Runbook and restore-test dates can be supplied to auto-evaluate their currency. |
| `Export-AdMonitoringAssessment.ps1` | MON-001..MON-010 | Per-DC (via WinRM): SIEM/WEF forwarding, advanced audit policy coverage, privileged-group change auditing, time synchronization, NTLM/legacy auth auditing. Domain-level: honeytoken accounts (by name pattern), and procedural detection-content controls (Kerberoasting/AS-REP/spray detection, retention, health monitoring, use-case review). |
| `Invoke-AdSecurityHardeningAssessment.ps1` | (orchestrator) | Runs the selected modules, writes each module's CSV into a per-run folder, and merges them into `AllControls.csv` with a `Module` column and a consolidated per-module and overall status summary. |

### Requirements

- PowerShell 7+ on a management host with the RSAT modules installed (`ActiveDirectory`, `DnsServer`).
- WinRM connectivity to the Domain Controllers / sampled computers for the modules that collect per-host facts (DC, DJ, MON).
- The Entra ID module requires the Microsoft Graph PowerShell SDK (`Install-Module Microsoft.Graph`) and a signed-in Graph session; it reads with `Policy.Read.All`, `Directory.Read.All`, `RoleManagement.Read.Directory`, `AccessReview.Read.All`, and `User.Read.All`.
- Recommended rights: a read-only/security-reader account on-premises and Global Reader or Security Reader in Entra ID.

### Key parameters

| Script | Parameter | Description |
| --- | --- | --- |
| `Export-DcHardeningAssessment.ps1` | `-Server`, `-KrbtgtMaxAgeDays`, `-BackupMaxAgeDays`, `-PatchAgeDays`, `-MaxDomainAdmins` | Target DCs (defaults to all) and the thresholds for krbtgt age, backup recency, patch age, and Domain Admins count. |
| `Export-DomainComputerHardeningAssessment.ps1` | `-ComputerName`, `-ComputersCsv`, `-SearchBase`, `-SampleSize`, `-BaselineGpoPattern`, `-LapsMinimumCoveragePercent` | Which computers to sample and how, plus the regex that identifies the security-baseline GPO (DJ-015) and the LAPS coverage threshold (DJ-003). |
| `Export-EntraIdHardeningAssessment.ps1` | `-BreakGlassUpn`, `-MaxGlobalAdmins` | Break-glass accounts to verify against CA exclusions (AAD-007) and the Global Admin count threshold (AAD-008). |
| `Export-DnsHardeningAssessment.ps1` | `-DnsServer`, `-MinimumSocketPoolSize` | Target DNS servers (defaults to all DCs) and the socket pool size floor (DNS-009). |
| `Export-AdcsHardeningAssessment.ps1` | `-Server`, `-SkipCaRegistryChecks` | AD server for directory queries; skip the certutil-based CA registry/web-enrollment probes when the CA hosts are unreachable. |
| `Export-AdBackupRecoveryAssessment.ps1` | `-BackupMaxAgeDays`, `-RunbookLastUpdated`, `-LastRestoreTest` | Backup recency threshold and optional runbook/restore-test dates that convert BR-004/BR-005 from Manual to an evaluated Pass/Warning. |
| `Export-AdMonitoringAssessment.ps1` | `-Server`, `-HoneytokenNamePattern` | Target DCs and the sAMAccountName wildcard used to verify honeytoken accounts (MON-007). |
| `Invoke-AdSecurityHardeningAssessment.ps1` | `-Module`, `-IncludeEntraId`, `-OutputFolder`, `-BreakGlassUpn`, `-ComputersCsv`, `-HoneytokenNamePattern` | Which modules to run (default: DC, DJ, DNS, ADCS, BR, MON), whether to add the Graph-dependent Entra ID module, and the per-module inputs to forward. |

Every module also accepts `-OutputPath` (a CSV) and, where it queries AD or uses remoting, `-Credential`.

### Usage

Run the whole on-premises suite and merge the results:

```powershell
./Invoke-AdSecurityHardeningAssessment.ps1 -OutputFolder ./HardeningRun
```

Include the Entra ID module (requires an existing `Connect-MgGraph` session) and verify break-glass accounts:

```powershell
Connect-MgGraph -Scopes Policy.Read.All,Directory.Read.All,RoleManagement.Read.Directory,AccessReview.Read.All,User.Read.All
./Invoke-AdSecurityHardeningAssessment.ps1 -IncludeEntraId -BreakGlassUpn bg1@contoso.com,bg2@contoso.com -HoneytokenNamePattern 'HNY-*'
```

Run a single module on its own:

```powershell
./Export-DcHardeningAssessment.ps1 -Server DC01,DC02 -OutputPath ./DcHardening.csv
./Export-EntraIdHardeningAssessment.ps1 -BreakGlassUpn bg1@contoso.com -OutputPath ./EntraHardening.csv
```

### Notes

- `Manual` is a first-class result, not a failure: the evidence field tells the assessor exactly what to confirm (SIEM onboarding, restore tests, ACL/DACL reviews, per-user MFA state). These controls cannot be proven from configuration alone.
- Patch-currency controls (DC-002, DJ-002) report the most recent installed hotfix date; cross-check missing updates against your patch management tooling (e.g. ManageEngine) for a definitive answer.
- The domain-joined computer module assesses a sample (default 25) — scope it with `-ComputersCsv`, `-SearchBase`, or `-SampleSize` for a full or targeted sweep.
- A module that throws is recorded as a single `Error` row by the orchestrator, so one failing module never aborts the run.

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

- `Get-EntraAdminAccounts.ps1` and `Get-EntraAdminMailboxInventory.ps1` read Microsoft Graph and Exchange Online data, then export CSV reports.
- `Find-ADDuplicateEmailProxyAddresses.ps1` is read-only against Active Directory and exports CSV reports.
- DNS scripts perform read-only DNS lookups.
- Review exported CSVs before sharing because they may contain user account identifiers, role assignments, group memberships, and mailbox information.
