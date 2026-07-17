# Script Reference

## Export-ExchangeObjectData.ps1

Exports Exchange Online recipient data, permission data, group memberships, and dashboard output. It builds a bulk recipient inventory (including proxy addresses) from Exchange Online, collects group memberships via Microsoft Graph `$batch` requests, gathers SendAs permissions in one org-wide sweep on `-All` runs, and looks up FullAccess permissions per mailbox by GUID.

### Common Uses

- Review mailbox proxy addresses and mailbox subtype.
- Audit FullAccess and SendAs permissions.
- Compare Exchange group membership and Entra ID group membership.
- Generate a searchable local dashboard for multi-object review.

### Requirements

- PowerShell.
- `ExchangeOnlineManagement` module.
- `Microsoft.Graph.Authentication` module unless `-SkipGraph` is used.
- Exchange Online read access for recipients and permissions.
- Graph delegated permissions such as `User.Read.All`, `Group.Read.All`, and `Directory.Read.All` for group membership collection.

Exchange Online RBAC varies by tenant, but the account must be allowed to run read cmdlets such as `Get-EXORecipient`, `Get-MailboxPermission`, and `Get-RecipientPermission`. Start with a read-only role group such as View-Only Organization Management where possible, and add only the additional recipient or permission-read roles required by your tenant policy.

### Parameters

| Parameter | Description |
| --- | --- |
| `-All` | Export all recipients matching `-RecipientType`. |
| `-InputCsv <path>` | Export recipients listed in a CSV. Needs one identity column (named by `-IdentityColumn`) with Exchange-resolvable values (UPN, primary SMTP, alias, or GUID); other columns are ignored. |
| `-Identity <value>` | Export one Exchange-resolvable recipient. |
| `-IdentityColumn <name>` | CSV column containing identities. Defaults to `Identity`. |
| `-RecipientType <type>` | `Mailbox`, `Group`, `MailUser`, `MailContact`, or `All`. Defaults to `Mailbox`. |
| `-OutputFolder <path>` | Destination folder for CSV and dashboard output. |
| `-SkipGraph` | Skip Microsoft Graph connection; `ExchangeGroupMemberships.csv` and `EntraGroupMemberships.csv` are emitted headers-only. |
| `-NoDashboard` | Suppress dashboard generation for multi-object runs. |

### Examples

```powershell
./Export-ExchangeObjectData.ps1 -All
```

```powershell
./Export-ExchangeObjectData.ps1 -All -RecipientType All -OutputFolder ./ExchangeExport
```

```powershell
./Export-ExchangeObjectData.ps1 -InputCsv ./objects.csv -IdentityColumn Identity -RecipientType Group
```

```powershell
./Export-ExchangeObjectData.ps1 -Identity user@contoso.com -SkipGraph
```

### Output Files

| File | Contents |
| --- | --- |
| `Objects-Consolidated.csv` | One row per recipient with semicolon-joined proxy, permission, and membership fields. |
| `Objects.csv` | Normalized recipient metadata including `MailboxType`. |
| `ProxyAddresses.csv` | One row per proxy address. |
| `FullAccess.csv` | One row per FullAccess permission entry. |
| `SendAs.csv` | One row per SendAs permission entry. |
| `ExchangeGroupMemberships.csv` | One row per Exchange group membership. |
| `EntraGroupMemberships.csv` | One row per Entra ID group membership. |
| `Errors.csv` | Lookup and export errors. Headers-only means no logged errors. |
| `Dashboard.html` | Static local dashboard for `-All` and `-InputCsv` runs unless suppressed. |

`ExchangeGroupMemberships.csv` `GroupIdentity` contains the group display name, with the group's mail address in parentheses when present (sourced from Microsoft Graph).

### Validation Notes

- Confirm `MailboxType` values for mailbox exports, especially shared, room, and equipment mailboxes.
- Check `Errors.csv` for skipped objects, permission lookup failures, and Graph lookup failures.
- For CSV input, confirm skipped recipient-type mismatches are logged with `RecipientTypeFilter`.
- If `-SkipGraph` is used, expect both `ExchangeGroupMemberships.csv` and `EntraGroupMemberships.csv` to be emitted headers-only.
- On `-All` runs, SendAs permissions are collected in one org-wide sweep joined to recipients by Exchange Identity; recipients sharing a display name can be misattributed, so verify SendAs row counts when duplicate display names exist in the tenant.

## Modular Exchange Export Jobs

Each stage of `Export-ExchangeObjectData.ps1` is also available as a standalone job script, plus a workbook collator and an orchestrator. The monolith is unchanged; both produce the same CSV names and columns.

### Common Uses

- Re-run a single stage (for example group memberships) without repeating the whole export.
- Split a large tenant export across sessions or time windows.
- Produce a single Excel workbook from an export folder for review and distribution.

### Requirements

- `Export-ExchangeObjectInventory.ps1`, `Export-ExchangeSendAsPermissions.ps1`, `Export-ExchangeFullAccessPermissions.ps1`: `ExchangeOnlineManagement`.
- `Export-ExchangeGroupMemberships.ps1`: `Microsoft.Graph.Authentication` only; no Exchange session.
- `Export-ExchangeObjectDataWorkbook.ps1`: `ImportExcel` (`Install-Module ImportExcel -Scope CurrentUser`); fully offline, Excel itself is not required.
- `Invoke-ExchangeObjectDataExport.ps1`: all of the above; the five job scripts must sit in the same folder as the orchestrator.

### Parameters

| Script | Parameters |
| --- | --- |
| `Export-ExchangeObjectInventory.ps1` | `-All` \| `-InputCsv` + `-IdentityColumn` \| `-Identity`, plus `-RecipientType`, `-OutputFolder`. The input CSV needs one identity column (named by `-IdentityColumn`, default `Identity`) with Exchange-resolvable values. |
| `Export-ExchangeGroupMemberships.ps1` | `-ObjectsCsv`, optional `-OutputFolder` (defaults to the objects CSV folder). Required `Objects.csv` columns: `Identity`, `PrimarySmtpAddress`, `ExternalDirectoryObjectId`, `RecipientTypeDetails`, `RecipientType`. |
| `Export-ExchangeSendAsPermissions.ps1` | `-ObjectsCsv`, optional `-OutputFolder`. Required `Objects.csv` columns: `Identity`, `PrimarySmtpAddress`, `RecipientTypeDetails`, `RecipientType`. |
| `Export-ExchangeFullAccessPermissions.ps1` | `-ObjectsCsv`, optional `-OutputFolder`. Required `Objects.csv` columns: `Identity`, `PrimarySmtpAddress`, `RecipientTypeDetails`, `RecipientType`. |
| `Export-ExchangeObjectDataWorkbook.ps1` | `-InputFolder`, optional `-OutputPath` (default `<InputFolder>/ExchangeObjectData.xlsx`), `-Force`. `Objects.csv` must exist in the folder; the other export CSVs and `Errors*.csv` are included when present. |
| `Invoke-ExchangeObjectDataExport.ps1` | Same selection parameters as the inventory script, plus `-Force` (passed to the workbook step). |

### Examples

```powershell
# Everything in one run
./Invoke-ExchangeObjectDataExport.ps1 -All -RecipientType Mailbox -OutputFolder ./export
```

```powershell
# Stage by stage
./Export-ExchangeObjectInventory.ps1 -All -OutputFolder ./export
./Export-ExchangeGroupMemberships.ps1 -ObjectsCsv ./export/Objects.csv
./Export-ExchangeSendAsPermissions.ps1 -ObjectsCsv ./export/Objects.csv
./Export-ExchangeFullAccessPermissions.ps1 -ObjectsCsv ./export/Objects.csv
./Export-ExchangeObjectDataWorkbook.ps1 -InputFolder ./export
```

### Output

The jobs write the same CSVs as the monolith (`Objects.csv`, `ProxyAddresses.csv`, `EntraGroupMemberships.csv`, `ExchangeGroupMemberships.csv`, `SendAs.csv`, `FullAccess.csv`), so the workbook script accepts a folder produced by either. Each job writes its own `Errors-<Job>.csv`.

The workbook contains a `Summary` sheet (row counts per dataset), a `Consolidated` sheet (one row per object with joined proxies, trustees, memberships, and a `HasErrors` flag), one sheet per non-empty dataset, and all `Errors*.csv` merged into one `Errors` sheet with the `Stage` column identifying the source job.

### Validation Notes

- Job scripts reuse an existing Exchange Online or Microsoft Graph session and leave it open, so a chained run authenticates once per service; disconnect manually when finished.
- The orchestrator aborts if the inventory fails; later job failures are warnings, the workbook is still built, and the exit code is 1 with the failed jobs named.
- The workbook script refuses to overwrite an existing `.xlsx` without `-Force` and rebuilds the file from scratch so stale sheets never survive.

## Export-ExchangeMigrationReadiness.ps1

Assesses mailboxes for migration readiness and classifies each as `Ready`, `Risky`, `Review`, or `Blocked` based on litigation hold, in-place/compliance holds, retention policy, mailbox type, and (optionally) license assignment. The script is read-only.

### Common Uses

- Identify mailboxes blocked from migration by litigation or compliance holds.
- Flag mailboxes with retention policies for pre-migration review.
- Surface uncommon mailbox types (discovery, legacy, linked) that need manual handling.
- Verify license assignment before migration with `-IncludeLicensing`.

### Requirements

- PowerShell.
- `ExchangeOnlineManagement` module.
- `Microsoft.Graph.Authentication` module only when `-IncludeLicensing` is used, with delegated `User.Read.All`.
- Exchange Online read access for recipients and mailboxes.

### Parameters

| Parameter | Description |
| --- | --- |
| `-All` | Assess all recipients in the tenant. |
| `-InputCsv <path>` | Assess identities listed in a CSV. Needs one identity column (named by `-IdentityColumn`) with Exchange-resolvable values; other columns are ignored. |
| `-IdentityColumn <name>` | Name of the identity column in `-InputCsv`. Defaults to `Identity`. |
| `-Identity <value>` | Assess one Exchange-resolvable recipient. |
| `-ObjectsCsv <path>` | Assess the objects in an `Objects.csv` inventory produced by `Export-ExchangeObjectInventory.ps1` or the monolith. Required columns: `Identity`, `DisplayName`, `RecipientType`; `PrimarySmtpAddress` and `ExternalDirectoryObjectId` are used when present. |
| `-OutputFolder <path>` | Destination folder. Defaults to the `Objects.csv` folder with `-ObjectsCsv`, otherwise a timestamped `ExchangeMigrationReadiness_<timestamp>` folder. |
| `-IncludeLicensing` | Check license assignment via Microsoft Graph; unlicensed users are marked `Blocked` and assigned license SKU names are written to the `Licenses` column. |

### Migration Status Logic

| Status | Triggered By |
| --- | --- |
| `Blocked` | Litigation hold enabled, in-place/compliance holds present, or (with `-IncludeLicensing`) no license assigned to a mailbox type that requires one. |
| `Risky` | A retention policy is assigned (and nothing blocks the mailbox). |
| `Review` | Mailbox type is `DiscoveryMailbox`, `LegacyMailbox`, `LinkedMailbox`, or `LinkedRoomMailbox`, or an online archive is present (`HasArchive`/`ArchiveState` columns) - converting a mailbox to a MailUser deletes its archive, so archives must be migrated or exported first. |
| `Ready` | None of the above. |

### Examples

```powershell
./Export-ExchangeMigrationReadiness.ps1 -All
```

```powershell
./Export-ExchangeMigrationReadiness.ps1 -ObjectsCsv ./export/Objects.csv -IncludeLicensing
```

```powershell
./Export-ExchangeMigrationReadiness.ps1 -InputCsv ./mailboxes.csv -IdentityColumn PrimarySmtpAddress
```

### Output Files

| File | Contents |
| --- | --- |
| `MigrationReadiness.csv` | One row per assessed object: identity, display name, recipient and mailbox type, hold and retention details, licensing, `Licenses` (semicolon-joined SKU names, for example `EMS; SPE_E5`), `HasArchive`/`ArchiveState`, `MigrationStatus`, and `BlockingReasons`. |
| `BlockedObjects.csv` | The subset with `MigrationStatus = Blocked`, including `Licenses`. |
| `Errors-MigrationReadiness.csv` | Objects that could not be assessed. Headers-only means no errors. |

### Validation Notes

- Mailbox lookups fall back through `PrimarySmtpAddress`, then `ExternalDirectoryObjectId`, then the raw `Identity`, so an inventory CSV with mixed identifier formats resolves correctly.
- Non-mailbox recipients (mail contacts, mail users, distribution groups) cannot be assessed by `Get-EXOMailbox` and are logged to `Errors-MigrationReadiness.csv`; this is expected when the input inventory includes non-mailbox objects.
- License checks require `User.Read.All`; a failed license lookup marks the user unlicensed and is logged to `Errors-MigrationReadiness.csv`, so verify `Blocked` objects whose only reason is `no license`.
- `Licenses` contains SKU part numbers as reported by Microsoft Graph (for example `SPE_E5`, `EXCHANGESTANDARD`), not marketing display names.
- The script reuses existing Exchange Online and Microsoft Graph sessions and leaves them open; disconnect manually when finished.

## Export-DomainCutoverAssessment.ps1

Assesses an accepted domain ahead of a tenant-to-tenant domain cutover: accepted-domain configuration, MX/SPF/Autodiscover/DKIM DNS records, and shared mailbox usage of the domain joined with hold status. The script is read-only.

### Common Uses

- Confirm whether a domain is authoritative and the tenant default before planning its removal.
- Identify third-party SPF senders and tenant-specific DKIM CNAMEs that a cutover must preserve or re-create.
- Enumerate which shared mailboxes use the domain as primary SMTP or alias.
- Surface litigation/compliance holds that pin mailboxes to the source tenant.

### Requirements

- PowerShell 7 on Windows, macOS, or Linux.
- `ExchangeOnlineManagement` module unless `-SkipExchange` or `-AcceptedDomainsCsv` is used.
- DNS lookups use `Resolve-DnsName` when available, otherwise `dig`.

### Parameters

| Parameter | Description |
| --- | --- |
| `-Domain <name>` | The accepted domain to assess. Mandatory. |
| `-ObjectsCsv <path>` | `Objects.csv` inventory for the shared mailbox analysis. Required columns: `Identity`, `DisplayName`, `RecipientTypeDetails`, `PrimarySmtpAddress`. |
| `-ProxyAddressesCsv <path>` | Proxy export; auto-discovered next to `Objects.csv` when omitted. Required columns: `Identity`, `AddressType`, `ProxyAddress`. |
| `-MigrationReadinessCsv <path>` | Readiness export for hold joins; auto-discovered next to `Objects.csv` when omitted. Columns used: `Identity`, `LitigationHold`, `ComplianceHolds`, `RetentionPolicy`, `MigrationStatus`, `BlockingReasons`, plus `HasArchive`/`ArchiveState` when present. |
| `-AcceptedDomainsCsv <path>` | A previously exported `AcceptedDomains.csv` (from this script or a `Get-AcceptedDomain` export), enabling offline runs. Columns used: `DomainName`, `DomainType`, `Default`, `InitialDomain`, `MatchSubDomains`, `PendingRemoval`, `WhenCreated`. |
| `-SkipExchange` | Do not connect to Exchange Online. |
| `-OutputFolder <path>` | Defaults to the `Objects.csv` folder, otherwise a timestamped folder. |

### Examples

```powershell
./Export-DomainCutoverAssessment.ps1 -Domain contoso.com -ObjectsCsv ./export/Objects.csv
```

```powershell
./Export-DomainCutoverAssessment.ps1 -Domain contoso.com -ObjectsCsv ./export/Objects.csv -AcceptedDomainsCsv ./AcceptedDomains.csv -SkipExchange
```

### Output Files

| File | Contents |
| --- | --- |
| `AcceptedDomains.csv` | All accepted domains with an `IsTargetDomain` flag. |
| `DnsRecords.csv` | MX, SPF, Autodiscover, and DKIM records, each with an assessment (EOP routing, third-party senders, tenant-bound DKIM). |
| `SharedMailboxDomainUsage.csv` | One row per shared mailbox: primary domain, target-domain aliases, other aliases, hold, archive (`HasArchive`/`ArchiveState`), and readiness status. |
| `CutoverImpactSummary.csv` | Metric/value pairs: domain type, default flag, MX/EOP status, SPF third parties, shared mailbox and hold counts. |
| `Errors-DomainCutover.csv` | Lookup failures. Headers-only means no errors. |

### Validation Notes

- The EOP MX hostname (`<domain>-<tld>.mail.protection.outlook.com`) is derived from the domain name and does not change when the domain moves tenants; routing follows domain verification.
- DKIM CNAME targets contain the tenant's onmicrosoft.com namespace and must be re-created in the target tenant after cutover.
- Hold columns are only populated when a `MigrationReadiness.csv` is available; re-run the readiness script first for current hold data.

## Export-TenantDomainDependencyInventory.ps1

Inventories distribution lists, mail-enabled security groups, shared resources, shared mailboxes, resource mailboxes, and address rewrite rules that depend on a target SMTP domain.

### Common Uses

- Identify groups and shared/resource mailboxes that use the target domain as primary SMTP or aliases.
- Capture owners, membership source, target-domain member samples, external sender access, moderation, forwarding, routing, and permissions.
- Flag items requiring migration, reconfiguration, owner decision, or exclusion.
- Generate downstream testing and communications needs for impacted groups and resources.

### Requirements

- PowerShell.
- `ExchangeOnlineManagement` module.
- `Microsoft.Graph.Authentication` module for Microsoft 365 group owner/member fallback when `Get-UnifiedGroupLinks` fails.
- Exchange Online read access for groups, recipients, mailboxes, mailbox permissions, recipient permissions, calendar processing, and address rewrite entries.
- Graph delegated scopes for fallback: `Group.Read.All`, `GroupMember.Read.All`, and `User.Read.All`.

### Parameters

| Parameter | Description |
| --- | --- |
| `-Domain <name>` | Target domain. Mandatory. |
| `-OutputFolder <path>` | Destination folder. Defaults to `./TenantDomainDependencyInventory_<timestamp>`. |
| `-IncludeAll` | Include all scoped objects, not just objects with target-domain dependencies. |
| `-SkipMemberExpansion` | Skip group member expansion for very large tenants. |
| `-SkipAddressRewriteRules` | Skip `Get-AddressRewriteEntry`. |

### Examples

```powershell
./Export-TenantDomainDependencyInventory.ps1 -Domain contoso.com
```

```powershell
./Export-TenantDomainDependencyInventory.ps1 -Domain contoso.com -OutputFolder ./ContosoDependencyInventory -IncludeAll
```

### Output Files

| File | Contents |
| --- | --- |
| `DomainDependencyInventory.csv` | Impacted groups and shared/resource mailboxes with owners, membership source, aliases/proxies, sender restrictions, moderation, forwarding/routing, permissions, disposition, and testing/comms needs. |
| `AddressRewriteRules.csv` | Address rewrite rules that reference the target domain when the cmdlet is available. |
| `Summary.csv` | Metric/value counts for migration, reconfiguration, owner decision, rewrite rows, and errors. |
| `Errors-DomainDependencyInventory.csv` | Recoverable lookup failures. |

### Validation Notes

- Use `-SkipMemberExpansion` first in very large tenants, then re-run without it for groups on the critical path.
- If `Get-UnifiedGroupLinks` fails with the Exchange Online `GetResponseHeader` error, the script falls back to Microsoft Graph for Microsoft 365 group owners/members.
- Treat `OwnerDecision` rows as business validation items, not technical blockers.
- Review `TestingAndCommsNeeds` to build a test script and owner communications plan.
- If `AddressRewriteRules.csv` is headers-only, check `Errors-DomainDependencyInventory.csv`; the cmdlet may not be available in the connected Exchange environment.

## Get-EntraAdminAccounts.ps1

Reports Entra ID user accounts with directory administrator access.

### Common Uses

- Review active directory role assignments.
- Include PIM-eligible assignments when permissions and licensing allow it.
- Expand role-assignable group membership to users.
- Export MFA registration status, license details, mailbox details, group memberships, and sign-in activity.

### Requirements

- PowerShell.
- Microsoft Graph PowerShell SDK.
- Recommended role: Global Reader or equivalent.
- Graph permissions: `Directory.Read.All`, `RoleManagement.Read.Directory`, `AuditLog.Read.All`, and `UserAuthenticationMethod.Read.All`.
- Optional `ExchangeOnlineManagement` module for `-IncludeMailbox`.

### Parameters

| Parameter | Description |
| --- | --- |
| `-IncludeMailbox` | Connect to Exchange Online and include mailbox type and primary SMTP address. |
| `-OutputPath <path>` | Destination CSV. Defaults to `./EntraAdminAccounts_<timestamp>.csv`. |

### Examples

```powershell
./Get-EntraAdminAccounts.ps1
```

```powershell
./Get-EntraAdminAccounts.ps1 -IncludeMailbox -OutputPath ./EntraAdmins.csv
```

### Output

Writes a CSV to `-OutputPath` and prints a console summary. The CSV includes admin identity, account status, role assignments, licenses, MFA registration details, mailbox information when requested, group membership, sign-in timestamps, and created date.

### Validation Notes

- Review warnings for tenants without Entra ID P2 or insufficient PIM permissions.
- Spot-check role membership against Entra admin center.
- Treat MFA fields as `Unknown` when registration lookup fails for a user.

## Get-EntraAdminMailboxInventory.ps1

Exports a focused mailbox/license inventory for Entra ID administrator users.

### Common Uses

- Find all users with active or PIM-eligible Entra ID directory roles.
- Include users assigned to roles through role-assignable groups.
- Check whether each admin has an Exchange Online mailbox.
- Export Exchange GUID, mailbox type, and assigned license SKU names for migration planning.

### Requirements

- PowerShell.
- Microsoft Graph PowerShell SDK: `Install-Module Microsoft.Graph`.
- `ExchangeOnlineManagement` module.
- Graph permissions: `Directory.Read.All`, `RoleManagement.Read.Directory`, and `User.Read.All`.
- Exchange Online read access to run `Get-EXOMailbox`.

### Parameters

| Parameter | Description |
| --- | --- |
| `-OutputPath <path>` | Destination CSV. Defaults to `./EntraAdminMailboxInventory_<timestamp>.csv`. |

### Examples

```powershell
./Get-EntraAdminMailboxInventory.ps1
```

```powershell
./Get-EntraAdminMailboxInventory.ps1 -OutputPath ./AdminMailboxInventory.csv
```

### Output

Writes a CSV with `DisplayName`, `UserPrincipalName`, `HasMailbox`, `ExchangeGuid`, `LicenseName`, and `MailboxType`.

### Validation Notes

- Review warnings for tenants without Entra ID P2 or insufficient PIM permissions.
- `LicenseName` contains Microsoft Graph SKU part numbers such as `SPE_E5` and `EXCHANGESTANDARD`.
- `HasMailbox=False` means `Get-EXOMailbox` did not find a mailbox for the admin user principal name.
- Spot-check a small sample against Entra admin center role assignments and Exchange admin center mailbox records.

## Get-EntraDirSyncProtectionSettings.ps1

Checks tenant-level Entra ID directory synchronization protection flags with Microsoft Graph v1.0.

### Common Uses

- Check whether hard-match cloud object takeover protection is enabled.
- Check whether soft match is blocked tenant-wide.
- Export a small CSV for tenant security evidence or migration readiness checks.

### Requirements

- PowerShell.
- `Microsoft.Graph.Authentication` module.
- Signed-in account must be Global Administrator for this Graph operation.
- Graph permission: `OnPremDirectorySynchronization.Read.All`.

### Parameters

| Parameter | Description |
| --- | --- |
| `-OutputPath <path>` | Optional CSV output path. If omitted, results are printed only. |

### Examples

```powershell
./Get-EntraDirSyncProtectionSettings.ps1
```

```powershell
./Get-EntraDirSyncProtectionSettings.ps1 -OutputPath ./DirSyncProtectionSettings.csv
```

### Output

Prints two rows and optionally exports them to CSV. Rows cover `BlockCloudObjectTakeover` mapped to Graph property `blockCloudObjectTakeoverThroughHardMatchEnabled`, and `BlockSoftMatch` mapped to `blockSoftMatchEnabled`.

### Validation Notes

- `RecommendedState` is `Enabled` for both flags.
- `Status=Unknown` means Graph did not return the property in `features`.
- The script uses `Invoke-MgGraphRequest` against `/directory/onPremisesSynchronization`; it does not use MSOnline cmdlets.

## Get-CrossTenantAdminRoleAssignments.ps1

Uses a source/target admin mapping CSV to export matching admin account role and license details from both tenants.

### Common Uses

- Compare source- and target-tenant admin accounts by shared `Prefix`.
- Review Entra ID active directory role assignments.
- Review PIM-eligible directory role assignments when permissions and licensing allow it.
- Export immutable IDs, actual UPNs, display names, and assigned license SKU part numbers.
- Resolve target accounts by source display name when the target UPN is blank or stale, preferring `Admin <name>` over the standard user display name.

### Requirements

- PowerShell.
- Microsoft Graph PowerShell SDK.
- Recommended role: Global Reader or equivalent in both tenants.
- Graph permissions: `Directory.Read.All` and `RoleManagement.Read.Directory`.

### Parameters

| Parameter | Description |
| --- | --- |
| `-InputCsv <path>` | Admin mapping CSV with `Prefix` plus the two UPN columns. Defaults to `./Admin.csv`. |
| `-OutputPath <path>` | Destination CSV. Defaults to `./CrossTenantAdminRoles_yyyyMMdd_HHmmss.csv`. |
| `-SourceUpnColumn <name>` | CSV column holding source-tenant UPNs. Defaults to `SourceUPN`. |
| `-TargetUpnColumn <name>` | CSV column holding target-tenant UPNs. Defaults to `TargetUPN`. |
| `-SourceLabel <name>` | Friendly source-tenant name used in prompts and the `Environment` column. Defaults to `Source`. |
| `-TargetLabel <name>` | Friendly target-tenant name. Defaults to `Target`. |
| `-SourceTenantId <id>` | Optional tenant ID/domain passed to `Connect-MgGraph` for the source sign-in. |
| `-TargetTenantId <id>` | Optional tenant ID/domain passed to `Connect-MgGraph` for the target sign-in. |

### Examples

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

### Output

Writes one combined CSV with one row per mapped account per environment. Columns include environment, prefix, input UPN, actual UPN, display name, immutable ID, active roles, PIM-eligible roles, all roles, assigned license SKU part numbers, lookup status, and error text. Target-tenant rows can use `LookupStatus=FoundByDisplayName` when the script strips `(Admin)` from the source display name and finds exactly one target user with `Admin <name>` or, if that is absent, `<name>`.

### Validation Notes

- Spot-check role assignments against the Entra admin center.
- Check `Error` for PIM eligibility warnings or license lookup failures.
- Check `Error` for display-name fallback details, especially ambiguous matches.
- Confirm expected missing users appear with `LookupStatus=NotFound`.

## Get-CrossTenantUserMapping.ps1

Exports all users from a source tenant and matches each one to a target tenant account via synced on-premises `extensionAttribute13`/`extensionAttribute14` values containing the source UPN.

### Common Uses

- Build a source-to-target account mapping ahead of a tenant-to-tenant migration.
- Verify that cross-tenant provisioning stamped the expected source UPN into extension attributes 13 or 14.
- Find source users with no corresponding target account (`MatchStatus=NotFound`).
- Surface duplicate stampings where multiple target accounts claim the same source UPN (`MatchStatus=Ambiguous`).

### Requirements

- PowerShell.
- Microsoft Graph PowerShell SDK (`Microsoft.Graph.Users`, `Microsoft.Graph.Authentication`).
- Recommended role: Global Reader or equivalent in both tenants.
- Graph permission in both tenants: `User.Read.All`.

### Parameters

| Parameter | Description |
| --- | --- |
| `-SourceTenantId <id>` | Optional source tenant ID or domain passed to `Connect-MgGraph`. |
| `-TargetTenantId <id>` | Optional target tenant ID or domain passed to `Connect-MgGraph`. |
| `-IncludeGuests` | Include guest users on both sides; defaults to members only. |
| `-OutputPath <path>` | Destination CSV. Defaults to `./CrossTenantUserMapping_yyyyMMdd_HHmmss.csv`. |

### Examples

```powershell
./Get-CrossTenantUserMapping.ps1
```

```powershell
./Get-CrossTenantUserMapping.ps1 -SourceTenantId contoso.onmicrosoft.com -TargetTenantId fabrikam.onmicrosoft.com -OutputPath ./Contoso-Fabrikam-UserMapping.csv
```

### Output

Writes one CSV row per source user: source UPN, display name, object ID, and enabled state; `MatchStatus` (`Matched`, `Ambiguous`, `NotFound`); `MatchType` (`Exact` for a full attribute value match, `Contains` when the attribute embeds the UPN in surrounding text); the matched attribute name(s); and the target account UPN, display name, object ID, enabled state, on-premises sync state, and both extension attribute values. Ambiguous rows join all candidates with `; `.

### Validation Notes

- Spot-check a few `Matched` rows against the target tenant's user properties in the Entra admin center.
- Review every `Ambiguous` row manually before using the mapping for migration.
- Confirm `NotFound` rows are expected (not-yet-provisioned or intentionally excluded accounts).
- Matching is case-insensitive and trims whitespace; a stale attribute value still matches, so verify recently renamed source UPNs separately.

## Export-AdminCaPimPosture.ps1

Maps Conditional Access policy applicability and PIM-eligible directory roles for in-scope admins into one Excel workbook.

### Common Uses

- Assess CA / PIM posture for admins ahead of a migration or security review.
- Find admins with no enabled Conditional Access policy applying to them.
- Find admins not covered by an MFA-requiring enabled CA policy.
- Surface admins explicitly excluded from CA policies.
- List role-scoped CA policies that only apply once a PIM-eligible role is activated.

### Requirements

- PowerShell.
- Microsoft Graph PowerShell SDK and ImportExcel (`Install-Module ImportExcel -Scope CurrentUser`).
- Recommended role: Global Reader or Security Reader.
- Graph permissions: `Policy.Read.All`, `RoleManagement.Read.Directory`, `Directory.Read.All`.

### Parameters

| Parameter | Description |
| --- | --- |
| `-InputCsv <path>` | Optional CSV with a `UserPrincipalName` column defining the in-scope admins. Omit to discover admins from active and PIM-eligible role assignments. |
| `-OutputPath <path>` | Destination workbook. Defaults to `./AdminCaPimPosture_yyyyMMdd_HHmmss.xlsx`. |

### Examples

```powershell
./Export-AdminCaPimPosture.ps1
```

```powershell
./Export-AdminCaPimPosture.ps1 -InputCsv ./InScopeAdmins.csv -OutputPath ./AdminCaPimPosture.xlsx
```

### Output

Writes one Excel workbook with three worksheets. `Summary` holds posture metrics (in-scope admins, CA policy count, PIM coverage, CA gaps, exclusions, lookup errors). `AdminPosture` has one row per admin: active directory roles, PIM-eligible roles (with group attribution and administrative unit scope), enabled and report-only CA policies applying, activation-only policies, exclusions, and whether any enabled applying policy requires MFA. `CaPolicyMap` has one row per admin per relevant CA policy: policy name, ID, state, applicability status (`Applies`, `Excluded`, `AppliesOnPimActivation`), MFA requirement, grant controls, and included applications.

### Validation Notes

- CA applicability is evaluated on user-scope conditions only (users, groups, directory roles); confirm application/platform/location conditions manually for policies that matter.
- Spot-check a few admins against the Entra admin center's Conditional Access "What If" tool.
- Review the `Summary` gap metrics first: admins with no enabled CA policy and admins without MFA-requiring CA are the primary posture findings.
- If the PIM warning appears, eligible-role mapping is incomplete; re-run with an account that can read role eligibility schedules.

## Find-ADDuplicateEmailProxyAddresses.ps1

Finds duplicate mail-related values across on-prem Active Directory objects.

### Common Uses

- Identify duplicate SMTP proxy addresses before migration.
- Check `mail` and `proxyAddresses` collisions.
- Include `targetAddress`, `userPrincipalName`, or other schema attributes when needed.

### Requirements

- Windows PowerShell or PowerShell on Windows with RSAT.
- Active Directory PowerShell module.
- Read access to the selected search base and schema.

### Parameters

| Parameter | Description |
| --- | --- |
| `-SearchBase <dn>` | Distinguished name where the scan starts. |
| `-Server <name>` | Domain controller or AD server to query. |
| `-OutputCsv <path>` | Pair-based duplicate output. |
| `-OccurrenceCsv <path>` | Occurrence-level duplicate output. |
| `-AdditionalAttributes <names>` | Extra AD attributes to scan after schema validation. |
| `-IncludeTargetAddress` | Include `targetAddress`. |
| `-IncludeUserPrincipalName` | Include `userPrincipalName`. |
| `-SmtpOnly` | Limit `proxyAddresses` checks to SMTP values. |
| `-NoOccurrenceCsv` | Suppress occurrence-level output. |

### Examples

```powershell
./Find-ADDuplicateEmailProxyAddresses.ps1
```

```powershell
./Find-ADDuplicateEmailProxyAddresses.ps1 -SearchBase "DC=contoso,DC=com" -Server "dc01.contoso.com"
```

```powershell
./Find-ADDuplicateEmailProxyAddresses.ps1 -IncludeTargetAddress -IncludeUserPrincipalName -SmtpOnly
```

### Output

Writes pair-based duplicate rows to `-OutputCsv`. Unless `-NoOccurrenceCsv` is used, also writes occurrence-level rows to `-OccurrenceCsv`. Pair output is optimized for remediation review; occurrence output is useful when a value appears on more than two objects.

### Validation Notes

- Verify duplicate pairs against AD before remediation.
- Review occurrence output when a value appears on more than two objects.
- Non-SMTP proxy prefixes are preserved to avoid false matches across address types.

## Get-MailDnsRecords.ps1

Checks MX, SPF, DMARC, and DKIM records using Windows DNS cmdlets.

### Requirements

- Windows with `Resolve-DnsName` available.
- Network access to the configured DNS resolver.

### Parameters

| Parameter | Description |
| --- | --- |
| `-Domain <domain>` | Domain or URL to inspect. |
| `-DkimSelectors <selectors>` | DKIM selectors to query. |
| `-Server <resolver>` | DNS resolver, such as `8.8.8.8` or `1.1.1.1`. |
| `-Json` | Emit structured JSON instead of formatted console output. |

### Examples

```powershell
./Get-MailDnsRecords.ps1 -Domain contoso.com
```

```powershell
./Get-MailDnsRecords.ps1 -Domain contoso.com -DkimSelectors selector1,selector2,google -Json
```

### Output

Prints grouped console sections for MX, SPF, DMARC, DKIM, and warnings. With `-Json`, emits a structured object containing the queried domain, timestamp, resolver, record collections, and warnings.

## get-mail-dns-records.sh

Checks MX, SPF, DMARC, and DKIM records on macOS/Linux using `dig`.

### Requirements

- Bash.
- `dig`.
- Network access to the configured DNS resolver.

### Options

| Option | Description |
| --- | --- |
| `<domain>` | Domain or URL to inspect. |
| `-d`, `--domain` | Domain or URL to inspect. |
| `-s`, `--server` | DNS resolver, such as `8.8.8.8` or `1.1.1.1`. |
| `-k`, `--dkim-selectors` | Comma-separated DKIM selectors. |
| `-h`, `--help` | Show help. |

### Examples

```bash
./get-mail-dns-records.sh contoso.com
```

```bash
./get-mail-dns-records.sh -d contoso.com -s 8.8.8.8 -k selector1,selector2
```

### Output

Prints grouped console sections for MX, SPF, DMARC, DKIM, and warnings. The Bash version does not write files by default; redirect output if you need to save it.

## DNS Validation Notes

- Multiple SPF records at the domain root are usually invalid and should be reviewed.
- Multiple DMARC records at `_dmarc.<domain>` should be reviewed.
- DKIM selectors are not reliably discoverable from DNS alone; pass known selectors when available.
- `TXT found; verify manually` means TXT records exist but do not clearly match a DKIM public key pattern.
