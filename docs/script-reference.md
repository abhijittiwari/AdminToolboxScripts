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
| `-InputCsv <path>` | Export recipients listed in a CSV. |
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
| `Export-ExchangeObjectInventory.ps1` | `-All` \| `-InputCsv` + `-IdentityColumn` \| `-Identity`, plus `-RecipientType`, `-OutputFolder`. |
| `Export-ExchangeGroupMemberships.ps1` | `-ObjectsCsv`, optional `-OutputFolder` (defaults to the objects CSV folder). |
| `Export-ExchangeSendAsPermissions.ps1` | `-ObjectsCsv`, optional `-OutputFolder`. |
| `Export-ExchangeFullAccessPermissions.ps1` | `-ObjectsCsv`, optional `-OutputFolder`. |
| `Export-ExchangeObjectDataWorkbook.ps1` | `-InputFolder`, optional `-OutputPath` (default `<InputFolder>/ExchangeObjectData.xlsx`), `-Force`. |
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
| `-InputCsv <path>` | Assess identities listed in a CSV. |
| `-IdentityColumn <name>` | CSV column containing identities. Defaults to `Identity`. |
| `-Identity <value>` | Assess one Exchange-resolvable recipient. |
| `-ObjectsCsv <path>` | Assess the objects in an `Objects.csv` inventory produced by `Export-ExchangeObjectInventory.ps1` or the monolith. |
| `-OutputFolder <path>` | Destination folder. Defaults to the `Objects.csv` folder with `-ObjectsCsv`, otherwise a timestamped `ExchangeMigrationReadiness_<timestamp>` folder. |
| `-IncludeLicensing` | Check license assignment via Microsoft Graph; unlicensed users are marked `Blocked` and assigned license SKU names are written to the `Licenses` column. |

### Migration Status Logic

| Status | Triggered By |
| --- | --- |
| `Blocked` | Litigation hold enabled, in-place/compliance holds present, or (with `-IncludeLicensing`) no license assigned. |
| `Risky` | A retention policy is assigned (and nothing blocks the mailbox). |
| `Review` | Mailbox type is `DiscoveryMailbox`, `LegacyMailbox`, `LinkedMailbox`, or `LinkedRoomMailbox`. |
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
| `MigrationReadiness.csv` | One row per assessed object: identity, display name, recipient and mailbox type, hold and retention details, licensing, `Licenses` (semicolon-joined SKU names, for example `EMS; SPE_E5`), `MigrationStatus`, and `BlockingReasons`. |
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
| `-ObjectsCsv <path>` | `Objects.csv` inventory for the shared mailbox analysis. |
| `-ProxyAddressesCsv <path>` | Proxy export; auto-discovered next to `Objects.csv` when omitted. |
| `-MigrationReadinessCsv <path>` | Readiness export for hold joins; auto-discovered next to `Objects.csv` when omitted. |
| `-AcceptedDomainsCsv <path>` | A previous `Get-AcceptedDomain` export, enabling offline runs. |
| `-SkipExchange` | Do not connect to Exchange Online. |
| `-OutputFolder <path>` | Defaults to the `Objects.csv` folder, otherwise a timestamped folder. |

### Examples

```powershell
./Export-DomainCutoverAssessment.ps1 -Domain wae.com -ObjectsCsv ./export/Objects.csv
```

```powershell
./Export-DomainCutoverAssessment.ps1 -Domain wae.com -ObjectsCsv ./export/Objects.csv -AcceptedDomainsCsv ./AcceptedDomains.csv -SkipExchange
```

### Output Files

| File | Contents |
| --- | --- |
| `AcceptedDomains.csv` | All accepted domains with an `IsTargetDomain` flag. |
| `DnsRecords.csv` | MX, SPF, Autodiscover, and DKIM records, each with an assessment (EOP routing, third-party senders, tenant-bound DKIM). |
| `SharedMailboxDomainUsage.csv` | One row per shared mailbox: primary domain, target-domain aliases, other aliases, hold and readiness status. |
| `CutoverImpactSummary.csv` | Metric/value pairs: domain type, default flag, MX/EOP status, SPF third parties, shared mailbox and hold counts. |
| `Errors-DomainCutover.csv` | Lookup failures. Headers-only means no errors. |

### Validation Notes

- The EOP MX hostname (`<domain>-<tld>.mail.protection.outlook.com`) is derived from the domain name and does not change when the domain moves tenants; routing follows domain verification.
- DKIM CNAME targets contain the tenant's onmicrosoft.com namespace and must be re-created in the target tenant after cutover.
- Hold columns are only populated when a `MigrationReadiness.csv` is available; re-run the readiness script first for current hold data.

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

## Get-CrossTenantAdminRoleAssignments.ps1

Uses a WAE/Fortescue admin mapping CSV to export matching admin account role and license details from both tenants.

### Common Uses

- Compare WAE and Fortescue admin accounts by shared `Prefix`.
- Review Entra ID active directory role assignments.
- Review PIM-eligible directory role assignments when permissions and licensing allow it.
- Export immutable IDs, actual UPNs, display names, and assigned license SKU part numbers.
- Resolve Fortescue accounts by WAE display name when the Fortescue UPN is blank or stale, preferring `Admin <name>` over the standard user display name.

### Requirements

- PowerShell.
- Microsoft Graph PowerShell SDK.
- Recommended role: Global Reader or equivalent in both tenants.
- Graph permissions: `Directory.Read.All` and `RoleManagement.Read.Directory`.

### Parameters

| Parameter | Description |
| --- | --- |
| `-InputCsv <path>` | Admin mapping CSV with `WAEUPN`, `Prefix`, and `FortescueUPN`. Defaults to `./Admin.csv`. |
| `-OutputPath <path>` | Destination CSV. Defaults to `./CrossTenantAdminRoles_yyyyMMdd_HHmmss.csv`. |

### Examples

```powershell
./Get-CrossTenantAdminRoleAssignments.ps1 -InputCsv ./Admin.csv
```

```powershell
./Get-CrossTenantAdminRoleAssignments.ps1 -InputCsv ./Admin.csv -OutputPath ./WAE-Fortescue-AdminRoles.csv
```

### Output

Writes one combined CSV with one row per mapped account per environment. Columns include environment, prefix, input UPN, actual UPN, display name, immutable ID, active roles, PIM-eligible roles, all roles, assigned license SKU part numbers, lookup status, and error text. Fortescue rows can use `LookupStatus=FoundByDisplayName` when the script strips `(Admin)` from the WAE display name and finds exactly one Fortescue user with `Admin <name>` or, if that is absent, `<name>`.

### Validation Notes

- Spot-check role assignments against the Entra admin center.
- Check `Error` for PIM eligibility warnings or license lookup failures.
- Check `Error` for display-name fallback details, especially ambiguous matches.
- Confirm expected missing users appear with `LookupStatus=NotFound`.

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
