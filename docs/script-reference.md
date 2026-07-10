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
