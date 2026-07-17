# Operations and Security

## Read-Only Intent

The scripts in this repository are designed for reporting and troubleshooting. They do not intentionally modify Microsoft 365, Entra ID, Exchange Online, Active Directory, or DNS records.

Before running scripts in production, review the source and validate the commands against your organization's change-management requirements.

## Permissions

Use the least-privileged administrative account that can read the required data.

| Area | Typical Access Needed |
| --- | --- |
| Exchange Online recipient exports | Read access to recipients, recipient permissions, and mailbox permissions. |
| Entra admin reporting | Graph permissions for directory roles, audit logs, users, authentication method registration, and groups. |
| Entra group membership enrichment | Graph directory, user, and group read permissions. |
| Active Directory duplicate scans | Read access to the selected AD search base and schema. |
| Security hardening assessments | AD read access (domain, configuration partition, group membership); WinRM to Domain Controllers and sampled computers for registry/service/audit-policy reads; Graph read scopes (`Policy.Read.All`, `Directory.Read.All`, `RoleManagement.Read.Directory`, `AccessReview.Read.All`, `User.Read.All`) for the Entra ID module. |
| Divestiture discovery inventories | AD/GPO/DNS/DHCP read access via RSAT; WinRM to DCs for time-source reads; local read on the Entra Connect and Exchange servers for those modules; Graph (`Domain.Read.All`, `Organization.Read.All`, `User.Read.All`), Exchange Online, Teams, SPO, and Security & Compliance read access for the cloud modules. |
| DNS record checks | Network access to the selected resolver. |

Some Microsoft Graph permissions may require admin consent depending on tenant policy.

For Exchange Online exports, confirm the account can run the read cmdlets used by the script before a broad export:

```powershell
Get-EXORecipient -ResultSize 1
Get-Recipient -Identity user@contoso.com
Get-MailboxPermission -Identity user@contoso.com
Get-RecipientPermission -Identity user@contoso.com
```

Use a test mailbox for permission checks. If these commands fail, review Exchange Online RBAC role group membership. A read-only role group such as View-Only Organization Management may be enough for recipient discovery, but permission cmdlets can require additional tenant-specific role assignments.

## Output Handling

Treat generated files as sensitive administrative data. Outputs may contain:

- User principal names and email addresses.
- Immutable IDs and object IDs.
- Admin role assignments.
- Group memberships.
- Mailbox permissions.
- Proxy addresses.
- Sign-in metadata.
- DNS security posture information.
- Security control pass/fail status and configuration weaknesses (hardening assessments).
- Forest topology, trust configuration, privileged identities, SIDHistory, and license assignments (discovery inventories).

Hardening assessment output is especially sensitive: a consolidated `AllControls.csv` describes exactly which controls fail on which hosts. Store it with the same care as a penetration-test report.

Recommended handling:

- Store exports in a restricted folder.
- Remove or redact output before sharing outside the admin team.
- Delete temporary exports when no longer needed.
- Avoid committing generated CSV or HTML output to source control.

## Running Safely

1. Run in a test tenant or against a small input CSV first when possible.
2. Use explicit `-OutputFolder` or `-OutputPath` values so outputs are easy to locate.
3. Review console warnings before acting on the data.
4. Validate a small sample against the source system.
5. Keep generated files out of public locations.

## Exchange Object Export Validation

For `Export-ExchangeObjectData.ps1`:

- Run a single-object export first:

```powershell
./Export-ExchangeObjectData.ps1 -Identity user@contoso.com -OutputFolder ./ValidationExport
```

- Confirm `Objects.csv` includes the expected recipient and `MailboxType`.
- Confirm `ProxyAddresses.csv` matches Exchange recipient proxy addresses.
- Confirm `FullAccess.csv` and `SendAs.csv` exclude default/self entries.
- Confirm `Errors.csv` is either headers-only or contains expected, explainable warnings.
- For multi-object runs, confirm `Dashboard.html` is present unless `-NoDashboard` is used.
- For single `-Identity` runs, confirm no dashboard is expected.

## Entra Admin Report Validation

For `Get-EntraAdminAccounts.ps1`:

- Spot-check admin users against Entra admin center role assignments.
- Review PIM warning messages, especially in tenants without Entra ID P2.
- Confirm MFA registration fields are populated or explicitly `Unknown` when lookup fails.
- If `-IncludeMailbox` is used, confirm the Exchange Online connection succeeds.

## Active Directory Duplicate Validation

For `Find-ADDuplicateEmailProxyAddresses.ps1`:

- Run against a narrow `-SearchBase` first.
- Review pair-based output and occurrence output together.
- Confirm duplicates in AD before remediation.
- Use `-SmtpOnly` when you only want SMTP address collision detection.

## Hardening Assessment Validation

For `Invoke-AdSecurityHardeningAssessment.ps1` and its modules:

- Run one module against a single host first (for example `./Export-DcHardeningAssessment.ps1 -Server DC01`) to confirm WinRM connectivity and rights before a full sweep.
- Treat `Manual` rows as a worklist, not a pass — the `Evidence` column names the artifact to confirm.
- Investigate `Error` rows before trusting the summary; they usually indicate an unreachable host or missing permission, not a passing control.
- Cross-check patch-currency results (DC-002, DJ-002) against your patch-management tooling.

## Discovery Inventory Validation

For `Invoke-AdDivestitureDiscovery.ps1` and its modules:

- Review `DiscoveryIndex.csv` first; any `Error` row means that module's data is missing, not empty.
- Spot-check volumetric counts (users, computers, mailbox sizes) against the source admin centres before using them for sizing or licensing decisions.
- Run `Export-EntraConnectConfiguration.ps1` on the sync server and `Export-ExchangeOnPremInventory.ps1` from an Exchange Management Shell; they report cleanly when run elsewhere but collect nothing.

## DNS Validation

For DNS scripts:

- Query both the system resolver and a public resolver when troubleshooting propagation.
- Review multiple SPF or DMARC warnings manually.
- Pass known DKIM selectors for your mail platforms.
- Treat missing DKIM results as inconclusive unless the selector list is known complete.

## Troubleshooting

| Symptom | Likely Cause | Action |
| --- | --- | --- |
| Graph connection fails | Missing module, consent, or delegated scope | Install Microsoft.Graph modules and confirm admin consent. |
| Graph authentication reports a missing MSAL method | Mixed or stale Microsoft.Graph/Microsoft.Identity.Client module versions | Run in PowerShell 7+, then update Microsoft Graph modules with `Update-Module Microsoft.Graph` or reinstall stale Graph modules. Use `-SkipGraph` if Entra membership enrichment is not required. |
| Exchange connection fails | Missing module or insufficient Exchange permission | Install ExchangeOnlineManagement and verify role assignment. |
| PIM eligibility warnings | Missing Entra ID P2 or Graph permission | Confirm licensing and `RoleManagement.Read.Directory`. |
| Empty detail CSV | No matching data or skipped enrichment | Check `Errors.csv`, input filters, and `-SkipGraph`. |
| Hardening/discovery rows show `Error` for a host | WinRM blocked, host offline, or insufficient rights | Test `Invoke-Command -ComputerName <host> { hostname }`, confirm firewall/WinRM policy, and re-run with `-Credential`. |
| Entra Connect module reports not installed | Script not run on the sync server | Run `Export-EntraConnectConfiguration.ps1` on the Entra Connect server itself. |
| On-prem Exchange module throws about missing cmdlets | No Exchange Management Shell session | Run from the Exchange Management Shell or import an Exchange remote session first. |
| DNS record not found | Propagation, resolver cache, or wrong selector | Retry with another resolver and known selector. |

## Repository Hygiene

Generated exports should not be committed. The repository ignores `.DS_Store` and local `gemini-debug-*.log` files, but it does not globally ignore generated CSV or HTML files because administrators may intentionally save sample output elsewhere. Review `git status` before committing.
