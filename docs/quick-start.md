# Quick Start

## 1. Clone the Repository

```bash
git clone https://github.com/abhijittiwari/AdminToolboxScripts.git
cd AdminToolboxScripts
```

## 2. Choose the Right Script

| Goal | Script |
| --- | --- |
| Export Exchange Online recipients, proxies, permissions, and memberships | `Export-ExchangeObjectData.ps1` |
| Run the same export as separate jobs and collate an Excel workbook | `Invoke-ExchangeObjectDataExport.ps1` (or the individual `Export-Exchange*` job scripts) |
| Assess migration readiness and identify blocked or risky mail objects | `Export-ExchangeMigrationReadiness.ps1` |
| Assess an accepted domain before a tenant-to-tenant cutover | `Export-DomainCutoverAssessment.ps1` |
| Report Entra ID admin users and role assignments | `Get-EntraAdminAccounts.ps1` |
| Assess AD / Entra ID / DNS / AD CS security hardening (Appendix B controls) | `Invoke-AdSecurityHardeningAssessment.ps1` (or the individual `Export-*HardeningAssessment.ps1` / `Export-*Assessment.ps1` modules) |
| Gather divestiture discovery evidence (forest, trusts, GPOs, hybrid identity, on-prem Exchange, M365 volumetrics) | `Invoke-AdDivestitureDiscovery.ps1` (or the individual `Export-*Inventory.ps1` modules) |
| Confirm the tenant's authentication model (Managed vs Federated/ADFS) | `Export-EntraAuthenticationModel.ps1` |
| Find duplicate AD mail/proxy addresses | `Find-ADDuplicateEmailProxyAddresses.ps1` |
| Check mail DNS records on Windows | `Get-MailDnsRecords.ps1` |
| Check mail DNS records on macOS/Linux | `get-mail-dns-records.sh` |

## 3. Install Prerequisites

### PowerShell Scripts

Install PowerShell 7 where possible. Windows PowerShell is also supported for scripts that rely on Windows-only modules.

Install only the modules needed for the script you plan to run.

| Script | Required Components |
| --- | --- |
| `Export-ExchangeObjectData.ps1` | `ExchangeOnlineManagement`; Microsoft Graph modules unless `-SkipGraph` is used. |
| `Invoke-ExchangeObjectDataExport.ps1` and the modular export jobs | `ExchangeOnlineManagement`; `Microsoft.Graph.Authentication`; `ImportExcel` for the workbook step. |
| `Export-ExchangeMigrationReadiness.ps1` | `ExchangeOnlineManagement`; `Microsoft.Graph.Authentication` only when using `-IncludeLicensing`. |
| `Export-DomainCutoverAssessment.ps1` | `ExchangeOnlineManagement` unless `-SkipExchange`/`-AcceptedDomainsCsv`; `Resolve-DnsName` or `dig` for DNS. |
| `Get-EntraAdminAccounts.ps1` | Microsoft Graph modules; `ExchangeOnlineManagement` only when using `-IncludeMailbox`. |
| `Find-ADDuplicateEmailProxyAddresses.ps1` | RSAT Active Directory PowerShell module on Windows. |
| `Invoke-AdSecurityHardeningAssessment.ps1` and the hardening modules | PowerShell 7; RSAT (`ActiveDirectory`, `DnsServer`); WinRM to Domain Controllers and sampled computers; Microsoft Graph SDK for the Entra ID module (`-IncludeEntraId`). |
| `Invoke-AdDivestitureDiscovery.ps1` and the discovery modules | PowerShell 7; RSAT (`ActiveDirectory`, `GroupPolicy`, `DnsServer`, `DhcpServer`); `ADSync` on the sync server for `Export-EntraConnectConfiguration.ps1`; Exchange Management Shell for `Export-ExchangeOnPremInventory.ps1`; Graph / `ExchangeOnlineManagement` / `MicrosoftTeams` / SPO / Security & Compliance sessions for the cloud modules (`-IncludeCloud`). |
| `Get-MailDnsRecords.ps1` | Windows `DnsClient` module with `Resolve-DnsName`. |

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Install-Module Microsoft.Graph -Scope CurrentUser
```

For on-prem Active Directory duplicate checks, install RSAT Active Directory tools on Windows so the `ActiveDirectory` module is available. For Windows DNS checks, confirm `Resolve-DnsName` is available before running the script.

### Bash DNS Script

The Bash DNS script requires `bash` and `dig`.

On macOS, `dig` is included by default. On many Linux distributions, install the package that provides DNS utilities, such as `dnsutils` or `bind-utils`.

## 4. Run Common Commands

### Exchange Object Export

```powershell
./Export-ExchangeObjectData.ps1 -All
```

```powershell
./Export-ExchangeObjectData.ps1 -InputCsv ./objects.csv -IdentityColumn Identity -RecipientType All
```

### Modular Exchange Export with Excel Workbook

```powershell
./Invoke-ExchangeObjectDataExport.ps1 -All -OutputFolder ./export
```

### Migration Readiness Assessment

```powershell
./Export-ExchangeMigrationReadiness.ps1 -ObjectsCsv ./export/Objects.csv -IncludeLicensing
```

### Entra Admin Account Report

```powershell
./Get-EntraAdminAccounts.ps1 -IncludeMailbox
```

### Security Hardening Assessment

```powershell
# On-premises modules (DC, domain computers, DNS, AD CS, backup, monitoring)
./Invoke-AdSecurityHardeningAssessment.ps1 -OutputFolder ./HardeningRun
```

```powershell
# Include the Entra ID module (requires a connected Graph session)
Connect-MgGraph -Scopes Policy.Read.All,Directory.Read.All,RoleManagement.Read.Directory,AccessReview.Read.All,User.Read.All
./Invoke-AdSecurityHardeningAssessment.ps1 -IncludeEntraId -BreakGlassUpn bg1@contoso.com
```

### Divestiture Discovery

```powershell
# On-premises discovery set (forest, sites, trusts, OU/GPO, populations, privileged, SIDHistory, dependencies)
./Invoke-AdDivestitureDiscovery.ps1 -OutputFolder ./DiscoveryRun
```

```powershell
# Add cloud modules (auth model, M365 volumetrics, licensing, Teams/Voice, Purview)
Connect-MgGraph -Scopes Domain.Read.All,Organization.Read.All,User.Read.All
./Invoke-AdDivestitureDiscovery.ps1 -IncludeCloud -OutputFolder ./DiscoveryRun
```

### AD Duplicate Proxy Check

```powershell
./Find-ADDuplicateEmailProxyAddresses.ps1 -SearchBase "DC=contoso,DC=com"
```

### DNS Checks

```powershell
./Get-MailDnsRecords.ps1 -Domain contoso.com
```

```bash
./get-mail-dns-records.sh -d contoso.com -s 8.8.8.8
```

## 5. Validate Results

- Review console warnings before using the output.
- Open generated CSVs and confirm expected headers are present.
- For Exchange object exports, review `Errors.csv` even when it is empty or headers-only.
- For dashboard output, open `Dashboard.html` locally and validate search, selection, expansion, and export behavior.
- Spot-check a small sample against the source system before sharing reports.

## 6. Protect Output Files

Generated CSV and HTML files may contain user identifiers, role assignments, group memberships, mailbox permissions, proxy addresses, or object identifiers. Store and share outputs as sensitive administrative data.
