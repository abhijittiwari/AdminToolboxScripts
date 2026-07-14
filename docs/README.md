# Admin Toolbox Scripts Documentation

This directory contains user-facing documentation for the scripts in this repository.

## Start Here

- [Quick Start](quick-start.md): install prerequisites, clone the repository, and run the scripts safely.
- [Script Reference](script-reference.md): parameters, examples, outputs, and validation notes for each script.
- [Operations and Security](operations-and-security.md): permissions, data handling, validation, and troubleshooting guidance.

## Script Inventory

| Script | Primary Use |
| --- | --- |
| `Export-ExchangeObjectData.ps1` | Exchange Online recipient export with permissions, proxy addresses, Exchange/Entra memberships, CSVs, and dashboard. |
| `Invoke-ExchangeObjectDataExport.ps1` | Runs the five modular export jobs in sequence and finishes with an Excel workbook. |
| `Export-ExchangeMigrationReadiness.ps1` | Identifies mail objects blocked, risky, or unsuitable for migration due to holds, retention, licensing, or compliance constraints. |
| `Export-DomainCutoverAssessment.ps1` | Pre-cutover assessment of an accepted domain: configuration, MX/SPF/DKIM DNS, shared mailbox domain usage, and hold impact. |
| `Export-ExchangeObjectInventory.ps1` | Modular job: recipient inventory (`Objects.csv`) and proxy addresses. |
| `Export-ExchangeGroupMemberships.ps1` | Modular job: Entra/Exchange group memberships via Microsoft Graph for an `Objects.csv` inventory. |
| `Export-ExchangeSendAsPermissions.ps1` | Modular job: SendAs permissions for an `Objects.csv` inventory. |
| `Export-ExchangeFullAccessPermissions.ps1` | Modular job: FullAccess mailbox permissions for an `Objects.csv` inventory. |
| `Export-ExchangeObjectDataWorkbook.ps1` | Collates export CSVs into one Excel workbook via ImportExcel (offline; Excel not required). |
| `Get-EntraAdminAccounts.ps1` | Entra ID administrator account and role assignment reporting. |
| `Find-ADDuplicateEmailProxyAddresses.ps1` | On-prem Active Directory duplicate mail/proxy address detection. |
| `Get-MailDnsRecords.ps1` | Windows PowerShell MX, SPF, DMARC, and DKIM lookup. |
| `get-mail-dns-records.sh` | macOS/Linux MX, SPF, DMARC, and DKIM lookup using `dig`. |

## Safety Model

The scripts are intended for reporting and troubleshooting. They do not remediate objects, change DNS, modify Exchange Online, modify Entra ID, or modify Active Directory. Output files may still contain sensitive administrative data and should be handled accordingly.
