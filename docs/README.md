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
| `Get-EntraAdminMailboxInventory.ps1` | Focused admin mailbox/license inventory including active and PIM-eligible directory role assignments. |
| `Get-CrossTenantUserMapping.ps1` | Maps source tenant users to target tenant accounts via synced on-premises extension attributes 13/14 containing the source UPN. |
| `Export-AdminCaPimPosture.ps1` | Conditional Access and PIM-eligible role mapping for in-scope admins, exported as a combined Excel workbook. |
| `Invoke-AdSecurityHardeningAssessment.ps1` | Runs the Appendix B hardening modules and merges their per-control results into one CSV with a consolidated summary. |
| `Export-DcHardeningAssessment.ps1` | Domain Controller hardening assessment (Appendix B controls DC-001..DC-022). |
| `Export-DomainComputerHardeningAssessment.ps1` | Domain-joined computer hardening assessment (Appendix B controls DJ-001..DJ-020). |
| `Export-EntraIdHardeningAssessment.ps1` | Entra ID tenant hardening assessment via Microsoft Graph (Appendix B controls AAD-001..AAD-015). |
| `Export-DnsHardeningAssessment.ps1` | Windows DNS Server hardening assessment (Appendix B controls DNS-001..DNS-010). |
| `Export-AdcsHardeningAssessment.ps1` | AD CS hardening assessment including ESC1-ESC8 (Appendix B controls ADCS-001..ADCS-010). |
| `Export-AdBackupRecoveryAssessment.ps1` | AD backup and forest-recovery assessment (Appendix B controls BR-001..BR-010). |
| `Export-AdMonitoringAssessment.ps1` | Monitoring and detection assessment (Appendix B controls MON-001..MON-010). |
| `Find-ADDuplicateEmailProxyAddresses.ps1` | On-prem Active Directory duplicate mail/proxy address detection. |
| `Get-MailDnsRecords.ps1` | Windows PowerShell MX, SPF, DMARC, and DKIM lookup. |
| `get-mail-dns-records.sh` | macOS/Linux MX, SPF, DMARC, and DKIM lookup using `dig`. |

## Safety Model

The scripts are intended for reporting and troubleshooting. They do not remediate objects, change DNS, modify Exchange Online, modify Entra ID, or modify Active Directory. Output files may still contain sensitive administrative data and should be handled accordingly.
