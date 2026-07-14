# Cross-Tenant Admin Role Export Design

## Purpose

Create a PowerShell script that uses the provided admin mapping CSV to report matching WAE and Fortescue admin accounts. The report will show each account's Entra ID directory role assignments, immutable ID, user principal name, display name, and assigned licenses in both environments.

The script is read-only. It will not change users, roles, licenses, or tenant configuration.

## Inputs

The script will accept an input CSV path. The CSV must contain these columns:

- `WAEUPN`
- `Prefix`
- `FortescueUPN`

Each row represents one admin identity pair. `Prefix` is preserved in the output so the WAE and Fortescue rows can be matched easily.

Default parameters:

- `-InputCsv ./Admin.csv`
- `-OutputPath ./CrossTenantAdminRoles_yyyyMMdd_HHmmss.csv`

## Output

The script will write one combined CSV containing both environments.

Output columns:

- `Environment`
- `Prefix`
- `InputUPN`
- `UserPrincipalName`
- `DisplayName`
- `ImmutableId`
- `ActiveDirectoryRoles`
- `PimEligibleDirectoryRoles`
- `AllDirectoryRoles`
- `LicensesAssigned`
- `LookupStatus`
- `Error`

`Environment` will be `WAE` or `Fortescue`. `InputUPN` will be the value read from `WAEUPN` or `FortescueUPN`; `UserPrincipalName` will be the actual UPN returned by Microsoft Graph.

Role and license fields will use semicolon-separated values to keep each account to one row per environment.

## Authentication Flow

The script will connect to Microsoft Graph separately for each tenant using delegated authentication.

Required Microsoft Graph scopes:

- `Directory.Read.All`
- `RoleManagement.Read.Directory`

Execution flow:

1. Connect to Microsoft Graph for WAE.
2. Collect WAE role and license data for all `WAEUPN` values.
3. Disconnect from Microsoft Graph.
4. Prompt the operator to sign into Fortescue.
5. Connect to Microsoft Graph for Fortescue.
6. Collect Fortescue role and license data for all `FortescueUPN` values.
7. Disconnect from Microsoft Graph.
8. Export the combined CSV.

This avoids app registration or certificate setup and fits ad hoc administrative use.

## Role Collection

The script will report Entra ID directory roles only, including active and PIM-eligible assignments.

Role sources:

- Active assignments from `Get-MgRoleManagementDirectoryRoleAssignment -All`
- PIM-eligible assignments from `Get-MgRoleManagementDirectoryRoleEligibilityScheduleInstance -All`
- Role display names from `Get-MgRoleManagementDirectoryRoleDefinition -All`

Assignments can be direct user assignments or group assignments. For group assignments, the script will expand transitive group members and mark the role with the group display name, for example:

`Global Administrator (Active via group: Admin Role Group)`

Service principals and other non-user principals will be ignored because the requested report is user-account based.

## User And License Lookup

For each UPN in the CSV, the script will call Microsoft Graph to retrieve:

- `DisplayName`
- `UserPrincipalName`
- `OnPremisesImmutableId`
- `Id`

Assigned licenses will be retrieved with `Get-MgUserLicenseDetail` and exported as SKU part numbers.

## Error Handling

The script will continue processing when individual lookups fail.

Expected statuses:

- `Found`: user was found and processed.
- `NotFound`: Graph could not find the input UPN.
- `Error`: user lookup or license lookup failed unexpectedly.

If PIM eligibility cannot be read because of missing permissions or licensing, the script will warn once for that environment, continue with active roles, leave `PimEligibleDirectoryRoles` blank for that environment, and include the PIM warning text in the `Error` column for rows from that environment.

## Validation

Before exporting, the script will validate that:

- The input CSV exists.
- Required columns are present.
- At least one row exists.

Manual validation after running:

- Spot-check a few WAE and Fortescue rows against the Entra admin center.
- Confirm PIM-eligible roles appear for accounts known to have eligible assignments.
- Confirm missing accounts are represented with `LookupStatus=NotFound` rather than silently skipped.

## Files To Update

Planned implementation files:

- Add `Get-CrossTenantAdminRoleAssignments.ps1`
- Update `README.md`
- Update `docs/script-reference.md`
- Add lightweight checks under `tests/` if existing script-check patterns can be reused without requiring a live Microsoft Graph connection.

## Out Of Scope

- Exchange Online RBAC role groups or management roles.
- MFA status, mailbox details, group membership reporting outside role-assignment expansion.
- Automated app-only authentication.
- Scheduled execution.
