# Exchange Export Progress and Verbose Output Design

## Goal

Make `Export-ExchangeObjectData.ps1` visibly active throughout long Exchange Online and Microsoft Graph runs, especially on PowerShell 7/macOS, without flooding normal output for large tenants.

## Selected Approach

Use phase-level `Write-Progress` and concise `Write-Host` status by default. Use `Write-Verbose` for per-object and diagnostic detail when the script is run with `-Verbose`.

This balances normal readability with troubleshooting depth for runs such as `./Export-ExchangeObjectData.ps1 -All -RecipientType Mailbox` against 1000+ recipients.

## Scope

Add progress and status coverage for these phases:

- Microsoft Graph connection
- ExchangeOnlineManagement import and Exchange Online connection
- Exchange recipient inventory loading
- proxy address derivation
- Graph group membership collection
- SendAs permission collection
- FullAccess permission collection
- consolidated row creation
- CSV export
- dashboard generation
- cleanup and disconnect

No export schema, permission logic, Graph query shape, or dashboard behavior should change.

## Default Output Behavior

Normal runs should show:

- current phase start and completion messages
- row/object counts after each phase
- progress bars for phases that can take noticeable time
- warning messages for recoverable failures, as today

Normal runs should not print every recipient identity unless the phase is already using progress status for the current item.

## Verbose Output Behavior

When run with `-Verbose`, long loops should emit useful diagnostics such as:

- current recipient identity
- lookup identity used for SendAs and FullAccess
- Graph batch request counts
- row counts produced by each phase
- fallback behavior, such as per-recipient SendAs lookups with `Get-EXORecipientPermission`

Verbose output must not expose secrets. It may include recipient identities, SMTP addresses, and object IDs because the script already exports those values.

## Implementation Notes

Use existing PowerShell primitives only:

- `Write-Progress` for progress bars
- `Write-Host` for concise phase status
- `Write-Verbose` for diagnostic details

Add small helper functions only if they reduce repeated progress boilerplate. Keep the implementation minimal and localized to existing phases.

## Testing

Update `tests/Export-ExchangeObjectData.Checks.ps1` to assert that key progress and verbose calls exist for previously silent phases, without requiring live Exchange or Graph connections.

Run the existing check script after implementation:

```powershell
pwsh -NoProfile -File "tests/Export-ExchangeObjectData.Checks.ps1"
```

## Copy Sync Requirement

After updating the source script in `AdminToolboxScripts`, ensure the working copy at `/Users/abhijittiwari/Documents/Fortescue/Export-ExchangeObjectData.ps1` matches it byte-for-byte.
