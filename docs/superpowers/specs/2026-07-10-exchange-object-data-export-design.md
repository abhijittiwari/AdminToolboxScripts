# Exchange Object Data Export Design

## Goal

Create a PowerShell script that exports Exchange Online object data for all mail-enabled recipients, multiple objects supplied by CSV, or a single object supplied directly. The export should include proxy addresses, mailbox FullAccess permissions, SendAs permissions, Exchange group memberships, Entra ID group memberships, normalized CSV outputs, a consolidated CSV, and a searchable HTML dashboard for multi-object runs.

## Script

Script name: `Export-ExchangeObjectData.ps1`

The script should follow the style of the existing repository scripts:

- PowerShell comment-based help at the top.
- Explicit requirements and examples.
- Read-only operations only.
- Progress and warnings for long-running lookups.
- CSV outputs suitable for review or downstream processing.

## Input Modes

The script supports three mutually exclusive input modes:

1. `-All`
   Exports all mail-enabled recipients Exchange Online can return.

2. `-InputCsv <path>`
   Reads multiple object identities from a CSV file.

3. `-Identity <value>`
   Exports one Exchange-resolvable object identity.

CSV input defaults to an `Identity` column. The `-IdentityColumn <name>` parameter allows callers to choose another column when needed. Identity values should be passed to Exchange recipient lookup as provided, so they may be UPNs, primary SMTP addresses, aliases, display names, GUIDs, or other Exchange-resolvable identifiers.

All input modes should accept `-RecipientType`, defaulting to `Mailbox`. The parameter controls which Exchange recipient object classes are exported or retained after lookup.

## Object Scope

By default, the script exports mailboxes only. Callers can broaden or narrow the object scope with `-RecipientType`.

Supported `-RecipientType` values:

- `Mailbox`: user, shared, room, equipment, discovery, and other mailbox recipient types Exchange returns as mailboxes.
- `Group`: mail-enabled security groups, distribution groups, dynamic distribution groups, and Microsoft 365 groups where Exchange returns them.
- `MailUser`: mail users.
- `MailContact`: mail contacts.
- `All`: all Exchange mail-enabled recipients returned by Exchange Online recipient discovery, including mailboxes, mail users, mail contacts, mail-enabled groups, distribution groups, dynamic groups, and Microsoft 365 groups when Exchange returns them.

For `-All`, recipient discovery should honor `-RecipientType`. For `-InputCsv` and `-Identity`, the script should resolve the supplied identities and then skip objects whose recipient type does not match `-RecipientType`, recording skipped objects in `Errors.csv` with a `RecipientTypeFilter` stage.

The script should normalize object records into a consistent schema even when a property is missing for a recipient type.

## Data Sources

Use an Exchange-first model with Graph enrichment.

Exchange Online PowerShell is the source for:

- Recipient identity and mail attributes.
- Proxy addresses.
- Mailbox FullAccess permissions where the object has mailbox permissions.
- SendAs permissions where available.
- Exchange recipient group memberships.

Microsoft Graph is the source for:

- Entra ID group memberships.
- Directory object identifiers needed to query Graph membership data.

Graph should be used only for features that need it. Entra group membership export is enabled by default because the design requires both Exchange and Entra memberships.

## Connections and Permissions

The script connects to Exchange Online with `Connect-ExchangeOnline`.

The script connects to Microsoft Graph for Entra group memberships. Required Graph scopes should include the least practical delegated scopes for reading users, groups, and memberships, expected to include:

- `User.Read.All`
- `Group.Read.All`
- `Directory.Read.All`

The README should document that Graph permissions may require admin consent depending on tenant policy.

## Exported Data

The script should produce both consolidated and normalized outputs.

### Consolidated Output

`Objects-Consolidated.csv`

One row per object. Multi-value fields are joined with semicolons. This file is optimized for quick review.

Columns:

- `Identity`
- `DisplayName`
- `RecipientType`
- `RecipientTypeDetails`
- `PrimarySmtpAddress`
- `Alias`
- `ExternalDirectoryObjectId`
- `ExchangeObjectId`
- `ProxyAddresses`
- `FullAccessTrustees`
- `SendAsTrustees`
- `ExchangeGroupMemberships`
- `EntraGroupMemberships`
- `HasErrors`

### Normalized Outputs

`Objects.csv`

One row per object with core identity and recipient metadata.

`ProxyAddresses.csv`

One row per proxy address.

`FullAccess.csv`

One row per FullAccess permission assignment.

`SendAs.csv`

One row per SendAs permission assignment.

`ExchangeGroupMemberships.csv`

One row per Exchange group membership.

`EntraGroupMemberships.csv`

One row per Entra ID group membership.

`Errors.csv`

One row per object-level or lookup-level failure.

## HTML Dashboard

The script generates `Dashboard.html` for multi-object runs only:

- Generated for `-All`.
- Generated for `-InputCsv`.
- Skipped for single `-Identity` runs.

The dashboard is a static local HTML file. It embeds the consolidated and normalized export data as JSON, so it does not require a web server.

Dashboard features:

- Search across display name, primary SMTP, recipient type, aliases, proxies, permissions, and group memberships.
- Checkbox selection per visible object.
- Select all visible and clear selection controls.
- Expandable detail views for proxies, FullAccess, SendAs, Exchange groups, and Entra groups.
- Browser-side CSV export for selected consolidated object rows.
- Browser-side detail export buttons for selected proxies, FullAccess, SendAs, Exchange group memberships, and Entra group memberships. Detail exports should include rows related to the selected objects only.

The dashboard and README must warn that the file can contain sensitive permissions, proxy addresses, membership data, and object identifiers.

## Error Handling

The script should continue on per-object and per-lookup failures.

For each failure, the script should:

- Write a warning to the console.
- Add a row to `Errors.csv`.
- Mark the affected object as having errors in the consolidated output when possible.

Error rows should include:

- `Identity`
- `Stage`
- `Operation`
- `Message`

Example stages:

- `RecipientLookup`
- `ProxyAddresses`
- `FullAccess`
- `SendAs`
- `ExchangeGroupMemberships`
- `EntraGroupMemberships`
- `Dashboard`

## Output Location

The script should accept `-OutputFolder`, defaulting to a timestamped folder such as `./ExchangeObjectData_<yyyyMMdd_HHmmss>`.

All CSVs and the dashboard should be written under that folder.

## Non-Goals

- No modification of Exchange or Entra objects.
- No permission remediation.
- No server-hosted dashboard.
- No scheduled task creation.
- No automatic upload of results.

## Testing and Verification

Minimum verification should include:

- PowerShell parser validation for `Export-ExchangeObjectData.ps1`.
- Static checks that expected output file names are present in the script.
- Static checks that all three input modes exist.
- Static checks that dashboard generation is limited to `-All` and `-InputCsv` runs.
- If possible without tenant access, a dry-run or helper-function test path for CSV identity parsing and dashboard HTML generation.

Tenant-dependent live validation should be manual and documented because Exchange Online and Graph access require tenant credentials and permissions.

## README Updates

Add a README section for `Export-ExchangeObjectData.ps1` covering:

- Purpose.
- Requirements.
- Exchange Online and Graph permissions.
- Input modes.
- Recipient type filtering and the default mailbox-only behavior.
- Output files.
- Dashboard behavior.
- Sensitive-data warning.
- Usage examples for `-All`, `-InputCsv`, and `-Identity`.
