# Export-ExchangeObjectData.ps1 Redesign — Graph-First Hybrid

Date: 2026-07-10
Status: Approved by user (brainstorming session)

## Problem

The current script makes up to five remote Exchange Online calls per recipient. At the
target scale (over 10,000 recipients per `-All` run) this is slow, and the per-recipient
`Get-Recipient`/`Get-Mailbox` proxy-address lookup layer has produced five consecutive
bug-fix attempts (commits e963904, 3a25277, 119aa14, 335b56e, plus an uncommitted retry
loop) because Exchange `Identity` values (display names) are ambiguous. The redesign
removes that layer entirely and moves per-recipient work to Microsoft Graph.

## Decisions (from brainstorming)

| Decision | Choice |
| --- | --- |
| FullAccess/SendAs (no Graph API exists) | Hybrid — always connect to both Graph and Exchange Online; keep both exports |
| Authentication | Interactive delegated sign-in for both services (as today) |
| External surface | Drop-in replacement: identical parameters, CSV filenames/columns, dashboard |
| Scale target | Over 10,000 recipients per `-All` run |
| Architecture | EXO spine + Graph batch engine |
| Accepted output-value change | `ExchangeGroupMemberships.GroupIdentity` holds Graph group display name (with group mail for disambiguation) instead of Exchange canonical name |

## Architecture: five-phase pipeline

Single script file, same parameter sets (`-All` / `-InputCsv` / `-Identity`,
`-RecipientType`, `-OutputFolder`, `-SkipGraph`, `-NoDashboard`).

### Phase 1 — Inventory (one paged EXO call)

`Get-EXORecipient -ResultSize Unlimited -Properties RecipientTypeDetails,
ExternalDirectoryObjectId, PrimarySmtpAddress, Alias, Guid, DistinguishedName,
EmailAddresses`, filtered by `-RecipientType`. CSV and single-identity modes call
`Get-EXORecipient -Identity <input>` per input row. The inventory is the authoritative
source for `Identity`, `RecipientTypeDetails`, `MailboxType`, `Alias`, and proxy
addresses.

### Phase 2 — Proxy address rows (in-memory, zero network)

Parsed from the `EmailAddresses` returned in Phase 1. The per-recipient proxy lookup
layer (and its identity-candidate retry machinery) is deleted.

### Phase 3 — Group memberships (Graph $batch)

For every recipient with a non-empty `ExternalDirectoryObjectId`:
`GET /directoryObjects/{id}/memberOf/microsoft.graph.group?$select=id,displayName,mail,
mailEnabled,securityEnabled` — issued through the `/$batch` endpoint, 20 sub-requests per
HTTP call (10,000 recipients ≈ 500 calls). One response feeds both CSVs:

- `EntraGroupMemberships.csv`: every group returned (same columns as today).
- `ExchangeGroupMemberships.csv`: the `mailEnabled -eq true` subset. `GroupIdentity`
  contains the group display name; the group mail address is appended in parentheses
  when present.

Direct membership (`memberOf`), not transitive, to match current semantics. Mail
contacts gain Entra membership coverage (today they are skipped).

### Phase 4 — Permissions (EXO; Graph has no API for these)

- SendAs: `-All` runs use one org-wide paged `Get-RecipientPermission -ResultSize
  Unlimited` and join in memory; CSV/single runs call per identity. Same self-trustee
  filtering as today (`NT AUTHORITY\SELF`, `SELF`, `S-1-5-10`).
- FullAccess: `Get-EXOMailboxPermission` per mailbox (no bulk mode exists), scoped to
  mailbox-type recipients only, always identified by immutable `Guid` — never display
  name or alias. Same inherited/self filtering as today.

### Phase 5 — Consolidate and export

Same eight CSVs (`Objects-Consolidated`, `Objects`, `ProxyAddresses`, `FullAccess`,
`SendAs`, `ExchangeGroupMemberships`, `EntraGroupMemberships`, `Errors`) and the same
HTML dashboard. Consolidation joins rows to recipients through a hashtable keyed by
`Identity` (the current `Where-Object` scan per recipient is quadratic at 10k+).

## Modules and authentication

- Imports: `ExchangeOnlineManagement` and `Microsoft.Graph.Authentication` only.
- All Graph traffic uses `Invoke-MgGraphRequest` (raw REST). No Graph SDK entity
  modules or cmdlets (no `Get-MgUser` etc.) — they lack `$batch` support and are slow
  and memory-heavy at this scale.
- `Connect-ExchangeOnline -ShowBanner:$false`; `Connect-MgGraph -Scopes User.Read.All,
  Group.Read.All, Directory.Read.All -NoWelcome`.
- `-SkipGraph` still supported: both membership CSVs are emitted headers-only with a
  logged notice, since both are now Graph-sourced.

## Key components

| Function | Responsibility |
| --- | --- |
| `Invoke-GraphBatch` | Chunk request descriptors into 20s, POST `/$batch`, map sub-responses back to recipients, surface per-item failures without failing the chunk |
| `Get-MembershipRows` | Build batch requests, follow `@odata.nextLink` for recipients in 100+ groups, split results into Exchange/Entra row sets |
| `Get-SendAsRows` | Org-wide bulk for `-All`; per-identity otherwise |
| `Get-FullAccessRows` | Per-mailbox by `Guid`, mailboxes only |
| `ConvertTo-ProxyRows` | Pure transform of Phase 1 `EmailAddresses` |
| `New-ConsolidatedRows` | Hashtable-indexed join |

`ConvertTo-ProxyRows`, `New-ConsolidatedRows`, batch chunking, and membership splitting
are pure functions testable offline with mock objects.

## Throttling and resilience

- Graph 429/503: honor `Retry-After`, exponential backoff, maximum 5 retries, then log
  the item to `Errors.csv` and continue. Failed sub-requests within a batch retry
  individually; the batch is not re-sent whole.
- EXO cmdlets rely on module built-in retry; each per-mailbox FullAccess call is
  try/catch-wrapped into `Errors.csv` (current contract).
- `Write-Progress` per phase with counts.
- No checkpoint/resume in this version; a failed run re-runs. Noted as a future
  enhancement only.

## Error handling contract (unchanged)

Every failure produces an `Errors.csv` row (`Identity, Stage, Operation, Message`) and a
console warning. One broken object never aborts the run.

## Testing

1. Offline: PowerShell parser validation plus a red-check script (pattern of
   `.superpowers/sdd/final-review-red-checks.ps1`) covering proxy parsing, batch
   chunking, membership splitting, consolidation joins, and empty-CSV header emission,
   using mock recipient objects.
2. Live, staged: single `-Identity` run, then a small `-InputCsv` batch, then full
   `-All`, comparing row counts against the old script's output on the same tenant.

## Out of scope

- App-only (certificate) authentication.
- Checkpoint/resume for interrupted runs.
- Public folders, dynamic distribution group expansion, soft-deleted recipients.
- Any change to CSV filenames, column sets, parameters, or dashboard behavior beyond
  the accepted `GroupIdentity` value change.
