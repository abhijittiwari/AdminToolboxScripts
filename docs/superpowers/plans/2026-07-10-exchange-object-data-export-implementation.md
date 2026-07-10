# Exchange Object Data Export Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build `Export-ExchangeObjectData.ps1`, a read-only Exchange Online and Graph reporting script that exports recipient object data, permissions, memberships, normalized CSVs, consolidated CSVs, and a searchable HTML dashboard.

**Architecture:** Implement one self-contained PowerShell script following the repository style, with small helper functions inside the script for input resolution, object normalization, permission collection, membership collection, CSV export, dashboard generation, and error logging. Use Exchange Online PowerShell as the primary data source and Microsoft Graph only for Entra group membership enrichment. Keep tenant-dependent validation manual; use parser and static tests for automated verification.

**Tech Stack:** PowerShell 5.1+/PowerShell 7, ExchangeOnlineManagement, Microsoft.Graph, static HTML/CSS/JavaScript, CSV exports.

## Global Constraints

- Script name must be `Export-ExchangeObjectData.ps1`.
- Script must be read-only and must not modify Exchange or Entra objects.
- Input modes are mutually exclusive: `-All`, `-InputCsv`, or `-Identity`.
- CSV input defaults to `Identity`; `-IdentityColumn` overrides it.
- `-RecipientType` defaults to `Mailbox`.
- Supported `-RecipientType` values are `Mailbox`, `Group`, `MailUser`, `MailContact`, and `All`.
- Outputs must include `Objects-Consolidated.csv`, `Objects.csv`, `ProxyAddresses.csv`, `FullAccess.csv`, `SendAs.csv`, `ExchangeGroupMemberships.csv`, `EntraGroupMemberships.csv`, and `Errors.csv`.
- `Dashboard.html` is generated for `-All` and `-InputCsv`, not for single `-Identity` runs.
- Dashboard must be a static local HTML file with embedded JSON and browser-side CSV export.
- The script must continue on per-object and per-lookup failures and write `Errors.csv`.
- README must warn that exports and dashboard may contain sensitive proxy, permission, membership, and object identifier data.

---

## File Structure

- Create `Export-ExchangeObjectData.ps1`: Main script with parameters, helper functions, Exchange/Graph collection, exports, dashboard generation, and cleanup.
- Modify `README.md`: Add table entry and full documentation section for the new script.
- No external module files: This repository currently keeps each script self-contained. Keep this script self-contained for portability.
- No test framework files: Use exact `pwsh` commands for parser and static checks because the repo has no test harness and tenant access is required for live integration.

---

### Task 1: Create Script Skeleton, Parameters, and Input Validation

**Files:**
- Create: `Export-ExchangeObjectData.ps1`

**Interfaces:**
- Produces script parameters: `-All`, `-InputCsv`, `-Identity`, `-IdentityColumn`, `-RecipientType`, `-OutputFolder`, `-SkipGraph`, `-NoDashboard`.
- Produces helper function: `Add-ExportError -Identity <string> -Stage <string> -Operation <string> -Message <string>`.
- Produces helper function: `Initialize-OutputFolder -OutputFolder <string> -> string`.
- Produces helper function: `Get-InputIdentityList -> string[]`.

- [ ] **Step 1: Write the failing parser/static test**

Run this before creating the script:

```bash
pwsh -NoProfile -Command '$path="./Export-ExchangeObjectData.ps1"; if (-not (Test-Path $path)) { throw "missing Export-ExchangeObjectData.ps1" }'
```

Expected: FAIL with `missing Export-ExchangeObjectData.ps1`.

- [ ] **Step 2: Create the minimal script skeleton**

Create `Export-ExchangeObjectData.ps1` with this initial content:

```powershell
<#
.SYNOPSIS
    Exports Exchange Online recipient object data, permissions, memberships, and dashboard output.

.DESCRIPTION
    Exports proxy addresses, FullAccess permissions, SendAs permissions, Exchange group memberships,
    and Entra ID group memberships for Exchange Online recipients. Supports all recipients of a
    selected type, identities from CSV, or a single identity.

.NOTES
    Requires: ExchangeOnlineManagement
    Optional Graph enrichment requires: Microsoft.Graph
    This script is read-only. It does not modify Exchange or Entra objects.

.EXAMPLE
    ./Export-ExchangeObjectData.ps1 -All

.EXAMPLE
    ./Export-ExchangeObjectData.ps1 -All -RecipientType All

.EXAMPLE
    ./Export-ExchangeObjectData.ps1 -InputCsv ./objects.csv -IdentityColumn Identity -RecipientType Group

.EXAMPLE
    ./Export-ExchangeObjectData.ps1 -Identity user@contoso.com
#>

[CmdletBinding(DefaultParameterSetName = 'All')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'All')]
    [switch]$All,

    [Parameter(Mandatory = $true, ParameterSetName = 'Csv')]
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
    [string]$InputCsv,

    [Parameter(Mandatory = $true, ParameterSetName = 'Identity')]
    [string]$Identity,

    [Parameter(Mandatory = $false, ParameterSetName = 'Csv')]
    [string]$IdentityColumn = 'Identity',

    [Parameter(Mandatory = $false)]
    [ValidateSet('Mailbox', 'Group', 'MailUser', 'MailContact', 'All')]
    [string]$RecipientType = 'Mailbox',

    [Parameter(Mandatory = $false)]
    [string]$OutputFolder = ".\ExchangeObjectData_$(Get-Date -Format 'yyyyMMdd_HHmmss')",

    [Parameter(Mandatory = $false)]
    [switch]$SkipGraph,

    [Parameter(Mandatory = $false)]
    [switch]$NoDashboard
)

$ErrorActionPreference = 'Stop'

$script:Errors = New-Object System.Collections.Generic.List[object]

function Add-ExportError {
    param(
        [Parameter(Mandatory = $false)] [string]$Identity,
        [Parameter(Mandatory = $true)] [string]$Stage,
        [Parameter(Mandatory = $true)] [string]$Operation,
        [Parameter(Mandatory = $true)] [string]$Message
    )

    $script:Errors.Add([pscustomobject]@{
        Identity  = $Identity
        Stage     = $Stage
        Operation = $Operation
        Message   = $Message
    }) | Out-Null

    Write-Warning "[$Stage/$Operation] $Identity $Message"
}

function Initialize-OutputFolder {
    param([Parameter(Mandatory = $true)] [string]$Path)

    $resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    if (-not (Test-Path -Path $resolvedPath)) {
        New-Item -Path $resolvedPath -ItemType Directory -Force | Out-Null
    }
    return $resolvedPath
}

function Get-InputIdentityList {
    if ($PSCmdlet.ParameterSetName -eq 'Identity') {
        return @($Identity)
    }

    if ($PSCmdlet.ParameterSetName -eq 'Csv') {
        $rows = @(Import-Csv -Path $InputCsv)
        if ($rows.Count -eq 0) {
            throw "Input CSV '$InputCsv' contains no rows."
        }

        if ($null -eq $rows[0].PSObject.Properties[$IdentityColumn]) {
            throw "Input CSV '$InputCsv' does not contain identity column '$IdentityColumn'."
        }

        return @(
            $rows |
                ForEach-Object { [string]$_.PSObject.Properties[$IdentityColumn].Value } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        )
    }

    return @()
}

$resolvedOutputFolder = Initialize-OutputFolder -Path $OutputFolder
$isMultiObjectRun = $PSCmdlet.ParameterSetName -in @('All', 'Csv')

Write-Host "Output folder: $resolvedOutputFolder" -ForegroundColor Cyan
Write-Host "Recipient type: $RecipientType" -ForegroundColor Cyan
```

- [ ] **Step 3: Verify parser and required parameters**

Run:

```bash
pwsh -NoProfile -Command '$tokens=$null; $errors=$null; [System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path "./Export-ExchangeObjectData.ps1"), [ref]$tokens, [ref]$errors) > $null; if ($errors.Count) { $errors | Format-List *; exit 1 }'
```

Expected: PASS with no output.

Run:

```bash
grep -n "ValidateSet('Mailbox', 'Group', 'MailUser', 'MailContact', 'All')" Export-ExchangeObjectData.ps1 && grep -n "\[string\]\$RecipientType = 'Mailbox'" Export-ExchangeObjectData.ps1
```

Expected: PASS and prints matching lines.

- [ ] **Step 4: Commit**

```bash
git add Export-ExchangeObjectData.ps1
git commit -m "Add Exchange object export script skeleton"
```

---

### Task 2: Add Recipient Discovery, Type Filtering, and Object Normalization

**Files:**
- Modify: `Export-ExchangeObjectData.ps1`

**Interfaces:**
- Consumes: `Get-InputIdentityList`, `$RecipientType`, `$PSCmdlet.ParameterSetName`, `Add-ExportError`.
- Produces: `Test-RecipientMatchesType -Recipient <object> -RecipientType <string> -> bool`.
- Produces: `Get-MailboxType -Recipient <object> -> string`.
- Produces: `ConvertTo-ExportObject -Recipient <object> -> pscustomobject`.
- Produces: `Get-TargetRecipients -> object[]`.

- [ ] **Step 1: Write the failing static checks**

Run before implementation:

```bash
grep -n "function Test-RecipientMatchesType" Export-ExchangeObjectData.ps1
```

Expected: FAIL with no output.

Run:

```bash
grep -n "MailboxType" Export-ExchangeObjectData.ps1
```

Expected: FAIL with no output.

- [ ] **Step 2: Add recipient helper functions before the main output lines**

Insert these functions before `$resolvedOutputFolder = Initialize-OutputFolder -Path $OutputFolder`:

```powershell
function Test-RecipientMatchesType {
    param(
        [Parameter(Mandatory = $true)] $Recipient,
        [Parameter(Mandatory = $true)] [string]$SelectedRecipientType
    )

    if ($SelectedRecipientType -eq 'All') {
        return $true
    }

    $typeDetails = [string]$Recipient.RecipientTypeDetails
    $type = [string]$Recipient.RecipientType

    switch ($SelectedRecipientType) {
        'Mailbox' {
            return ($typeDetails -match 'Mailbox$' -or $type -eq 'UserMailbox')
        }
        'Group' {
            return ($typeDetails -match 'Group$' -or $typeDetails -match 'DistributionGroup' -or $typeDetails -eq 'MailUniversalSecurityGroup')
        }
        'MailUser' {
            return ($typeDetails -eq 'MailUser' -or $type -eq 'MailUser')
        }
        'MailContact' {
            return ($typeDetails -eq 'MailContact' -or $type -eq 'MailContact')
        }
    }

    return $false
}

function Get-MailboxType {
    param([Parameter(Mandatory = $true)] $Recipient)

    $typeDetails = [string]$Recipient.RecipientTypeDetails
    if ($typeDetails -match 'Mailbox$') {
        return $typeDetails
    }

    if ([string]$Recipient.RecipientType -eq 'UserMailbox') {
        return 'UserMailbox'
    }

    return ''
}

function ConvertTo-ExportObject {
    param([Parameter(Mandatory = $true)] $Recipient)

    [pscustomobject]@{
        Identity                  = [string]$Recipient.Identity
        DisplayName               = [string]$Recipient.DisplayName
        RecipientType             = [string]$Recipient.RecipientType
        RecipientTypeDetails      = [string]$Recipient.RecipientTypeDetails
        MailboxType               = Get-MailboxType -Recipient $Recipient
        PrimarySmtpAddress        = [string]$Recipient.PrimarySmtpAddress
        Alias                     = [string]$Recipient.Alias
        ExternalDirectoryObjectId = [string]$Recipient.ExternalDirectoryObjectId
        ExchangeObjectId          = [string]$Recipient.Guid
        DistinguishedName         = [string]$Recipient.DistinguishedName
    }
}

function Get-TargetRecipients {
    if ($PSCmdlet.ParameterSetName -eq 'All') {
        Write-Host "Loading Exchange recipients..." -ForegroundColor Cyan
        return @(
            Get-EXORecipient -ResultSize Unlimited -Properties RecipientTypeDetails,ExternalDirectoryObjectId,PrimarySmtpAddress,Alias,Guid,DistinguishedName |
                Where-Object { Test-RecipientMatchesType -Recipient $_ -SelectedRecipientType $RecipientType }
        )
    }

    $identities = @(Get-InputIdentityList)
    $recipients = New-Object System.Collections.Generic.List[object]

    foreach ($inputIdentity in $identities) {
        try {
            $recipient = Get-EXORecipient -Identity $inputIdentity -Properties RecipientTypeDetails,ExternalDirectoryObjectId,PrimarySmtpAddress,Alias,Guid,DistinguishedName -ErrorAction Stop
            if (Test-RecipientMatchesType -Recipient $recipient -SelectedRecipientType $RecipientType) {
                $recipients.Add($recipient) | Out-Null
            }
            else {
                Add-ExportError -Identity $inputIdentity -Stage 'RecipientTypeFilter' -Operation 'FilterRecipient' -Message "Recipient '$($recipient.DisplayName)' does not match selected recipient type '$RecipientType'."
            }
        }
        catch {
            Add-ExportError -Identity $inputIdentity -Stage 'RecipientLookup' -Operation 'Get-EXORecipient' -Message $_.Exception.Message
        }
    }

    return @($recipients)
}
```

- [ ] **Step 3: Add connection and recipient loading**

Append after the existing `Write-Host "Recipient type: $RecipientType"` line:

```powershell
Import-Module ExchangeOnlineManagement -ErrorAction Stop
Connect-ExchangeOnline -ShowBanner:$false

$targetRecipients = @(Get-TargetRecipients)
$objects = @($targetRecipients | ForEach-Object { ConvertTo-ExportObject -Recipient $_ })

Write-Host "Objects selected: $($objects.Count)" -ForegroundColor Green
```

- [ ] **Step 4: Verify parser and static checks**

Run:

```bash
pwsh -NoProfile -Command '$tokens=$null; $errors=$null; [System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path "./Export-ExchangeObjectData.ps1"), [ref]$tokens, [ref]$errors) > $null; if ($errors.Count) { $errors | Format-List *; exit 1 }'
```

Expected: PASS with no output.

Run:

```bash
grep -n "function Test-RecipientMatchesType" Export-ExchangeObjectData.ps1 && grep -n "MailboxType" Export-ExchangeObjectData.ps1 && grep -n "RecipientTypeFilter" Export-ExchangeObjectData.ps1
```

Expected: PASS and prints matching lines.

- [ ] **Step 5: Commit**

```bash
git add Export-ExchangeObjectData.ps1
git commit -m "Add Exchange recipient discovery and filtering"
```

---

### Task 3: Add Proxy Address, FullAccess, SendAs, and Exchange Group Membership Collection

**Files:**
- Modify: `Export-ExchangeObjectData.ps1`

**Interfaces:**
- Consumes: `$targetRecipients`, `Add-ExportError`.
- Produces: `$proxyRows`, `$fullAccessRows`, `$sendAsRows`, `$exchangeGroupMembershipRows`.
- Produces: functions `Get-ProxyAddressRows`, `Get-FullAccessRows`, `Get-SendAsRows`, `Get-ExchangeGroupMembershipRows`.

- [ ] **Step 1: Write failing static checks**

Run:

```bash
grep -n "function Get-FullAccessRows" Export-ExchangeObjectData.ps1
```

Expected: FAIL with no output.

Run:

```bash
grep -n "Get-RecipientPermission" Export-ExchangeObjectData.ps1
```

Expected: FAIL with no output.

- [ ] **Step 2: Add Exchange detail collection functions before `Get-TargetRecipients`**

```powershell
function Get-ProxyAddressRows {
    param([Parameter(Mandatory = $true)] $Recipient)

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($address in @($Recipient.EmailAddresses)) {
        if ($null -eq $address) { continue }
        $value = [string]$address
        $parts = $value -split ':', 2
        $prefix = if ($parts.Count -eq 2) { $parts[0] } else { '' }
        $addressValue = if ($parts.Count -eq 2) { $parts[1] } else { $value }

        $rows.Add([pscustomobject]@{
            Identity           = [string]$Recipient.Identity
            PrimarySmtpAddress = [string]$Recipient.PrimarySmtpAddress
            AddressType        = $prefix
            ProxyAddress       = $addressValue
            RawProxyAddress    = $value
            IsPrimarySmtp      = $value.StartsWith('SMTP:')
        }) | Out-Null
    }
    return @($rows)
}

function Get-FullAccessRows {
    param([Parameter(Mandatory = $true)] $Recipient)

    $mailboxType = Get-MailboxType -Recipient $Recipient
    if ([string]::IsNullOrWhiteSpace($mailboxType)) {
        return @()
    }

    try {
        return @(
            Get-EXOMailboxPermission -Identity $Recipient.Identity -ErrorAction Stop |
                Where-Object { -not $_.IsInherited -and $_.User -notmatch '^NT AUTHORITY\\SELF$' -and ($_.AccessRights -contains 'FullAccess') } |
                ForEach-Object {
                    [pscustomobject]@{
                        Identity           = [string]$Recipient.Identity
                        PrimarySmtpAddress = [string]$Recipient.PrimarySmtpAddress
                        Trustee            = [string]$_.User
                        AccessRights       = ($_.AccessRights -join '; ')
                        Deny               = [bool]$_.Deny
                        IsInherited        = [bool]$_.IsInherited
                    }
                }
        )
    }
    catch {
        Add-ExportError -Identity ([string]$Recipient.Identity) -Stage 'FullAccess' -Operation 'Get-EXOMailboxPermission' -Message $_.Exception.Message
        return @()
    }
}

function Get-SendAsRows {
    param([Parameter(Mandatory = $true)] $Recipient)

    try {
        return @(
            Get-RecipientPermission -Identity $Recipient.Identity -ErrorAction Stop |
                Where-Object { -not $_.IsInherited -and ($_.AccessRights -contains 'SendAs') } |
                ForEach-Object {
                    [pscustomobject]@{
                        Identity           = [string]$Recipient.Identity
                        PrimarySmtpAddress = [string]$Recipient.PrimarySmtpAddress
                        Trustee            = [string]$_.Trustee
                        AccessRights       = ($_.AccessRights -join '; ')
                        IsInherited        = [bool]$_.IsInherited
                    }
                }
        )
    }
    catch {
        Add-ExportError -Identity ([string]$Recipient.Identity) -Stage 'SendAs' -Operation 'Get-RecipientPermission' -Message $_.Exception.Message
        return @()
    }
}

function Get-ExchangeGroupMembershipRows {
    param([Parameter(Mandatory = $true)] $Recipient)

    try {
        return @(
            Get-EXORecipient -Identity $Recipient.Identity -Properties MemberOfGroup -ErrorAction Stop |
                Select-Object -ExpandProperty MemberOfGroup -ErrorAction SilentlyContinue |
                ForEach-Object {
                    [pscustomobject]@{
                        Identity           = [string]$Recipient.Identity
                        PrimarySmtpAddress = [string]$Recipient.PrimarySmtpAddress
                        GroupIdentity      = [string]$_
                    }
                }
        )
    }
    catch {
        Add-ExportError -Identity ([string]$Recipient.Identity) -Stage 'ExchangeGroupMemberships' -Operation 'Get-EXORecipient MemberOfGroup' -Message $_.Exception.Message
        return @()
    }
}
```

- [ ] **Step 3: Ensure recipient discovery loads EmailAddresses**

In `Get-TargetRecipients`, update both `Get-EXORecipient` calls so the `-Properties` list includes `EmailAddresses`.

Use this exact property list in both places:

```powershell
-Properties RecipientTypeDetails,ExternalDirectoryObjectId,PrimarySmtpAddress,Alias,Guid,DistinguishedName,EmailAddresses,MemberOfGroup
```

- [ ] **Step 4: Add collection loop after `$objects` is built**

Append after `Write-Host "Objects selected: $($objects.Count)" -ForegroundColor Green`:

```powershell
$proxyRows = New-Object System.Collections.Generic.List[object]
$fullAccessRows = New-Object System.Collections.Generic.List[object]
$sendAsRows = New-Object System.Collections.Generic.List[object]
$exchangeGroupMembershipRows = New-Object System.Collections.Generic.List[object]

$index = 0
foreach ($recipient in $targetRecipients) {
    $index++
    Write-Progress -Activity 'Collecting Exchange object data' -Status "$index of $($targetRecipients.Count): $($recipient.Identity)" -PercentComplete (($index / [Math]::Max($targetRecipients.Count, 1)) * 100)

    foreach ($row in @(Get-ProxyAddressRows -Recipient $recipient)) { $proxyRows.Add($row) | Out-Null }
    foreach ($row in @(Get-FullAccessRows -Recipient $recipient)) { $fullAccessRows.Add($row) | Out-Null }
    foreach ($row in @(Get-SendAsRows -Recipient $recipient)) { $sendAsRows.Add($row) | Out-Null }
    foreach ($row in @(Get-ExchangeGroupMembershipRows -Recipient $recipient)) { $exchangeGroupMembershipRows.Add($row) | Out-Null }
}
Write-Progress -Activity 'Collecting Exchange object data' -Completed
```

- [ ] **Step 5: Verify parser and static checks**

Run:

```bash
pwsh -NoProfile -Command '$tokens=$null; $errors=$null; [System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path "./Export-ExchangeObjectData.ps1"), [ref]$tokens, [ref]$errors) > $null; if ($errors.Count) { $errors | Format-List *; exit 1 }'
```

Expected: PASS with no output.

Run:

```bash
grep -n "function Get-FullAccessRows" Export-ExchangeObjectData.ps1 && grep -n "Get-RecipientPermission" Export-ExchangeObjectData.ps1 && grep -n "EmailAddresses,MemberOfGroup" Export-ExchangeObjectData.ps1
```

Expected: PASS and prints matching lines.

- [ ] **Step 6: Commit**

```bash
git add Export-ExchangeObjectData.ps1
git commit -m "Collect Exchange permissions and memberships"
```

---

### Task 4: Add Graph Entra Group Membership Collection

**Files:**
- Modify: `Export-ExchangeObjectData.ps1`

**Interfaces:**
- Consumes: `$targetRecipients`, `$SkipGraph`, `Add-ExportError`.
- Produces: `$entraGroupMembershipRows`.
- Produces function: `Connect-GraphIfNeeded`.
- Produces function: `Get-EntraGroupMembershipRows -Recipient <object> -> object[]`.

- [ ] **Step 1: Write failing static checks**

Run:

```bash
grep -n "function Get-EntraGroupMembershipRows" Export-ExchangeObjectData.ps1
```

Expected: FAIL with no output.

- [ ] **Step 2: Add Graph functions before the main connection block**

```powershell
function Connect-GraphIfNeeded {
    if ($SkipGraph) {
        Write-Host 'Skipping Microsoft Graph connection because -SkipGraph was specified.' -ForegroundColor Yellow
        return $false
    }

    try {
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
        Import-Module Microsoft.Graph.Groups -ErrorAction Stop
        Connect-MgGraph -Scopes @('User.Read.All', 'Group.Read.All', 'Directory.Read.All') -NoWelcome
        return $true
    }
    catch {
        Add-ExportError -Identity '' -Stage 'GraphConnect' -Operation 'Connect-MgGraph' -Message $_.Exception.Message
        return $false
    }
}

function Get-EntraGroupMembershipRows {
    param([Parameter(Mandatory = $true)] $Recipient)

    $externalId = [string]$Recipient.ExternalDirectoryObjectId
    if ([string]::IsNullOrWhiteSpace($externalId)) {
        return @()
    }

    try {
        return @(
            Get-MgUserMemberOf -UserId $externalId -All -ErrorAction Stop |
                Where-Object { $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group' } |
                ForEach-Object {
                    [pscustomobject]@{
                        Identity           = [string]$Recipient.Identity
                        PrimarySmtpAddress = [string]$Recipient.PrimarySmtpAddress
                        GroupId            = [string]$_.Id
                        GroupDisplayName   = [string]$_.AdditionalProperties.displayName
                        GroupMail          = [string]$_.AdditionalProperties.mail
                        GroupSecurityEnabled = [string]$_.AdditionalProperties.securityEnabled
                    }
                }
        )
    }
    catch {
        Add-ExportError -Identity ([string]$Recipient.Identity) -Stage 'EntraGroupMemberships' -Operation 'Get-MgUserMemberOf' -Message $_.Exception.Message
        return @()
    }
}
```

- [ ] **Step 3: Connect Graph after Exchange recipient selection**

Add after `Write-Host "Objects selected: $($objects.Count)" -ForegroundColor Green` and before the Exchange collection loop:

```powershell
$graphConnected = Connect-GraphIfNeeded
$entraGroupMembershipRows = New-Object System.Collections.Generic.List[object]
```

- [ ] **Step 4: Collect Entra memberships inside the existing recipient loop**

Inside the loop from Task 3, after Exchange membership collection, add:

```powershell
if ($graphConnected) {
    foreach ($row in @(Get-EntraGroupMembershipRows -Recipient $recipient)) { $entraGroupMembershipRows.Add($row) | Out-Null }
}
```

- [ ] **Step 5: Verify parser and static checks**

Run:

```bash
pwsh -NoProfile -Command '$tokens=$null; $errors=$null; [System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path "./Export-ExchangeObjectData.ps1"), [ref]$tokens, [ref]$errors) > $null; if ($errors.Count) { $errors | Format-List *; exit 1 }'
```

Expected: PASS with no output.

Run:

```bash
grep -n "function Get-EntraGroupMembershipRows" Export-ExchangeObjectData.ps1 && grep -n "Connect-MgGraph" Export-ExchangeObjectData.ps1 && grep -n "Directory.Read.All" Export-ExchangeObjectData.ps1
```

Expected: PASS and prints matching lines.

- [ ] **Step 6: Commit**

```bash
git add Export-ExchangeObjectData.ps1
git commit -m "Collect Entra group memberships"
```

---

### Task 5: Add Consolidation and CSV Export

**Files:**
- Modify: `Export-ExchangeObjectData.ps1`

**Interfaces:**
- Consumes: `$objects`, `$proxyRows`, `$fullAccessRows`, `$sendAsRows`, `$exchangeGroupMembershipRows`, `$entraGroupMembershipRows`, `$script:Errors`, `$resolvedOutputFolder`.
- Produces: `Objects-Consolidated.csv`, `Objects.csv`, `ProxyAddresses.csv`, `FullAccess.csv`, `SendAs.csv`, `ExchangeGroupMemberships.csv`, `EntraGroupMemberships.csv`, `Errors.csv`.
- Produces function: `Join-Values`.
- Produces function: `New-ConsolidatedRows`.
- Produces function: `Export-ReportCsvs`.

- [ ] **Step 1: Write failing static check**

Run:

```bash
grep -n "Objects-Consolidated.csv" Export-ExchangeObjectData.ps1
```

Expected: FAIL with no output.

- [ ] **Step 2: Add consolidation/export functions before the main connection block**

```powershell
function Join-Values {
    param([Parameter(Mandatory = $false)] $Values)

    return (@($Values) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Sort-Object -Unique) -join '; '
}

function New-ConsolidatedRows {
    param(
        [Parameter(Mandatory = $true)] [object[]]$Objects,
        [Parameter(Mandatory = $true)] [object[]]$ProxyRows,
        [Parameter(Mandatory = $true)] [object[]]$FullAccessRows,
        [Parameter(Mandatory = $true)] [object[]]$SendAsRows,
        [Parameter(Mandatory = $true)] [object[]]$ExchangeGroupRows,
        [Parameter(Mandatory = $true)] [object[]]$EntraGroupRows,
        [Parameter(Mandatory = $true)] [object[]]$ErrorRows
    )

    foreach ($object in $Objects) {
        $identity = [string]$object.Identity
        [pscustomobject]@{
            Identity                  = $object.Identity
            DisplayName               = $object.DisplayName
            RecipientType             = $object.RecipientType
            RecipientTypeDetails      = $object.RecipientTypeDetails
            MailboxType               = $object.MailboxType
            PrimarySmtpAddress        = $object.PrimarySmtpAddress
            Alias                     = $object.Alias
            ExternalDirectoryObjectId = $object.ExternalDirectoryObjectId
            ExchangeObjectId          = $object.ExchangeObjectId
            ProxyAddresses            = Join-Values (($ProxyRows | Where-Object { $_.Identity -eq $identity }).RawProxyAddress)
            FullAccessTrustees        = Join-Values (($FullAccessRows | Where-Object { $_.Identity -eq $identity }).Trustee)
            SendAsTrustees            = Join-Values (($SendAsRows | Where-Object { $_.Identity -eq $identity }).Trustee)
            ExchangeGroupMemberships  = Join-Values (($ExchangeGroupRows | Where-Object { $_.Identity -eq $identity }).GroupIdentity)
            EntraGroupMemberships     = Join-Values (($EntraGroupRows | Where-Object { $_.Identity -eq $identity }).GroupDisplayName)
            HasErrors                 = [bool](@($ErrorRows | Where-Object { $_.Identity -eq $identity }).Count -gt 0)
        }
    }
}

function Export-ReportCsvs {
    param(
        [Parameter(Mandatory = $true)] [string]$Folder,
        [Parameter(Mandatory = $true)] [object[]]$Objects,
        [Parameter(Mandatory = $true)] [object[]]$ConsolidatedRows,
        [Parameter(Mandatory = $true)] [object[]]$ProxyRows,
        [Parameter(Mandatory = $true)] [object[]]$FullAccessRows,
        [Parameter(Mandatory = $true)] [object[]]$SendAsRows,
        [Parameter(Mandatory = $true)] [object[]]$ExchangeGroupRows,
        [Parameter(Mandatory = $true)] [object[]]$EntraGroupRows,
        [Parameter(Mandatory = $true)] [object[]]$ErrorRows
    )

    $ConsolidatedRows | Export-Csv -Path (Join-Path $Folder 'Objects-Consolidated.csv') -NoTypeInformation -Encoding UTF8
    $Objects | Export-Csv -Path (Join-Path $Folder 'Objects.csv') -NoTypeInformation -Encoding UTF8
    $ProxyRows | Export-Csv -Path (Join-Path $Folder 'ProxyAddresses.csv') -NoTypeInformation -Encoding UTF8
    $FullAccessRows | Export-Csv -Path (Join-Path $Folder 'FullAccess.csv') -NoTypeInformation -Encoding UTF8
    $SendAsRows | Export-Csv -Path (Join-Path $Folder 'SendAs.csv') -NoTypeInformation -Encoding UTF8
    $ExchangeGroupRows | Export-Csv -Path (Join-Path $Folder 'ExchangeGroupMemberships.csv') -NoTypeInformation -Encoding UTF8
    $EntraGroupRows | Export-Csv -Path (Join-Path $Folder 'EntraGroupMemberships.csv') -NoTypeInformation -Encoding UTF8
    $ErrorRows | Export-Csv -Path (Join-Path $Folder 'Errors.csv') -NoTypeInformation -Encoding UTF8
}
```

- [ ] **Step 3: Add export calls after collection loop**

Append after the collection loop:

```powershell
$consolidatedRows = @(
    New-ConsolidatedRows `
        -Objects $objects `
        -ProxyRows @($proxyRows) `
        -FullAccessRows @($fullAccessRows) `
        -SendAsRows @($sendAsRows) `
        -ExchangeGroupRows @($exchangeGroupMembershipRows) `
        -EntraGroupRows @($entraGroupMembershipRows) `
        -ErrorRows @($script:Errors)
)

Export-ReportCsvs `
    -Folder $resolvedOutputFolder `
    -Objects $objects `
    -ConsolidatedRows $consolidatedRows `
    -ProxyRows @($proxyRows) `
    -FullAccessRows @($fullAccessRows) `
    -SendAsRows @($sendAsRows) `
    -ExchangeGroupRows @($exchangeGroupMembershipRows) `
    -EntraGroupRows @($entraGroupMembershipRows) `
    -ErrorRows @($script:Errors)

Write-Host "CSV exports written to: $resolvedOutputFolder" -ForegroundColor Green
```

- [ ] **Step 4: Verify parser and output file names**

Run:

```bash
pwsh -NoProfile -Command '$tokens=$null; $errors=$null; [System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path "./Export-ExchangeObjectData.ps1"), [ref]$tokens, [ref]$errors) > $null; if ($errors.Count) { $errors | Format-List *; exit 1 }'
```

Expected: PASS with no output.

Run:

```bash
for name in Objects-Consolidated.csv Objects.csv ProxyAddresses.csv FullAccess.csv SendAs.csv ExchangeGroupMemberships.csv EntraGroupMemberships.csv Errors.csv; do grep -n "$name" Export-ExchangeObjectData.ps1 || exit 1; done
```

Expected: PASS and prints matching lines for all names.

- [ ] **Step 5: Commit**

```bash
git add Export-ExchangeObjectData.ps1
git commit -m "Export Exchange object data to CSV"
```

---

### Task 6: Add Static HTML Dashboard Generation

**Files:**
- Modify: `Export-ExchangeObjectData.ps1`

**Interfaces:**
- Consumes: `$isMultiObjectRun`, `$NoDashboard`, `$consolidatedRows`, detail row arrays, `$resolvedOutputFolder`.
- Produces: `Dashboard.html` for multi-object runs only.
- Produces function: `Export-DashboardHtml`.

- [ ] **Step 1: Write failing static checks**

Run:

```bash
grep -n "function Export-DashboardHtml" Export-ExchangeObjectData.ps1
```

Expected: FAIL with no output.

- [ ] **Step 2: Add dashboard generation function before main connection block**

```powershell
function ConvertTo-JsonLiteral {
    param([Parameter(Mandatory = $false)] $Value)
    return ($Value | ConvertTo-Json -Depth 10 -Compress)
}

function Export-DashboardHtml {
    param(
        [Parameter(Mandatory = $true)] [string]$Folder,
        [Parameter(Mandatory = $true)] [object[]]$ConsolidatedRows,
        [Parameter(Mandatory = $true)] [object[]]$ProxyRows,
        [Parameter(Mandatory = $true)] [object[]]$FullAccessRows,
        [Parameter(Mandatory = $true)] [object[]]$SendAsRows,
        [Parameter(Mandatory = $true)] [object[]]$ExchangeGroupRows,
        [Parameter(Mandatory = $true)] [object[]]$EntraGroupRows
    )

    $objectsJson = ConvertTo-JsonLiteral -Value $ConsolidatedRows
    $proxyJson = ConvertTo-JsonLiteral -Value $ProxyRows
    $fullAccessJson = ConvertTo-JsonLiteral -Value $FullAccessRows
    $sendAsJson = ConvertTo-JsonLiteral -Value $SendAsRows
    $exchangeGroupsJson = ConvertTo-JsonLiteral -Value $ExchangeGroupRows
    $entraGroupsJson = ConvertTo-JsonLiteral -Value $EntraGroupRows

    $html = @"
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Exchange Object Data Dashboard</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 24px; color: #1f2933; background: #f7f9fb; }
    h1 { margin-bottom: 8px; }
    .warning { background: #fff3cd; border: 1px solid #ffe69c; padding: 12px; border-radius: 6px; margin: 12px 0; }
    .toolbar { display: flex; gap: 8px; flex-wrap: wrap; margin: 16px 0; }
    input[type="search"] { padding: 8px; min-width: 320px; }
    button { padding: 8px 10px; cursor: pointer; }
    table { border-collapse: collapse; width: 100%; background: white; }
    th, td { border: 1px solid #d9e2ec; padding: 8px; text-align: left; vertical-align: top; }
    th { background: #e6f0ff; position: sticky; top: 0; }
    details { margin-top: 6px; }
    .muted { color: #627d98; }
  </style>
</head>
<body>
  <h1>Exchange Object Data Dashboard</h1>
  <p class="muted">Search, select objects, and export selected data from this local static report.</p>
  <div class="warning">This dashboard may contain sensitive proxy addresses, mailbox permissions, group memberships, and object identifiers. Handle it as sensitive administrative data.</div>
  <div class="toolbar">
    <input id="search" type="search" placeholder="Search objects, proxies, permissions, groups">
    <button onclick="selectVisible(true)">Select visible</button>
    <button onclick="selectVisible(false)">Clear visible</button>
    <button onclick="exportSelected('objects')">Export selected objects</button>
    <button onclick="exportSelected('proxies')">Export selected proxies</button>
    <button onclick="exportSelected('fullaccess')">Export selected FullAccess</button>
    <button onclick="exportSelected('sendas')">Export selected SendAs</button>
    <button onclick="exportSelected('exchangegroups')">Export selected Exchange groups</button>
    <button onclick="exportSelected('entragroups')">Export selected Entra groups</button>
  </div>
  <div id="summary"></div>
  <table>
    <thead><tr><th>Select</th><th>Object</th><th>Type</th><th>Primary SMTP</th><th>Details</th></tr></thead>
    <tbody id="rows"></tbody>
  </table>
<script>
const objects = $objectsJson;
const proxies = $proxyJson;
const fullaccess = $fullAccessJson;
const sendas = $sendAsJson;
const exchangegroups = $exchangeGroupsJson;
const entragroups = $entraGroupsJson;
const selected = new Set();

function textForObject(o) {
  return Object.values(o).join(' ').toLowerCase();
}

function relatedRows(data, identity) {
  return data.filter(r => String(r.Identity) === String(identity));
}

function render() {
  const q = document.getElementById('search').value.toLowerCase();
  const body = document.getElementById('rows');
  body.innerHTML = '';
  const visible = objects.filter(o => textForObject(o).includes(q));
  document.getElementById('summary').textContent = visible.length + ' visible of ' + objects.length + ' objects';
  for (const o of visible) {
    const id = String(o.Identity);
    const tr = document.createElement('tr');
    tr.dataset.identity = id;
    tr.innerHTML = '<td><input type="checkbox" ' + (selected.has(id) ? 'checked' : '') + ' onchange="toggleSelected(this, \'' + escapeHtml(id) + '\')"></td>' +
      '<td><strong>' + escapeHtml(o.DisplayName || '') + '</strong><br><span class="muted">' + escapeHtml(id) + '</span></td>' +
      '<td>' + escapeHtml(o.RecipientTypeDetails || o.RecipientType || '') + '<br>' + escapeHtml(o.MailboxType || '') + '</td>' +
      '<td>' + escapeHtml(o.PrimarySmtpAddress || '') + '</td>' +
      '<td>' + detailsHtml(o) + '</td>';
    body.appendChild(tr);
  }
}

function detailsHtml(o) {
  const id = String(o.Identity);
  return section('Proxies', relatedRows(proxies, id), 'RawProxyAddress') +
    section('FullAccess', relatedRows(fullaccess, id), 'Trustee') +
    section('SendAs', relatedRows(sendas, id), 'Trustee') +
    section('Exchange Groups', relatedRows(exchangegroups, id), 'GroupIdentity') +
    section('Entra Groups', relatedRows(entragroups, id), 'GroupDisplayName');
}

function section(title, rows, field) {
  const values = rows.map(r => r[field]).filter(Boolean).map(escapeHtml).join('<br>');
  return '<details><summary>' + title + ' (' + rows.length + ')</summary>' + (values || '<span class="muted">None</span>') + '</details>';
}

function toggleSelected(cb, id) {
  if (cb.checked) selected.add(id); else selected.delete(id);
}

function selectVisible(value) {
  document.querySelectorAll('#rows tr').forEach(tr => { if (value) selected.add(tr.dataset.identity); else selected.delete(tr.dataset.identity); });
  render();
}

function exportSelected(kind) {
  const ids = selected;
  let data = objects;
  if (kind === 'proxies') data = proxies;
  if (kind === 'fullaccess') data = fullaccess;
  if (kind === 'sendas') data = sendas;
  if (kind === 'exchangegroups') data = exchangegroups;
  if (kind === 'entragroups') data = entragroups;
  const rows = data.filter(r => ids.has(String(r.Identity)));
  downloadCsv(kind + '-selected.csv', rows);
}

function downloadCsv(filename, rows) {
  if (!rows.length) { alert('No selected rows to export.'); return; }
  const headers = Object.keys(rows[0]);
  const csv = [headers.join(',')].concat(rows.map(row => headers.map(h => csvCell(row[h])).join(','))).join('\n');
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = filename;
  link.click();
  URL.revokeObjectURL(link.href);
}

function csvCell(value) {
  const s = value == null ? '' : String(value);
  return '"' + s.replace(/"/g, '""') + '"';
}

function escapeHtml(value) {
  return String(value == null ? '' : value).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

document.getElementById('search').addEventListener('input', render);
render();
</script>
</body>
</html>
"@

    Set-Content -Path (Join-Path $Folder 'Dashboard.html') -Value $html -Encoding UTF8
}
```

- [ ] **Step 3: Add dashboard call after CSV export**

Append after `Write-Host "CSV exports written to: $resolvedOutputFolder" -ForegroundColor Green`:

```powershell
if ($isMultiObjectRun -and -not $NoDashboard) {
    try {
        Export-DashboardHtml `
            -Folder $resolvedOutputFolder `
            -ConsolidatedRows $consolidatedRows `
            -ProxyRows @($proxyRows) `
            -FullAccessRows @($fullAccessRows) `
            -SendAsRows @($sendAsRows) `
            -ExchangeGroupRows @($exchangeGroupMembershipRows) `
            -EntraGroupRows @($entraGroupMembershipRows)
        Write-Host "Dashboard written to: $(Join-Path $resolvedOutputFolder 'Dashboard.html')" -ForegroundColor Green
    }
    catch {
        Add-ExportError -Identity '' -Stage 'Dashboard' -Operation 'Export-DashboardHtml' -Message $_.Exception.Message
        @($script:Errors) | Export-Csv -Path (Join-Path $resolvedOutputFolder 'Errors.csv') -NoTypeInformation -Encoding UTF8
    }
}
```

- [ ] **Step 4: Verify parser and dashboard gating**

Run:

```bash
pwsh -NoProfile -Command '$tokens=$null; $errors=$null; [System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path "./Export-ExchangeObjectData.ps1"), [ref]$tokens, [ref]$errors) > $null; if ($errors.Count) { $errors | Format-List *; exit 1 }'
```

Expected: PASS with no output.

Run:

```bash
grep -n "function Export-DashboardHtml" Export-ExchangeObjectData.ps1 && grep -n "Dashboard.html" Export-ExchangeObjectData.ps1 && grep -n "if (\$isMultiObjectRun -and -not \$NoDashboard)" Export-ExchangeObjectData.ps1
```

Expected: PASS and prints matching lines.

- [ ] **Step 5: Commit**

```bash
git add Export-ExchangeObjectData.ps1
git commit -m "Add Exchange object dashboard export"
```

---

### Task 7: Add Cleanup, Final Summary, and README Documentation

**Files:**
- Modify: `Export-ExchangeObjectData.ps1`
- Modify: `README.md`

**Interfaces:**
- Consumes all final script outputs.
- Produces README table entry and documentation section.

- [ ] **Step 1: Write failing README static check**

Run:

```bash
grep -n "Export-ExchangeObjectData.ps1" README.md
```

Expected: FAIL with no output if README has not yet been updated.

- [ ] **Step 2: Add disconnect/final summary to script**

Append at the end of `Export-ExchangeObjectData.ps1`:

```powershell
Write-Host "Objects exported: $($objects.Count)" -ForegroundColor Green
Write-Host "Errors logged: $($script:Errors.Count)" -ForegroundColor Yellow

try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch { }
if ($graphConnected) {
    try { Disconnect-MgGraph | Out-Null } catch { }
}
```

- [ ] **Step 3: Update README scripts table**

Add this row to the table under `## Scripts`:

```markdown
| `Export-ExchangeObjectData.ps1` | PowerShell | Exports Exchange Online recipient proxies, permissions, Exchange/Entra group memberships, CSV reports, and a searchable HTML dashboard. |
```

- [ ] **Step 4: Add README section before `## Get-EntraAdminAccounts.ps1`**

```markdown
## Export-ExchangeObjectData.ps1

Exports Exchange Online recipient data for mailboxes by default, or for selected recipient object types. The script can process all matching recipients, multiple identities from a CSV, or a single identity. It exports proxy addresses, FullAccess permissions, SendAs permissions, Exchange group memberships, Entra ID group memberships, normalized CSV files, a consolidated CSV, and a searchable HTML dashboard for multi-object runs.

### Requirements

- PowerShell.
- ExchangeOnlineManagement module: `Install-Module ExchangeOnlineManagement`.
- Microsoft.Graph modules for Entra group membership enrichment.
- Exchange Online permission to read recipients and permissions.
- Microsoft Graph delegated scopes for Entra group memberships:
  - `User.Read.All`
  - `Group.Read.All`
  - `Directory.Read.All`

Graph permissions may require admin consent depending on tenant policy.

### Parameters

| Parameter | Type | Default | Description |
| --- | --- | --- | --- |
| `-All` | Switch | Off | Exports all recipients matching `-RecipientType`. |
| `-InputCsv` | String | None | CSV file containing object identities to export. |
| `-Identity` | String | None | Single Exchange-resolvable identity to export. |
| `-IdentityColumn` | String | `Identity` | CSV column containing identities. |
| `-RecipientType` | String | `Mailbox` | Object type filter: `Mailbox`, `Group`, `MailUser`, `MailContact`, or `All`. |
| `-OutputFolder` | String | Timestamped folder | Folder where CSVs and dashboard are written. |
| `-SkipGraph` | Switch | Off | Skips Microsoft Graph connection and Entra group membership export. |
| `-NoDashboard` | Switch | Off | Suppresses `Dashboard.html` generation for multi-object runs. |

### Usage

```powershell
./Export-ExchangeObjectData.ps1 -All
```

```powershell
./Export-ExchangeObjectData.ps1 -All -RecipientType All
```

```powershell
./Export-ExchangeObjectData.ps1 -InputCsv ./objects.csv -IdentityColumn Identity -RecipientType Group
```

```powershell
./Export-ExchangeObjectData.ps1 -Identity user@contoso.com
```

### Output Files

- `Objects-Consolidated.csv`
- `Objects.csv`
- `ProxyAddresses.csv`
- `FullAccess.csv`
- `SendAs.csv`
- `ExchangeGroupMemberships.csv`
- `EntraGroupMemberships.csv`
- `Errors.csv`
- `Dashboard.html` for `-All` and `-InputCsv` runs unless `-NoDashboard` is used

### Dashboard

The dashboard is a static local HTML file. It supports search, checkbox selection, expandable object details, and browser-side CSV export of selected objects and selected detail rows.

### Notes

- The default recipient type is `Mailbox`.
- Mailbox exports include `MailboxType`, such as `UserMailbox`, `SharedMailbox`, `RoomMailbox`, or `EquipmentMailbox`.
- The script is read-only.
- If a lookup fails for an object or permission type, the script continues and records the failure in `Errors.csv`.
- Export files and the dashboard may contain sensitive proxy addresses, mailbox permissions, group memberships, and object identifiers. Handle them as sensitive administrative data.
```

- [ ] **Step 5: Verify parser and README docs**

Run:

```bash
pwsh -NoProfile -Command '$tokens=$null; $errors=$null; [System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path "./Export-ExchangeObjectData.ps1"), [ref]$tokens, [ref]$errors) > $null; if ($errors.Count) { $errors | Format-List *; exit 1 }'
```

Expected: PASS with no output.

Run:

```bash
grep -n "Export-ExchangeObjectData.ps1" README.md && grep -n "MailboxType" README.md && grep -n "Dashboard.html" README.md
```

Expected: PASS and prints matching lines.

- [ ] **Step 6: Commit**

```bash
git add Export-ExchangeObjectData.ps1 README.md
git commit -m "Document Exchange object data export"
```

---

### Task 8: Final Verification and Push

**Files:**
- Verify: `Export-ExchangeObjectData.ps1`
- Verify: `README.md`

**Interfaces:**
- Consumes final implementation from Tasks 1-7.
- Produces pushed GitHub commits.

- [ ] **Step 1: Run full parser check**

```bash
pwsh -NoProfile -Command '$tokens=$null; $errors=$null; [System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path "./Export-ExchangeObjectData.ps1"), [ref]$tokens, [ref]$errors) > $null; if ($errors.Count) { $errors | Format-List *; exit 1 }'
```

Expected: PASS with no output.

- [ ] **Step 2: Run required content checks**

```bash
for pattern in "Objects-Consolidated.csv" "ProxyAddresses.csv" "FullAccess.csv" "SendAs.csv" "ExchangeGroupMemberships.csv" "EntraGroupMemberships.csv" "Errors.csv" "Dashboard.html" "RecipientTypeFilter" "MailboxType"; do grep -n "$pattern" Export-ExchangeObjectData.ps1 || exit 1; done
```

Expected: PASS and prints matching lines for every required script pattern.

```bash
for pattern in "Export-ExchangeObjectData.ps1" "RecipientType" "MailboxType" "Dashboard.html" "sensitive"; do grep -n "$pattern" README.md || exit 1; done
```

Expected: PASS and prints matching lines for every required README pattern.

- [ ] **Step 3: Inspect working tree and recent history**

```bash
git status --short --branch
git log --oneline -10
```

Expected: working tree clean and implementation commits present on `main`.

- [ ] **Step 4: Push**

```bash
git push origin main
```

Expected: push succeeds.

- [ ] **Step 5: Verify remote commit**

```bash
git status --short --branch
```

Expected: `## main...origin/main` with no modified files.
