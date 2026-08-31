<#
.SYNOPSIS
    Read-only tenant migration discovery for on-premises Active Directory and
    Microsoft 365, with correlation into a single Excel workbook.

.DESCRIPTION
    Get-TenantMigrationDiscovery collects migration-relevant configuration and
    dependency data from two broad sources:

      1. On-premises Active Directory / Group Policy
         - Organizational Units
         - Users and Exchange-related AD attributes
         - Computers / servers and Kerberos/delegation attributes
         - Contacts
         - Groups and recursively expanded membership
         - GPO inventory, links, and optional setting-level detail

      2. Microsoft 365 / Entra ID
         - Entra users, groups, memberships and Conditional Access
         - Licensing
         - App registrations, service principals and consent grants
         - Exchange Online mailboxes, archives, permissions and distribution lists
         - Intune devices and policy state
         - OneDrive
         - SharePoint Online
         - Teams
         - Purview / Security & Compliance

    The script is designed for discovery only. It does not intentionally modify
    source tenant, Active Directory, Exchange, Intune, SharePoint, Teams, or
    Purview configuration.

    COLLECTION MODES
    ----------------
    -IncludeOnPrem
        Runs the on-premises Active Directory and GPO collectors.

    -IncludeCloud
        Runs the selected Microsoft 365 cloud collectors.

    Both switches can be used together on a suitable Windows/domain-connected
    workstation, or the two halves can be run separately and merged into the same
    workbook with -InputWorkbook.

    MERGE / OVERWRITE BEHAVIOUR
    ---------------------------
    When -InputWorkbook is supplied, that .xlsx file is used as both the discovery
    source and the destination workbook.

    For each worksheet produced by the CURRENT run:
      - New rows replace the existing worksheet; they are not appended.
      - If a collector successfully produces zero rows, an older worksheet with
        the same name is removed so stale data is not left behind.
      - Worksheets not produced by the current run remain untouched.
      - 00_Summary is rebuilt from the final workbook contents.

    Therefore a rerun acts as a refresh of the selected workloads, not as a
    historical append operation. Make a copy of an input workbook first if you
    need an immutable snapshot.

    SCOPED MIGRATION-WAVE MODE
    --------------------------
    -ScopeToInputUsers limits user/device-oriented cloud collection to identities
    already present in the input workbook. The script builds matching candidates
    from UPN, mail, primary SMTP, secondary SMTP aliases, otherMails and available
    cloud aliases. This mode requires -InputWorkbook.

    Scope safety is fail-closed: if scoped mode is requested but no usable users
    can be derived from the workbook, the script stops rather than accidentally
    running a tenant-wide collection.

    PLATFORM NOTES
    --------------
    Windows:
      - Required for the on-prem ActiveDirectory / GroupPolicy collectors.
      - Microsoft.Online.SharePoint.PowerShell (SPO module) is Windows-oriented.

    macOS / Linux / PowerShell 7:
      - Cloud workloads can run if their modules support the platform.
      - Use -SharePointModule PnP with -PnPClientId for SharePoint collection.

.PARAMETER Scope
    Controls on-premises AD search scope.

    Domain (default)
        Enumerate the entire directory search base.

    OU
        Restrict on-premises collection to -OUPath and descendants.

.PARAMETER OUPath
    Distinguished name of the OU used when -Scope OU.

    Example:
        "OU=Finance,OU=Users,DC=contoso,DC=com"

.PARAMETER Server
    Optional domain controller / directory server to use for AD queries.

    Example:
        dc01.contoso.com

.PARAMETER IncludeOnPrem
    Enables the on-premises Active Directory and Group Policy collectors.

.PARAMETER IncludeCloud
    Enables Microsoft 365 cloud discovery.

.PARAMETER SkipGpoSettings
    Skips the setting-by-setting GPO_Settings extraction while still collecting
    GPO inventory and links. Useful when you need a faster inventory pass.

.PARAMETER SkipIntuneDevicePolicies
    Skips per-device Intune configuration/compliance state calls while retaining
    managed-device inventory and tenant configuration-profile inventory.

.PARAMETER GpoLinksOnly
    Fastest GPO mode. Reads GPO-to-container links directly from gPLink and avoids
    Get-GPOReport. Produces GPOs and GPO_Links only; detailed settings and enabled
    client-side extension summaries are omitted.

.PARAMETER CloudWorkload
    One or more cloud collectors to run.

    Valid values:
        Entra
        Licensing
        Apps
        Exchange
        SharePoint
        OneDrive
        Teams
        Intune
        Purview

    Default: all values above.

.PARAMETER OutputPath
    Folder used for a newly created workbook, logs, and temporary files.
    Defaults to the script directory.

    When -InputWorkbook is supplied, the workbook itself is updated in place, but
    logs still use OutputPath.

.PARAMETER SharePointAdminUrl
    SharePoint tenant admin endpoint.

    Example:
        https://contoso-admin.sharepoint.com

    The script also normalizes the common copy/paste Markdown form:
        [https://contoso-admin.sharepoint.com](https://contoso-admin.sharepoint.com)

.PARAMETER SharePointModule
    SharePoint provider.

    SPO (default)
        Microsoft.Online.SharePoint.PowerShell. Best suited to Windows.

    PnP
        PnP.PowerShell. Recommended for PowerShell 7/macOS/Linux. Requires
        -PnPClientId.

.PARAMETER PnPClientId
    Entra application/client ID used by PnP.PowerShell authentication.
    Required when -SharePointModule PnP.

.PARAMETER LoginMethod
    Cloud sign-in preference.

    Interactive (default)
        Uses interactive modern authentication. EXO/Purview use -DisableWAM when
        the installed module supports it.

    DeviceCode
        Uses device-code authentication for workloads whose installed modules
        expose it. Current Security & Compliance PowerShell does not expose a
        Connect-IPPSSession -Device switch, so Purview falls back to interactive
        modern authentication.

.PARAMETER InputWorkbook
    Existing .xlsx discovery workbook to merge into and refresh.

    Important:
        The file is modified in place. Worksheets collected in the current run are
        replaced; they are not appended.

.PARAMETER ScopeToInputUsers
    Restricts user/device-oriented cloud collection to the migration identities
    found in -InputWorkbook.

    Requires -InputWorkbook.

.OUTPUTS
    An Excel workbook containing whichever worksheets were produced by the chosen
    workloads, plus 00_Summary and Dependency_Map.

    ON-PREM WORKSHEETS
      OrganizationalUnits
      Users
      Devices
      Contacts
      Groups
      GroupMembership
      GPOs
      GPO_Links
      GPO_Settings               (unless skipped / links-only mode)

    ENTRA / LICENSING / APP WORKSHEETS
      Entra_Users
      Entra_Groups
      Entra_GroupMembership
      CA_Policies
      CA_NamedLocations
      License_Summary             (tenant-wide SKU capacity + reconciliation)
      License_UserDetail          (wave-scoped when requested; direct/group source)
      License_GroupAssignments    (tenant-wide group -> SKU configuration)
      Entra_AppRegistrations
      Entra_EnterpriseApps
      Entra_OAuthGrants
      Entra_AppRoleGrants
      Entra_UserAppAssignments

    EXCHANGE WORKSHEETS
      Exchange_Mailboxes
      Exchange_Archives
      Exchange_MailboxPermissions
      Exchange_DistLists
      Exchange_DistListMembers

    INTUNE / COLLABORATION WORKSHEETS
      Intune_Devices
      Intune_DevicePolicies
      Intune_ConfigProfiles
      OneDrive
      SharePoint_Sites
      Teams
      Teams_Channels
      Teams_Members

    PURVIEW WORKSHEETS
      Purview_SensitivityLabels
      Purview_LabelPolicies
      Purview_RetentionLabels
      Purview_RetentionPolicies
      Purview_DLPPolicies
      Purview_eDiscoveryHolds

    CORRELATION
      Dependency_Map
      00_Summary

.EXAMPLE
    # Show the built-in documentation without running discovery.
    Get-Help ./Get-TenantMigrationDiscovery.fixed9-documented.ps1 -Full

.EXAMPLE
    # Full on-premises discovery for the entire domain.
    ./Get-TenantMigrationDiscovery.fixed9-documented.ps1 `
        -IncludeOnPrem

.EXAMPLE
    # Restrict on-premises discovery to one OU tree.
    ./Get-TenantMigrationDiscovery.fixed9-documented.ps1 `
        -IncludeOnPrem `
        -Scope OU `
        -OUPath "OU=Finance,OU=Users,DC=contoso,DC=com"

.EXAMPLE
    # Fast on-premises discovery when you only need GPO inventory and links.
    ./Get-TenantMigrationDiscovery.fixed9-documented.ps1 `
        -IncludeOnPrem `
        -GpoLinksOnly

.EXAMPLE
    # Cloud discovery with every cloud workload except SharePoint.
    ./Get-TenantMigrationDiscovery.fixed9-documented.ps1 `
        -IncludeCloud `
        -CloudWorkload Entra,Licensing,Apps,Exchange,OneDrive,Teams,Intune,Purview

.EXAMPLE
    # Exchange + Entra only. Useful when validating mailbox discovery quickly.
    ./Get-TenantMigrationDiscovery.fixed9-documented.ps1 `
        -IncludeCloud `
        -CloudWorkload Entra,Exchange

.EXAMPLE
    # First pass: collect on-premises data and create a new workbook.
    ./Get-TenantMigrationDiscovery.fixed9-documented.ps1 `
        -IncludeOnPrem `
        -OutputPath "C:\MigrationDiscovery"

.EXAMPLE
    # Second pass: refresh cloud data into the existing on-prem workbook.
    # The supplied workbook is modified in place.
    ./Get-TenantMigrationDiscovery.fixed9-documented.ps1 `
        -IncludeCloud `
        -InputWorkbook "C:\MigrationDiscovery\S32-MigrationDiscovery-20260825-073000.xlsx" `
        -SharePointAdminUrl "https://contoso-admin.sharepoint.com"

.EXAMPLE
    # Migration-wave mode: collect cloud dependencies only for users/devices
    # represented in the existing workbook.
    ./Get-TenantMigrationDiscovery.fixed9-documented.ps1 `
        -IncludeCloud `
        -ScopeToInputUsers `
        -InputWorkbook "C:\MigrationDiscovery\Wave01.xlsx" `
        -SharePointAdminUrl "https://contoso-admin.sharepoint.com"

.EXAMPLE
    # macOS / PowerShell 7: use PnP.PowerShell for SharePoint.
    ./Get-TenantMigrationDiscovery.fixed9-documented.ps1 `
        -IncludeCloud `
        -ScopeToInputUsers `
        -InputWorkbook "/Users/me/MigrationDiscovery/Wave01.xlsx" `
        -SharePointAdminUrl "https://contoso-admin.sharepoint.com" `
        -SharePointModule PnP `
        -PnPClientId "00000000-0000-0000-0000-000000000000"

.EXAMPLE
    # Console/headless-style sign-in where supported by the installed modules.
    ./Get-TenantMigrationDiscovery.fixed9-documented.ps1 `
        -IncludeCloud `
        -LoginMethod DeviceCode `
        -CloudWorkload Entra,Licensing,Apps,Exchange,Teams,Intune

.EXAMPLE
    # Faster Intune pass: inventory devices and profiles, but skip per-device
    # config/compliance state calls.
    ./Get-TenantMigrationDiscovery.fixed9-documented.ps1 `
        -IncludeCloud `
        -CloudWorkload Entra,Intune `
        -SkipIntuneDevicePolicies

.EXAMPLE
    # Combined run from a Windows domain-connected management workstation with
    # the required on-prem and cloud modules installed.
    ./Get-TenantMigrationDiscovery.fixed9-documented.ps1 `
        -IncludeOnPrem `
        -IncludeCloud `
        -SharePointAdminUrl "https://contoso-admin.sharepoint.com"

.NOTES
    SOURCE SAFETY
      Discovery is intended to be read-only against AD and Microsoft 365.

    WORKBOOK SAFETY
      -InputWorkbook is updated in place. If you need historical snapshots, copy
      the workbook before rerunning the script.

    REQUIRED MODULES
      On-premises:
        ActiveDirectory
        GroupPolicy

      Cloud core:
        Microsoft.Graph.Authentication
        Microsoft.Graph.Users
        Microsoft.Graph.Groups
        Microsoft.Graph.Identity.DirectoryManagement
        Microsoft.Graph.Identity.SignIns
        Microsoft.Graph.Applications
        Microsoft.Graph.DeviceManagement
        ExchangeOnlineManagement
        MicrosoftTeams

      SharePoint:
        Microsoft.Online.SharePoint.PowerShell
          OR
        PnP.PowerShell

      Excel output:
        ImportExcel

    AUTHORIZATION
      The script cannot bypass tenant permissions. A collector can only return
      properties that the signed-in account and tenant policies allow it to read.
      Permission/API failures are logged and, where possible, isolated so other
      workloads can continue.

    LOGGING
      A timestamped S32-MigrationDiscovery-*.log file is written to OutputPath.
      INFO, WARN, ERROR and detailed progress entries are included.

    Author: S32 migration workstream
#>

[CmdletBinding()]
param(
    [ValidateSet('Domain', 'OU')]
    [string]$Scope = 'Domain',

    [string]$OUPath,

    [string]$Server,

    [switch]$IncludeOnPrem,

    [switch]$IncludeCloud,

    # Skip the deep per-setting GPO extraction (GPO_Settings sheet). GPOs, links,
    # and enabled-extension summary are still produced. Use for a faster pass when
    # you only need the GPO inventory, not the setting-level Intune mapping.
    [switch]$SkipGpoSettings,

    # Skip the per-device Intune policy pass (config + compliance policy state per
    # device = 2 Graph calls each). Device inventory incl. MACs still collected.
    [switch]$SkipIntuneDevicePolicies,

    # Fastest GPO mode: read GPO->OU links directly from each container's gPLink
    # attribute -- NO Get-GPOReport at all, so no timeout risk (near-instant).
    # Produces GPOs + GPO_Links only (no settings, no enabled-extension summary).
    [switch]$GpoLinksOnly,

    [ValidateSet('Entra', 'Licensing', 'Apps', 'Exchange', 'SharePoint', 'OneDrive', 'Teams', 'Intune', 'Purview')]
    [string[]]$CloudWorkload = @('Entra', 'Licensing', 'Apps', 'Exchange', 'SharePoint', 'OneDrive', 'Teams', 'Intune', 'Purview'),

    [string]$OutputPath = $PSScriptRoot,

    [string]$SharePointAdminUrl,

    # SharePoint/OneDrive provider. 'SPO' = Microsoft.Online.SharePoint.PowerShell
    # (interactive browser only; needs Windows PowerShell 5.1). 'PnP' =
    # PnP.PowerShell, which supports device-code and PS7 -- use it on WAM-blocked
    # hosts. PnP requires -PnPClientId (an Entra app registration; the old shared
    # PnP app was retired).
    [ValidateSet('SPO', 'PnP')]
    [string]$SharePointModule = 'SPO',

    # Entra app (client) ID for PnP.PowerShell sign-in. Required when
    # -SharePointModule PnP. A delegated app with Sites.FullControl.All (or read)
    # granted; if you lack consent rights, have an admin register/consent one.
    [string]$PnPClientId,

    # Cloud sign-in method:
    #   Interactive (default) = browser/WAM sign-in. EXO/IPPS use -DisableWAM when
    #       the installed module exposes it; Graph auto-falls back to device code if
    #       the WAM broker can't get a window handle.
    #   DeviceCode            = device-code flow for Graph/EXO/Teams/PnP where the
    #       installed module exposes it. Current Connect-IPPSSession has no -Device
    #       switch, so Purview falls back to interactive modern authentication.
    [ValidateSet('Interactive', 'DeviceCode')]
    [string]$LoginMethod = 'Interactive',

    # Existing workbook to read as a SOURCE and merge INTO. Use this to run the
    # halves separately and land them in one correlated file:
    #   Pass 1 (DC):     -IncludeOnPrem                       -> new workbook
    #   Pass 2 (laptop): -IncludeCloud -InputWorkbook <pass1> -> merges + correlates
    # Works in either order (on-prem first or cloud first). Prior sheets are
    # read back to hydrate the identity index so Dependency_Map spans both halves.
    [string]$InputWorkbook,

    # Restrict cloud pulls to ONLY what's in the -InputWorkbook (your migration
    # wave), instead of the whole tenant:
    #   - per-user data (Entra users, mailboxes, OneDrive, per-user licenses,
    #     Teams membership) is limited to the workbook's Users/Entra_Users.
    #   - Intune devices are limited to the workbook's Devices sheet (by device
    #     name); if there's no Devices sheet, they fall back to the user scope.
    # Tenant-level datasets (group catalog, CA, apps, SKU totals, site/Teams
    # inventory, Purview) are still collected in full as dependency context.
    # Requires -InputWorkbook.
    [switch]$ScopeToInputUsers
)

# =============================================================================
# OPERATOR QUICK REFERENCE
# =============================================================================
# 1. This script is discovery/read-only against the source environments.
# 2. -InputWorkbook is different: that Excel file IS updated in place.
# 3. Existing sheets collected in the current run are replaced, not appended.
# 4. Use -ScopeToInputUsers for migration-wave cloud collection.
# 5. Use -CloudWorkload to reduce runtime while testing a specific workload.
# 6. Use the timestamped log file to identify permission/API-specific failures.
# 7. Run `Get-Help <script> -Full` for examples and the full worksheet catalog.
# =============================================================================

#region ---------------------------------------------------------------- Setup

if (-not $IncludeOnPrem -and -not $IncludeCloud) {
    throw "Specify -IncludeOnPrem and/or -IncludeCloud. Nothing to collect otherwise."
}
if ($Scope -eq 'OU' -and [string]::IsNullOrWhiteSpace($OUPath)) {
    throw "-Scope OU requires -OUPath (a distinguished name)."
}
# Internal flags consumed by cloud connect blocks.
$UseDeviceCode = ($LoginMethod -eq 'DeviceCode')
$RunningOnWindows = (($PSVersionTable.PSEdition -eq 'Desktop') -or ($PSVersionTable.Platform -eq 'Win32NT'))
# Set true only after Connect-MgGraph succeeds; the Graph helpers no-op when false
# so a failed sign-in doesn't trigger thousands of pointless batch retries.
$GraphConnected = $false
if (-not (Test-Path $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$Stamp   = Get-Date -Format 'yyyyMMdd-HHmmss'
$LogPath = Join-Path $OutputPath "S32-MigrationDiscovery-$Stamp.log"

if ($PSBoundParameters.ContainsKey('InputWorkbook') -and -not [string]::IsNullOrWhiteSpace($InputWorkbook)) {
    if (-not (Test-Path $InputWorkbook)) { throw "-InputWorkbook not found: $InputWorkbook" }
    if ([IO.Path]::GetExtension($InputWorkbook) -ne '.xlsx') { throw "-InputWorkbook must be an .xlsx file." }
    # Merge INTO the supplied workbook; new sheets append, existing sheets refresh.
    $WorkbookPath = (Resolve-Path $InputWorkbook).Path
    $Merging = $true
}
else {
    $WorkbookPath = Join-Path $OutputPath "S32-MigrationDiscovery-$Stamp.xlsx"
    $Merging = $false
}

# FUNCTION: Write-Log
# Purpose : Write a timestamped INFO/WARN/ERROR message to both the console and
#           the discovery log file. Centralizing log output keeps every workload
#           consistent and makes long-running failures easier to diagnose.
function Write-Log {
    param([string]$Message, [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO')
    $line = "{0} [{1}] {2}" -f (Get-Date -Format 's'), $Level, $Message
    switch ($Level) {
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'ERROR' { Write-Host $line -ForegroundColor Red }
        default { Write-Host $line -ForegroundColor Cyan }
    }
    Add-Content -Path $LogPath -Value $line
}

# FUNCTION: Write-Verbose2
# Purpose : Emit always-visible low-priority progress/detail messages. This is
#           intentionally separate from PowerShell -Verbose so operators can see
#           movement during long tenant scans without enabling debug verbosity.
function Write-Verbose2 {
    # Grey, always-on detail line so the console keeps moving during long steps.
    param([string]$Message)
    Write-Host ("  {0} {1}" -f (Get-Date -Format 'HH:mm:ss'), $Message) -ForegroundColor DarkGray
    Add-Content -Path $LogPath -Value ("{0} [DTL ] {1}" -f (Get-Date -Format 's'), $Message)
}

# Chat/copy-paste sometimes turns a URL into Markdown: [https://...](https://...).
# Normalize that form before any SharePoint cmdlet sees it.
if ($SharePointAdminUrl -and $SharePointAdminUrl -match '^\[(https://[^\]]+)\]\((https://[^\)]+)\)$') {
    $SharePointAdminUrl = $Matches[2]
    Write-Log "Normalized Markdown-formatted -SharePointAdminUrl to '$SharePointAdminUrl'." 'WARN'
}

# FUNCTION: Write-Activity
# Purpose : Maintain a reusable Write-Progress bar. Supports known totals and
#           indeterminate/running-count operations.
function Write-Activity {
    # Draw/refresh a progress bar. $Total 0 = indeterminate (show running count).
    param(
        [Parameter(Mandatory)][string]$Activity,
        [int]$Current = 0, [int]$Total = 0, [string]$Item = '', [int]$Id = 1
    )
    if ($Total -gt 0) {
        $pct = [int](([math]::Min($Current, $Total) / $Total) * 100)
        Write-Progress -Id $Id -Activity $Activity -Status "$Current of $Total  $Item" -PercentComplete $pct
    }
    else {
        Write-Progress -Id $Id -Activity $Activity -Status "$Current processed  $Item"
    }
}
# FUNCTION: Complete-Activity
# Purpose : Mark a Write-Progress activity as completed and remove its progress bar.
function Complete-Activity { param([string]$Activity = ' ', [int]$Id = 1) Write-Progress -Id $Id -Activity $Activity -Completed }

# --- Memory safety for large scoped runs --------------------------------------
# The typed Microsoft.Graph Get-Mg* cmdlets leak memory across many calls, so the
# per-user loops below use Invoke-MgGraphRequest (raw REST -> hashtables) instead,
# and force a collection every few hundred iterations. Also cap PS history growth.
$MaximumHistoryCount = 64

# GET a Graph resource as a hashtable; follows @odata.nextLink for collections.
# Returns an array (single-object GETs come back as a 1-element array).
# FUNCTION: Invoke-GraphGetAll
# Purpose : Perform a Microsoft Graph GET using Invoke-MgGraphRequest, follow
#           @odata.nextLink until the collection is exhausted, and return simple
#           hashtable-backed objects. Retries transient failures before giving up.
# Why     : Raw REST is substantially more memory-efficient than thousands of typed
#           Get-Mg* calls in large scoped migrations.
function Invoke-GraphGetAll {
    param([string]$Uri)
    if (-not $GraphConnected) { return , @() }
    $items = [System.Collections.Generic.List[object]]::new()
    $next = $Uri
    while ($next) {
        $resp = $null
        for ($attempt = 1; $attempt -le 4; $attempt++) {
            try {
                $resp = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType Hashtable -ErrorAction Stop
                break
            }
            catch {
                if ($attempt -ge 4) { throw }
                Start-Sleep -Seconds ([math]::Min(20, 5 * $attempt))
            }
        }
        if ($resp -is [System.Collections.IDictionary] -and $resp.ContainsKey('value')) {
            foreach ($v in $resp['value']) { $items.Add($v) }
            $next = $resp['@odata.nextLink']
        }
        else { $items.Add($resp); $next = $null }
    }
    , $items.ToArray()
}

# Reclaim memory during long per-user loops (mitigates the Graph SDK leak).
# FUNCTION: Invoke-LoopGc
# Purpose : Periodically request garbage collection during very large loops and
#           trim avoidable process/history growth. This is a memory-safety helper,
#           not part of discovery logic.
function Invoke-LoopGc {
    param([int]$Index, [int]$Every = 250)
    if ($Index -gt 0 -and ($Index % $Every) -eq 0) {
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
    }
}

# Batch many GET requests via Graph /$batch (20 per HTTP call) instead of one call
# each -- essential at scale (thousands of users): ~20x fewer calls, far less SDK
# memory churn, much faster. $Requests = @(@{Key=..; Url="/users/.."}, ...); returns
# a hashtable Key -> response body (only 2xx). Retries 429/5xx sub-responses.
# FUNCTION: Invoke-GraphBatch
# Purpose : Execute Microsoft Graph JSON batch GET requests in chunks (maximum 20),
#           retry throttled/transient subrequests, surface non-retryable HTTP errors,
#           and follow subrequest @odata.nextLink pagination where present.
# Safety  : The body is explicitly converted to JSON to avoid PowerShell adapter
#           self-reference serialization failures seen with nested PSObjects.
function Invoke-GraphBatch {
    param(
        [object[]]$Requests,
        [string]$Activity = 'Graph batch',
        [ValidateRange(1,20)][int]$ChunkSize = 20,
        [int[]]$QuietStatuses = @()
    )
    $out = @{}
    if (-not $GraphConnected) { return $out }

    # Force a real object array once. Avoid passing PowerShell adapter/indexer
    # properties from Hashtable/PSObject wrappers into the Graph SDK serializer.
    $requestArray = @($Requests)
    $total = $requestArray.Count
    if ($total -eq 0) { return $out }

    # Aggregate non-success statuses so 400/403/404 responses cannot silently
    # turn into missing discovery data. Callers can explicitly quiet expected
    # statuses (OneDrive uses 404 for users whose drive is not provisioned).
    $errorCounts  = @{}
    $errorSamples = @{}

    for ($i = 0; $i -lt $total; $i += $ChunkSize) {
        $end = [math]::Min($i + $ChunkSize - 1, $total - 1)
        Write-Activity -Activity $Activity -Current ($end + 1) -Total $total -Item "batch $([int]($i/$ChunkSize)+1)"

        $idMap = @{}
        $reqList = [System.Collections.Generic.List[object]]::new()
        $n = 0

        foreach ($r in $requestArray[$i..$end]) {
            if ($r -is [System.Collections.IDictionary]) {
                $requestKey = $r['Key']
                $requestUrl = $r['Url']
            }
            else {
                $requestKey = $r.Key
                $requestUrl = $r.Url
            }

            if ($null -eq $requestKey -or [string]::IsNullOrWhiteSpace($requestKey.ToString()) -or -not $requestUrl) {
                Write-Log "  Skipping malformed batch request (Key='$requestKey', Url='$requestUrl')." 'WARN'
                continue
            }

            $sid = [string]$n
            $idMap[$sid] = $requestKey.ToString()
            $reqList.Add([ordered]@{
                id     = $sid
                method = 'GET'
                url    = [string]$requestUrl
            })
            $n++
        }

        if ($reqList.Count -eq 0) { continue }

        $attempt = 0
        $go = $true
        while ($go -and $attempt -lt 4) {
            $attempt++
            $go = $false
            try {
                $batchBody = @{
                    requests = @($reqList.ToArray())
                } | ConvertTo-Json -Depth 8 -Compress

                $resp = Invoke-MgGraphRequest `
                    -Method POST `
                    -Uri 'v1.0/$batch' `
                    -Body $batchBody `
                    -ContentType 'application/json' `
                    -OutputType Hashtable `
                    -ErrorAction Stop

                $throttled = [System.Collections.Generic.List[string]]::new()
                foreach ($sub in @($resp['responses'])) {
                    $st = [int]$sub['status']
                    $subId = [string]$sub['id']
                    if ($st -ge 200 -and $st -lt 300) {
                        if ($idMap.ContainsKey($subId)) {
                            $body = $sub['body']

                            # A successful batch subresponse can itself be paged.
                            # Follow @odata.nextLink so memberOf/app assignments/policy
                            # state are not silently truncated at the first page.
                            if ($body -is [System.Collections.IDictionary] -and $body.ContainsKey('value') -and $body['@odata.nextLink']) {
                                $combined = [System.Collections.Generic.List[object]]::new()
                                foreach ($v in @($body['value'])) { $combined.Add($v) }
                                $nextLink = $body['@odata.nextLink']
                                $pageGuard = 0
                                while ($nextLink -and $pageGuard -lt 100) {
                                    $pageGuard++
                                    try {
                                        $page = Invoke-MgGraphRequest -Method GET -Uri $nextLink -OutputType Hashtable -ErrorAction Stop
                                        foreach ($v in @($page['value'])) { $combined.Add($v) }
                                        $nextLink = $page['@odata.nextLink']
                                    }
                                    catch {
                                        Write-Log "  '$Activity' pagination failed for key '$($idMap[$subId])': $($_.Exception.Message)" 'WARN'
                                        $nextLink = $null
                                    }
                                }
                                $body['value'] = $combined.ToArray()
                                if ($body.ContainsKey('@odata.nextLink')) { $body.Remove('@odata.nextLink') }
                            }

                            $out[$idMap[$subId]] = $body
                        }
                    }
                    elseif ($st -eq 429 -or $st -ge 500) {
                        if ($attempt -lt 4) {
                            $throttled.Add($subId)
                        }
                        elseif ($QuietStatuses -notcontains $st) {
                            $err = $sub['body']; $code = 'retry-exhausted'; $msg = ''
                            if ($err -is [System.Collections.IDictionary] -and $err['error']) {
                                $e = $err['error']
                                if ($e -is [System.Collections.IDictionary]) { if ($e['code']) { $code = $e['code'] }; if ($e['message']) { $msg = $e['message'] } }
                            }
                            $bucket = "$st|$code"
                            if (-not $errorCounts.ContainsKey($bucket)) { $errorCounts[$bucket] = 0; $errorSamples[$bucket] = "$($idMap[$subId]): $msg" }
                            $errorCounts[$bucket]++
                        }
                    }
                    elseif ($QuietStatuses -notcontains $st) {
                        $err = $sub['body']; $code = 'GraphError'; $msg = ''
                        if ($err -is [System.Collections.IDictionary] -and $err['error']) {
                            $e = $err['error']
                            if ($e -is [System.Collections.IDictionary]) { if ($e['code']) { $code = $e['code'] }; if ($e['message']) { $msg = $e['message'] } }
                        }
                        $bucket = "$st|$code"
                        if (-not $errorCounts.ContainsKey($bucket)) { $errorCounts[$bucket] = 0; $errorSamples[$bucket] = "$($idMap[$subId]): $msg" }
                        $errorCounts[$bucket]++
                    }
                }

                if ($throttled.Count -gt 0 -and $attempt -lt 4) {
                    Start-Sleep -Seconds ([math]::Min(20, 5 * $attempt))
                    $retryIds = @($throttled.ToArray())
                    $newReqList = [System.Collections.Generic.List[object]]::new()
                    foreach ($req in $reqList) {
                        $reqId = if ($req -is [System.Collections.IDictionary]) { [string]$req['id'] } else { [string]$req.id }
                        if ($retryIds -contains $reqId) { $newReqList.Add($req) }
                    }
                    $reqList = $newReqList
                    $go = ($reqList.Count -gt 0)
                }
            }
            catch {
                if ($attempt -lt 4) {
                    Start-Sleep -Seconds (5 * $attempt)
                    $go = $true
                }
                else {
                    Write-Log "  Batch call failed after retries at script line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)" 'WARN'
                }
            }
        }
        Invoke-LoopGc ($i + 1) 200
    }

    foreach ($bucket in ($errorCounts.Keys | Sort-Object)) {
        $parts = $bucket -split '\|', 2
        Write-Log "  '$Activity': $($errorCounts[$bucket]) subrequest(s) failed with HTTP $($parts[0]) $($parts[1]). Sample: $($errorSamples[$bucket])" 'WARN'
    }

    Complete-Activity $Activity
    $out
}

# Project a Graph /users hashtable (from Invoke-MgGraphRequest) to a slim record,
# and feed the identity index. Keeps only the slim object in memory.
# FUNCTION: ConvertFrom-GraphUser
# Purpose : Convert a raw Graph user hashtable into the lightweight Entra_Users
#           report shape and feed core identity attributes into Dependency_Map.
function ConvertFrom-GraphUser {
    param($H)   # untyped: batch bodies aren't always [hashtable]
    $lic = ''
    if ($H['assignedLicenses']) { $lic = (@($H['assignedLicenses'] | ForEach-Object { $_['skuId'] } | Where-Object { $_ }) -join ';') }
    $rec = [pscustomobject]@{
        DisplayName       = $H['displayName']
        UserPrincipalName = $H['userPrincipalName']
        Mail              = $H['mail']
        AccountEnabled    = $H['accountEnabled']
        UserType          = $H['userType']
        DirSynced         = $H['onPremisesSyncEnabled']
        ImmutableId       = $H['onPremisesImmutableId']
        Department        = $H['department']
        JobTitle          = $H['jobTitle']
        EmployeeType      = $H['employeeType']
        OtherMails        = (@($H['otherMails']) -join ';')
        ProxyAddresses    = (@($H['proxyAddresses']) -join ';')
        LicenseSkuIds     = $lic
        ObjectId          = $H['id']
        CreatedDateTime   = $H['createdDateTime']
    }
    $key = if ($H['userPrincipalName']) { $H['userPrincipalName'] } elseif ($H['mail']) { $H['mail'] } else { $null }
    if ($key) {
        $idr = Get-IdentityRecord $key
        if ($idr) { $idr.EntraObjectId = $H['id']; $idr.EntraEnabled = $H['accountEnabled']; $idr.EntraDirSynced = $H['onPremisesSyncEnabled']; $idr.Licenses = $lic }
    }
    # EmployeeType is useful in Exchange migration reports but is not an Exchange
    # mailbox property. Cache it by every cloud identity key so the Exchange collector
    # can join it without another directory lookup. On-prem Users data can overwrite
    # these keys later and is treated as the preferred source.
    if ($H['employeeType']) {
        $et = $H['employeeType'].ToString()
        $etKeys = @($H['userPrincipalName'], $H['mail']) + @($H['otherMails'])
        foreach ($pa in @($H['proxyAddresses'])) {
            if ($pa -and "$pa" -match '(?i)^smtp:(.+)$') { $etKeys += $Matches[1] }
        }
        foreach ($ek in $etKeys) {
            if ($ek) {
                $elk = $ek.ToString().Trim().ToLowerInvariant()
                if (-not $EmployeeTypeByKey.ContainsKey($elk)) { $EmployeeTypeByKey[$elk] = $et }
            }
        }
    }
    $rec
}

# FUNCTION: Test-Module
# Purpose : Return $true when a required PowerShell module can be discovered.
#           Workload blocks use this to skip cleanly instead of crashing globally.
function Test-Module {
    param([string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Write-Log "Module '$Name' is not installed. Install it (Install-Module $Name -Scope CurrentUser) and re-run." 'ERROR'
        return $false
    }
    return $true
}

# Ordered collector for all worksheets: name -> array of objects
# $Report is the in-memory workbook model. Each key becomes an Excel worksheet.
# Collectors only add keys they actually ran, which is important to merge semantics:
# untouched workloads remain untouched in an existing workbook.
$Report = [ordered]@{}

# Master identity index used to build the dependency sheet.
# Key = canonical lower-case identity; value = correlated record. IdentityAliases
# maps alternate UPN/mail/SMTP values to that canonical key so a changed cloud UPN
# does not create a second dependency row for the same person.
# $Identity is the cross-workload correlation index used to build Dependency_Map.
# Keys are normalized user identifiers; aliases can be registered/merged so an
# on-prem UPN, cloud UPN and SMTP alias still represent one logical person.
$Identity = @{}
$IdentityAliases = @{}

# FUNCTION: Resolve-IdentityKey
# Purpose : Resolve a UPN/SMTP/alias to the canonical key used by the master
#           dependency identity index.
function Resolve-IdentityKey {
    param([string]$Key)
    if ([string]::IsNullOrWhiteSpace($Key)) { return $null }
    $k = $Key.Trim().ToLowerInvariant()
    $guard = 0
    while ($IdentityAliases.ContainsKey($k) -and $IdentityAliases[$k] -ne $k -and $guard -lt 20) {
        $k = $IdentityAliases[$k]
        $guard++
    }
    return $k
}

# FUNCTION: Merge-IdentityRecordKeys
# Purpose : Merge two dependency identity records when later evidence proves they
#           represent the same person (for example, source UPN vs cloud UPN).
function Merge-IdentityRecordKeys {
    param([string]$FromKey, [string]$ToKey)
    $from = Resolve-IdentityKey $FromKey
    $to   = Resolve-IdentityKey $ToKey
    if (-not $from -or -not $to -or $from -eq $to -or -not $Identity.ContainsKey($from)) { return }

    if (-not $Identity.ContainsKey($to)) {
        # Re-key the existing ordered record without reconstructing it.
        $record = $Identity[$from]
        $record.Key = $to
        $Identity[$to] = $record
        $Identity.Remove($from)
        $IdentityAliases[$from] = $to
        return
    }

    $source = $Identity[$from]
    $target = $Identity[$to]
    foreach ($prop in @('OnPremDN','OnPremEnabled','OnPremOU','SamAccountName','EntraObjectId','EntraEnabled','EntraDirSynced','Licenses',
                         'MailboxType','MailboxSizeMB','ArchiveEnabled','ArchiveSizeMB','ArchiveItemCount','AutoExpandingArchive',
                         'LitigationHold','InPlaceHolds','OneDriveUrl','OneDriveSizeMB')) {
        $sv = $source[$prop]
        $tv = $target[$prop]
        if (($null -eq $tv -or ([string]$tv -eq '')) -and $null -ne $sv -and ([string]$sv -ne '')) { $target[$prop] = $sv }
    }
    foreach ($prop in @('TeamsMemberOf','SharePointSites','IntuneDevices','EntraGroups','OnPremGroups')) {
        foreach ($v in @($source[$prop])) { Add-UniqueString -Collection $target[$prop] -Value $v }
    }
    $Identity.Remove($from)
    $IdentityAliases[$from] = $to
}

# FUNCTION: Register-IdentityAlias
# Purpose : Register an alternate identity value (UPN, mail or SMTP alias) against
#           one canonical dependency-map identity.
function Register-IdentityAlias {
    param([string]$Alias, [string]$Canonical)
    if ([string]::IsNullOrWhiteSpace($Alias) -or [string]::IsNullOrWhiteSpace($Canonical)) { return }
    $a = $Alias.Trim().ToLowerInvariant()
    $c = Resolve-IdentityKey $Canonical
    $old = Resolve-IdentityKey $a
    if ($old -and $old -ne $c) { Merge-IdentityRecordKeys -FromKey $old -ToKey $c }
    $IdentityAliases[$a] = $c
}

# Cross-source EmployeeType lookup used to enrich Exchange_Mailboxes. Keys are
# lower-case UPN/mail/SMTP addresses. On-prem Users values take precedence when present.
$EmployeeTypeByKey = @{}

# FUNCTION: Get-IdentityRecord
# Purpose : Retrieve or lazily create the canonical dependency record for a user.
#           All workload collectors add their footprint to this common structure.
function Get-IdentityRecord {
    param([string]$Key)
    if ([string]::IsNullOrWhiteSpace($Key)) { return $null }
    $k = Resolve-IdentityKey $Key
    if (-not $Identity.ContainsKey($k)) {
        $Identity[$k] = [ordered]@{
            Key             = $k
            OnPremDN        = $null
            OnPremEnabled   = $null
            OnPremOU        = $null
            SamAccountName  = $null
            EntraObjectId   = $null
            EntraEnabled    = $null
            EntraDirSynced  = $null
            Licenses        = $null
            MailboxType     = $null
            MailboxSizeMB   = $null
            ArchiveEnabled  = $null
            ArchiveSizeMB   = $null
            ArchiveItemCount = $null
            AutoExpandingArchive = $null
            LitigationHold  = $null
            InPlaceHolds    = $null
            OneDriveUrl     = $null
            OneDriveSizeMB  = $null
            TeamsMemberOf   = [System.Collections.Generic.List[string]]::new()
            SharePointSites = [System.Collections.Generic.List[string]]::new()
            IntuneDevices   = [System.Collections.Generic.List[string]]::new()
            EntraGroups     = [System.Collections.Generic.List[string]]::new()
            OnPremGroups    = [System.Collections.Generic.List[string]]::new()
        }
    }
    return $Identity[$k]
}

# Given a DN, return the parent OU/container DN (the object's location).
# FUNCTION: Get-ParentOU
# Purpose : Derive the parent OU/container distinguished name from an object DN.
function Get-ParentOU {
    param([string]$DistinguishedName)
    if ([string]::IsNullOrWhiteSpace($DistinguishedName)) { return $null }
    ($DistinguishedName -split '(?<!\\),', 2)[1]
}


# Identity dependency collections use List[string] for broad PowerShell/.NET
# compatibility. Keep them de-duplicated explicitly so merge/re-run hydration
# cannot double Teams/groups/devices/sites in Dependency_Map.
# FUNCTION: Add-UniqueString
# Purpose : Add a non-empty string to a collection only once. Prevents dependency
#           counts from inflating during merged/repeated discovery runs.
function Add-UniqueString {
    param($Collection, $Value)
    if ($null -eq $Collection -or [string]::IsNullOrWhiteSpace([string]$Value)) { return }
    $text = $Value.ToString().Trim()
    if (-not ($Collection -contains $text)) { $null = $Collection.Add($text) }
}

# SharePoint owner values can be plain UPNs or claims strings such as
# i:0#.f|membership|user@contoso.com. Return the email-like identity when present.
# FUNCTION: Get-PrincipalEmail
# Purpose : Normalize principal strings into an SMTP/UPN-like value, including
#           SharePoint claims-form identities such as i:0#.f|membership|user@domain.
function Get-PrincipalEmail {
    param($Value)
    if ($null -eq $Value) { return $null }
    $text = $Value.ToString().Trim()
    if ($text -match '(?i)([^|;\s]+@[^|;\s]+)') { return $Matches[1] }
    return $text
}

# Convert a raw AD FileTime attribute (accountExpires/pwdLastSet/lastLogonTimestamp)
# to a DateTime, returning $null for the 0 / 9223372036854775807 "never" sentinels
# instead of throwing "Not a valid Win32 FileTime".
# Return the inner text of a namespace-agnostic child element, or $null.
# FUNCTION: Get-Inner
# Purpose : Safely read a named child element from GPO XML without assuming the
#           extension-specific XML namespace.
function Get-Inner {
    param($Node, $Local)
    if (-not $Node) { return $null }
    $n = $Node.SelectSingleNode("*[local-name()='$Local']")
    if ($n) { $n.InnerText } else { $null }
}

# FUNCTION: ConvertTo-SafeDate
# Purpose : Convert AD file-time/date values while tolerating null, zero and
#           never-set sentinel values that otherwise throw conversion exceptions.
function ConvertTo-SafeDate {
    param($FileTimeValue)
    if ($null -eq $FileTimeValue) { return $null }
    try {
        $l = [int64]$FileTimeValue
        if ($l -le 0 -or $l -ge 9223372036854775807) { return $null }
        return [DateTime]::FromFileTimeUtc($l)
    }
    catch { return $null }
}

# Coerce a value read back from Excel (often the string "True"/"False") to bool.
# FUNCTION: ConvertTo-Bool
# Purpose : Normalize booleans imported from Excel/text or returned by APIs.
function ConvertTo-Bool {
    param($Value)
    if ($null -eq $Value) { return $false }
    if ($Value -is [bool]) { return $Value }
    return ($Value.ToString().Trim() -in @('True', '1', 'Yes'))
}


# Convert Exchange ByteQuantifiedSize / REST string representations to MB.
# Handles native ToBytes(), nested Value.ToBytes(), and localized display strings
# that include a numeric byte count in parentheses.
# FUNCTION: Convert-ExchangeSizeToMB
# Purpose : Normalize Exchange size representations (objects or formatted strings)
#           into a numeric MB value for reporting/correlation.
function Convert-ExchangeSizeToMB {
    param($Value)
    if ($null -eq $Value) { return $null }
    try {
        if ($Value.PSObject.Methods['ToBytes']) {
            return [math]::Round([double]$Value.ToBytes() / 1MB, 1)
        }
        if ($Value.PSObject.Properties['Value'] -and $Value.Value -and $Value.Value.PSObject.Methods['ToBytes']) {
            return [math]::Round([double]$Value.Value.ToBytes() / 1MB, 1)
        }
        $text = $Value.ToString()
        if ($text -match '\(([0-9,\. ]+) bytes\)') {
            $digits = ($Matches[1] -replace '[^0-9]', '')
            if ($digits) { return [math]::Round([double]$digits / 1MB, 1) }
        }
    }
    catch {}
    return $null
}


# Flatten SID/SIDHistory values to migration-friendly text. AD normally returns
# SecurityIdentifier objects, but tolerate byte[]/other representations too.
# FUNCTION: ConvertTo-SidText
# Purpose : Convert SID objects, strings or binary SID values into readable SID text.
#           Used for migration-critical SID/SIDHistory reporting.
function ConvertTo-SidText {
    param($Value)
    if ($null -eq $Value) { return $null }
    $vals = [System.Collections.Generic.List[string]]::new()
    # A single binary SID is itself IEnumerable; wrapping it with @() would split
    # it into individual bytes. Put values into an object-list explicitly so byte[]
    # remains one SID while normal multi-valued SIDHistory still enumerates entries.
    $sidValues = [System.Collections.Generic.List[object]]::new()
    if ($Value -is [byte[]]) { $sidValues.Add([object]$Value) }
    else { foreach ($entry in @($Value)) { $sidValues.Add($entry) } }
    foreach ($v in $sidValues) {
        if ($null -eq $v) { continue }
        try {
            if ($v -is [System.Security.Principal.SecurityIdentifier]) {
                $vals.Add($v.Value)
            }
            elseif ($v -is [byte[]]) {
                $vals.Add(([System.Security.Principal.SecurityIdentifier]::new($v, 0)).Value)
            }
            elseif ($v.PSObject.Properties['Value'] -and $v.Value) {
                $vals.Add($v.Value.ToString())
            }
            else {
                $vals.Add($v.ToString())
            }
        }
        catch { $vals.Add($v.ToString()) }
    }
    if ($vals.Count -eq 0) { return $null }
    return ($vals.ToArray() -join ';')
}

# Read prior worksheets from an existing workbook back into the identity index,
# so a later half (cloud after on-prem, or vice versa) produces a Dependency_Map
# that spans BOTH halves. Only reconstructs fields that live in a per-identity
# sheet; Teams membership and group lists are re-derived only when that half is
# collected in the current run (they have no per-user source sheet to read back).
# FUNCTION: Import-PriorWorkbookIdentity
# Purpose : Hydrate the dependency identity index from an existing workbook before
#           new collectors run. This is what lets split on-prem/cloud passes remain
#           correlated across repeated merge runs.
function Import-PriorWorkbookIdentity {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return }
    $sheets = @()
    try { $sheets = @((Get-ExcelSheetInfo -Path $Path).Name) }
    catch { Write-Log "Could not read sheet list from '$Path': $($_.Exception.Message)" 'WARN'; return }

    $samToUpn = @{}   # to re-attach group membership (keyed by SamAccountName) to identities

    if ($sheets -contains 'Users') {
        foreach ($u in (Import-Excel -Path $Path -WorksheetName 'Users')) {
            $key = if ($u.UserPrincipalName) { $u.UserPrincipalName } elseif ($u.Mail) { $u.Mail } else { $null }
            if ($key) {
                $id = Get-IdentityRecord $key
                $id.OnPremDN       = $u.DistinguishedName
                $id.OnPremEnabled  = ConvertTo-Bool $u.Enabled
                $id.OnPremOU       = $u.OU
                $id.SamAccountName = $u.SamAccountName
                if ($u.EmployeeType) {
                    $et = $u.EmployeeType.ToString()
                    $etKeys = @($u.UserPrincipalName, $u.Mail, $u.PrimarySMTP)
                    if ($u.ProxyAddresses) {
                        foreach ($pa in ($u.ProxyAddresses.ToString() -split ';')) {
                            if ($pa -match '(?i)^smtp:(.+)$') { $etKeys += $Matches[1] }
                        }
                    }
                    foreach ($ek in $etKeys) {
                        if ($ek) { $EmployeeTypeByKey[$ek.ToString().Trim().ToLowerInvariant()] = $et }
                    }
                }
                if ($u.SamAccountName) { $samToUpn[$u.SamAccountName] = $key }
            }
        }
    }
    if ($sheets -contains 'Entra_Users') {
        foreach ($u in (Import-Excel -Path $Path -WorksheetName 'Entra_Users')) {
            if ($u.UserPrincipalName) {
                $id = Get-IdentityRecord $u.UserPrincipalName
                $id.EntraObjectId = $u.ObjectId
                $id.EntraEnabled  = ConvertTo-Bool $u.AccountEnabled
                if ($null -ne $u.DirSynced -and -not [string]::IsNullOrWhiteSpace($u.DirSynced.ToString())) { $id.EntraDirSynced = ConvertTo-Bool $u.DirSynced }
                if (-not $id.Licenses) { $id.Licenses = $u.LicenseSkuIds }
                if ($u.EmployeeType) {
                    foreach ($ek in @($u.UserPrincipalName, $u.Mail)) {
                        if (-not $ek) { continue }
                        $elk = $ek.ToString().Trim().ToLowerInvariant()
                        if (-not $EmployeeTypeByKey.ContainsKey($elk)) { $EmployeeTypeByKey[$elk] = $u.EmployeeType.ToString() }
                    }
                }
            }
        }
    }
    if ($sheets -contains 'License_UserDetail') {
        foreach ($u in (Import-Excel -Path $Path -WorksheetName 'License_UserDetail')) {
            if ($u.UserPrincipalName) { (Get-IdentityRecord $u.UserPrincipalName).Licenses = $u.Licenses }
        }
    }
    if ($sheets -contains 'Exchange_Mailboxes') {
        foreach ($m in (Import-Excel -Path $Path -WorksheetName 'Exchange_Mailboxes')) {
            $mailIdentityKey = if ($m.UserPrincipalName) { $m.UserPrincipalName } elseif ($m.PrimarySmtpAddress) { $m.PrimarySmtpAddress } else { $null }
            if ($mailIdentityKey) {
                $id = Get-IdentityRecord $mailIdentityKey
                $id.MailboxType    = $m.MailboxType
                $id.MailboxSizeMB  = $m.SizeMB
                $archiveGuidText = if ($m.ArchiveGuid) { $m.ArchiveGuid.ToString() } else { '' }
                $id.ArchiveEnabled = (($m.ArchiveEnabled -and (ConvertTo-Bool $m.ArchiveEnabled)) -or
                                      ($m.ArchiveStatus -and $m.ArchiveStatus -ne 'None') -or
                                      ($archiveGuidText -and $archiveGuidText -ne '00000000-0000-0000-0000-000000000000'))
                if ($null -ne $m.ArchiveSizeMB) { $id.ArchiveSizeMB = $m.ArchiveSizeMB }
                if ($null -ne $m.ArchiveItemCount) { $id.ArchiveItemCount = $m.ArchiveItemCount }
                if ($null -ne $m.AutoExpandingArchiveEnabled) { $id.AutoExpandingArchive = ConvertTo-Bool $m.AutoExpandingArchiveEnabled }
                $id.LitigationHold = ConvertTo-Bool $m.LitigationHoldEnabled
                $id.InPlaceHolds   = $m.InPlaceHolds
            }
        }
    }
    if ($sheets -contains 'OneDrive') {
        foreach ($o in (Import-Excel -Path $Path -WorksheetName 'OneDrive')) {
            if ($o.Owner) {
                $id = Get-IdentityRecord $o.Owner
                $id.OneDriveUrl    = $o.Url
                $id.OneDriveSizeMB = $o.StorageUsedMB
            }
        }
    }
    if ($sheets -contains 'Intune_Devices') {
        foreach ($d in (Import-Excel -Path $Path -WorksheetName 'Intune_Devices')) {
            if ($d.UserPrincipalName -and $d.DeviceName) {
                Add-UniqueString -Collection ((Get-IdentityRecord $d.UserPrincipalName).IntuneDevices) -Value $d.DeviceName
            }
        }
    }
    if ($sheets -contains 'Teams_Members') {
        foreach ($m in (Import-Excel -Path $Path -WorksheetName 'Teams_Members')) {
            if ($m.User -and $m.Team) { Add-UniqueString -Collection ((Get-IdentityRecord $m.User).TeamsMemberOf) -Value $m.Team }
        }
    }
    $priorGroupUsers = @{}
    if ($sheets -contains 'Entra_GroupMembership') {
        foreach ($m in (Import-Excel -Path $Path -WorksheetName 'Entra_GroupMembership')) {
            if ($m.UserPrincipalName -and $m.GroupName) { Add-UniqueString -Collection ((Get-IdentityRecord $m.UserPrincipalName).EntraGroups) -Value $m.GroupName }
            if ($m.GroupId -and $m.UserPrincipalName) {
                $gid = $m.GroupId.ToString().Trim().ToLowerInvariant()
                if (-not $priorGroupUsers.ContainsKey($gid)) { $priorGroupUsers[$gid] = [System.Collections.Generic.List[string]]::new() }
                Add-UniqueString -Collection $priorGroupUsers[$gid] -Value $m.UserPrincipalName
            }
        }
    }
    if ($sheets -contains 'SharePoint_Sites') {
        foreach ($site in (Import-Excel -Path $Path -WorksheetName 'SharePoint_Sites')) {
            if (-not $site.Url) { continue }
            $ownerUpn = Get-PrincipalEmail $site.Owner
            if ($ownerUpn -and $ownerUpn -like '*@*') { Add-UniqueString -Collection ((Get-IdentityRecord $ownerUpn).SharePointSites) -Value $site.Url }
            if ($site.GroupId) {
                $gid = $site.GroupId.ToString().Trim().ToLowerInvariant()
                if ($priorGroupUsers.ContainsKey($gid)) {
                    foreach ($upn in $priorGroupUsers[$gid]) { Add-UniqueString -Collection ((Get-IdentityRecord $upn).SharePointSites) -Value $site.Url }
                }
            }
        }
    }
    if ($sheets -contains 'GroupMembership') {
        foreach ($gm in (Import-Excel -Path $Path -WorksheetName 'GroupMembership')) {
            if ($gm.MemberClass -eq 'user' -and $gm.MemberSam -and $gm.GroupName) {
                $upn = $samToUpn[$gm.MemberSam]
                if ($upn) { Add-UniqueString -Collection ((Get-IdentityRecord $upn).OnPremGroups) -Value $gm.GroupName }
            }
        }
    }
    Write-Log "Hydrated identity index from prior workbook ($($Identity.Count) identities so far)."
}

# Flatten a GPO's XML report to individual settings (deep GPO -> Intune mapping).
# Uses local-name() XPath so the per-extension XML namespaces don't matter.
# FUNCTION: Expand-GpoSettings
# Purpose : Flatten GPO XML into one row per individual setting. Handles multiple
#           extension namespaces and produces a migration-friendly setting inventory.
function Expand-GpoSettings {
    param($GpoNode, [string]$GpoName)

    $rows = [System.Collections.Generic.List[object]]::new()
    function New-SettingRow { param($Scope,$Type,$Category,$Name,$State,$Value)
        $rows.Add([pscustomobject]@{
            GpoName     = $GpoName
            Scope       = $Scope
            SettingType = $Type
            Category    = $Category
            Name        = $Name
            State       = $State
            Value       = $Value
        })
    }

    foreach ($scope in 'Computer','User') {
        $scopeNode = $GpoNode.SelectSingleNode("*[local-name()='$scope']")
        if (-not $scopeNode) { continue }

        # 1) Administrative Templates (registry.pol policies) -- primary Intune target
        foreach ($p in @($scopeNode.SelectNodes(".//*[local-name()='Policy']"))) {
            $name  = Get-Inner $p 'Name'
            $state = Get-Inner $p 'State'
            $cat   = Get-Inner $p 'Category'
            # any value-bearing child controls (dropdowns, text, numeric, checkboxes)
            $vals = foreach ($c in @($p.SelectNodes(".//*[local-name()='DropDownList' or local-name()='EditText' or local-name()='Numeric' or local-name()='Checkbox' or local-name()='Text' or local-name()='ListBox' or local-name()='MultiText']"))) {
                $cn = Get-Inner $c 'Name'
                $cv = (Get-Inner $c 'Value'); if (-not $cv) { $cv = Get-Inner $c 'State' }
                if ($cn -or $cv) { "$cn=$cv" }
            }
            New-SettingRow $scope 'AdminTemplate' $cat $name $state (($vals | Where-Object {$_}) -join ' | ')
        }

        # 2) Scripts (logon/logoff/startup/shutdown)
        foreach ($s in @($scopeNode.SelectNodes(".//*[local-name()='Script']"))) {
            New-SettingRow $scope 'Script' (Get-Inner $s 'Type') (Get-Inner $s 'Command') $null (Get-Inner $s 'Parameters')
        }

        # 3) Security -- SecurityOptions, Account (password/lockout), UserRights, Audit
        foreach ($so in @($scopeNode.SelectNodes(".//*[local-name()='SecurityOptions']"))) {
            $disp = $so.SelectSingleNode(".//*[local-name()='Name']")
            $val  = (Get-Inner $so 'SettingNumber'); if (-not $val) { $val = Get-Inner $so 'SettingString' }
            if (-not $val) { $val = Get-Inner $so 'SettingBoolean' }
            New-SettingRow $scope 'SecurityOption' 'SecurityOptions' ($(if($disp){$disp.InnerText})) $null $val
        }
        foreach ($ac in @($scopeNode.SelectNodes(".//*[local-name()='Account']"))) {
            New-SettingRow $scope 'AccountPolicy' (Get-Inner $ac 'Type') (Get-Inner $ac 'Name') $null (Get-Inner $ac 'SettingNumber')
        }
        foreach ($ur in @($scopeNode.SelectNodes(".//*[local-name()='UserRightsAssignment']"))) {
            $members = @($ur.SelectNodes(".//*[local-name()='Name']") | ForEach-Object { $_.InnerText }) -join '; '
            New-SettingRow $scope 'UserRight' 'UserRightsAssignment' (Get-Inner $ur 'Name') $null $members
        }

        # 4) Group Policy Preferences -- registry items and drive maps (common carry-overs)
        foreach ($r in @($scopeNode.SelectNodes(".//*[local-name()='Registry']/*[local-name()='Properties']"))) {
            New-SettingRow $scope 'GPP_Registry' $r.hive "$($r.key)\$($r.name)" $r.action $r.value
        }
        foreach ($d in @($scopeNode.SelectNodes(".//*[local-name()='Drive']/*[local-name()='Properties']"))) {
            New-SettingRow $scope 'GPP_DriveMap' $d.letter $d.path $d.action $d.label
        }
        foreach ($fr in @($scopeNode.SelectNodes(".//*[local-name()='Folder_Redirection']//*[local-name()='Location']"))) {
            New-SettingRow $scope 'FolderRedirect' (Get-Inner $fr 'DestinationPath') $null $null $null
        }
    }
    ,$rows.ToArray()
}

if ($Merging) {
    if (-not (Test-Module ImportExcel)) { throw "ImportExcel is required to use -InputWorkbook." }
    Import-Module ImportExcel -ErrorAction Stop
    Write-Log "Merge mode: reading prior workbook as source -> '$WorkbookPath'."
}

# Optional scoping: limit cloud pulls to the users AND devices already in the workbook.
$TargetUserKeys   = $null   # every user match value (UPN/mail/SMTP), lowercased -> $true
$TargetUserIds    = $null   # per user: ordered candidate ids to query cloud with
$TargetDeviceKeys = $null   # device names (and DNS short names), lowercased -> $true
$WaveGroupIds     = $null   # Entra group object ids the wave belongs to (for SPO GroupId)
$WaveGroupMails   = $null   # those groups' SMTPs (for distribution-list scoping)
$WaveUserIds      = $null   # wave users' Entra object ids (for CA / app-consent scoping)
if ($ScopeToInputUsers) {
    if (-not $Merging) { throw "-ScopeToInputUsers requires -InputWorkbook." }
    $TargetUserKeys   = @{}
    $TargetUserIds    = [System.Collections.Generic.List[string[]]]::new()
    $TargetDeviceKeys = @{}
    $targetUserKeyToIndex = @{}
    $wbSheets = @()
    try { $wbSheets = @((Get-ExcelSheetInfo -Path $WorkbookPath).Name) } catch {}

    # Read the users straight from the workbook so we keep UPN, mail AND primary
    # SMTP -- on-prem UPN often differs from the cloud UPN, so we must be able to
    # fall back to mail/SMTP when looking each user up in the cloud.
    foreach ($sheet in @('Users', 'Entra_Users')) {
        if ($wbSheets -notcontains $sheet) { continue }
        foreach ($u in (Import-Excel -Path $WorkbookPath -WorksheetName $sheet)) {
            $smtp = if ($sheet -eq 'Users' -and $u.PrimarySMTP) { ($u.PrimarySMTP -split ';')[0] } else { $null }
            $extraAliases = @()
            # Use every SMTP alias from the workbook, not just PrimarySMTP. This is
            # critical when the source UPN/primary address changed during coexistence
            # but an older source address survives as a secondary proxy in Entra.
            if ($u.ProxyAddresses) {
                foreach ($pa in ($u.ProxyAddresses.ToString() -split ';')) {
                    if ($pa -match '(?i)^smtp:(.+)$') { $extraAliases += $Matches[1] }
                }
            }
            if ($sheet -eq 'Entra_Users' -and $u.OtherMails) { $extraAliases += @($u.OtherMails.ToString() -split ';') }
            $cands = @($u.UserPrincipalName, $u.Mail, $smtp) + @($extraAliases) |
                Where-Object { $_ } | ForEach-Object { "$_".Trim() } | Select-Object -Unique
            if (@($cands).Count -eq 0) { continue }

            # The merge workbook can contain both Users and Entra_Users. Treat rows
            # with any overlapping identity key as the same person, merge their
            # candidate IDs, and avoid issuing duplicate cloud requests on re-runs.
            $existingIndex = $null
            foreach ($v in $cands) {
                $vk = $v.ToLowerInvariant()
                if ($targetUserKeyToIndex.ContainsKey($vk)) { $existingIndex = [int]$targetUserKeyToIndex[$vk]; break }
            }
            if ($null -eq $existingIndex) {
                $TargetUserIds.Add([string[]]$cands)
                $existingIndex = $TargetUserIds.Count - 1
            }
            else {
                $mergedCands = @($TargetUserIds[$existingIndex]) + @($cands) | Where-Object { $_ } | Select-Object -Unique
                $TargetUserIds[$existingIndex] = [string[]]$mergedCands
            }
            foreach ($v in @($TargetUserIds[$existingIndex])) {
                $vk = $v.ToLowerInvariant()
                $TargetUserKeys[$vk] = $true
                $targetUserKeyToIndex[$vk] = $existingIndex
            }
        }
    }
    # A bridge row can occasionally join two previously separate candidate sets
    # (for example, an on-prem UPN in Users and a changed cloud UPN in Entra_Users).
    # Collapse all overlapping sets transitively so a rerun cannot issue duplicate
    # per-user cloud calls or overstate the migration-wave user count.
    if ($TargetUserIds.Count -gt 1) {
        $collapsedTargets = [System.Collections.Generic.List[string[]]]::new()
        $collapsedKeyToIndex = @{}
        foreach ($candidateSet in @($TargetUserIds.ToArray())) {
            $normalized = @($candidateSet | Where-Object { $_ } | ForEach-Object { $_.ToString().Trim() } | Select-Object -Unique)
            if ($normalized.Count -eq 0) { continue }

            $linkedIndices = @($normalized | ForEach-Object {
                $nk = $_.ToLowerInvariant()
                if ($collapsedKeyToIndex.ContainsKey($nk)) { [int]$collapsedKeyToIndex[$nk] }
            } | Select-Object -Unique)

            if ($linkedIndices.Count -eq 0) {
                $collapsedTargets.Add([string[]]$normalized)
                $idx = $collapsedTargets.Count - 1
            }
            else {
                $idx = [int]($linkedIndices | Select-Object -First 1)
                $all = @($collapsedTargets[$idx]) + @($normalized)
                foreach ($other in @($linkedIndices | Where-Object { $_ -ne $idx })) {
                    $otherIndex = [int]$other
                    $all += @($collapsedTargets[$otherIndex])
                    $collapsedTargets[$otherIndex] = [string[]]@()
                }
                $collapsedTargets[$idx] = [string[]](@($all | Where-Object { $_ } | Select-Object -Unique))
            }

            foreach ($v in @($collapsedTargets[$idx])) {
                $collapsedKeyToIndex[$v.ToLowerInvariant()] = $idx
            }
        }

        # Compact any entries that were absorbed into another candidate set and
        # rebuild the key->index map so subsequent Entra matching has stable indices.
        $TargetUserIds = [System.Collections.Generic.List[string[]]]::new()
        $targetUserKeyToIndex = @{}
        foreach ($candidateSet in $collapsedTargets) {
            if (@($candidateSet).Count -eq 0) { continue }
            $TargetUserIds.Add([string[]]$candidateSet)
            $idx = $TargetUserIds.Count - 1
            foreach ($v in @($candidateSet)) {
                $vk = $v.ToLowerInvariant()
                $TargetUserKeys[$vk] = $true
                $targetUserKeyToIndex[$vk] = $idx
            }
        }
    }

    # Device scope from the workbook's Devices sheet (name + DNS short name).
    if ($wbSheets -contains 'Devices') {
        foreach ($d in (Import-Excel -Path $WorkbookPath -WorksheetName 'Devices')) {
            $names = @($d.Name, $(if ($d.DNSHostName) { ($d.DNSHostName -split '\.')[0] } ))
            foreach ($n in $names) { if ($n) { $TargetDeviceKeys[$n.ToString().Trim().ToLowerInvariant()] = $true } }
        }
    }
    if ($TargetUserKeys.Count -eq 0) {
        throw "-ScopeToInputUsers was requested but no users were found in the workbook (need a populated 'Users' or 'Entra_Users' sheet). Refusing to fall back to tenant-wide collection."
    }
    Write-Log "Scoping per-user cloud collection to $($TargetUserIds.Count) users ($($TargetUserKeys.Count) match keys) from the input workbook."
    if ($TargetDeviceKeys.Count -gt 0) {
        Write-Log "Scoping Intune device collection to $($TargetDeviceKeys.Count) devices from the workbook's Devices sheet."
    }

    # The first candidate is normally the on-prem UPN because Users is processed
    # before Entra_Users. Treat every UPN/mail/SMTP candidate in the merged set as
    # an alias of that canonical identity before hydrating prior worksheets.
    foreach ($candidateSet in $TargetUserIds) {
        if (@($candidateSet).Count -eq 0) { continue }
        $canonical = $candidateSet[0]
        foreach ($alias in $candidateSet) { Register-IdentityAlias -Alias $alias -Canonical $canonical }
    }
}

# Hydrate prior data only after identity aliases are known, otherwise a changed
# cloud UPN and the original on-prem UPN can become two dependency records.
if ($Merging) { Import-PriorWorkbookIdentity -Path $WorkbookPath }

# Membership test used by the scoped cloud collectors ($true when not scoping).
# FUNCTION: Test-TargetUser
# Purpose : Determine whether a user identity belongs to the migration-wave scope.
#           Returns $true for every user when scoped mode is not enabled.
function Test-TargetUser {
    param([string]$Key)
    if ($null -eq $TargetUserKeys) { return $true }
    if ([string]::IsNullOrWhiteSpace($Key)) { return $false }
    return $TargetUserKeys.ContainsKey($Key.Trim().ToLowerInvariant())
}
# Intune device scope: keep a device if it's in the workbook's Devices list OR its
# primary user is in the wave (so we get both the listed machines and everything
# else those users have enrolled). $true when not scoping.
# FUNCTION: Test-TargetDevice
# Purpose : Determine whether an Intune device is in scope by workbook device name
#           or by association to a scoped migration-wave user.
function Test-TargetDevice {
    param($Device)
    if ($TargetDeviceKeys -and $TargetDeviceKeys.Count -gt 0) {
        $n = if ($Device.DeviceName) { $Device.DeviceName } else { '' }
        if ($n -and $TargetDeviceKeys.ContainsKey($n.ToString().Trim().ToLowerInvariant())) { return $true }
    }
    return (Test-TargetUser $Device.UserPrincipalName)
}

# Resolve a workbook user to a cloud object by trying its candidate ids in order
# (UPN, then mail, then SMTP). $Getter is a scriptblock taking one id string.
# FUNCTION: Resolve-CloudUser
# Purpose : Try a sequence of candidate identities (UPN/mail/SMTP aliases) against
#           a supplied cloud lookup until one resolves successfully.
function Resolve-CloudUser {
    param([string[]]$Candidates, [scriptblock]$Getter)
    foreach ($idv in $Candidates) {
        try { $obj = & $Getter $idv; if ($obj) { return $obj } } catch {}
    }
    return $null
}

# Apps and OneDrive need Entra object IDs even when the caller deliberately omits
# the Entra workload. Resolve the wave once and cache the slim result so subset
# workload runs stay correctly scoped instead of falling back to the first on-prem
# UPN candidate (which may not be the cloud UPN).
$ScopedGraphUsersCache = $null
# FUNCTION: Get-ScopedGraphUsersForWave
# Purpose : Resolve scoped workbook identities to lightweight Graph user objects.
#           Used by cloud workloads that need Entra ObjectIds even when Entra user
#           inventory is not itself selected as a workload.
function Get-ScopedGraphUsersForWave {
    if ($null -ne $script:ScopedGraphUsersCache) { return @($script:ScopedGraphUsersCache) }
    $resolved = [System.Collections.Generic.List[object]]::new()

    if ($null -ne $TargetUserIds) {
        foreach ($h in (Invoke-GraphGetAll "v1.0/users?`$select=id,userPrincipalName,mail,otherMails,proxyAddresses&`$top=999")) {
            $keys = @($h['userPrincipalName'], $h['mail']) + @($h['otherMails'])
            foreach ($pa in @($h['proxyAddresses'])) {
                if ($pa -and "$pa" -match '(?i)^smtp:(.+)$') { $keys += $Matches[1] }
            }
            $match = $false
            foreach ($k in $keys) {
                if ($k -and $TargetUserKeys.ContainsKey($k.ToString().Trim().ToLowerInvariant())) { $match = $true; break }
            }
            if ($match -and $h['id']) {
                $resolved.Add([pscustomobject]@{ UserPrincipalName = $h['userPrincipalName']; ObjectId = $h['id']; Upn = $h['userPrincipalName']; Id = $h['id'] })
            }
        }
    }
    else {
        foreach ($h in (Invoke-GraphGetAll "v1.0/users?`$select=id,userPrincipalName&`$top=999")) {
            if ($h['id']) { $resolved.Add([pscustomobject]@{ UserPrincipalName = $h['userPrincipalName']; ObjectId = $h['id']; Upn = $h['userPrincipalName']; Id = $h['id'] }) }
        }
    }

    $script:ScopedGraphUsersCache = $resolved.ToArray()
    return @($script:ScopedGraphUsersCache)
}

# True if any value in $Values (lowercased) is a key of hashtable $Set.
# FUNCTION: Test-AnyIn
# Purpose : Return $true when any supplied value exists in a normalized hashtable set.
function Test-AnyIn {
    param($Values, $Set)
    if (-not $Set) { return $false }
    foreach ($v in @($Values)) { if ($v -and $Set.ContainsKey($v.ToString().ToLowerInvariant())) { return $true } }
    return $false
}

# Does a Conditional Access policy target the wave (by included users OR groups)?
# "All" users means it affects everyone, so it's in scope. $true when not scoping.
# FUNCTION: Test-CaInWave
# Purpose : Decide whether a Conditional Access policy is relevant to the scoped
#           migration wave based on included/excluded users/groups and All-users rules.
function Test-CaInWave {
    param($Conditions)
    if ($null -eq $TargetUserKeys) { return $true }
    $u = $Conditions.Users
    if (@($u.IncludeUsers) -contains 'All') { return $true }
    if (Test-AnyIn $u.IncludeUsers  $WaveUserIds)  { return $true }
    if (Test-AnyIn $u.IncludeGroups $WaveGroupIds) { return $true }
    return $false
}

# Purview policy relevance to the wave, by its location arrays. Locations are
# name/URL-based (not object ids), so this is a best-effort FLAG, not a filter:
# org-wide ("All") is always in scope; otherwise we match a location string against
# the wave's user identifiers (UPN/mail/SMTP) and group SMTPs. Returns $true when
# not scoping. Used to set an AppliesToWave column (rows are never dropped).
# FUNCTION: Test-PurviewInWave
# Purpose : Best-effort relevance test for Purview records when scoped discovery is
#           active. Tenant-wide policies remain context where precise attribution
#           is not available.
function Test-PurviewInWave {
    param([object[]]$LocationArrays)
    if ($null -eq $TargetUserKeys) { return $true }
    foreach ($arr in $LocationArrays) {
        foreach ($loc in @($arr)) {
            if (-not $loc) { continue }
            $s = $loc.ToString().Trim()
            if ($s -eq 'All') { return $true }
            $sl = $s.ToLowerInvariant()
            if ($TargetUserKeys.ContainsKey($sl)) { return $true }
            if ($WaveGroupMails -and $WaveGroupMails.ContainsKey($sl)) { return $true }
        }
    }
    return $false
}

#endregion

#region ------------------------------------------------------------- On-prem AD
# ON-PREMISES DATA CONTRACT
# -------------------------
# Source modules : ActiveDirectory and GroupPolicy (RSAT)
# Source access  : Read-only queries
# Main outputs   : OrganizationalUnits, Users, Devices, Contacts, Groups,
#                  GroupMembership, GPOs, GPO_Links, GPO_Settings
# Correlation    : User/device/group identifiers are also fed into Dependency_Map.
# Performance    : GPO setting extraction is the most expensive on-prem operation;
#                  use -SkipGpoSettings or -GpoLinksOnly for faster passes.

if ($IncludeOnPrem) {
    if (-not (Test-Module ActiveDirectory) -or -not (Test-Module GroupPolicy)) {
        throw "On-prem collection requires the ActiveDirectory and GroupPolicy RSAT modules."
    }
    Import-Module ActiveDirectory -ErrorAction Stop
    Import-Module GroupPolicy     -ErrorAction Stop

    $adCommon = @{ ErrorAction = 'Stop' }
    if ($Server) { $adCommon['Server'] = $Server }

    # Scope: subtree search always walks nested OUs.
    if ($Scope -eq 'OU') {
        $searchScope = @{ SearchBase = $OUPath; SearchScope = 'Subtree' }
        Write-Log "AD scope: OU subtree under '$OUPath'."
    }
    else {
        $searchScope = @{}
        Write-Log "AD scope: entire domain."
    }

    # -- Organizational Units (structure map) --------------------------------
    # Source : Get-ADOrganizationalUnit
    # Output : OrganizationalUnits
    # Captures OU hierarchy, management metadata, GPO-link presence and stable IDs.
    Write-Log "Collecting Organizational Units..."
    try {
        $ouProps = 'Name', 'DistinguishedName', 'Description', 'ManagedBy', 'gPLink', 'whenCreated', 'whenChanged', 'CanonicalName'
        $ous = Get-ADOrganizationalUnit -Filter * @searchScope @adCommon -Properties $ouProps |
            ForEach-Object {
                [pscustomobject]@{
                    Name              = $_.Name
                    ObjectGUID        = if ($_.ObjectGUID) { $_.ObjectGUID.ToString() } else { $null }
                    ObjectClass       = $_.ObjectClass
                    CanonicalName     = $_.CanonicalName
                    DistinguishedName = $_.DistinguishedName
                    ParentOU          = Get-ParentOU $_.DistinguishedName
                    Description       = $_.Description
                    ManagedBy         = $_.ManagedBy
                    HasLinkedGPO      = [bool]$_.gPLink
                    WhenCreated       = $_.whenCreated
                    WhenChanged       = $_.whenChanged
                }
            }
        $Report['OrganizationalUnits'] = @($ous)
        Write-Log "  $($ous.Count) OUs."
    }
    catch { Write-Log "OU collection failed: $($_.Exception.Message)" 'ERROR' }

    # -- Users ----------------------------------------------------------------
    # Source : Get-ADUser
    # Output : Users
    # Captures identity, HR attributes, mail/proxy data, Exchange hybrid attributes,
    # SID/GUID/SIDHistory, extension attributes, timestamps and OU placement.
    Write-Log "Collecting Users..."
    try {
        # Use RAW FileTime attributes (accountExpires/pwdLastSet/lastLogonTimestamp)
        # and convert them with ConvertTo-SafeDate -- the module's synthesized
        # AccountExpirationDate/PasswordLastSet/LastLogonDate throw on "never" sentinels.
        $userProps = 'DisplayName', 'GivenName', 'Surname', 'UserPrincipalName', 'SamAccountName',
                     'mail', 'proxyAddresses', 'Enabled', 'Department', 'Title', 'Manager',
                     'Company', 'physicalDeliveryOfficeName', 'telephoneNumber', 'mobile',
                     'employeeID', 'employeeType', 'accountExpires', 'pwdLastSet',
                     'lastLogonTimestamp', 'whenCreated', 'whenChanged', 'DistinguishedName',
                     'SIDHistory', 'userAccountControl',
                     'adminCount', 'primaryGroupID', 'CanonicalName', 'msDS-PrincipalName',
                     'msExchMailboxGuid', 'targetAddress', 'mailNickname', 'legacyExchangeDN',
                     'msExchRecipientTypeDetails', 'msExchRemoteRecipientType',
                     'extensionAttribute1','extensionAttribute2','extensionAttribute3','extensionAttribute4','extensionAttribute5',
                     'extensionAttribute6','extensionAttribute7','extensionAttribute8','extensionAttribute9','extensionAttribute10',
                     'extensionAttribute11','extensionAttribute12','extensionAttribute13','extensionAttribute14','extensionAttribute15'
        Write-Verbose2 "Querying Active Directory for user objects (this can take a moment on a large domain)..."
        $ui = 0
        $users = Get-ADUser -Filter * @searchScope @adCommon -Properties $userProps |
            ForEach-Object {
                $ui++
                if ($ui % 100 -eq 0) { Write-Activity -Activity 'Collecting users' -Current $ui -Item $_.SamAccountName }
                $rec = [pscustomobject]@{
                    DisplayName       = $_.DisplayName
                    GivenName         = $_.GivenName
                    Surname           = $_.Surname
                    UserPrincipalName = $_.UserPrincipalName
                    SamAccountName    = $_.SamAccountName
                    PrimarySMTP       = ($_.proxyAddresses | Where-Object { $_ -clike 'SMTP:*' } | ForEach-Object { $_ -replace '^SMTP:' }) -join ';'
                    Mail              = $_.mail
                    ProxyAddresses    = ($_.proxyAddresses -join ';')
                    Enabled           = $_.Enabled
                    Department        = $_.Department
                    Title             = $_.Title
                    Company           = $_.Company
                    Office            = $_.physicalDeliveryOfficeName
                    Manager           = $_.Manager
                    Phone             = $_.telephoneNumber
                    Mobile            = $_.mobile
                    EmployeeID        = $_.employeeID
                    EmployeeType      = $_.employeeType
                    ObjectGUID        = if ($_.ObjectGUID) { $_.ObjectGUID.ToString() } else { $null }
                    ObjectSID         = ConvertTo-SidText $_.SID
                    SIDHistory        = ConvertTo-SidText $_.SIDHistory
                    ObjectClass       = $_.ObjectClass
                    UserAccountControl= $_.userAccountControl
                    AdminCount        = $_.adminCount
                    PrimaryGroupID    = $_.primaryGroupID
                    CanonicalName     = $_.CanonicalName
                    MsDSPrincipalName = $_.'msDS-PrincipalName'
                    TargetAddress     = $_.targetAddress
                    MailNickname      = $_.mailNickname
                    LegacyExchangeDN  = $_.legacyExchangeDN
                    MsExchRecipientTypeDetails = $_.msExchRecipientTypeDetails
                    MsExchRemoteRecipientType  = $_.msExchRemoteRecipientType
                    ExtensionAttribute1  = $_.extensionAttribute1
                    ExtensionAttribute2  = $_.extensionAttribute2
                    ExtensionAttribute3  = $_.extensionAttribute3
                    ExtensionAttribute4  = $_.extensionAttribute4
                    ExtensionAttribute5  = $_.extensionAttribute5
                    ExtensionAttribute6  = $_.extensionAttribute6
                    ExtensionAttribute7  = $_.extensionAttribute7
                    ExtensionAttribute8  = $_.extensionAttribute8
                    ExtensionAttribute9  = $_.extensionAttribute9
                    ExtensionAttribute10 = $_.extensionAttribute10
                    ExtensionAttribute11 = $_.extensionAttribute11
                    ExtensionAttribute12 = $_.extensionAttribute12
                    ExtensionAttribute13 = $_.extensionAttribute13
                    ExtensionAttribute14 = $_.extensionAttribute14
                    ExtensionAttribute15 = $_.extensionAttribute15
                    MsExchMailboxGuid = if ($_.msExchMailboxGuid -is [byte[]]) { ([Guid]::new($_.msExchMailboxGuid)).ToString() } elseif ($_.msExchMailboxGuid) { $_.msExchMailboxGuid.ToString() } else { $null }
                    HasExchMailbox    = [bool]$_.msExchMailboxGuid
                    AccountExpires    = ConvertTo-SafeDate $_.accountExpires
                    PasswordLastSet   = ConvertTo-SafeDate $_.pwdLastSet
                    LastLogonDate     = ConvertTo-SafeDate $_.lastLogonTimestamp
                    WhenCreated       = $_.whenCreated
                    WhenChanged       = $_.whenChanged
                    OU                = Get-ParentOU $_.DistinguishedName
                    DistinguishedName = $_.DistinguishedName
                }
                # Cache EmployeeType by UPN/mail/SMTP so Exchange rows can be enriched
                # whether on-prem and cloud are collected in one run or merged across runs.
                if ($_.employeeType) {
                    $et = $_.employeeType.ToString()
                    $etKeys = @($_.UserPrincipalName, $_.mail)
                    foreach ($pa in @($_.proxyAddresses)) {
                        if ($pa -and "$pa" -match '(?i)^smtp:(.+)$') { $etKeys += $Matches[1] }
                    }
                    foreach ($ek in $etKeys) {
                        if ($ek) { $EmployeeTypeByKey[$ek.ToString().Trim().ToLowerInvariant()] = $et }
                    }
                }

                # feed the identity index
                $key = if ($_.UserPrincipalName) { $_.UserPrincipalName } elseif ($_.mail) { $_.mail } else { $null }
                if ($key) {
                    $id = Get-IdentityRecord $key
                    $id.OnPremDN       = $_.DistinguishedName
                    $id.OnPremEnabled  = $_.Enabled
                    $id.OnPremOU       = $rec.OU
                    $id.SamAccountName = $_.SamAccountName
                }
                $rec
            }
        Complete-Activity 'Collecting users'
        $Report['Users'] = @($users)
        Write-Log "  $($users.Count) users."
    }
    catch { Write-Log "User collection failed: $($_.Exception.Message)" 'ERROR' }

    # -- Devices / Computers --------------------------------------------------
    # Source : Get-ADComputer
    # Output : Devices
    # Captures OS, logon/password timestamps, SID/GUID, SPNs, encryption settings
    # and Kerberos delegation attributes relevant to server/device migration.
    Write-Log "Collecting Computers/Devices..."
    try {
        $compProps = 'Name', 'DNSHostName', 'OperatingSystem', 'OperatingSystemVersion',
                     'Enabled', 'lastLogonTimestamp', 'pwdLastSet', 'whenCreated', 'whenChanged',
                     'DistinguishedName', 'IPv4Address', 'Description',
                     'SIDHistory', 'userAccountControl', 'adminCount', 'primaryGroupID',
                     'CanonicalName', 'msDS-PrincipalName', 'servicePrincipalName',
                     'msDS-SupportedEncryptionTypes', 'TrustedForDelegation',
                     'TrustedToAuthForDelegation', 'msDS-AllowedToDelegateTo'
        Write-Verbose2 "Querying Active Directory for computer objects..."
        $ci = 0
        $computers = Get-ADComputer -Filter * @searchScope @adCommon -Properties $compProps |
            ForEach-Object {
                $ci++
                if ($ci % 100 -eq 0) { Write-Activity -Activity 'Collecting computers' -Current $ci -Item $_.Name }
                [pscustomobject]@{
                    Name              = $_.Name
                    DNSHostName       = $_.DNSHostName
                    OperatingSystem   = $_.OperatingSystem
                    OSVersion         = $_.OperatingSystemVersion
                    Enabled           = $_.Enabled
                    IPv4Address       = $_.IPv4Address
                    Description       = $_.Description
                    ObjectGUID        = if ($_.ObjectGUID) { $_.ObjectGUID.ToString() } else { $null }
                    ObjectSID         = ConvertTo-SidText $_.SID
                    SIDHistory        = ConvertTo-SidText $_.SIDHistory
                    ObjectClass       = $_.ObjectClass
                    UserAccountControl= $_.userAccountControl
                    AdminCount        = $_.adminCount
                    PrimaryGroupID    = $_.primaryGroupID
                    CanonicalName     = $_.CanonicalName
                    MsDSPrincipalName = $_.'msDS-PrincipalName'
                    ServicePrincipalNames = ($_.servicePrincipalName -join ';')
                    SupportedEncryptionTypes = $_.'msDS-SupportedEncryptionTypes'
                    TrustedForDelegation = $_.TrustedForDelegation
                    TrustedToAuthForDelegation = $_.TrustedToAuthForDelegation
                    AllowedToDelegateTo = ($_.'msDS-AllowedToDelegateTo' -join ';')
                    LastLogonDate     = ConvertTo-SafeDate $_.lastLogonTimestamp
                    PasswordLastSet   = ConvertTo-SafeDate $_.pwdLastSet
                    WhenCreated       = $_.whenCreated
                    WhenChanged       = $_.whenChanged
                    OU                = Get-ParentOU $_.DistinguishedName
                    DistinguishedName = $_.DistinguishedName
                }
            }
        Complete-Activity 'Collecting computers'
        $Report['Devices'] = @($computers)
        Write-Log "  $($computers.Count) computers."
    }
    catch { Write-Log "Computer collection failed: $($_.Exception.Message)" 'ERROR' }

    # -- Contacts -------------------------------------------------------------
    # Source : Get-ADObject (objectClass=contact)
    # Output : Contacts
    # Captures mail routing/legacy Exchange attributes and stable directory IDs.
    Write-Log "Collecting Contacts..."
    try {
        $contacts = Get-ADObject -LDAPFilter '(objectClass=contact)' @searchScope @adCommon `
                        -Properties displayName, mail, proxyAddresses, telephoneNumber, whenCreated, whenChanged, distinguishedName, objectGUID, objectClass, canonicalName, targetAddress, mailNickname, legacyExchangeDN |
            ForEach-Object {
                [pscustomobject]@{
                    DisplayName       = $_.displayName
                    Mail              = $_.mail
                    ProxyAddresses    = ($_.proxyAddresses -join ';')
                    Phone             = $_.telephoneNumber
                    ObjectGUID        = if ($_.ObjectGUID) { $_.ObjectGUID.ToString() } else { $null }
                    ObjectClass       = $_.ObjectClass
                    CanonicalName     = $_.CanonicalName
                    TargetAddress     = $_.targetAddress
                    MailNickname      = $_.mailNickname
                    LegacyExchangeDN  = $_.legacyExchangeDN
                    WhenCreated       = $_.whenCreated
                    WhenChanged       = $_.whenChanged
                    OU                = Get-ParentOU $_.distinguishedName
                    DistinguishedName = $_.distinguishedName
                }
            }
        $Report['Contacts'] = @($contacts)
        Write-Log "  $(@($contacts).Count) contacts."
    }
    catch { Write-Log "Contact collection failed: $($_.Exception.Message)" 'ERROR' }

    # -- Groups + nested membership ------------------------------------------
    # Source : Get-ADGroup + Get-ADGroupMember -Recursive
    # Output : Groups, GroupMembership
    # Membership is flattened to leaf members for dependency analysis while the
    # group report retains direct member count and migration-critical SID/GUID data.
    Write-Log "Collecting Groups and (nested) membership..."
    try {
        $groupProps = 'Name', 'SamAccountName', 'GroupCategory', 'GroupScope', 'mail',
                      'Description', 'ManagedBy', 'member', 'whenCreated', 'whenChanged', 'DistinguishedName',
                      'SIDHistory', 'adminCount', 'CanonicalName',
                      'msDS-PrincipalName', 'proxyAddresses', 'mailNickname'
        Write-Verbose2 "Querying Active Directory for group objects..."
        $groups = @(Get-ADGroup -Filter * @searchScope @adCommon -Properties $groupProps)
        $groupRows      = [System.Collections.Generic.List[object]]::new()
        $membershipRows = [System.Collections.Generic.List[object]]::new()

        # SamAccountName -> UPN lookup, built once (avoids an O(n^2) scan per member).
        $samToUpnLive = @{}
        foreach ($u in @($Report['Users'])) {
            if ($u.SamAccountName) { $samToUpnLive[$u.SamAccountName] = $u.UserPrincipalName }
        }

        $gi = 0; $gTotal = $groups.Count
        foreach ($g in $groups) {
            $gi++
            Write-Activity -Activity 'Expanding group membership' -Current $gi -Total $gTotal -Item $g.Name
            if ($gi % 250 -eq 0) { Write-Verbose2 "  ...expanded $gi of $gTotal groups" }
            $groupRows.Add([pscustomobject]@{
                Name              = $g.Name
                SamAccountName    = $g.SamAccountName
                Category          = $g.GroupCategory      # Security / Distribution
                Scope             = $g.GroupScope         # DomainLocal / Global / Universal
                Mail              = $g.mail
                ProxyAddresses    = ($g.proxyAddresses -join ';')
                MailNickname      = $g.mailNickname
                Description       = $g.Description
                ManagedBy         = $g.ManagedBy
                ObjectGUID        = if ($g.ObjectGUID) { $g.ObjectGUID.ToString() } else { $null }
                ObjectSID         = ConvertTo-SidText $g.SID
                SIDHistory        = ConvertTo-SidText $g.SIDHistory
                ObjectClass       = $g.ObjectClass
                AdminCount        = $g.adminCount
                CanonicalName     = $g.CanonicalName
                MsDSPrincipalName = $g.'msDS-PrincipalName'
                DirectMemberCount = @($g.member).Count
                OU                = Get-ParentOU $g.DistinguishedName
                DistinguishedName = $g.DistinguishedName
                WhenCreated       = $g.whenCreated
                WhenChanged       = $g.whenChanged
            })

            # Recursive expansion flattens nested groups to leaf members.
            try {
                $members = Get-ADGroupMember -Identity $g.DistinguishedName -Recursive @adCommon
            }
            catch {
                Write-Log "  Could not expand '$($g.Name)': $($_.Exception.Message)" 'WARN'
                $members = @()
            }
            foreach ($m in $members) {
                $membershipRows.Add([pscustomobject]@{
                    GroupName         = $g.Name
                    GroupScope        = $g.GroupScope
                    GroupCategory     = $g.GroupCategory
                    GroupObjectGUID   = if ($g.ObjectGUID) { $g.ObjectGUID.ToString() } else { $null }
                    GroupSID          = ConvertTo-SidText $g.SID
                    MemberName        = $m.name
                    MemberSam         = $m.SamAccountName
                    MemberClass       = $m.objectClass
                    MemberObjectGUID  = if ($m.ObjectGUID) { $m.ObjectGUID.ToString() } else { $null }
                    MemberSID         = ConvertTo-SidText $m.SID
                    MemberDN          = $m.distinguishedName
                    GroupDN           = $g.DistinguishedName
                })
                if ($m.objectClass -eq 'user' -and $m.SamAccountName) {
                    $key = $samToUpnLive[$m.SamAccountName]
                    if ($key) { Add-UniqueString -Collection ((Get-IdentityRecord $key).OnPremGroups) -Value $g.Name }
                }
            }
        }
        Complete-Activity 'Expanding group membership'
        $Report['Groups']          = $groupRows.ToArray()
        $Report['GroupMembership'] = $membershipRows.ToArray()
        Write-Log "  $($groupRows.Count) groups, $($membershipRows.Count) membership rows (nested expanded)."
    }
    catch { Write-Log "Group collection failed: $($_.Exception.Message)" 'ERROR' }

    # -- GPOs + links + settings report --------------------------------------
    # Source : Get-GPO / Get-GPOReport or direct gPLink in links-only mode
    # Output : GPOs, GPO_Links, optionally GPO_Settings
    # Combined XML is preferred for efficiency; resilient per-GPO fallback is used
    # if a domain-wide report times out.
    Write-Log "Collecting GPOs..."
    try {
        $gpoParams = @{ All = $true }
        if ($Server) { $gpoParams['Server'] = $Server }
        $allGpo = Get-GPO @gpoParams
        Write-Log "  $(@($allGpo).Count) GPOs found."

        if ($GpoLinksOnly) {
            # Fast mode: read GPO->OU links straight from each container's gPLink
            # attribute. No Get-GPOReport, so no timeout risk -- near-instant, but
            # no settings and no enabled-extension summary.
            $guidToName = @{}
            foreach ($g in $allGpo) { $guidToName[$g.Id.ToString().ToLowerInvariant()] = $g.DisplayName }

            $gpoRows = foreach ($g in $allGpo) {
                [pscustomobject]@{
                    Name            = $g.DisplayName
                    Id              = $g.Id
                    GpoStatus       = $g.GpoStatus
                    Created         = $g.CreationTime
                    Modified        = $g.ModificationTime
                    ComputerEnabled = $g.Computer.Enabled
                    UserEnabled     = $g.User.Enabled
                    WmiFilter       = $g.WmiFilter.Name
                }
            }

            # gPLink packs ordered links as [LDAP://cn={GUID},...;<options>] where
            # options bit0=1 disabled, bit1=2 enforced.
            $linkRows = [System.Collections.Generic.List[object]]::new()
            $linked = Get-ADObject -LDAPFilter '(gPLink=*)' @searchScope @adCommon -Properties gPLink, distinguishedName
            foreach ($c in $linked) {
                $order = 0
                foreach ($m in [regex]::Matches([string]$c.gPLink, '\[LDAP://cn=\{(?<guid>[0-9A-Fa-f-]+)\},[^;\]]+;(?<opt>\d+)\]')) {
                    $order++
                    $opt  = [int]$m.Groups['opt'].Value
                    $guid = $m.Groups['guid'].Value.ToLowerInvariant()
                    $linkRows.Add([pscustomobject]@{
                        GpoName     = $(if ($guidToName.ContainsKey($guid)) { $guidToName[$guid] } else { "(unknown $guid)" })
                        GpoId       = $guid
                        LinkedTo    = $c.distinguishedName
                        LinkOrder   = $order
                        LinkEnabled = (-not ($opt -band 1))
                        Enforced    = [bool]($opt -band 2)
                    })
                }
            }
            $Report['GPOs']      = @($gpoRows)
            $Report['GPO_Links'] = $linkRows.ToArray()
            Write-Log "  $(@($gpoRows).Count) GPOs, $($linkRows.Count) links (links-only mode; settings skipped)."
        }
        else {

        # Prefer ONE domain-wide XML report (far fewer round trips). On a large
        # domain that single report can exceed the GPMC RPC timeout (0x800705B4),
        # so fall back to resilient per-GPO reports if the combined call fails.
        $serverParam = @{}
        if ($Server) { $serverParam['Server'] = $Server }
        $gpoNodes   = @{}
        $combinedOk = $false
        try {
            [xml]$reportXml = Get-GPOReport @serverParam -All -ReportType Xml
            foreach ($n in $reportXml.SelectNodes("//*[local-name()='GPO']")) {
                $idNode = $n.SelectSingleNode("*[local-name()='Identifier']/*[local-name()='Identifier']")
                if ($idNode -and $idNode.InnerText) {
                    $gpoNodes[$idNode.InnerText.Trim('{', '}').ToLowerInvariant()] = $n
                }
            }
            $combinedOk = $true
            Write-Log "  Combined report parsed ($($gpoNodes.Count) GPOs). Extracting..."
        }
        catch {
            Write-Log "  Combined report failed ($($_.Exception.Message)). Falling back to per-GPO reports (slower, resilient)." 'WARN'
        }

        $gpoRows      = [System.Collections.Generic.List[object]]::new()
        $linkRows     = [System.Collections.Generic.List[object]]::new()
        $settingRows  = [System.Collections.Generic.List[object]]::new()
        $i = 0; $total = @($allGpo).Count

        foreach ($gpo in $allGpo) {
            $i++
            Write-Progress -Activity "Processing GPOs" -Status "$i / $total : $($gpo.DisplayName)" `
                -PercentComplete ([int](($i / [math]::Max($total,1)) * 100))

            # Combined path: look up the prebuilt node. Fallback path: fetch this
            # GPO's report on its own (small, won't hit the domain-wide timeout).
            $node = $null
            if ($combinedOk) {
                $node = $gpoNodes[$gpo.Id.ToString().ToLowerInvariant()]
            }
            else {
                try {
                    [xml]$one = Get-GPOReport @serverParam -Guid $gpo.Id -ReportType Xml
                    $node = $one.SelectSingleNode("//*[local-name()='GPO']")
                }
                catch { Write-Log "  Report failed for '$($gpo.DisplayName)': $($_.Exception.Message)" 'WARN' }
            }

            # Enabled client-side extensions per scope (correct child-<Name> read).
            $computerExt = @(); $userExt = @()
            if ($node) {
                $computerExt = @($node.SelectNodes("*[local-name()='Computer']/*[local-name()='ExtensionData']/*[local-name()='Name']") | ForEach-Object { $_.InnerText })
                $userExt     = @($node.SelectNodes("*[local-name()='User']/*[local-name()='ExtensionData']/*[local-name()='Name']") | ForEach-Object { $_.InnerText })
            }
            $links = if ($node) { @($node.SelectNodes("*[local-name()='LinksTo']")) } else { @() }

            $gpoRows.Add([pscustomobject]@{
                Name              = $gpo.DisplayName
                Id                = $gpo.Id
                GpoStatus         = $gpo.GpoStatus
                Created           = $gpo.CreationTime
                Modified          = $gpo.ModificationTime
                ComputerEnabled   = $gpo.Computer.Enabled
                UserEnabled       = $gpo.User.Enabled
                ComputerExtensions= ($computerExt -join '; ')
                UserExtensions    = ($userExt -join '; ')
                WmiFilter         = $gpo.WmiFilter.Name
                LinkedOUCount     = $links.Count
            })

            foreach ($link in $links) {
                $linkRows.Add([pscustomobject]@{
                    GpoName     = $gpo.DisplayName
                    GpoId       = $gpo.Id
                    LinkedTo    = (Get-Inner $link 'SOMPath')
                    LinkEnabled = (Get-Inner $link 'Enabled')
                    Enforced    = (Get-Inner $link 'NoOverride')
                })
            }

            # Deep setting-level extraction (GPO -> Intune configuration mapping)
            if (-not $SkipGpoSettings -and $node) {
                try {
                    foreach ($row in (Expand-GpoSettings -GpoNode $node -GpoName $gpo.DisplayName)) {
                        $settingRows.Add($row)
                    }
                }
                catch { Write-Log "  Setting extraction failed for '$($gpo.DisplayName)': $($_.Exception.Message)" 'WARN' }
            }
        }
        Write-Progress -Activity "Processing GPOs" -Completed

        $Report['GPOs']      = $gpoRows.ToArray()
        $Report['GPO_Links'] = $linkRows.ToArray()
        if (-not $SkipGpoSettings) { $Report['GPO_Settings'] = $settingRows.ToArray() }
        Write-Log "  $($gpoRows.Count) GPOs, $($linkRows.Count) links$(if (-not $SkipGpoSettings) { ", $($settingRows.Count) individual settings" })."
        }  # end else (report-based path)
    }
    catch { Write-Log "GPO collection failed: $($_.Exception.Message)" 'ERROR' }
}

#endregion

#region -------------------------------------------------------------- Cloud M365
# CLOUD DATA CONTRACT
# -------------------
# Graph is connected first because Microsoft.Graph and ExchangeOnlineManagement can
# load different MSAL versions into the same PowerShell process. Workload failures
# are isolated with try/catch so one missing permission/module does not normally
# prevent unrelated workloads from being exported.

if ($IncludeCloud) {

    # ---- Connection order: GRAPH FIRST ---------------------------------
    # ExchangeOnlineManagement and Microsoft.Graph bundle conflicting versions of
    # MSAL (Microsoft.Identity.Client); whichever loads first wins the process.
    # If EXO loads first, its older MSAL breaks Connect-MgGraph ("Method not found
    # ...WithLogging"). Graph is the more fragile one AND used by most workloads,
    # so we import Graph.Authentication + connect Graph BEFORE touching EXO/IPPS,
    # which then connect lazily in their own blocks below.
    $ExoConnected = $false; $IppsConnected = $false

    # ---- Entra ID (Microsoft Graph) ------------------------------------
    # Establishes the delegated Graph session used by Entra, licensing, apps,
    # Intune and OneDrive collectors. The script deliberately avoids requesting
    # consent dynamically; tenant/user permissions determine readable data.
    if ($CloudWorkload -contains 'Entra' -or $CloudWorkload -contains 'Intune' -or $CloudWorkload -contains 'Apps' -or $CloudWorkload -contains 'Licensing' -or $CloudWorkload -contains 'OneDrive') {
        if (Test-Module Microsoft.Graph.Authentication) {
            Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
            try {
                # NOTE (S32): NO -Scopes. The signed-in account has read-only
                # directory roles and cannot consent to app scopes. Graph uses
                # whatever the role already permits.
                Write-Log "Connecting to Microsoft Graph (delegated, no -Scopes)..."
                if ($UseDeviceCode) {
                    Connect-MgGraph -UseDeviceCode -NoWelcome -ErrorAction Stop
                }
                else {
                    try {
                        Connect-MgGraph -NoWelcome -ErrorAction Stop
                    }
                    catch {
                        # WAM broker with no window handle surfaces as an NRE here;
                        # device code needs no window, so fall back to it automatically.
                        Write-Log "  Interactive Graph sign-in failed ($($_.Exception.Message)). Retrying with device code..." 'WARN'
                        Connect-MgGraph -UseDeviceCode -NoWelcome -ErrorAction Stop
                    }
                }
                $ctx = Get-MgContext
                if ($ctx) { $GraphConnected = $true; Write-Log "  Connected as $($ctx.Account)." }
            }
            catch { Write-Log "Graph connect failed: $($_.Exception.Message)" 'ERROR' }
            if (-not $GraphConnected) {
                Write-Log "Graph is not connected -- skipping all Graph workloads. This is almost always a stale PowerShell session: CLOSE the terminal, open a NEW one, and re-run (assembly conflicts can't be cleared any other way)." 'ERROR'
            }
        }
    }

    # If Graph sign-in failed, do not let every Graph collector emit misleading
    # empty sheets or cascaded errors. Non-Graph workloads (Exchange, Teams, SPO,
    # Purview) can still continue independently.
    $graphDependentWorkloads = @('Entra','Licensing','Apps','Intune','OneDrive')
    if (-not $GraphConnected) {
        $requestedGraphWorkloads = @($CloudWorkload | Where-Object { $_ -in $graphDependentWorkloads })
        if ($requestedGraphWorkloads.Count -gt 0) {
            Write-Log "Skipping Graph-dependent workloads because Graph is not connected: $($requestedGraphWorkloads -join ', ')." 'WARN'
            $CloudWorkload = @($CloudWorkload | Where-Object { $_ -notin $graphDependentWorkloads })
        }
    }

    # ---- Entra users / groups / devices --------------------------------
    # Outputs include Entra_Users, Entra_Groups and Entra_GroupMembership. Scoped
    # mode performs alias-aware matching so source and cloud UPN differences can
    # still correlate to the same migration identity.
    if ($CloudWorkload -contains 'Entra') {
        Write-Log "Collecting Entra users..."
        try {
            # NOTE: signInActivity intentionally omitted -- it needs AuditLog.Read.All
            # (and Entra P1) which the read-only account may lack, and requesting it
            # makes the whole Get-MgUser call fail. It isn't used in the output.
            # Raw REST via Invoke-MgGraphRequest (NOT Get-MgUser) -- the typed cmdlet
            # leaks memory across thousands of calls and OOMs on large waves. We
            # project each user to a slim object immediately so nothing heavy is held.
            $sel = 'id,displayName,userPrincipalName,mail,otherMails,proxyAddresses,accountEnabled,userType,onPremisesSyncEnabled,onPremisesImmutableId,department,jobTitle,employeeType,assignedLicenses,createdDateTime'
            $entraUsers = [System.Collections.Generic.List[object]]::new()
            if ($null -ne $TargetUserIds) {
                # Scoped: fetch the wave's users via Graph /$batch (20 per call), trying
                # UPN then mail then SMTP across rounds for the ones that miss.
                # Scan the directory once (paged, memory-safe) and keep users whose
                # UPN, mail, or otherMails is in the wave set. Reliable even when the
                # on-prem UPN differs from the cloud UPN -- Graph's /users/{key} path
                # only resolves UPN/id, so per-user lookups miss mail-only matches.
                Write-Verbose2 "Matching $($TargetUserIds.Count) wave users against the directory..."
                $ep = 0
                foreach ($h in (Invoke-GraphGetAll "v1.0/users?`$select=$sel&`$top=999")) {
                    $ep++
                    if ($ep % 500 -eq 0) { Write-Verbose2 "  ...scanned $ep directory users, matched $($entraUsers.Count)"; Invoke-LoopGc $ep }
                    $isMatch = $false

                    # Cross-tenant/source identities commonly survive as SMTP aliases
                    # even when the Entra UPN and mail attributes have changed. Include
                    # proxyAddresses in the match set, but only SMTP entries (not X500,
                    # SIP, etc.).
                    $cloudMatchKeys = @($h['userPrincipalName'], $h['mail']) + @($h['otherMails'])
                    foreach ($pa in @($h['proxyAddresses'])) {
                        if ($pa -and "$pa" -match '(?i)^smtp:(.+)$') { $cloudMatchKeys += $Matches[1] }
                    }
                    $matchedScopeKeys = [System.Collections.Generic.List[string]]::new()
                    foreach ($kk in $cloudMatchKeys) {
                        if ($kk -and $TargetUserKeys.ContainsKey("$kk".Trim().ToLowerInvariant())) {
                            $isMatch = $true
                            Add-UniqueString -Collection $matchedScopeKeys -Value $kk
                        }
                    }
                    if ($isMatch) {
                        $canonical = $null
                        foreach ($mk in $matchedScopeKeys) {
                            $rk = Resolve-IdentityKey $mk
                            if ($Identity.ContainsKey($rk) -and $Identity[$rk].OnPremDN) { $canonical = $rk; break }
                        }
                        if (-not $canonical -and $matchedScopeKeys.Count -gt 0) { $canonical = Resolve-IdentityKey $matchedScopeKeys[0] }
                        if ($canonical) {
                            foreach ($mk in $matchedScopeKeys) { Register-IdentityAlias -Alias $mk -Canonical $canonical }
                            foreach ($ck in $cloudMatchKeys) { if ($ck) { Register-IdentityAlias -Alias $ck -Canonical $canonical } }
                        }
                        $entraUsers.Add((ConvertFrom-GraphUser $h))
                    }
                }
                $miss = [math]::Max(0, ($TargetUserIds.Count - $entraUsers.Count))
                if ($miss -gt 0) { Write-Log "  $miss of $($TargetUserIds.Count) wave users not matched in Entra (not present, or different identity)." 'WARN' }
                elseif ($entraUsers.Count -gt $TargetUserIds.Count) { Write-Log "  Entra matching returned more directory objects than scoped identities; review duplicate/alias matches." 'WARN' }
            }
            else {
                Write-Verbose2 "Downloading the full Entra user list (paged REST)..."
                $ep = 0
                foreach ($h in (Invoke-GraphGetAll "v1.0/users?`$select=$sel&`$top=999")) {
                    $ep++
                    if ($ep % 200 -eq 0) { Write-Activity -Activity 'Processing Entra users' -Current $ep -Item $h['userPrincipalName']; Invoke-LoopGc $ep }
                    $entraUsers.Add((ConvertFrom-GraphUser $h))
                }
                Complete-Activity 'Processing Entra users'
            }
            $Report['Entra_Users'] = $entraUsers.ToArray()
            if ($null -ne $TargetUserIds) {
                $WaveUserIds = @{}
                foreach ($eu in $entraUsers) { if ($eu.ObjectId) { $WaveUserIds[$eu.ObjectId.ToString().ToLowerInvariant()] = $true } }
            }
            Write-Log "  $($entraUsers.Count) Entra users."
        }
        catch { Write-Log "Entra user collection failed at script line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)" 'ERROR' }

        Write-Log "Collecting Entra groups..."
        try {
            if ($null -ne $TargetUserIds) {
                # Scoped: only the groups the wave users belong to, via each user's
                # memberOf. Also produces per-user Entra_GroupMembership and fills
                # the Dependency_Map EntraGroupCount.
                Write-Verbose2 "Scoped: reading Entra group membership per wave user (memberOf, batched)..."
                $grpById = @{}
                $entraMemRows = [System.Collections.Generic.List[object]]::new()
                $srcU = @($Report['Entra_Users'])
                # One batched pass: /users/{id}/memberOf for every wave user.
                $moReqs = foreach ($u in $srcU) { if ($u.ObjectId) { @{ Key = $u.UserPrincipalName; Url = "/users/$($u.ObjectId)/memberOf?`$select=id,displayName,mail,groupTypes,securityEnabled,mailEnabled,onPremisesSyncEnabled,membershipRule&`$top=999" } } }
                $moRes = Invoke-GraphBatch @($moReqs | Where-Object { $_ }) 'Reading Entra group membership'
                foreach ($u in $srcU) {
                    $body = $moRes[$u.UserPrincipalName]
                    if (-not $body) { continue }
                    foreach ($o in @($body['value'])) {
                        if ($o['@odata.type'] -ne '#microsoft.graph.group') { continue }   # skip roles/admin units
                        $gname = $o['displayName']; $gid = $o['id']
                        $entraMemRows.Add([pscustomobject]@{
                            UserPrincipalName = $u.UserPrincipalName
                            GroupName         = $gname
                            GroupId           = $gid
                            GroupMail         = $o['mail']
                        })
                        if ($u.UserPrincipalName -and $gname) { Add-UniqueString -Collection ((Get-IdentityRecord $u.UserPrincipalName).EntraGroups) -Value $gname }
                        if ($gid -and -not $grpById.ContainsKey($gid)) { $grpById[$gid] = $o }
                    }
                }
                # Distinct groups the wave touches -> the (scoped) catalog sheet.
                $entraGroups = foreach ($g in $grpById.Values) {
                    $gt = @($g['groupTypes'])
                    [pscustomobject]@{
                        DisplayName     = $g['displayName']
                        Mail            = $g['mail']
                        GroupTypes      = ($gt -join ';')
                        IsM365Group     = ($gt -contains 'Unified')
                        IsDynamic       = ($gt -contains 'DynamicMembership')
                        SecurityEnabled = $g['securityEnabled']
                        MailEnabled     = $g['mailEnabled']
                        DirSynced       = $g['onPremisesSyncEnabled']
                        MembershipRule  = $g['membershipRule']
                        ObjectId        = $g['id']
                    }
                }
                $Report['Entra_Groups']          = @($entraGroups)
                $Report['Entra_GroupMembership'] = $entraMemRows.ToArray()
                # Wave group id/mail sets -> used to scope SharePoint sites and DLs.
                $WaveGroupIds = @{}; $WaveGroupMails = @{}
                foreach ($g in $grpById.Values) {
                    if ($g['id'])   { $WaveGroupIds[$g['id'].ToString().ToLowerInvariant()] = $true }
                    if ($g['mail']) { $WaveGroupMails[$g['mail'].ToString().ToLowerInvariant()] = $true }
                }
                Write-Log "  $(@($entraGroups).Count) Entra groups (scoped to wave), $($entraMemRows.Count) memberships."
            }
            else {
                # Unscoped: full tenant group catalog via paged REST (no per-user expansion).
                $entraGroups = foreach ($g in (Invoke-GraphGetAll "v1.0/groups?`$select=id,displayName,mail,groupTypes,securityEnabled,mailEnabled,onPremisesSyncEnabled,membershipRule&`$top=999")) {
                    $gt = @($g['groupTypes'])
                    [pscustomobject]@{
                        DisplayName     = $g['displayName']
                        Mail            = $g['mail']
                        GroupTypes      = ($gt -join ';')
                        IsM365Group     = ($gt -contains 'Unified')
                        IsDynamic       = ($gt -contains 'DynamicMembership')
                        SecurityEnabled = $g['securityEnabled']
                        MailEnabled     = $g['mailEnabled']
                        DirSynced       = $g['onPremisesSyncEnabled']
                        MembershipRule  = $g['membershipRule']
                        ObjectId        = $g['id']
                    }
                }
                $Report['Entra_Groups'] = @($entraGroups)
                Write-Log "  $(@($entraGroups).Count) Entra groups."
            }
        }
        catch { Write-Log "Entra group collection failed: $($_.Exception.Message)" 'ERROR' }

        # -- Conditional Access policies + Named Locations -------------------
        # Output : CA_Policies, CA_NamedLocations
        # In scoped mode, policies are retained when they apply to the wave through
        # explicit users/groups or broad All-user targeting.
        if (Test-Module Microsoft.Graph.Identity.SignIns) {
            Import-Module Microsoft.Graph.Identity.SignIns -ErrorAction SilentlyContinue
            Write-Log "Collecting Conditional Access policies..."
            try {
                # When scoping, keep only policies that target the wave (included
                # users or groups), plus any that apply to "All" users.
                $caPolicies = Get-MgIdentityConditionalAccessPolicy -All -ErrorAction Stop |
                    Where-Object { Test-CaInWave $_.Conditions } |
                    ForEach-Object {
                        $c = $_.Conditions
                        [pscustomobject]@{
                            DisplayName        = $_.DisplayName
                            State              = $_.State
                            IncludeUsers       = ($c.Users.IncludeUsers      -join ';')
                            ExcludeUsers       = ($c.Users.ExcludeUsers      -join ';')
                            IncludeGroups      = ($c.Users.IncludeGroups     -join ';')
                            ExcludeGroups      = ($c.Users.ExcludeGroups     -join ';')
                            IncludeRoles       = ($c.Users.IncludeRoles      -join ';')
                            IncludeApps        = ($c.Applications.IncludeApplications -join ';')
                            ExcludeApps        = ($c.Applications.ExcludeApplications -join ';')
                            UserActions        = ($c.Applications.IncludeUserActions   -join ';')
                            ClientAppTypes     = ($c.ClientAppTypes         -join ';')
                            Platforms          = ($c.Platforms.IncludePlatforms -join ';')
                            IncludeLocations   = ($c.Locations.IncludeLocations -join ';')
                            ExcludeLocations   = ($c.Locations.ExcludeLocations -join ';')
                            SignInRiskLevels   = ($c.SignInRiskLevels       -join ';')
                            UserRiskLevels     = ($c.UserRiskLevels         -join ';')
                            GrantControls      = ($_.GrantControls.BuiltInControls -join ';')
                            GrantOperator      = $_.GrantControls.Operator
                            SessionControls    = (($_.SessionControls.PSObject.Properties | Where-Object { $_.Value }).Name -join ';')
                            PolicyId           = $_.Id
                        }
                    }
                $Report['CA_Policies'] = @($caPolicies)
                Write-Log "  $(@($caPolicies).Count) CA policies$(if ($null -ne $TargetUserKeys) { ' (scoped to wave users/groups + All-user policies)' })."
            }
            catch { Write-Log "Conditional Access collection failed: $($_.Exception.Message)" 'ERROR' }

            Write-Log "Collecting Named Locations..."
            try {
                $namedLocations = Get-MgIdentityConditionalAccessNamedLocation -All -ErrorAction Stop |
                    ForEach-Object {
                        # AdditionalProperties can be null; guard before indexing.
                        $ap = $_.AdditionalProperties
                        if ($null -eq $ap) { $ap = @{} }
                        $odata = if ($ap.ContainsKey('@odata.type')) { $ap['@odata.type'] -replace '#microsoft.graph.', '' } else { $null }
                        $ips = @()
                        if ($ap.ContainsKey('ipRanges') -and $ap['ipRanges']) {
                            $ips = foreach ($r in @($ap['ipRanges'])) {
                                if ($r -is [System.Collections.IDictionary]) { $r['cidrAddress'] }
                                elseif ($r.PSObject.Properties['cidrAddress']) { $r.cidrAddress }
                            }
                        }
                        $countries = if ($ap.ContainsKey('countriesAndRegions') -and $ap['countriesAndRegions']) { @($ap['countriesAndRegions']) } else { @() }
                        [pscustomobject]@{
                            DisplayName  = $_.DisplayName
                            LocationType = $odata
                            IsTrusted    = if ($ap.ContainsKey('isTrusted')) { $ap['isTrusted'] } else { $null }
                            IpRanges     = (($ips | Where-Object { $_ }) -join ';')
                            Countries    = ($countries -join ';')
                            LocationId   = $_.Id
                        }
                    }
                $Report['CA_NamedLocations'] = @($namedLocations)
                Write-Log "  $(@($namedLocations).Count) named locations."
            }
            catch { Write-Log "Named Location collection failed: $($_.Exception.Message)" 'ERROR' }
        }
        else {
            Write-Log "Microsoft.Graph.Identity.SignIns not installed; skipping Conditional Access." 'WARN'
        }
    }

    # ---- Licensing (tenant SKU capacity + reconciled assignment detail) --
    # Output : License_Summary, License_UserDetail, License_GroupAssignments
    #
    # IMPORTANT LICENSING MODEL
    # -------------------------
    # Get-MgSubscribedSku is the authoritative tenant subscription/capacity view.
    # Its ConsumedUnits value already includes licenses effectively consumed by users
    # through both direct assignment and group-based licensing. The older version of
    # this script exposed ConsumedUnits but did not explain where those assignments
    # came from, which made the summary difficult to reconcile with the M365/Entra
    # admin portals when group-based licensing was in use.
    #
    # To make the report auditable, this block now performs a lightweight tenant-wide
    # user scan of assignedLicenses + licenseAssignmentStates and independently counts
    # effective assignments by SKU. licenseAssignmentStates.assignedByGroup identifies
    # direct vs inherited/group-based assignments and also exposes assignment errors.
    #
    # Scope behavior:
    #   * License_Summary is ALWAYS tenant-wide because it represents subscription
    #     capacity/consumption and must not change when -ScopeToInputUsers is used.
    #   * License_GroupAssignments is ALWAYS tenant-wide because licensing groups are
    #     tenant configuration objects.
    #   * License_UserDetail is scoped to the migration wave when -ScopeToInputUsers
    #     is used; otherwise it contains every licensed user in the tenant.
    if ($CloudWorkload -contains 'Licensing') {
        Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction SilentlyContinue

        # Normalize every SkuId to lower-case string form. Graph SDK objects often expose
        # subscribed SKU ids as [Guid] while REST user assignments expose strings; using
        # a normalized key prevents friendly-name lookups from silently missing.
        $skuMap = @{}
        $skuRaw = @()
        Write-Log "Collecting tenant SKUs..."
        try {
            $skuRaw = @(Get-MgSubscribedSku -All -ErrorAction Stop)
            foreach ($sku in $skuRaw) {
                if ($sku.SkuId) {
                    $skuKey = $sku.SkuId.ToString().Trim().ToLowerInvariant()
                    $skuMap[$skuKey] = $sku.SkuPartNumber
                }
            }
            Write-Log "  $($skuRaw.Count) subscribed SKUs."
        }
        catch {
            Write-Log "SKU inventory failed: $($_.Exception.Message)" 'ERROR'
            $skuRaw = @()
        }

        # -----------------------------------------------------------------
        # Tenant-wide effective user assignment reconciliation.
        # -----------------------------------------------------------------
        # This deliberately runs tenant-wide even during a scoped migration-wave run.
        # Otherwise a tenant-level License_Summary would incorrectly compare the SKU's
        # global ConsumedUnits against only the users in the current migration wave.
        $tenantLicenseUsers = [System.Collections.Generic.List[object]]::new()
        $tenantSkuUsage = @{}
        $licenseReconciliationAvailable = $false
        $licenseAssignmentStateAvailable = $true

        Write-Log "Reconciling tenant license assignments (direct + group-based)..."
        try {
            $licenseUserUri = "v1.0/users?`$select=id,displayName,userPrincipalName,mail,otherMails,proxyAddresses,assignedLicenses,licenseAssignmentStates&`$top=999"
            try {
                $tenantGraphLicenseUsers = @(Invoke-GraphGetAll $licenseUserUri)
            }
            catch {
                # Some restricted readers can enumerate assignedLicenses but cannot read
                # licenseAssignmentStates. Preserve an accurate effective-assignment count
                # and mark assignment-source detail as unavailable rather than failing the
                # whole licensing workload.
                $licenseAssignmentStateAvailable = $false
                Write-Log "  licenseAssignmentStates unavailable ($($_.Exception.Message)); retrying with effective assignedLicenses only." 'WARN'
                $tenantGraphLicenseUsers = @(Invoke-GraphGetAll "v1.0/users?`$select=id,displayName,userPrincipalName,mail,otherMails,proxyAddresses,assignedLicenses&`$top=999")
            }

            $tu = 0; $tuTotal = $tenantGraphLicenseUsers.Count
            foreach ($h in $tenantGraphLicenseUsers) {
                $tu++
                if ($tu % 500 -eq 0) { Write-Activity -Activity 'Reconciling tenant licenses' -Current $tu -Total $tuTotal -Item $h['userPrincipalName']; Invoke-LoopGc $tu }

                $assignedSkuKeys = @(
                    @($h['assignedLicenses']) |
                        ForEach-Object {
                            $sid = if ($_ -is [System.Collections.IDictionary]) { $_['skuId'] } else { $_.skuId }
                            if ($sid) { $sid.ToString().Trim().ToLowerInvariant() }
                        } |
                        Where-Object { $_ } |
                        Sort-Object -Unique
                )
                if ($assignedSkuKeys.Count -eq 0) { continue }

                $states = if ($licenseAssignmentStateAvailable) { @($h['licenseAssignmentStates']) } else { @() }
                $assignmentRecords = [System.Collections.Generic.List[object]]::new()

                foreach ($skuKey in $assignedSkuKeys) {
                    $direct = $false
                    $groupIds = [System.Collections.Generic.List[string]]::new()
                    $stateNames = [System.Collections.Generic.List[string]]::new()
                    $errorNames = [System.Collections.Generic.List[string]]::new()

                    foreach ($st in $states) {
                        $stateSku = if ($st -is [System.Collections.IDictionary]) { $st['skuId'] } else { $st.skuId }
                        if (-not $stateSku -or $stateSku.ToString().Trim().ToLowerInvariant() -ne $skuKey) { continue }

                        $assignedByGroup = if ($st -is [System.Collections.IDictionary]) { $st['assignedByGroup'] } else { $st.assignedByGroup }
                        if ([string]::IsNullOrWhiteSpace([string]$assignedByGroup)) {
                            $direct = $true
                        }
                        else {
                            Add-UniqueString -Collection $groupIds -Value $assignedByGroup.ToString().Trim().ToLowerInvariant()
                        }

                        $stateName = if ($st -is [System.Collections.IDictionary]) { $st['state'] } else { $st.state }
                        if ($stateName) { Add-UniqueString -Collection $stateNames -Value $stateName.ToString() }

                        $stateError = if ($st -is [System.Collections.IDictionary]) { $st['error'] } else { $st.error }
                        if ($stateError) { Add-UniqueString -Collection $errorNames -Value $stateError.ToString() }
                        if ($stateName -and $stateName.ToString() -in @('Error','ActiveWithError')) {
                            if (-not $stateError) { Add-UniqueString -Collection $errorNames -Value $stateName.ToString() }
                        }
                    }

                    $groupBased = ($groupIds.Count -gt 0)
                    $source = if (-not $licenseAssignmentStateAvailable) { 'Unknown' }
                              elseif ($direct -and $groupBased) { 'Direct+Group' }
                              elseif ($groupBased) { 'Group' }
                              elseif ($direct) { 'Direct' }
                              else { 'Unknown' }

                    if (-not $tenantSkuUsage.ContainsKey($skuKey)) {
                        $tenantSkuUsage[$skuKey] = [ordered]@{
                            EffectiveAssignedUsers       = 0
                            DirectAssignedUsers          = 0
                            GroupAssignedUsers           = 0
                            DirectOnlyUsers              = 0
                            GroupOnlyUsers               = 0
                            DirectAndGroupUsers          = 0
                            AssignmentSourceUnknownUsers = 0
                            UsersWithAssignmentErrors    = 0
                            GroupIds                     = @{}
                        }
                    }
                    $usage = $tenantSkuUsage[$skuKey]
                    $usage.EffectiveAssignedUsers++
                    if ($direct)     { $usage.DirectAssignedUsers++ }
                    if ($groupBased) { $usage.GroupAssignedUsers++ }
                    if ($direct -and $groupBased)       { $usage.DirectAndGroupUsers++ }
                    elseif ($direct)                    { $usage.DirectOnlyUsers++ }
                    elseif ($groupBased)                { $usage.GroupOnlyUsers++ }
                    else                                { $usage.AssignmentSourceUnknownUsers++ }
                    if ($errorNames.Count -gt 0)        { $usage.UsersWithAssignmentErrors++ }
                    foreach ($gid in $groupIds)         { $usage.GroupIds[$gid] = $true }

                    $assignmentRecords.Add([pscustomobject]@{
                        SkuKey          = $skuKey
                        AssignmentSource= $source
                        Direct          = $direct
                        GroupIds        = @($groupIds.ToArray())
                        States          = @($stateNames.ToArray())
                        Errors          = @($errorNames.ToArray())
                    })
                }

                $tenantLicenseUsers.Add([pscustomobject]@{
                    ObjectId          = $h['id']
                    DisplayName       = $h['displayName']
                    UserPrincipalName = $h['userPrincipalName']
                    Mail              = $h['mail']
                    OtherMails        = @($h['otherMails'])
                    ProxyAddresses    = @($h['proxyAddresses'])
                    Assignments       = @($assignmentRecords.ToArray())
                })
            }
            Complete-Activity 'Reconciling tenant licenses'
            $licenseReconciliationAvailable = $true
            Write-Log "  Reconciled $($tenantLicenseUsers.Count) licensed tenant users."
        }
        catch {
            Write-Log "Tenant license reconciliation failed: $($_.Exception.Message)" 'WARN'
            $tenantLicenseUsers = [System.Collections.Generic.List[object]]::new()
            $tenantSkuUsage = @{}
            $licenseReconciliationAvailable = $false
        }

        # -----------------------------------------------------------------
        # Group-based license configuration.
        # -----------------------------------------------------------------
        # A group license consumes units through its users, not as one unit for the
        # group object itself. This sheet records the configuration relationship:
        # group -> assigned SKU. The per-user licenseAssignmentStates above shows which
        # users actually inherit the SKU from which groups.
        $licenseGroupRows = [System.Collections.Generic.List[object]]::new()
        $licenseGroupNameById = @{}
        $groupCountBySku = @{}
        Write-Log "Collecting group-based license assignments..."
        try {
            $groupUri = "v1.0/groups?`$select=id,displayName,mail,groupTypes,securityEnabled,mailEnabled,onPremisesSyncEnabled,assignedLicenses&`$top=999"
            foreach ($g in (Invoke-GraphGetAll $groupUri)) {
                $gid = if ($g['id']) { $g['id'].ToString().Trim().ToLowerInvariant() } else { $null }
                if ($gid) { $licenseGroupNameById[$gid] = $g['displayName'] }

                foreach ($al in @($g['assignedLicenses'])) {
                    $sid = if ($al -is [System.Collections.IDictionary]) { $al['skuId'] } else { $al.skuId }
                    if (-not $sid) { continue }
                    $skuKey = $sid.ToString().Trim().ToLowerInvariant()
                    $disabledPlans = if ($al -is [System.Collections.IDictionary]) { @($al['disabledPlans']) } else { @($al.disabledPlans) }
                    $gt = @($g['groupTypes'])

                    if (-not $groupCountBySku.ContainsKey($skuKey)) { $groupCountBySku[$skuKey] = @{} }
                    if ($gid) { $groupCountBySku[$skuKey][$gid] = $true }

                    $licenseGroupRows.Add([pscustomobject]@{
                        GroupDisplayName          = $g['displayName']
                        GroupId                   = $g['id']
                        GroupMail                 = $g['mail']
                        IsM365Group               = ($gt -contains 'Unified')
                        IsDynamic                 = ($gt -contains 'DynamicMembership')
                        SecurityEnabled           = $g['securityEnabled']
                        MailEnabled               = $g['mailEnabled']
                        DirSynced                 = $g['onPremisesSyncEnabled']
                        SkuPartNumber             = $(if ($skuMap.ContainsKey($skuKey)) { $skuMap[$skuKey] } else { $skuKey })
                        SkuId                     = $sid
                        DisabledServicePlanCount  = @($disabledPlans).Count
                        DisabledServicePlanIds    = (@($disabledPlans) -join ';')
                    })
                }
            }
            $Report['License_GroupAssignments'] = $licenseGroupRows.ToArray()
            Write-Log "  $($licenseGroupRows.Count) group-to-SKU license assignments."
        }
        catch {
            # Do not write an empty report on an API/permission failure: in merge mode
            # that would delete a previously valid worksheet and make a collection
            # failure look like a genuine zero-result refresh.
            Write-Log "Group-based license configuration collection failed: $($_.Exception.Message)" 'WARN'
        }

        # -----------------------------------------------------------------
        # Tenant SKU summary with independent reconciliation.
        # -----------------------------------------------------------------
        try {
            $skus = foreach ($sku in $skuRaw) {
                $skuKey = if ($sku.SkuId) { $sku.SkuId.ToString().Trim().ToLowerInvariant() } else { '' }
                $enabled   = [int]$sku.PrepaidUnits.Enabled
                $consumed  = [int]$sku.ConsumedUnits
                $warning   = [int]$sku.PrepaidUnits.Warning
                $suspended = [int]$sku.PrepaidUnits.Suspended
                $lockedOut = 0
                if ($null -ne $sku.PrepaidUnits.LockedOut) { $lockedOut = [int]$sku.PrepaidUnits.LockedOut }
                $available = $enabled - $consumed

                $usage = if ($skuKey -and $tenantSkuUsage.ContainsKey($skuKey)) { $tenantSkuUsage[$skuKey] } else { $null }
                $effective = if ($licenseReconciliationAvailable) { if ($usage) { [int]$usage.EffectiveAssignedUsers } else { 0 } } else { $null }
                $delta = $null
                $reconciliationStatus = 'Unavailable'
                if ($licenseReconciliationAvailable) {
                    if ($sku.AppliesTo -and $sku.AppliesTo.ToString() -ne 'User') {
                        $reconciliationStatus = 'NotApplicable'
                    }
                    else {
                        $delta = $consumed - $effective
                        $reconciliationStatus = $(if ($delta -eq 0) { 'Matched' } else { 'Mismatch' })
                    }
                }

                $subscriptionIds = @($sku.SubscriptionIds)
                $configuredGroups = if ($skuKey -and $groupCountBySku.ContainsKey($skuKey)) { @($groupCountBySku[$skuKey].Keys).Count } else { 0 }
                $effectiveGroups = if ($licenseReconciliationAvailable -and $usage) { @($usage.GroupIds.Keys).Count } elseif ($licenseReconciliationAvailable) { 0 } else { $null }

                [pscustomobject]@{
                    SkuPartNumber                = $sku.SkuPartNumber
                    SkuId                        = $sku.SkuId
                    AppliesTo                    = $sku.AppliesTo
                    CapabilityStatus             = $sku.CapabilityStatus
                    AccountName                  = $sku.AccountName
                    Enabled                      = $enabled
                    Consumed                     = $consumed
                    Available                    = $available
                    Warning                      = $warning
                    Suspended                    = $suspended
                    LockedOut                    = $lockedOut
                    OverAllocated                = ($available -lt 0)
                    EffectiveAssignedUsers       = $effective
                    DirectAssignedUsers          = $(if ($licenseReconciliationAvailable -and $usage) { [int]$usage.DirectAssignedUsers } elseif ($licenseReconciliationAvailable) { 0 } else { $null })
                    GroupAssignedUsers           = $(if ($licenseReconciliationAvailable -and $usage) { [int]$usage.GroupAssignedUsers } elseif ($licenseReconciliationAvailable) { 0 } else { $null })
                    DirectOnlyUsers              = $(if ($licenseReconciliationAvailable -and $usage) { [int]$usage.DirectOnlyUsers } elseif ($licenseReconciliationAvailable) { 0 } else { $null })
                    GroupOnlyUsers               = $(if ($licenseReconciliationAvailable -and $usage) { [int]$usage.GroupOnlyUsers } elseif ($licenseReconciliationAvailable) { 0 } else { $null })
                    DirectAndGroupUsers          = $(if ($licenseReconciliationAvailable -and $usage) { [int]$usage.DirectAndGroupUsers } elseif ($licenseReconciliationAvailable) { 0 } else { $null })
                    AssignmentSourceUnknownUsers = $(if ($licenseReconciliationAvailable -and $usage) { [int]$usage.AssignmentSourceUnknownUsers } elseif ($licenseReconciliationAvailable) { 0 } else { $null })
                    UsersWithAssignmentErrors    = $(if ($licenseReconciliationAvailable -and $usage) { [int]$usage.UsersWithAssignmentErrors } elseif ($licenseReconciliationAvailable) { 0 } else { $null })
                    LicensingGroupCount          = $configuredGroups
                    EffectiveLicensingGroupCount = $effectiveGroups
                    ReconciliationDelta          = $delta
                    ReconciliationStatus         = $reconciliationStatus
                    SubscriptionCount            = $subscriptionIds.Count
                    SubscriptionIds              = ($subscriptionIds -join ';')
                    ServicePlanCount              = @($sku.ServicePlans).Count
                }
            }
            $Report['License_Summary'] = @($skus)
            $mismatchCount = @($skus | Where-Object { $_.ReconciliationStatus -eq 'Mismatch' }).Count
            Write-Log "  License summary built for $(@($skus).Count) SKUs; $mismatchCount reconciliation mismatch(es)."
        }
        catch { Write-Log "SKU summary build failed: $($_.Exception.Message)" 'ERROR' }

        # -----------------------------------------------------------------
        # Per-user license detail.
        # -----------------------------------------------------------------
        # When scoping, this is the migration-wave view. The tenant-wide scan above is
        # still used as its source so group/direct attribution is consistent with the
        # tenant summary and no separate per-user Graph calls are required.
        Write-Log "Building per-user license detail..."
        if (-not $licenseReconciliationAvailable) {
            Write-Log "  Per-user license detail not refreshed because tenant license enumeration failed." 'WARN'
        }
        else {
        try {
            $licDetail = foreach ($u in $tenantLicenseUsers) {
                $inScope = ($null -eq $TargetUserKeys)
                if (-not $inScope) {
                    $candidateKeys = @($u.UserPrincipalName, $u.Mail) + @($u.OtherMails)
                    foreach ($pa in @($u.ProxyAddresses)) {
                        if ($pa -and "$pa" -match '(?i)^smtp:(.+)$') { $candidateKeys += $Matches[1] }
                    }
                    foreach ($ck in $candidateKeys) {
                        if ($ck -and (Test-TargetUser $ck)) { $inScope = $true; break }
                    }
                }
                if (-not $inScope) { continue }

                $friendly = [System.Collections.Generic.List[string]]::new()
                $directLicenses = [System.Collections.Generic.List[string]]::new()
                $groupLicenses = [System.Collections.Generic.List[string]]::new()
                $mixedLicenses = [System.Collections.Generic.List[string]]::new()
                $licensingGroups = [System.Collections.Generic.List[string]]::new()
                $assignmentErrors = [System.Collections.Generic.List[string]]::new()

                foreach ($a in @($u.Assignments)) {
                    $skuName = if ($skuMap.ContainsKey($a.SkuKey)) { $skuMap[$a.SkuKey] } else { $a.SkuKey }
                    Add-UniqueString -Collection $friendly -Value $skuName
                    if ($a.AssignmentSource -eq 'Direct')      { Add-UniqueString -Collection $directLicenses -Value $skuName }
                    elseif ($a.AssignmentSource -eq 'Group')  { Add-UniqueString -Collection $groupLicenses -Value $skuName }
                    elseif ($a.AssignmentSource -eq 'Direct+Group') {
                        Add-UniqueString -Collection $directLicenses -Value $skuName
                        Add-UniqueString -Collection $groupLicenses -Value $skuName
                        Add-UniqueString -Collection $mixedLicenses -Value $skuName
                    }

                    foreach ($gid in @($a.GroupIds)) {
                        $gname = if ($licenseGroupNameById.ContainsKey($gid)) { $licenseGroupNameById[$gid] } else { $gid }
                        Add-UniqueString -Collection $licensingGroups -Value "$gname [$skuName]"
                    }
                    foreach ($err in @($a.Errors)) { Add-UniqueString -Collection $assignmentErrors -Value "$skuName`: $err" }
                }

                $key = $u.UserPrincipalName
                if ($key) {
                    $licIdentity = Get-IdentityRecord $key
                    if ($licIdentity) { $licIdentity.Licenses = ($friendly -join ';') }
                }

                [pscustomobject]@{
                    DisplayName             = $u.DisplayName
                    UserPrincipalName       = $u.UserPrincipalName
                    LicenseCount            = $friendly.Count
                    Licenses                = ($friendly -join '; ')
                    DirectLicenseCount      = $directLicenses.Count
                    DirectLicenses          = ($directLicenses -join '; ')
                    GroupBasedLicenseCount  = $groupLicenses.Count
                    GroupBasedLicenses      = ($groupLicenses -join '; ')
                    DirectAndGroupCount     = $mixedLicenses.Count
                    DirectAndGroupLicenses  = ($mixedLicenses -join '; ')
                    LicensingGroups         = ($licensingGroups -join '; ')
                    LicenseAssignmentErrors = ($assignmentErrors -join '; ')
                }
            }
            $Report['License_UserDetail'] = @($licDetail)
            Write-Log "  $(@($licDetail).Count) licensed users$(if ($null -ne $TargetUserKeys) { ' in migration scope' } else { ' tenant-wide' })."
        }
        catch { Write-Log "Per-user license detail failed: $($_.Exception.Message)" 'ERROR' }
        }
    }

    # ---- App registrations, enterprise apps, consents ------------------
    # Output : Entra_AppRegistrations, Entra_EnterpriseApps, Entra_OAuthGrants,
    #          Entra_AppRoleGrants, Entra_UserAppAssignments
    # Captures credentials/expiry, enterprise apps, delegated consent, application
    # permission grants and user-to-app assignments.
    # None of these move in a T2T: app IDs, secrets and certs are per-tenant,
    # and every delegated/application consent must be re-granted in the target.
    # Expiring secrets/certs are a hard cutover dependency.
    if ($CloudWorkload -contains 'Apps') {
        if (Test-Module Microsoft.Graph.Applications) {
            Import-Module Microsoft.Graph.Applications -ErrorAction SilentlyContinue
            $now = Get-Date

            # Scoped pre-pass: which enterprise apps do the wave users actually use?
            # (per-user app role assignments). Drives the app/consent scoping below.
            $usedSpIds = $null
            if ($null -ne $TargetUserIds) {
                Import-Module Microsoft.Graph.Users -ErrorAction SilentlyContinue
                Write-Verbose2 "Scoped: reading app assignments per wave user..."
                $usedSpIds = @{}
                $uaRows = [System.Collections.Generic.List[object]]::new()
                $srcAU = if ($Report['Entra_Users']) { @($Report['Entra_Users']) } else { @(Get-ScopedGraphUsersForWave) }
                # Apps-only scoped runs still need the wave's object-id set for
                # delegated-consent filtering later in this block.
                if (-not $WaveUserIds) {
                    $WaveUserIds = @{}
                    foreach ($wu in $srcAU) { if ($wu.ObjectId) { $WaveUserIds[$wu.ObjectId.ToString().ToLowerInvariant()] = $true } }
                }
                # One batched pass: /users/{id}/appRoleAssignments for every wave user.
                $arReqs = foreach ($u in $srcAU) { if ($u.ObjectId) { @{ Key = $u.UserPrincipalName; Url = "/users/$($u.ObjectId)/appRoleAssignments" } } }
                $arRes = Invoke-GraphBatch @($arReqs | Where-Object { $_ }) 'Reading user app assignments'
                foreach ($u in $srcAU) {
                    $body = $arRes[$u.UserPrincipalName]
                    if (-not $body) { continue }
                    foreach ($a in @($body['value'])) {
                        $rid = $a['resourceId']
                        $uaRows.Add([pscustomobject]@{
                            UserPrincipalName = $u.UserPrincipalName
                            AppDisplayName    = $a['resourceDisplayName']
                            ResourceSpId      = $rid
                            AppRoleId         = $a['appRoleId']
                            AssignedOn        = $a['createdDateTime']
                        })
                        if ($rid) { $usedSpIds[$rid.ToString().ToLowerInvariant()] = $true }
                    }
                }
                $Report['Entra_UserAppAssignments'] = $uaRows.ToArray()
                Write-Log "  $($uaRows.Count) user->app assignments; $($usedSpIds.Count) enterprise apps used by the wave."
            }

            # The service principals we care about: scoped to wave-used apps (batched
            # by id), else all (paged REST). Hashtables; reused for EnterpriseApps + app-roles.
            $spSel = 'id,displayName,appId,servicePrincipalType,accountEnabled,appOwnerOrganizationId,homepage,replyUrls,tags'
            $spObjs = @()
            try {
                if ($null -ne $usedSpIds) {
                    $spReqs = foreach ($id in $usedSpIds.Keys) { @{ Key = $id; Url = "/servicePrincipals/$id`?`$select=$spSel" } }
                    $spObjs = @((Invoke-GraphBatch @($spReqs) 'Fetching enterprise apps').Values)
                }
                else {
                    Write-Verbose2 "Enumerating all service principals (paged REST)..."
                    $spObjs = @(Invoke-GraphGetAll "v1.0/servicePrincipals?`$select=$spSel&`$top=999")
                }
            }
            catch { Write-Log "Service principal enumeration failed: $($_.Exception.Message)" 'ERROR' }
            $spObjs = @($spObjs)

            Write-Log "Collecting app registrations (owned apps)..."
            try {
                $appSel = 'id,appId,displayName,signInAudience,passwordCredentials,keyCredentials,requiredResourceAccess,web,spa,publicClient,createdDateTime'
                $apps = foreach ($h in (Invoke-GraphGetAll "v1.0/applications?`$select=$appSel&`$top=999")) {
                    $allExp = @(@($h['passwordCredentials'] | ForEach-Object { $_['endDateTime'] }) + @($h['keyCredentials'] | ForEach-Object { $_['endDateTime'] }) | Where-Object { $_ })
                    $nextExpRaw = ($allExp | Sort-Object | Select-Object -First 1)
                    $nextExp = $null; if ($nextExpRaw) { try { $nextExp = [datetime]$nextExpRaw } catch { $nextExp = $nextExpRaw } }
                    # API permissions requested (delegated + application), by resource
                    $perms = foreach ($ra in @($h['requiredResourceAccess'])) {
                        $roleCount  = @($ra['resourceAccess'] | Where-Object { $_['type'] -eq 'Role' }).Count
                        $scopeCount = @($ra['resourceAccess'] | Where-Object { $_['type'] -eq 'Scope' }).Count
                        "$($ra['resourceAppId']): $roleCount app / $scopeCount delegated"
                    }
                    $redir = @()
                    foreach ($seg in 'web','spa','publicClient') { if ($h[$seg]) { $redir += @($h[$seg]['redirectUris']) } }
                    [pscustomobject]@{
                        DisplayName        = $h['displayName']
                        AppId              = $h['appId']
                        ObjectId           = $h['id']
                        SignInAudience     = $h['signInAudience']
                        SecretCount        = @($h['passwordCredentials']).Count
                        CertCount          = @($h['keyCredentials']).Count
                        NextCredentialExpiry = $nextExp
                        CredentialExpired  = ($nextExp -is [datetime] -and $nextExp -lt $now)
                        RedirectUris       = ((@($redir) | Where-Object { $_ }) -join ';')
                        RequiredApiPerms   = ($perms -join ' | ')
                        CreatedDateTime    = $h['createdDateTime']
                    }
                }
                $Report['Entra_AppRegistrations'] = @($apps)
                Write-Log "  $(@($apps).Count) app registrations$(if ($null -ne $TargetUserKeys) { ' (tenant-wide dev artifacts; not wave-scoped)' })."
            }
            catch { Write-Log "App registration collection failed: $($_.Exception.Message)" 'ERROR' }

            Write-Log "Collecting enterprise apps (service principals)..."
            try {
                $sps = $spObjs | ForEach-Object {
                    [pscustomobject]@{
                        DisplayName        = $_['displayName']
                        AppId              = $_['appId']
                        ObjectId           = $_['id']
                        SpType             = $_['servicePrincipalType']
                        AccountEnabled     = $_['accountEnabled']
                        AppOwnerOrgId      = $_['appOwnerOrganizationId']
                        # first-party MS apps share this well-known tenant id
                        IsThirdParty       = ($_['appOwnerOrganizationId'] -and $_['appOwnerOrganizationId'] -ne 'f8cdef31-a31e-4b4a-93e4-5f571e91255a')
                        Homepage           = $_['homepage']
                        ReplyUrls          = (@($_['replyUrls']) -join ';')
                        Tags               = (@($_['tags']) -join ';')
                    }
                }
                $Report['Entra_EnterpriseApps'] = @($sps)
                Write-Log "  $(@($sps).Count) service principals$(if ($null -ne $usedSpIds) { ' (scoped to wave-used apps)' })."
            }
            catch { Write-Log "Enterprise app collection failed: $($_.Exception.Message)" 'ERROR' }

            Write-Log "Collecting delegated consents (OAuth2 grants)..."
            try {
                # Scoped: admin-wide (AllPrincipals) consents + consents by wave users.
                $grants = foreach ($h in (Invoke-GraphGetAll "v1.0/oauth2PermissionGrants?`$top=999")) {
                    if (-not (($null -eq $TargetUserKeys) -or $h['consentType'] -eq 'AllPrincipals' -or (Test-AnyIn $h['principalId'] $WaveUserIds))) { continue }
                    [pscustomobject]@{
                        ClientSpId    = $h['clientId']
                        ResourceSpId  = $h['resourceId']
                        ConsentType   = $h['consentType']      # AllPrincipals (admin) / Principal (user)
                        PrincipalId   = $h['principalId']
                        Scopes        = $h['scope']
                    }
                }
                $Report['Entra_OAuthGrants'] = @($grants)
                Write-Log "  $(@($grants).Count) delegated consent grants$(if ($null -ne $TargetUserKeys) { ' (admin + wave-user)' })."
            }
            catch { Write-Log "OAuth2 grant collection failed: $($_.Exception.Message)" 'ERROR' }

            Write-Log "Collecting application-permission consents (app role assignments)..."
            try {
                # Batch the in-scope service principals' app-role assignments.
                $appRoles = [System.Collections.Generic.List[object]]::new()
                $arReqs = foreach ($sp in $spObjs) { @{ Key = $sp['id']; Url = "/servicePrincipals/$($sp['id'])/appRoleAssignments" } }
                $arRes = Invoke-GraphBatch @($arReqs) 'Reading app-role assignments'
                foreach ($sp in $spObjs) {
                    $body = $arRes[$sp['id']]
                    if (-not $body) { continue }
                    foreach ($a in @($body['value'])) {
                        $appRoles.Add([pscustomobject]@{
                            ClientApp     = $sp['displayName']
                            ClientSpId    = $sp['id']
                            Resource      = $a['resourceDisplayName']
                            ResourceSpId  = $a['resourceId']
                            AppRoleId     = $a['appRoleId']
                            AssignedOn    = $a['createdDateTime']
                        })
                    }
                }
                $Report['Entra_AppRoleGrants'] = $appRoles.ToArray()
                Write-Log "  $($appRoles.Count) application-permission grants."
            }
            catch { Write-Log "App role assignment collection failed: $($_.Exception.Message)" 'ERROR' }
        }
        else { Write-Log "Microsoft.Graph.Applications not installed; skipping Apps." 'WARN' }
    }

    # ---- Intune (managed devices + config profiles) --------------------
    # Output : Intune_Devices, Intune_DevicePolicies, Intune_ConfigProfiles
    # Device config/compliance states use Microsoft Graph JSON batching. Use
    # -SkipIntuneDevicePolicies when a quick inventory is sufficient.
    if ($CloudWorkload -contains 'Intune') {
        if (Test-Module Microsoft.Graph.DeviceManagement) {
            Import-Module Microsoft.Graph.DeviceManagement -ErrorAction SilentlyContinue
            Write-Log "Collecting Intune managed devices..."
            try {
                Write-Verbose2 "Downloading managed device list (paged REST)..."
                # Paged REST (not the typed -All cmdlet, which leaks) -> slim objects.
                $devSel = 'id,deviceName,userPrincipalName,operatingSystem,osVersion,managedDeviceOwnerType,complianceState,enrolledDateTime,lastSyncDateTime,manufacturer,model,serialNumber,ethernetMacAddress,wiFiMacAddress,azureADDeviceId'
                $rawDevices = foreach ($h in (Invoke-GraphGetAll "v1.0/deviceManagement/managedDevices?`$select=$devSel&`$top=999")) {
                    [pscustomobject]@{
                        DeviceName = $h['deviceName']; UserPrincipalName = $h['userPrincipalName']
                        OperatingSystem = $h['operatingSystem']; OsVersion = $h['osVersion']; ManagedDeviceOwnerType = $h['managedDeviceOwnerType']
                        ComplianceState = $h['complianceState']; EnrolledDateTime = $h['enrolledDateTime']; LastSyncDateTime = $h['lastSyncDateTime']
                        Manufacturer = $h['manufacturer']; Model = $h['model']; SerialNumber = $h['serialNumber']
                        EthernetMacAddress = $h['ethernetMacAddress']; WiFiMacAddress = $h['wiFiMacAddress']
                        AzureAdDeviceId = $h['azureADDeviceId']; Id = $h['id']
                    }
                }
                # Keep devices in the workbook's list OR owned by a wave user.
                $rawDevices = @($rawDevices | Where-Object { Test-TargetDevice $_ })
                Write-Log ("  $($rawDevices.Count) devices in scope" + $(if (-not $SkipIntuneDevicePolicies) { '; reading per-device policies (batched)...' } else { '.' }))

                # Per-device config + compliance policy state, batched (2 batch sets).
                $cfgRes = @{}; $cmpRes = @{}
                if (-not $SkipIntuneDevicePolicies -and $rawDevices.Count -gt 0) {
                    $cfgReqs = foreach ($d in $rawDevices) { @{ Key = $d.Id; Url = "/deviceManagement/managedDevices/$($d.Id)/deviceConfigurationStates" } }
                    $cfgRes = Invoke-GraphBatch @($cfgReqs) 'Device config-policy state'
                    $cmpReqs = foreach ($d in $rawDevices) { @{ Key = $d.Id; Url = "/deviceManagement/managedDevices/$($d.Id)/deviceCompliancePolicyStates" } }
                    $cmpRes = Invoke-GraphBatch @($cmpReqs) 'Device compliance-policy state'
                }

                $policyRows = [System.Collections.Generic.List[object]]::new()
                $devices = foreach ($d in $rawDevices) {
                    $cfg = @(); $cmp = @()
                    if (-not $SkipIntuneDevicePolicies) {
                        $cb = $cfgRes[$d.Id]; if ($cb) { $cfg = @($cb['value']) }
                        $mb = $cmpRes[$d.Id]; if ($mb) { $cmp = @($mb['value']) }
                        foreach ($s in $cfg) { $policyRows.Add([pscustomobject]@{ DeviceName = $d.DeviceName; UserPrincipalName = $d.UserPrincipalName; PolicyType = 'Configuration'; PolicyName = $s['displayName']; State = $s['state'] }) }
                        foreach ($s in $cmp) { $policyRows.Add([pscustomobject]@{ DeviceName = $d.DeviceName; UserPrincipalName = $d.UserPrincipalName; PolicyType = 'Compliance';    PolicyName = $s['displayName']; State = $s['state'] }) }
                    }
                    $nonComp = @(($cfg + $cmp) | Where-Object { $_['state'] -and $_['state'] -notin @('compliant', 'notApplicable') } | ForEach-Object { $_['displayName'] })

                    if ($d.UserPrincipalName) { Add-UniqueString -Collection ((Get-IdentityRecord $d.UserPrincipalName).IntuneDevices) -Value $d.DeviceName }

                    [pscustomobject]@{
                        DeviceName           = $d.DeviceName
                        UserPrincipalName    = $d.UserPrincipalName
                        OS                   = $d.OperatingSystem
                        OSVersion            = $d.OsVersion
                        Ownership            = $d.ManagedDeviceOwnerType
                        Compliance           = $d.ComplianceState
                        Enrolled             = $d.EnrolledDateTime
                        LastSync             = $d.LastSyncDateTime
                        Manufacturer         = $d.Manufacturer
                        Model                = $d.Model
                        Serial               = $d.SerialNumber
                        EthernetMac          = $d.EthernetMacAddress
                        WiFiMac              = $d.WiFiMacAddress
                        PolicyCount          = (@($cfg).Count + @($cmp).Count)
                        NonCompliantPolicies = ($nonComp -join '; ')
                        AzureAdDeviceId      = $d.AzureAdDeviceId
                        ManagedDeviceId      = $d.Id
                    }
                }
                $Report['Intune_Devices'] = @($devices)
                if (-not $SkipIntuneDevicePolicies) {
                    $Report['Intune_DevicePolicies'] = $policyRows.ToArray()
                    Write-Log "  $(@($devices).Count) Intune devices, $($policyRows.Count) device-policy states."
                }
                else {
                    Write-Log "  $(@($devices).Count) Intune devices (per-device policies skipped)."
                }

                # Configuration profiles (Intune counterpart to GPOs) via paged REST.
                $profiles = foreach ($h in (Invoke-GraphGetAll "v1.0/deviceManagement/deviceConfigurations?`$select=id,displayName,createdDateTime,lastModifiedDateTime")) {
                    [pscustomobject]@{
                        DisplayName = $h['displayName']
                        Type        = if ($h['@odata.type']) { $h['@odata.type'] -replace '#microsoft.graph.', '' } else { $null }
                        Created     = $h['createdDateTime']
                        Modified    = $h['lastModifiedDateTime']
                        Id          = $h['id']
                    }
                }
                $Report['Intune_ConfigProfiles'] = @($profiles)
                Write-Log "  $(@($profiles).Count) Intune config profiles."
            }
            catch { Write-Log "Intune collection failed: $($_.Exception.Message)" 'ERROR' }
        }
        else {
            Write-Log "Microsoft.Graph.DeviceManagement not installed; skipping Intune." 'WARN'
        }
    }

    # ---- OneDrive (via Graph: /users/{id}/drive) -----------------------
    # Output : OneDrive
    # Reads each in-scope user's drive metadata/quota via Graph and attaches the
    # OneDrive footprint to Dependency_Map.
    # Per-user drive incl. quota; scoped to the wave via Entra_Users. No SPO module.
    if ($CloudWorkload -contains 'OneDrive') {
        Write-Log "Collecting OneDrive (Graph /users/{id}/drive)..."
        try {
            # Users to check: the collected Entra users if present (already scoped),
            # else the wave candidate UPNs, else a minimal all-users pull.
            $odSrc =
                if ($Report['Entra_Users']) { @($Report['Entra_Users']) | ForEach-Object { [pscustomobject]@{ Upn = $_.UserPrincipalName; Id = $_.ObjectId } } }
                elseif ($null -ne $TargetUserIds) { @(Get-ScopedGraphUsersForWave) | ForEach-Object { [pscustomobject]@{ Upn = $_.UserPrincipalName; Id = $_.ObjectId } } }
                else { Invoke-GraphGetAll "v1.0/users?`$select=id,userPrincipalName&`$top=999" | ForEach-Object { [pscustomobject]@{ Upn = $_['userPrincipalName']; Id = $_['id'] } } }
            $odSrc = @($odSrc)
            $odRows = [System.Collections.Generic.List[object]]::new()
            # One batched pass: /users/{id}/drive for every user.
            $odReqs = foreach ($ou in $odSrc) { if ($ou.Id) { @{ Key = $ou.Upn; Url = "/users/$([uri]::EscapeDataString($ou.Id))/drive?`$select=webUrl,quota,driveType,lastModifiedDateTime" } } }
            $odRes = Invoke-GraphBatch @($odReqs | Where-Object { $_ }) 'Collecting OneDrive' -QuietStatuses 404
            foreach ($ou in $odSrc) {
                $d = $odRes[$ou.Upn]
                if (-not $d) { continue }   # user has no OneDrive provisioned (404) -> skip
                $q = $d['quota']
                $usedMb = if ($q -and $null -ne $q['used'])  { [math]::Round([double]$q['used']  / 1MB, 1) } else { $null }
                $totMb  = if ($q -and $null -ne $q['total']) { [math]::Round([double]$q['total'] / 1MB, 1) } else { $null }
                $odRows.Add([pscustomobject]@{
                    Owner         = $ou.Upn
                    Url           = $d['webUrl']
                    StorageUsedMB = $usedMb
                    StorageQuotaMB= $totMb
                    DriveType     = $d['driveType']
                    LastModified  = $d['lastModifiedDateTime']
                    QuotaState    = $(if ($q) { $q['state'] })
                })
                if ($ou.Upn) {
                    $idr = Get-IdentityRecord $ou.Upn
                    if ($idr) { $idr.OneDriveUrl = $d['webUrl']; $idr.OneDriveSizeMB = $usedMb }
                }
            }
            $Report['OneDrive'] = $odRows.ToArray()
            Write-Log "  $($odRows.Count) OneDrive sites."
        }
        catch { Write-Log "OneDrive (Graph) collection failed: $($_.Exception.Message)" 'ERROR' }
    }

    # ---- Exchange Online -----------------------------------------------
    # Outputs: Exchange_Mailboxes, Exchange_Archives, Exchange_MailboxPermissions,
    #          Exchange_DistLists, Exchange_DistListMembers
    # Includes primary/archive statistics, forwarding, retention/hold metadata,
    # CustomAttribute1-15, EmployeeType correlation, FullAccess/SendAs/SendOnBehalf
    # delegation and distribution-list dependencies.
    # Connect lazily, AFTER Graph, so Graph's MSAL is already loaded.
    if ($CloudWorkload -contains 'Exchange') {
        if (Test-Module ExchangeOnlineManagement) {
            Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue
            try {
                Write-Log "Connecting to Exchange Online..."
                $exoConnect = @{ ShowBanner = $false; ErrorAction = 'Stop' }
                $exoCommand = Get-Command Connect-ExchangeOnline -ErrorAction Stop
                if ($UseDeviceCode -and $exoCommand.Parameters.ContainsKey('Device')) {
                    $exoConnect['Device'] = $true
                }
                elseif ($exoCommand.Parameters.ContainsKey('DisableWAM')) {
                    $exoConnect['DisableWAM'] = $true
                }
                Connect-ExchangeOnline @exoConnect
                $ExoConnected = $true
            }
            catch { Write-Log "Exchange Online connect failed: $($_.Exception.Message)" 'ERROR' }
        }
        else { Write-Log "ExchangeOnlineManagement not installed; skipping Exchange." 'WARN' }

        if (-not $ExoConnected) {
            Write-Log "Exchange Online not connected; skipping mailbox/DL collection." 'WARN'
        }
        else {
            try {
                Write-Log "Collecting mailboxes..."
                $mbxProps = 'DisplayName','UserPrincipalName','Alias','PrimarySmtpAddress','ExchangeGuid','ExternalDirectoryObjectId','RecipientTypeDetails',
                            'ForwardingAddress','ForwardingSmtpAddress','DeliverToMailboxAndForward','GrantSendOnBehalfTo',
                            'ArchiveStatus','ArchiveState','ArchiveGuid','ArchiveName','ArchiveQuota','ArchiveWarningQuota',
                            'ArchiveDatabase','ArchiveDomain','ArchiveRelease','AutoExpandingArchiveEnabled',
                            'DisabledArchiveGuid','DisabledArchiveDatabase',
                            'LitigationHoldEnabled','LitigationHoldDuration','InPlaceHolds',
                            'RetentionPolicy','RetentionHoldEnabled','EmailAddressPolicyEnabled','HiddenFromAddressListsEnabled',
                            'CustomAttribute1','CustomAttribute2','CustomAttribute3','CustomAttribute4','CustomAttribute5',
                            'CustomAttribute6','CustomAttribute7','CustomAttribute8','CustomAttribute9','CustomAttribute10',
                            'CustomAttribute11','CustomAttribute12','CustomAttribute13','CustomAttribute14','CustomAttribute15'
                # Scoped: one mailbox lookup per workbook user (UPN, then mail, then
                # SMTP; users without a mailbox are skipped). Unscoped: all mailboxes.
                $rawMbx =
                    if ($null -ne $TargetUserIds) {
                        Write-Verbose2 "Fetching mailboxes for $($TargetUserIds.Count) wave users..."
                        $fi = 0; $fTot = $TargetUserIds.Count
                        foreach ($cands in $TargetUserIds) {
                            $fi++
                            Write-Activity -Activity 'Fetching mailboxes (scoped)' -Current $fi -Total $fTot -Item $cands[0] -Id 2
                            if ($fi % 100 -eq 0) { Write-Verbose2 "  ...fetched $fi of $fTot" }
                            Invoke-LoopGc $fi
                            Resolve-CloudUser $cands { param($x) Get-EXOMailbox -Identity $x -Properties $mbxProps -ErrorAction Stop }
                        }
                        Complete-Activity 'Fetching mailboxes (scoped)' 2
                    }
                    else {
                        Write-Verbose2 "Enumerating all mailboxes (Get-EXOMailbox -ResultSize Unlimited)..."
                        Get-EXOMailbox -ResultSize Unlimited -Properties $mbxProps
                    }
                $mbx = @($rawMbx)

                # Mailbox size/items: prefer ONE Graph usage report over the slow
                # per-mailbox Get-EXOMailboxStatistics loop. Falls back to per-mailbox
                # if Graph isn't connected or the tenant de-identifies report names.
                $usageByKey = $null
                $tmpCsv = $null
                try {
                    $gctx = $null; try { $gctx = Get-MgContext } catch {}
                    if ($gctx) {
                        Write-Verbose2 "Fetching mailbox usage report (Graph getMailboxUsageDetail)..."
                        $tmpCsv = Join-Path $OutputPath "mbxusage-$Stamp.csv"
                        Invoke-MgGraphRequest -Method GET -Uri "v1.0/reports/getMailboxUsageDetail(period='D7')" -OutputFilePath $tmpCsv -ErrorAction Stop
                        $usageRows = @(Import-Csv $tmpCsv)
                        $sample = @($usageRows | Select-Object -First 5)
                        if ($sample.Count -eq 0) {
                            Write-Log "  Usage report returned no rows; using per-mailbox statistics instead." 'WARN'
                            $usageByKey = $null
                        }
                        elseif (-not @($sample | Where-Object { $_.'User Principal Name' -like '*@*' })) {
                            Write-Log "  Usage report is de-identified (tenant privacy setting); using per-mailbox statistics instead." 'WARN'
                            $usageByKey = $null
                        }
                        else {
                            $usageByKey = @{}
                            foreach ($r in $usageRows) { if ($r.'User Principal Name') { $usageByKey[$r.'User Principal Name'.ToString().Trim().ToLowerInvariant()] = $r } }
                            if ($usageByKey.Count -eq 0) {
                                Write-Log "  Usage report had no usable UPN rows; using per-mailbox statistics instead." 'WARN'
                                $usageByKey = $null
                            }
                            else {
                                Write-Log "  Mailbox sizes from the usage report ($($usageByKey.Count) mailboxes); statistics fallback will be used for any missing rows."
                            }
                        }
                    }
                }
                catch { Write-Log "  Usage report unavailable ($($_.Exception.Message)); using per-mailbox statistics." 'WARN'; $usageByKey = $null }
                finally { if ($tmpCsv -and (Test-Path $tmpCsv)) { Remove-Item $tmpCsv -Force -ErrorAction SilentlyContinue } }

                Write-Log ("  Building rows for $($mbx.Count) mailboxes (archive statistics are collected automatically for archive-enabled mailboxes)" + $(if ($usageByKey) { '...' } else { ' (primary + archive per-mailbox statistics can be slow)...' }))
                $archiveRows = [System.Collections.Generic.List[object]]::new()
                $archiveStatsFailures = [System.Collections.Generic.List[string]]::new()
                $mi = 0; $mTot = $mbx.Count
                $mbx = $mbx |
                    ForEach-Object {
                        $mi++
                        Write-Activity -Activity 'Building mailbox rows' -Current $mi -Total $mTot -Item $_.PrimarySmtpAddress -Id 3
                        if ($mi % 100 -eq 0) { Write-Verbose2 "  ...$mi of $mTot mailboxes" }
                        Invoke-LoopGc $mi
                        $sizeMb = $null; $itemCount = $null
                        $ur = $null
                        if ($usageByKey) {
                            if ($_.UserPrincipalName)  { $ur = $usageByKey[$_.UserPrincipalName.ToString().ToLowerInvariant()] }
                            if (-not $ur -and $_.PrimarySmtpAddress) { $ur = $usageByKey[$_.PrimarySmtpAddress.ToString().ToLowerInvariant()] }
                            if ($ur) {
                                try { $sizeMb = [math]::Round([double]$ur.'Storage Used (Byte)' / 1MB, 1) } catch { $sizeMb = $null }
                                $itemCount = $ur.'Item Count'
                            }
                        }
                        # If Graph usage is unavailable OR this mailbox is missing from
                        # the report, fall back to authoritative EXO mailbox statistics.
                        if ((-not $ur -or $null -eq $sizeMb -or $null -eq $itemCount) -and $_.PrimarySmtpAddress) {
                            $stat = try { Get-EXOMailboxStatistics -Identity $_.PrimarySmtpAddress -ErrorAction Stop } catch { $null }
                            if ($stat) {
                                if ($null -eq $sizeMb) { $sizeMb = Convert-ExchangeSizeToMB $stat.TotalItemSize }
                                if ($null -eq $itemCount) { $itemCount = $stat.ItemCount }
                            }
                        }
                        # ArchiveStatus can occasionally report None even when a usable archive
                        # exists, so treat a non-empty ArchiveGuid as authoritative too.
                        $archiveGuidText = if ($_.ArchiveGuid) { $_.ArchiveGuid.ToString() } else { '' }
                        $archiveEnabled = (($_.ArchiveStatus -and $_.ArchiveStatus.ToString() -ne 'None') -or
                                           ($archiveGuidText -and $archiveGuidText -ne '00000000-0000-0000-0000-000000000000'))

                        $archiveItemCount = $null; $archiveSizeMb = $null
                        $archiveDeletedItemCount = $null; $archiveDeletedSizeMb = $null
                        $archiveLastLogon = $null; $archiveStorageLimitStatus = $null
                        if ($archiveEnabled -and $_.PrimarySmtpAddress) {
                            $archiveSmtp = $_.PrimarySmtpAddress.ToString()
                            try {
                                $archiveStat = Get-EXOMailboxStatistics -Identity $archiveSmtp -Archive `
                                    -Properties LastLogonTime,StorageLimitStatus -ErrorAction Stop
                                if ($archiveStat) {
                                    $archiveItemCount = $archiveStat.ItemCount
                                    $archiveDeletedItemCount = $archiveStat.DeletedItemCount
                                    $archiveLastLogon = $archiveStat.LastLogonTime
                                    $archiveStorageLimitStatus = $archiveStat.StorageLimitStatus
                                    $archiveSizeMb = Convert-ExchangeSizeToMB $archiveStat.TotalItemSize
                                    $archiveDeletedSizeMb = Convert-ExchangeSizeToMB $archiveStat.TotalDeletedItemSize
                                }
                            }
                            catch { $archiveStatsFailures.Add($archiveSmtp) }
                        }

                        $inPlace = ($_.InPlaceHolds -join ';')
                        $archiveName = if ($_.ArchiveName) { ($_.ArchiveName -join ';') } else { $null }

                        # EmployeeType is sourced from the merged on-prem Users sheet when
                        # available, with Entra employeeType as fallback. Match by UPN first,
                        # then primary SMTP because those can differ during tenant migration.
                        $employeeType = $null
                        foreach ($ek in @($_.UserPrincipalName, $_.PrimarySmtpAddress)) {
                            if (-not $ek) { continue }
                            $elk = $ek.ToString().Trim().ToLowerInvariant()
                            if ($EmployeeTypeByKey.ContainsKey($elk)) {
                                $employeeType = $EmployeeTypeByKey[$elk]
                                break
                            }
                        }

                        $rec = [pscustomobject]@{
                            DisplayName                 = $_.DisplayName
                            UserPrincipalName           = $_.UserPrincipalName
                            Alias                       = $_.Alias
                            PrimarySmtpAddress          = $_.PrimarySmtpAddress
                            EmployeeType                = $employeeType
                            ExchangeGuid                = $_.ExchangeGuid
                            ExternalDirectoryObjectId   = $_.ExternalDirectoryObjectId
                            RecipientTypeDetails        = $_.RecipientTypeDetails
                            MailboxType                 = $_.RecipientTypeDetails
                            ForwardingAddress           = if ($_.ForwardingAddress) { $_.ForwardingAddress.ToString() } else { $null }
                            ForwardingSmtpAddress       = if ($_.ForwardingSmtpAddress) { $_.ForwardingSmtpAddress.ToString() } else { $null }
                            DeliverToMailboxAndForward  = $_.DeliverToMailboxAndForward
                            GrantSendOnBehalfTo          = if ($_.GrantSendOnBehalfTo) { ($_.GrantSendOnBehalfTo | ForEach-Object { $_.ToString() }) -join ';' } else { $null }
                            ItemCount                   = $itemCount
                            SizeMB                      = $sizeMb
                            ArchiveEnabled              = $archiveEnabled
                            ArchiveStatus               = $_.ArchiveStatus
                            ArchiveState                = $_.ArchiveState
                            ArchiveGuid                 = $_.ArchiveGuid
                            ArchiveName                 = $archiveName
                            ArchiveSizeMB               = $archiveSizeMb
                            ArchiveItemCount            = $archiveItemCount
                            ArchiveDeletedItemCount     = $archiveDeletedItemCount
                            ArchiveDeletedSizeMB        = $archiveDeletedSizeMb
                            ArchiveLastLogonTime        = $archiveLastLogon
                            ArchiveStorageLimitStatus   = $archiveStorageLimitStatus
                            ArchiveQuota                = $_.ArchiveQuota
                            ArchiveWarningQuota         = $_.ArchiveWarningQuota
                            AutoExpandingArchiveEnabled = $_.AutoExpandingArchiveEnabled
                            ArchiveDatabase             = $_.ArchiveDatabase
                            ArchiveDomain               = $_.ArchiveDomain
                            ArchiveRelease              = $_.ArchiveRelease
                            DisabledArchiveGuid         = $_.DisabledArchiveGuid
                            DisabledArchiveDatabase     = $_.DisabledArchiveDatabase
                            LitigationHoldEnabled       = $_.LitigationHoldEnabled
                            LitigationHoldDays          = $_.LitigationHoldDuration
                            InPlaceHolds                = $inPlace
                            RetentionHoldEnabled        = $_.RetentionHoldEnabled
                            RetentionPolicy             = $_.RetentionPolicy
                            EmailAddressPolicyEnabled   = $_.EmailAddressPolicyEnabled
                            HiddenFromAddressListsEnabled = $_.HiddenFromAddressListsEnabled
                            CustomAttribute1            = $_.CustomAttribute1
                            CustomAttribute2            = $_.CustomAttribute2
                            CustomAttribute3            = $_.CustomAttribute3
                            CustomAttribute4            = $_.CustomAttribute4
                            CustomAttribute5            = $_.CustomAttribute5
                            CustomAttribute6            = $_.CustomAttribute6
                            CustomAttribute7            = $_.CustomAttribute7
                            CustomAttribute8            = $_.CustomAttribute8
                            CustomAttribute9            = $_.CustomAttribute9
                            CustomAttribute10           = $_.CustomAttribute10
                            CustomAttribute11           = $_.CustomAttribute11
                            CustomAttribute12           = $_.CustomAttribute12
                            CustomAttribute13           = $_.CustomAttribute13
                            CustomAttribute14           = $_.CustomAttribute14
                            CustomAttribute15           = $_.CustomAttribute15
                        }

                        if ($archiveEnabled) {
                            $archiveRows.Add([pscustomobject]@{
                                DisplayName                 = $_.DisplayName
                                UserPrincipalName           = $_.UserPrincipalName
                                Alias                       = $_.Alias
                                PrimarySmtpAddress          = $_.PrimarySmtpAddress
                                ExchangeGuid                = $_.ExchangeGuid
                                ExternalDirectoryObjectId   = $_.ExternalDirectoryObjectId
                                RecipientTypeDetails        = $_.RecipientTypeDetails
                                MailboxType                 = $_.RecipientTypeDetails
                                ArchiveEnabled              = $archiveEnabled
                                ArchiveStatus               = $_.ArchiveStatus
                                ArchiveState                = $_.ArchiveState
                                ArchiveGuid                 = $_.ArchiveGuid
                                ArchiveName                 = $archiveName
                                ArchiveSizeMB               = $archiveSizeMb
                                ArchiveItemCount            = $archiveItemCount
                                ArchiveDeletedItemCount     = $archiveDeletedItemCount
                                ArchiveDeletedSizeMB        = $archiveDeletedSizeMb
                                ArchiveLastLogonTime        = $archiveLastLogon
                                ArchiveStorageLimitStatus   = $archiveStorageLimitStatus
                                ArchiveQuota                = $_.ArchiveQuota
                                ArchiveWarningQuota         = $_.ArchiveWarningQuota
                                AutoExpandingArchiveEnabled = $_.AutoExpandingArchiveEnabled
                                ArchiveDatabase             = $_.ArchiveDatabase
                                ArchiveDomain               = $_.ArchiveDomain
                                ArchiveRelease              = $_.ArchiveRelease
                            })
                        }

                        $mailIdentityKey = if ($_.UserPrincipalName) { $_.UserPrincipalName } elseif ($_.PrimarySmtpAddress) { $_.PrimarySmtpAddress } else { $null }
                        $id = Get-IdentityRecord $mailIdentityKey
                        if ($id) {
                            $id.MailboxType          = $_.RecipientTypeDetails
                            $id.MailboxSizeMB        = $sizeMb
                            $id.ArchiveEnabled       = $archiveEnabled
                            $id.ArchiveSizeMB        = $archiveSizeMb
                            $id.ArchiveItemCount     = $archiveItemCount
                            $id.AutoExpandingArchive = $_.AutoExpandingArchiveEnabled
                            $id.LitigationHold       = $_.LitigationHoldEnabled
                            $id.InPlaceHolds         = $inPlace
                        }
                        $rec
                    }
                Complete-Activity 'Building mailbox rows' 3
                $Report['Exchange_Mailboxes'] = @($mbx)
                $Report['Exchange_Archives'] = $archiveRows.ToArray()
                Write-Log "  $(@($mbx).Count) mailboxes; $($archiveRows.Count) archive mailboxes."
                if ($archiveStatsFailures.Count -gt 0) {
                    $sampleFailures = @($archiveStatsFailures | Select-Object -First 10) -join ', '
                    Write-Log "  Archive properties were captured, but archive statistics failed for $($archiveStatsFailures.Count) mailbox(es). Sample: $sampleFailures" 'WARN'
                }

                # Mailbox delegation is migration-critical. Capture the three common
                # delegation models separately but normalize them into one worksheet:
                # FullAccess (mailbox ACL), SendAs (recipient ACL), SendOnBehalf (mailbox property).
                Write-Log "Collecting mailbox permissions (FullAccess, SendAs, SendOnBehalf)..."
                $mailboxPermissionRows = [System.Collections.Generic.List[object]]::new()
                $mailboxPermissionFailures = [System.Collections.Generic.List[string]]::new()
                $permIdx = 0; $permTotal = @($mbx).Count
                foreach ($m in @($mbx)) {
                    $permIdx++
                    $smtp = if ($m.PrimarySmtpAddress) { $m.PrimarySmtpAddress.ToString() } else { $null }
                    if (-not $smtp) { continue }
                    Write-Activity -Activity 'Reading mailbox permissions' -Current $permIdx -Total $permTotal -Item $smtp -Id 4
                    Invoke-LoopGc $permIdx

                    # FullAccess. Suppress inherited/system/self ACLs because they are not
                    # migration delegations and would swamp the report with boilerplate.
                    try {
                        foreach ($p in @(Get-EXOMailboxPermission -Identity $smtp -ResultSize Unlimited -ErrorAction Stop)) {
                            $trustee = if ($p.User) { $p.User.ToString() } else { $null }
                            $rights = @($p.AccessRights | ForEach-Object { $_.ToString() })
                            if (-not ($rights -contains 'FullAccess')) { continue }
                            if ($p.IsInherited) { continue }
                            if (-not $trustee -or $trustee -match '^(NT AUTHORITY\\SELF|S-1-5-10)$') { continue }
                            $mailboxPermissionRows.Add([pscustomobject]@{
                                MailboxDisplayName       = $m.DisplayName
                                MailboxUserPrincipalName = $m.UserPrincipalName
                                MailboxPrimarySmtpAddress= $smtp
                                MailboxType              = $m.RecipientTypeDetails
                                PermissionType           = 'FullAccess'
                                Trustee                  = $trustee
                                AccessRights             = ($rights -join ';')
                                IsInherited              = [bool]$p.IsInherited
                                Deny                     = [bool]$p.Deny
                            })
                        }
                    }
                    catch { $mailboxPermissionFailures.Add("FullAccess|$smtp|$($_.Exception.Message)") }

                    # SendAs lives on the recipient permission ACL.
                    try {
                        foreach ($p in @(Get-EXORecipientPermission -Identity $smtp -AccessRights SendAs -ResultSize Unlimited -ErrorAction Stop)) {
                            $trustee = if ($p.Trustee) { $p.Trustee.ToString() } else { $null }
                            $rights = @($p.AccessRights | ForEach-Object { $_.ToString() })
                            if (-not ($rights -contains 'SendAs')) { continue }
                            if (-not $trustee -or $trustee -match '^(NT AUTHORITY\\SELF|S-1-5-10)$') { continue }
                            $mailboxPermissionRows.Add([pscustomobject]@{
                                MailboxDisplayName       = $m.DisplayName
                                MailboxUserPrincipalName = $m.UserPrincipalName
                                MailboxPrimarySmtpAddress= $smtp
                                MailboxType              = $m.RecipientTypeDetails
                                PermissionType           = 'SendAs'
                                Trustee                  = $trustee
                                AccessRights             = ($rights -join ';')
                                IsInherited              = $null
                                Deny                     = $null
                            })
                        }
                    }
                    catch { $mailboxPermissionFailures.Add("SendAs|$smtp|$($_.Exception.Message)") }

                    # SendOnBehalf is a mailbox property, already fetched with the mailbox.
                    foreach ($delegate in @($m.GrantSendOnBehalfTo -split ';' | Where-Object { $_ })) {
                        $mailboxPermissionRows.Add([pscustomobject]@{
                            MailboxDisplayName       = $m.DisplayName
                            MailboxUserPrincipalName = $m.UserPrincipalName
                            MailboxPrimarySmtpAddress= $smtp
                            MailboxType              = $m.RecipientTypeDetails
                            PermissionType           = 'SendOnBehalf'
                            Trustee                  = $delegate.Trim()
                            AccessRights             = 'SendOnBehalf'
                            IsInherited              = $false
                            Deny                     = $false
                        })
                    }
                }
                Complete-Activity 'Reading mailbox permissions' 4
                $Report['Exchange_MailboxPermissions'] = $mailboxPermissionRows.ToArray()
                Write-Log "  $($mailboxPermissionRows.Count) mailbox permission/delegation rows."
                if ($mailboxPermissionFailures.Count -gt 0) {
                    $samplePermissionFailures = @($mailboxPermissionFailures | Select-Object -First 10) -join ' | '
                    Write-Log "  Mailbox permission lookup failed for $($mailboxPermissionFailures.Count) operation(s). Sample: $samplePermissionFailures" 'WARN'
                }

                Write-Log "Collecting distribution / mail-enabled groups..."
                # When scoped, keep DLs that the wave belongs to (by SMTP from the Entra
                # group membership pass). Without Entra group scoping, retain the existing
                # behavior and enumerate all DLs rather than silently returning none.
                $rawDls = @(Get-DistributionGroup -ResultSize Unlimited | Where-Object {
                    ($null -eq $TargetUserKeys) -or ($null -eq $WaveGroupMails) -or
                    ($_.PrimarySmtpAddress -and $WaveGroupMails.ContainsKey($_.PrimarySmtpAddress.ToString().ToLowerInvariant()))
                })
                $dls = @($rawDls | ForEach-Object {
                    [pscustomobject]@{
                        DisplayName                        = $_.DisplayName
                        Alias                              = $_.Alias
                        PrimarySmtpAddress                 = $_.PrimarySmtpAddress
                        RecipientTypeDetails               = $_.RecipientTypeDetails
                        ExternalDirectoryObjectId          = $_.ExternalDirectoryObjectId
                        ManagedBy                          = ($_.ManagedBy -join ';')
                        MemberJoinRestriction              = $_.MemberJoinRestriction
                        MemberDepartRestriction            = $_.MemberDepartRestriction
                        HiddenFromAddressListsEnabled      = $_.HiddenFromAddressListsEnabled
                        RequireSenderAuthenticationEnabled = $_.RequireSenderAuthenticationEnabled
                        ModerationEnabled                  = $_.ModerationEnabled
                        ModeratedBy                        = ($_.ModeratedBy -join ';')
                        SendModerationNotifications        = $_.SendModerationNotifications
                    }
                })
                $Report['Exchange_DistLists'] = $dls
                Write-Log "  $($dls.Count) distribution/mail-enabled groups$(if ($null -ne $TargetUserKeys -and $null -ne $WaveGroupMails) { ' (scoped to wave)' })."

                # Membership is flattened to one row per group/member relationship. This
                # intentionally keeps non-user recipients (contacts, nested groups, etc.).
                Write-Log "Collecting distribution-list members..."
                $dlMemberRows = [System.Collections.Generic.List[object]]::new()
                $dlMemberFailures = [System.Collections.Generic.List[string]]::new()
                $dlIdx = 0; $dlTotal = $rawDls.Count
                foreach ($dl in $rawDls) {
                    $dlIdx++
                    $dlIdentity = if ($dl.PrimarySmtpAddress) { $dl.PrimarySmtpAddress.ToString() } else { $dl.Identity.ToString() }
                    Write-Activity -Activity 'Reading distribution-list members' -Current $dlIdx -Total $dlTotal -Item $dlIdentity -Id 5
                    Invoke-LoopGc $dlIdx
                    try {
                        foreach ($member in @(Get-DistributionGroupMember -Identity $dlIdentity -ResultSize Unlimited -ErrorAction Stop)) {
                            $dlMemberRows.Add([pscustomobject]@{
                                GroupDisplayName                 = $dl.DisplayName
                                GroupAlias                       = $dl.Alias
                                GroupPrimarySmtpAddress          = $dl.PrimarySmtpAddress
                                GroupRecipientTypeDetails        = $dl.RecipientTypeDetails
                                GroupExternalDirectoryObjectId   = $dl.ExternalDirectoryObjectId
                                MemberDisplayName                = $member.DisplayName
                                MemberAlias                      = $member.Alias
                                MemberPrimarySmtpAddress         = $member.PrimarySmtpAddress
                                MemberRecipientType              = $member.RecipientType
                                MemberRecipientTypeDetails       = $member.RecipientTypeDetails
                                MemberExternalDirectoryObjectId  = $member.ExternalDirectoryObjectId
                                MemberIdentity                   = if ($member.Identity) { $member.Identity.ToString() } else { $null }
                            })
                        }
                    }
                    catch { $dlMemberFailures.Add("$dlIdentity|$($_.Exception.Message)") }
                }
                Complete-Activity 'Reading distribution-list members' 5
                $Report['Exchange_DistListMembers'] = $dlMemberRows.ToArray()
                Write-Log "  $($dlMemberRows.Count) distribution-list membership rows."
                if ($dlMemberFailures.Count -gt 0) {
                    $sampleDlFailures = @($dlMemberFailures | Select-Object -First 10) -join ' | '
                    Write-Log "  Distribution-list member lookup failed for $($dlMemberFailures.Count) group(s). Sample: $sampleDlFailures" 'WARN'
                }
            }
            catch { Write-Log "Exchange Online collection failed: $($_.Exception.Message)" 'ERROR' }
        }
    }

    # ---- SharePoint Online (SPO/PnP admin) -----------------------------
    # Output : SharePoint_Sites
    # Provider is selected by -SharePointModule. On non-Windows PowerShell 7 hosts,
    # PnP.PowerShell is the practical option. Scoped attribution uses group-backed
    # site membership and owner matching where available.
    # OneDrive is collected via Graph above; this block is SharePoint sites only.
    if ($CloudWorkload -contains 'SharePoint') {
        if ([string]::IsNullOrWhiteSpace($SharePointAdminUrl)) {
            Write-Log "SharePoint in scope but -SharePointAdminUrl not supplied; skipping." 'WARN'
        }
        else {
            # Enumerate sites into a NORMALIZED shape so SPO and PnP feed the same
            # downstream logic.
            $allSpo = $null
            try {
                if ($SharePointModule -eq 'PnP') {
                    if (-not (Test-Module PnP.PowerShell)) { throw "PnP.PowerShell not installed (Install-Module PnP.PowerShell)." }
                    if ([string]::IsNullOrWhiteSpace($PnPClientId)) { throw "-SharePointModule PnP requires -PnPClientId (an Entra app registration client id)." }
                    Import-Module PnP.PowerShell -ErrorAction Stop
                    Write-Log "Connecting to SharePoint via PnP.PowerShell$(if ($UseDeviceCode) { ' (device code)' } else { ' (interactive)' })..."
                    $pnp = @{ Url = $SharePointAdminUrl; ClientId = $PnPClientId; ErrorAction = 'Stop' }
                    if ($UseDeviceCode) { $pnp['DeviceLogin'] = $true } else { $pnp['Interactive'] = $true }
                    Connect-PnPOnline @pnp
                    Write-Log "Enumerating sites (PnP, incl. OneDrive)..."
                    $allSpo = Get-PnPTenantSite -IncludeOneDriveSites -ErrorAction Stop | ForEach-Object {
                        [pscustomobject]@{
                            Title = $_.Title; Url = $_.Url; Owner = $_.Owner; Template = $_.Template
                            StorageUsedMB = [math]::Round(($_.StorageUsage), 1); StorageQuotaMB = $_.StorageQuota
                            SharingCapability = $_.SharingCapability; LastModified = $_.LastContentModifiedDate
                            GroupId = $_.GroupId; Status = $_.Status
                        }
                    }
                }
                else {
                    if (-not $RunningOnWindows) { throw "The Microsoft.Online.SharePoint.PowerShell (SPO) module is Windows-only. On macOS/Linux use -SharePointModule PnP -PnPClientId <client-id>." }
                    if (-not (Test-Module Microsoft.Online.SharePoint.PowerShell)) { throw "Microsoft.Online.SharePoint.PowerShell not installed (or use -SharePointModule PnP)." }
                    Import-Module Microsoft.Online.SharePoint.PowerShell -ErrorAction Stop
                    Write-Log "Connecting to SharePoint Online (interactive; SPO module has no device-code option)..."
                    Connect-SPOService -Url $SharePointAdminUrl -ErrorAction Stop
                    Write-Log "Enumerating SPO sites (incl. personal/OneDrive)..."
                    $allSpo = Get-SPOSite -Limit All -IncludePersonalSite:$true | ForEach-Object {
                        [pscustomobject]@{
                            Title = $_.Title; Url = $_.Url; Owner = $_.Owner; Template = $_.Template
                            StorageUsedMB = [math]::Round($_.StorageUsageCurrent, 1); StorageQuotaMB = $_.StorageQuota
                            SharingCapability = $_.SharingCapability; LastModified = $_.LastContentModifiedDate
                            GroupId = $_.GroupId; Status = $_.Status
                        }
                    }
                }
            }
            catch {
                $spLevel = if (-not $RunningOnWindows -and $SharePointModule -eq 'SPO') { 'WARN' } else { 'ERROR' }
                Write-Log "SharePoint connection failed: $($_.Exception.Message)" $spLevel
            }

            if ($allSpo) {
                if ($CloudWorkload -contains 'SharePoint') {
                    $sites = @($allSpo | Where-Object { $_.Template -notlike 'SPSPERS*' })
                    if ($null -ne $TargetUserKeys) {
                        # Scope to sites backed by a group the wave belongs to, or owned
                        # by a wave user. (Classic non-group sites can't be attributed to
                        # users without per-site enumeration, so they're excluded.)
                        $before = $sites.Count
                        $sites = @($sites | Where-Object {
                            ($WaveGroupIds -and $_.GroupId -and $WaveGroupIds.ContainsKey($_.GroupId.ToString().ToLowerInvariant())) -or (Test-TargetUser (Get-PrincipalEmail $_.Owner))
                        })
                        if (-not $WaveGroupIds) { Write-Log "  SharePoint scoping used owner-match only (add Entra to the workload for group-backed site scoping)." 'WARN' }
                        Write-Log "  Scoped SharePoint sites to the wave: $($sites.Count) of $before."
                    }
                    $Report['SharePoint_Sites'] = @($sites)

                    # Feed SharePoint dependencies back into the per-identity map.
                    # Group-backed sites are attached to each wave user who belongs to
                    # that Entra group; classic/owner-attributed sites are attached by owner.
                    $spGroupUsers = @{}
                    foreach ($gm in @($Report['Entra_GroupMembership'])) {
                        if (-not $gm.GroupId -or -not $gm.UserPrincipalName) { continue }
                        $gid = $gm.GroupId.ToString().Trim().ToLowerInvariant()
                        if (-not $spGroupUsers.ContainsKey($gid)) { $spGroupUsers[$gid] = [System.Collections.Generic.List[string]]::new() }
                        Add-UniqueString -Collection $spGroupUsers[$gid] -Value $gm.UserPrincipalName
                    }
                    foreach ($site in $sites) {
                        $ownerUpn = Get-PrincipalEmail $site.Owner
                        if ($ownerUpn -and $ownerUpn -like '*@*') { Add-UniqueString -Collection ((Get-IdentityRecord $ownerUpn).SharePointSites) -Value $site.Url }
                        if ($site.GroupId) {
                            $gid = $site.GroupId.ToString().Trim().ToLowerInvariant()
                            if ($spGroupUsers.ContainsKey($gid)) {
                                foreach ($upn in $spGroupUsers[$gid]) { Add-UniqueString -Collection ((Get-IdentityRecord $upn).SharePointSites) -Value $site.Url }
                            }
                        }
                    }
                    Write-Log "  $(@($sites).Count) team/communication sites."
                }
            }
        }
    }

    # ---- Teams ----------------------------------------------------------
    # Output : Teams, Teams_Channels, Teams_Members
    # Captures team inventory, channel inventory and team-level membership. Private
    # or shared channel membership may require separate deeper discovery.
    if ($CloudWorkload -contains 'Teams') {
        if (Test-Module MicrosoftTeams) {
            Import-Module MicrosoftTeams -ErrorAction Stop
            try {
                Write-Log "Connecting to Microsoft Teams$(if ($UseDeviceCode) { ' (device authentication)...' } else { ' (interactive)...' })"
                $teamsConnect = @{ ErrorAction = 'Stop' }
                $teamsCommand = Get-Command Connect-MicrosoftTeams -ErrorAction Stop
                if ($UseDeviceCode -and $teamsCommand.Parameters.ContainsKey('UseDeviceAuthentication')) {
                    $teamsConnect['UseDeviceAuthentication'] = $true
                }
                elseif (-not $UseDeviceCode -and $teamsCommand.Parameters.ContainsKey('DisableWAM')) {
                    $teamsConnect['DisableWAM'] = $true
                }
                elseif ($UseDeviceCode) {
                    Write-Log "  Installed MicrosoftTeams module does not expose -UseDeviceAuthentication; using its default interactive flow." 'WARN'
                }
                Connect-MicrosoftTeams @teamsConnect | Out-Null
                Write-Verbose2 "Enumerating all Teams..."
                $teamRows    = [System.Collections.Generic.List[object]]::new()
                $channelRows = [System.Collections.Generic.List[object]]::new()
                $memberRows  = [System.Collections.Generic.List[object]]::new()
                $allTeams = @(Get-Team)
                Write-Log "Collecting channels/membership for $($allTeams.Count) teams..."
                $ti = 0; $tTot = $allTeams.Count
                foreach ($t in $allTeams) {
                    $ti++
                    Write-Activity -Activity 'Collecting Teams' -Current $ti -Total $tTot -Item $t.DisplayName
                    # Buffer per-team rows; when scoping we only KEEP the team if a
                    # wave user is a member (so the Teams list scopes to the wave too).
                    $tChannels = [System.Collections.Generic.List[object]]::new()
                    $tMembers  = [System.Collections.Generic.List[object]]::new()
                    $waveCount = 0
                    try {
                        foreach ($c in (Get-TeamChannel -GroupId $t.GroupId -ErrorAction Stop)) {
                            $tChannels.Add([pscustomobject]@{
                                Team        = $t.DisplayName
                                Channel     = $c.DisplayName
                                MembershipType = $c.MembershipType
                            })
                        }
                        foreach ($u in (Get-TeamUser -GroupId $t.GroupId -ErrorAction Stop)) {
                            if ($u.User -and (Test-TargetUser $u.User)) {
                                $tMembers.Add([pscustomobject]@{
                                    Team    = $t.DisplayName
                                    GroupId = $t.GroupId
                                    User    = $u.User
                                    Name    = $u.Name
                                    Role    = $u.Role
                                })
                                $waveCount++
                            }
                        }
                    } catch { Write-Log "  Team '$($t.DisplayName)' detail failed: $($_.Exception.Message)" 'WARN' }

                    # Include the team when not scoping, or when it has a wave member.
                    if ($null -eq $TargetUserKeys -or $waveCount -gt 0) {
                        $teamRows.Add([pscustomobject]@{
                            DisplayName  = $t.DisplayName
                            GroupId      = $t.GroupId
                            Visibility   = $t.Visibility
                            Archived     = $t.Archived
                            MailNickname = $t.MailNickName
                            Description  = $t.Description
                        })
                        foreach ($cr in $tChannels) { $channelRows.Add($cr) }
                        foreach ($mr in $tMembers)  { $memberRows.Add($mr); Add-UniqueString -Collection ((Get-IdentityRecord $mr.User).TeamsMemberOf) -Value $t.DisplayName }
                    }
                }
                Complete-Activity 'Collecting Teams'
                $Report['Teams']          = $teamRows.ToArray()
                $Report['Teams_Channels'] = $channelRows.ToArray()
                $Report['Teams_Members']  = $memberRows.ToArray()
                Write-Log "  $($teamRows.Count) teams, $($channelRows.Count) channels, $($memberRows.Count) memberships."
            }
            catch { Write-Log "Teams collection failed: $($_.Exception.Message)" 'ERROR' }
        }
        else { Write-Log "MicrosoftTeams not installed; skipping Teams." 'WARN' }
    }

    # ---- Purview / Security & Compliance -------------------------------
    # Outputs: sensitivity/retention labels and policies, DLP policies and
    #          eDiscovery/mailbox-hold context. Connect-IPPSSession is reused from
    #          ExchangeOnlineManagement and can have different auth switches than EXO.
    # Labels, DLP, retention and eDiscovery holds are all rebuilt (not moved)
    # in the target tenant, and container/content labels gate what can be
    # migrated at all -- so they are first-class dependency inputs.
    # Connect lazily, AFTER Graph.
    if ($CloudWorkload -contains 'Purview') {
        if (Test-Module ExchangeOnlineManagement) {
            Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue
            try {
                Write-Log "Connecting to Security & Compliance (Connect-IPPSSession)..."
                $ippsConnect = @{ ShowBanner = $false; ErrorAction = 'Stop' }
                $ippsCommand = Get-Command Connect-IPPSSession -ErrorAction Stop
                if ($ippsCommand.Parameters.ContainsKey('DisableWAM')) { $ippsConnect['DisableWAM'] = $true }
                if ($UseDeviceCode) {
                    Write-Log "  Connect-IPPSSession does not expose a -Device switch; Purview will use its interactive modern-auth flow instead." 'WARN'
                }
                Connect-IPPSSession @ippsConnect
                $IppsConnected = $true
            }
            catch { Write-Log "Security & Compliance connect failed: $($_.Exception.Message)" 'ERROR' }
        }
        else { Write-Log "ExchangeOnlineManagement not installed; skipping Purview." 'WARN' }

        if (-not $IppsConnected) {
            Write-Log "Security & Compliance not connected; skipping Purview." 'WARN'
        }
        else {
            try {
                Write-Log "Collecting sensitivity labels..."
                try {
                    $labels = Get-Label -ErrorAction Stop | ForEach-Object {
                        [pscustomobject]@{
                            DisplayName   = $_.DisplayName
                            Name          = $_.Name
                            Priority      = $_.Priority
                            Enabled       = $_.Enabled
                            ContentType   = ($_.ContentType -join ';')   # File/Email/Site/UnifiedGroup...
                            ParentLabel   = $_.ParentLabelDisplayName
                            Guid          = $_.Guid
                        }
                    }
                    $Report['Purview_SensitivityLabels'] = @($labels)
                    Write-Log "  $(@($labels).Count) sensitivity labels."
                } catch { Write-Log "  Sensitivity labels failed: $($_.Exception.Message)" 'WARN' }

                try {
                    $labelPolicies = Get-LabelPolicy -ErrorAction Stop | ForEach-Object {
                        [pscustomobject]@{
                            Name          = $_.Name
                            Enabled       = $_.Enabled
                            Labels        = ($_.Labels -join ';')
                            Mode          = $_.Mode
                        }
                    }
                    $Report['Purview_LabelPolicies'] = @($labelPolicies)
                    Write-Log "  $(@($labelPolicies).Count) label policies."
                } catch { Write-Log "  Label policies failed: $($_.Exception.Message)" 'WARN' }

                Write-Log "Collecting retention labels & policies..."
                try {
                    $retLabels = Get-ComplianceTag -ErrorAction Stop | ForEach-Object {
                        [pscustomobject]@{
                            Name           = $_.Name
                            RetentionAction= $_.RetentionAction
                            RetentionDays  = $_.RetentionDuration
                            IsRecordLabel  = $_.IsRecordLabel
                            Regulatory     = $_.Regulatory
                        }
                    }
                    $Report['Purview_RetentionLabels'] = @($retLabels)
                    Write-Log "  $(@($retLabels).Count) retention labels."
                } catch { Write-Log "  Retention labels failed: $($_.Exception.Message)" 'WARN' }

                try {
                    $retPolicies = Get-RetentionCompliancePolicy -ErrorAction Stop | ForEach-Object {
                        [pscustomobject]@{
                            Name              = $_.Name
                            Enabled           = $_.Enabled
                            Mode              = $_.Mode
                            Workloads         = $_.Workload
                            AppliesToWave     = (Test-PurviewInWave @($_.ExchangeLocation, $_.ModernGroupLocation))
                            ExchangeLocation  = ($_.ExchangeLocation  -join ';')
                            SharePointLocation= ($_.SharePointLocation -join ';')
                            OneDriveLocation  = ($_.OneDriveLocation  -join ';')
                            TeamsChannelLocation = ($_.TeamsChannelLocation -join ';')
                        }
                    }
                    $Report['Purview_RetentionPolicies'] = @($retPolicies)
                    Write-Log "  $(@($retPolicies).Count) retention policies$(if ($null -ne $TargetUserKeys) { " ($((@($retPolicies) | Where-Object AppliesToWave).Count) flagged AppliesToWave)" })."
                } catch { Write-Log "  Retention policies failed: $($_.Exception.Message)" 'WARN' }

                Write-Log "Collecting DLP policies..."
                try {
                    $dlp = Get-DlpCompliancePolicy -ErrorAction Stop | ForEach-Object {
                        [pscustomobject]@{
                            Name              = $_.Name
                            Enabled           = ($_.Mode -like 'Enable*')
                            Mode              = $_.Mode
                            Workloads         = $_.Workload
                            AppliesToWave     = (Test-PurviewInWave @($_.ExchangeLocation))
                            ExchangeLocation  = ($_.ExchangeLocation  -join ';')
                            SharePointLocation= ($_.SharePointLocation -join ';')
                            OneDriveLocation  = ($_.OneDriveLocation  -join ';')
                            TeamsLocation     = ($_.TeamsLocation     -join ';')
                        }
                    }
                    $Report['Purview_DLPPolicies'] = @($dlp)
                    Write-Log "  $(@($dlp).Count) DLP policies$(if ($null -ne $TargetUserKeys) { " ($((@($dlp) | Where-Object AppliesToWave).Count) flagged AppliesToWave)" })."
                } catch { Write-Log "  DLP policies failed: $($_.Exception.Message)" 'WARN' }

                Write-Log "Collecting eDiscovery / case holds..."
                try {
                    $holds = Get-CaseHoldPolicy -ErrorAction Stop | ForEach-Object {
                        [pscustomobject]@{
                            Name             = $_.Name
                            Enabled          = $_.Enabled
                            Workloads        = $_.Workload
                            ExchangeLocation = ($_.ExchangeLocation  -join ';')
                            SharePointLocation = ($_.SharePointLocation -join ';')
                        }
                    }
                    $Report['Purview_eDiscoveryHolds'] = @($holds)
                    Write-Log "  $(@($holds).Count) case hold policies."
                } catch { Write-Log "  Case holds failed (may need eDiscovery role): $($_.Exception.Message)" 'WARN' }
            }
            catch { Write-Log "Security & Compliance collection failed: $($_.Exception.Message)" 'ERROR' }
        }
    }
}

#endregion

#region ---------------------------------------------- Device <-> Intune correlation
# CORRELATION STEP
# ----------------
# Adds current Intune enrollment information back onto the on-prem Devices report.
# Either side may come from this run or from a previously merged workbook.

# Flag each on-prem AD device with its Intune enrollment status. Both datasets may
# come from THIS run or from the -InputWorkbook (split runs), so we source each from
# $Report first, falling back to the workbook. Match is on device name.
$intuneList = if ($Report['Intune_Devices']) { @($Report['Intune_Devices']) }
              elseif ($Merging) { try { @(Import-Excel -Path $WorkbookPath -WorksheetName 'Intune_Devices') } catch { @() } }
              else { @() }

$adDevices = if ($Report['Devices']) { @($Report['Devices']) }
             elseif ($Merging) { try { @(Import-Excel -Path $WorkbookPath -WorksheetName 'Devices') } catch { @() } }
             else { @() }

if (@($adDevices).Count -gt 0 -and @($intuneList).Count -gt 0) {
    Write-Log "Correlating $($adDevices.Count) AD devices with $($intuneList.Count) Intune-enrolled devices..."
    $intuneByName = @{}
    foreach ($d in $intuneList) {
        if ($d.DeviceName) { $intuneByName[$d.DeviceName.ToString().Trim().ToUpperInvariant()] = $d }
    }
    # Columns we add; excluded from the source first so re-runs stay idempotent.
    # (OperatingSystem/OSVersion are re-added too, coalesced: AD value else Intune.)
    $intuneCols = 'IsIntuneEnrolled','IntuneCompliance','IntuneOwnership','IntuneLastSync','IntuneEnrolled',
                  'IntunePrimaryUser','Serial','Model','Manufacturer','EthernetMac','WiFiMac',
                  'IntunePolicyCount','IntuneNonCompliantPolicies','AzureAdDeviceId','OperatingSystem','OSVersion'
    $matched = 0
    $enriched = foreach ($c in $adDevices) {
        $nm = if ($c.Name) { $c.Name } elseif ($c.DNSHostName) { ($c.DNSHostName -split '\.')[0] } else { $null }
        $hit = if ($nm) { $intuneByName[$nm.ToString().Trim().ToUpperInvariant()] } else { $null }
        if ($hit) { $matched++ }
        $c | Select-Object * -ExcludeProperty $intuneCols | Select-Object *,
            @{n = 'OperatingSystem';  e = { if ($c.OperatingSystem) { $c.OperatingSystem } elseif ($hit) { $hit.OS } } },
            @{n = 'OSVersion';        e = { if ($c.OSVersion) { $c.OSVersion } elseif ($hit) { $hit.OSVersion } } },
            @{n = 'IsIntuneEnrolled'; e = { [bool]$hit } },
            @{n = 'IntuneCompliance'; e = { if ($hit) { $hit.Compliance } } },
            @{n = 'IntuneOwnership';  e = { if ($hit) { $hit.Ownership } } },
            @{n = 'IntuneLastSync';   e = { if ($hit) { $hit.LastSync } } },
            @{n = 'IntuneEnrolled';   e = { if ($hit) { $hit.Enrolled } } },
            @{n = 'IntunePrimaryUser';e = { if ($hit) { $hit.UserPrincipalName } } },
            @{n = 'Serial';           e = { if ($hit) { $hit.Serial } } },
            @{n = 'Model';            e = { if ($hit) { $hit.Model } } },
            @{n = 'Manufacturer';     e = { if ($hit) { $hit.Manufacturer } } },
            @{n = 'EthernetMac';      e = { if ($hit) { $hit.EthernetMac } } },
            @{n = 'WiFiMac';          e = { if ($hit) { $hit.WiFiMac } } },
            @{n = 'IntunePolicyCount';e = { if ($hit) { $hit.PolicyCount } } },
            @{n = 'IntuneNonCompliantPolicies'; e = { if ($hit) { $hit.NonCompliantPolicies } } },
            @{n = 'AzureAdDeviceId';  e = { if ($hit) { $hit.AzureAdDeviceId } } }
    }
    $Report['Devices'] = @($enriched)
    Write-Log "  $matched of $($adDevices.Count) AD devices are Intune-enrolled ($([int]($matched / [math]::Max($adDevices.Count,1) * 100))%)."
}
elseif (@($adDevices).Count -gt 0) {
    Write-Log "Skipping device/Intune correlation: no Intune data available (collect Intune, or merge a workbook that has it)." 'WARN'
}

#endregion

#region --------------------------------------------------------- Dependency map
# DEPENDENCY MAP
# --------------
# One logical row per correlated identity. This section converts the master identity
# index accumulated by every collector into migration-planning counts/flags such as
# mailbox/archive state, group memberships, apps, Teams, OneDrive, SharePoint and
# Intune dependencies.

Write-Log "Building per-identity dependency map..."
$depRows = foreach ($k in $Identity.Keys) {
    $r = $Identity[$k]
    [pscustomobject]@{
        Identity          = $r.Key
        SamAccountName    = $r.SamAccountName
        OnPremOU          = $r.OnPremOU
        OnPremEnabled     = $r.OnPremEnabled
        EntraEnabled      = $r.EntraEnabled
        DirSynced         = $r.EntraDirSynced
        MatchedOnPremAndEntra = ([bool]$r.EntraObjectId -and [bool]$r.OnPremDN)
        Licenses          = $r.Licenses
        MailboxType       = $r.MailboxType
        MailboxSizeMB     = $r.MailboxSizeMB
        ArchiveEnabled    = $r.ArchiveEnabled
        ArchiveSizeMB     = $r.ArchiveSizeMB
        ArchiveItemCount  = $r.ArchiveItemCount
        AutoExpandingArchive = $r.AutoExpandingArchive
        LitigationHold    = $r.LitigationHold
        InPlaceHolds      = $r.InPlaceHolds
        OneDriveUrl       = $r.OneDriveUrl
        OneDriveSizeMB    = $r.OneDriveSizeMB
        OnPremGroupCount  = $r.OnPremGroups.Count
        EntraGroupCount   = $r.EntraGroups.Count
        TeamsCount        = $r.TeamsMemberOf.Count
        SharePointCount   = $r.SharePointSites.Count
        IntuneDeviceCount = $r.IntuneDevices.Count
        OnPremGroups      = ($r.OnPremGroups      -join '; ')
        EntraGroups       = ($r.EntraGroups       -join '; ')
        Teams             = ($r.TeamsMemberOf     -join '; ')
        IntuneDevices     = ($r.IntuneDevices     -join '; ')
        # Migration workload flags
        NeedsExchangeMove = [bool]$r.MailboxType
        NeedsArchiveMove  = [bool]$r.ArchiveEnabled
        NeedsOneDriveMove = [bool]$r.OneDriveUrl
        NeedsTeamsMove    = ($r.TeamsMemberOf.Count -gt 0)
        NeedsIntuneMove   = ($r.IntuneDevices.Count -gt 0)
        # Holds block a clean cutover -- flag them for legal/compliance sign-off
        HasComplianceHold = ([bool]$r.LitigationHold -or -not [string]::IsNullOrWhiteSpace($r.InPlaceHolds))
        MigrationComplexity = $(
            $score = 0
            if ($r.MailboxType)               { $score++ }
            if ($r.ArchiveEnabled)            { $score++ }
            if ($r.OneDriveUrl)               { $score++ }
            if ($r.TeamsMemberOf.Count -gt 0) { $score++ }
            if ($r.IntuneDevices.Count -gt 0) { $score++ }
            if ($r.OnPremGroups.Count -gt 10) { $score++ }
            # a hold forces coordination regardless of anything else
            if ($r.LitigationHold -or -not [string]::IsNullOrWhiteSpace($r.InPlaceHolds)) { $score += 2 }
            switch ($score) { {$_ -ge 4} {'High'; break} {$_ -ge 2} {'Medium'; break} default {'Low'} }
        )
    }
}
$Report['Dependency_Map'] = @($depRows)
Write-Log "  $(@($depRows).Count) correlated identities."

#endregion

#region -------------------------------------------------------------- Excel out
# EXCEL WRITE SEMANTICS
# ---------------------
# - Every sheet produced in THIS run is a refresh operation.
# - Non-empty results replace the existing sheet via -ClearSheet.
# - An explicitly empty refreshed result removes a stale prior sheet.
# - Sheets not produced by the current run are preserved.
# - 00_Summary is rebuilt after all writes.
# - When -InputWorkbook is used, that file is the destination and is modified in place.

if (-not (Test-Module ImportExcel)) {
    throw "ImportExcel is required to write the workbook. Install-Module ImportExcel -Scope CurrentUser"
}
Import-Module ImportExcel -ErrorAction Stop

Write-Log ("{0} workbook: {1}" -f $(if ($Merging) { 'Merging into' } else { 'Writing' }), $WorkbookPath)

# Write each data sheet. -ClearSheet makes a re-run REPLACE the sheet rather
# than append rows -- essential so merging a half (or re-running one) refreshes
# cleanly instead of duplicating data.
$wi = 0; $wTot = $Report.Keys.Count
foreach ($name in $Report.Keys) {
    $wi++
    $data = @($Report[$name])
    Write-Activity -Activity 'Writing worksheets' -Current $wi -Total $wTot -Item $name
    if ($data.Count -eq 0) {
        # $Report only contains sheets produced by collectors in THIS run. If one of
        # those collectors now returns zero rows, retaining an older populated sheet
        # would make the merged workbook stale/misleading. Remove the old sheet.
        if (Test-Path $WorkbookPath) {
            try {
                $emptyPkg = Open-ExcelPackage -Path $WorkbookPath
                $oldSheet = $emptyPkg.Workbook.Worksheets[$name]
                if ($oldSheet) {
                    $emptyPkg.Workbook.Worksheets.Delete($name)
                    Close-ExcelPackage $emptyPkg -Save
                    Write-Verbose2 "Removed stale sheet '$name' because the refreshed result is empty."
                }
                else { Close-ExcelPackage $emptyPkg -NoSave }
            }
            catch { Write-Log "Could not clear empty worksheet '$name': $($_.Exception.Message)" 'WARN' }
        }
        continue
    }
    Write-Verbose2 "Writing sheet '$name' ($($data.Count) rows)..."
    $data | Export-Excel -Path $WorkbookPath -WorksheetName $name -ClearSheet `
        -AutoSize -BoldTopRow -FreezeTopRow -AutoFilter -TableStyle Medium2
}
Complete-Activity 'Writing worksheets'

# Rebuild 00_Summary from the FINAL workbook contents (spans prior + new sheets),
# reading row counts from each worksheet's dimension rather than re-importing data.
$summary = $null
try {
    $pkg = Open-ExcelPackage -Path $WorkbookPath
    $summary = foreach ($ws in $pkg.Workbook.Worksheets) {
        if ($ws.Name -eq '00_Summary') { continue }
        [pscustomobject]@{
            Worksheet = $ws.Name
            RowCount  = $(if ($ws.Dimension) { $ws.Dimension.Rows - 1 } else { 0 })
        }
    }
    Close-ExcelPackage $pkg -NoSave
}
catch {
    Write-Log "Could not enumerate sheets for summary: $($_.Exception.Message)" 'WARN'
    $summary = foreach ($name in $Report.Keys) {
        [pscustomobject]@{ Worksheet = $name; RowCount = @($Report[$name]).Count }
    }
}
$summary | Sort-Object Worksheet |
    Export-Excel -Path $WorkbookPath -WorksheetName '00_Summary' -ClearSheet -MoveToStart `
        -AutoSize -BoldTopRow -FreezeTopRow -TableStyle Medium2

Write-Log "Done. Workbook: $WorkbookPath"
Write-Log "Log: $LogPath"

# Best-effort disconnects
if ($IncludeCloud) {
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}
    try { Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue } catch {}
}

Write-Host "`nDiscovery complete -> $WorkbookPath" -ForegroundColor Green

#endregion
