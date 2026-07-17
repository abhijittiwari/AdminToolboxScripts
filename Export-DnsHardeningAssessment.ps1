<#
.SYNOPSIS
    Assesses Windows DNS Server hardening controls DNS-001 to DNS-010 (Appendix B baseline).

.DESCRIPTION
    Read-only assessment of AD-integrated DNS servers against the security hardening
    baseline: AD-integrated zones with secure dynamic updates, restricted zone transfers,
    aging/scavenging, forwarder restrictions, recursion and root-hint review, DNSSEC,
    query/change logging, the global query block list (WPAD/ISATAP), socket pool and
    cache locking, and zone/record permission restrictions.

    Uses the DnsServer RSAT module against each targeted DNS server. Each control is
    emitted as one row per server with Status: Pass / Fail / Warning / Manual / Error.
    Judgment controls (trusted forwarders, DNSSEC evaluation, zone ACL review) are
    reported as Manual with the evidence needed for the review.

    The script makes no changes to any system.

.PARAMETER DnsServer
    One or more DNS server names to assess. Defaults to all Domain Controllers in the domain.

.PARAMETER Credential
    Optional credential for AD discovery (DnsServer cmdlets use the current session identity).

.PARAMETER MinimumSocketPoolSize
    Minimum socket pool size before DNS-009 fails. Default 2500.

.PARAMETER OutputPath
    Optional CSV output path. If omitted, results are only written to the pipeline/table.

.EXAMPLE
    .\Export-DnsHardeningAssessment.ps1 -OutputPath .\DnsHardening.csv

.EXAMPLE
    .\Export-DnsHardeningAssessment.ps1 -DnsServer DC01,DC02
#>

[CmdletBinding()]
param(
    [string[]]$DnsServer = @(),
    [pscredential]$Credential,
    [int]$MinimumSocketPoolSize = 2500,
    [string]$OutputPath = ''
)

$ErrorActionPreference = 'Stop'

function New-ControlRow {
    param(
        [Parameter(Mandatory = $true)] [string]$ControlId,
        [Parameter(Mandatory = $true)] [string]$Control,
        [Parameter(Mandatory = $true)] [string]$Target,
        [Parameter(Mandatory = $true)] [ValidateSet('Pass','Fail','Warning','Manual','Error')] [string]$Status,
        [Parameter(Mandatory = $false)] [string]$Evidence = ''
    )

    return [pscustomobject]@{
        ControlId = $ControlId
        Control   = $Control
        Target    = $Target
        Status    = $Status
        Evidence  = $Evidence
    }
}

function Get-DnsServerFacts {
    param([Parameter(Mandatory = $true)] [string]$ComputerName)

    Import-Module DnsServer -ErrorAction Stop
    $facts = @{}

    $zones = @(Get-DnsServerZone -ComputerName $ComputerName)
    $facts.Zones = @()
    foreach ($zone in $zones) {
        $aging = $null
        if ($zone.ZoneType -eq 'Primary' -and -not $zone.IsAutoCreated) {
            try { $aging = Get-DnsServerZoneAging -ComputerName $ComputerName -Name $zone.ZoneName -ErrorAction Stop } catch { }
        }
        $facts.Zones += @{
            Name              = $zone.ZoneName
            ZoneType          = [string]$zone.ZoneType
            IsDsIntegrated    = [bool]$zone.IsDsIntegrated
            IsAutoCreated     = [bool]$zone.IsAutoCreated
            IsReverseLookupZone = [bool]$zone.IsReverseLookupZone
            IsSigned          = [bool]$zone.IsSigned
            DynamicUpdate     = [string]$zone.DynamicUpdate
            SecureSecondaries = [string]$zone.SecureSecondaries
            MasterServers     = @($zone.MasterServers | ForEach-Object { [string]$_ })
            AgingEnabled      = if ($aging) { [bool]$aging.AgingEnabled } else { $null }
        }
    }

    $scavenging = Get-DnsServerScavenging -ComputerName $ComputerName
    $facts.ScavengingEnabled = [bool]$scavenging.ScavengingState
    $facts.ScavengingIntervalDays = if ($scavenging.ScavengingInterval) { [int]$scavenging.ScavengingInterval.TotalDays } else { 0 }

    $forwarder = Get-DnsServerForwarder -ComputerName $ComputerName
    $facts.Forwarders = @($forwarder.IPAddress | ForEach-Object { [string]$_ })
    $facts.UseRootHint = [bool]$forwarder.UseRootHint

    $recursion = Get-DnsServerRecursion -ComputerName $ComputerName
    $facts.RecursionEnabled = [bool]$recursion.Enable

    try {
        $facts.RootHintCount = @((Get-DnsServerRootHint -ComputerName $ComputerName -ErrorAction Stop)).Count
    } catch { $facts.RootHintCount = 0 }

    $diagnostics = Get-DnsServerDiagnostics -ComputerName $ComputerName
    $facts.QueryLoggingEnabled = [bool]($diagnostics.Queries -and $diagnostics.Answers)
    $facts.EventLogLevel = [int]$diagnostics.EventLogLevel

    $blockList = Get-DnsServerGlobalQueryBlockList -ComputerName $ComputerName
    $facts.GlobalQueryBlockListEnabled = [bool]$blockList.Enable
    $facts.GlobalQueryBlockList = @($blockList.List | ForEach-Object { [string]$_ })

    $setting = Get-DnsServerSetting -ComputerName $ComputerName -All
    $facts.SocketPoolSize = [int]$setting.SocketPoolSize

    $cache = Get-DnsServerCache -ComputerName $ComputerName
    $facts.CacheLockingPercent = [int]$cache.LockingPercent

    return $facts
}

function Test-DnsServerControls {
    param(
        [Parameter(Mandatory = $true)] [string]$ServerName,
        [Parameter(Mandatory = $true)] $Facts,
        [int]$MinimumSocketPoolSize = 2500
    )

    $rows = [System.Collections.Generic.List[object]]::new()
    $add = {
        param($Id, $Control, $Status, $Evidence)
        $rows.Add((New-ControlRow -ControlId $Id -Control $Control -Target $ServerName -Status $Status -Evidence $Evidence))
    }

    $primaryZones = @(@($Facts.Zones) | Where-Object { $_.ZoneType -eq 'Primary' -and -not $_.IsAutoCreated -and $_.Name -ne 'TrustAnchors' })

    # DNS-001 AD-integrated zones, secure updates only
    $insecureZones = @($primaryZones | Where-Object { -not $_.IsDsIntegrated -or $_.DynamicUpdate -ne 'Secure' })
    if ($primaryZones.Count -eq 0) {
        & $add 'DNS-001' 'AD-integrated zones shall be used with secure dynamic updates only' 'Warning' 'No primary zones hosted on this server'
    }
    elseif ($insecureZones.Count -eq 0) {
        & $add 'DNS-001' 'AD-integrated zones shall be used with secure dynamic updates only' 'Pass' ("{0} primary zone(s) AD-integrated with secure dynamic updates" -f $primaryZones.Count)
    }
    else {
        & $add 'DNS-001' 'AD-integrated zones shall be used with secure dynamic updates only' 'Fail' ("Non-compliant zones: {0}" -f (($insecureZones | ForEach-Object { "{0} (DsIntegrated={1}, DynamicUpdate={2})" -f $_.Name, $_.IsDsIntegrated, $_.DynamicUpdate }) -join '; '))
    }

    # DNS-002 zone transfers restricted
    $openTransferZones = @($primaryZones | Where-Object { $_.SecureSecondaries -eq 'TransferAnyServer' })
    if ($openTransferZones.Count -eq 0) {
        & $add 'DNS-002' 'Zone transfers shall be restricted to authorized servers' 'Pass' 'No zone allows transfer to any server'
    }
    else {
        & $add 'DNS-002' 'Zone transfers shall be restricted to authorized servers' 'Fail' ("Zones allowing transfer to any server: {0}" -f (($openTransferZones | ForEach-Object Name) -join ', '))
    }

    # DNS-003 aging and scavenging
    $agingDisabledZones = @($primaryZones | Where-Object { -not $_.IsReverseLookupZone -and $_.AgingEnabled -eq $false })
    if (-not $Facts.ScavengingEnabled) {
        & $add 'DNS-003' 'Aging and scavenging of stale records shall be configured' 'Fail' 'Server-level scavenging is disabled'
    }
    elseif ($agingDisabledZones.Count -gt 0) {
        & $add 'DNS-003' 'Aging and scavenging of stale records shall be configured' 'Warning' ("Scavenging enabled (interval {0} day(s)) but aging disabled on: {1}" -f $Facts.ScavengingIntervalDays, (($agingDisabledZones | ForEach-Object Name) -join ', '))
    }
    else {
        & $add 'DNS-003' 'Aging and scavenging of stale records shall be configured' 'Pass' ("Scavenging enabled, interval {0} day(s); aging enabled on forward primary zones" -f $Facts.ScavengingIntervalDays)
    }

    # DNS-004 forwarders restricted to trusted resolvers
    $conditionalForwarders = @(@($Facts.Zones) | Where-Object { $_.ZoneType -eq 'Forwarder' })
    $forwarderEvidence = "Forwarders: {0}; conditional forwarders: {1}" -f `
        ($(if (@($Facts.Forwarders).Count -gt 0) { @($Facts.Forwarders) -join ', ' } else { '(none)' })), `
        ($(if ($conditionalForwarders.Count -gt 0) { ($conditionalForwarders | ForEach-Object { "{0} -> {1}" -f $_.Name, ($_.MasterServers -join '/') }) -join '; ' } else { '(none)' }))
    & $add 'DNS-004' 'Forwarders and conditional forwarders shall be restricted to trusted resolvers' 'Manual' ($forwarderEvidence + ' - confirm each resolver is an approved/trusted service')

    # DNS-005 recursion and root hints
    if (-not $Facts.RecursionEnabled) {
        & $add 'DNS-005' 'Recursion and root-hint settings shall be reviewed and hardened' 'Pass' 'Recursion disabled'
    }
    else {
        & $add 'DNS-005' 'Recursion and root-hint settings shall be reviewed and hardened' 'Manual' ("Recursion enabled (expected for internal resolvers); root hints: {0}, UseRootHint={1} - confirm the server is not internet-facing" -f $Facts.RootHintCount, $Facts.UseRootHint)
    }

    # DNS-006 DNSSEC evaluation
    $signedZones = @($primaryZones | Where-Object { $_.IsSigned })
    if ($signedZones.Count -gt 0) {
        & $add 'DNS-006' 'DNSSEC shall be evaluated and implemented for critical zones' 'Pass' ("Signed zones: {0}" -f (($signedZones | ForEach-Object Name) -join ', '))
    }
    else {
        & $add 'DNS-006' 'DNSSEC shall be evaluated and implemented for critical zones' 'Manual' ("No zones are DNSSEC-signed ({0} primary zone(s)); document the DNSSEC evaluation outcome for critical zones" -f $primaryZones.Count)
    }

    # DNS-007 query and change logging
    if ($Facts.QueryLoggingEnabled) {
        & $add 'DNS-007' 'DNS query and change logging shall be enabled and monitored' 'Pass' 'Diagnostic query/answer logging enabled; confirm logs are collected and monitored'
    }
    else {
        & $add 'DNS-007' 'DNS query and change logging shall be enabled and monitored' 'Warning' ("Diagnostic query logging disabled (EventLogLevel={0}); enable the DNS analytic log or debug logging and forward to the SIEM" -f $Facts.EventLogLevel)
    }

    # DNS-008 global query block list
    $blockList = @($Facts.GlobalQueryBlockList | ForEach-Object { $_.ToLowerInvariant() })
    if ($Facts.GlobalQueryBlockListEnabled -and $blockList -contains 'wpad' -and $blockList -contains 'isatap') {
        & $add 'DNS-008' 'The global query block list shall be enabled (WPAD/ISATAP)' 'Pass' ("Enabled; list: {0}" -f ($blockList -join ', '))
    }
    else {
        & $add 'DNS-008' 'The global query block list shall be enabled (WPAD/ISATAP)' 'Fail' ("Enabled={0}; list: {1} (must include wpad and isatap)" -f $Facts.GlobalQueryBlockListEnabled, ($blockList -join ', '))
    }

    # DNS-009 socket pool and cache locking
    if ($Facts.SocketPoolSize -ge $MinimumSocketPoolSize -and $Facts.CacheLockingPercent -eq 100) {
        & $add 'DNS-009' 'DNS socket pool and cache locking shall be configured to mitigate cache poisoning' 'Pass' ("SocketPoolSize={0}, CacheLockingPercent=100" -f $Facts.SocketPoolSize)
    }
    else {
        & $add 'DNS-009' 'DNS socket pool and cache locking shall be configured to mitigate cache poisoning' 'Fail' ("SocketPoolSize={0} (minimum {1}), CacheLockingPercent={2} (expected 100)" -f $Facts.SocketPoolSize, $MinimumSocketPoolSize, $Facts.CacheLockingPercent)
    }

    # DNS-010 zone/record permissions
    & $add 'DNS-010' 'Permissions to modify DNS zones and records shall be restricted to authorized administrators' 'Manual' 'Review DACLs on the MicrosoftDNS containers and zone objects, and DnsAdmins group membership, against the administrative tiering model'

    return $rows.ToArray()
}

function Get-TargetDnsServers {
    param(
        [string[]]$DnsServer,
        [pscredential]$Credential
    )

    if ($DnsServer.Count -gt 0) { return $DnsServer }
    Import-Module ActiveDirectory -ErrorAction Stop
    $adParams = @{ Filter = '*' }
    if ($Credential) { $adParams.Credential = $Credential }
    return @(Get-ADDomainController @adParams | ForEach-Object HostName)
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
$allRows = [System.Collections.Generic.List[object]]::new()
$targets = @(Get-TargetDnsServers -DnsServer $DnsServer -Credential $Credential)
Write-Host ("Assessing {0} DNS server(s): {1}" -f $targets.Count, ($targets -join ', ')) -ForegroundColor Cyan

foreach ($target in $targets) {
    Write-Host ("  Collecting DNS facts from {0}..." -f $target) -ForegroundColor Cyan
    try {
        $facts = Get-DnsServerFacts -ComputerName $target
        foreach ($row in (Test-DnsServerControls -ServerName $target -Facts $facts -MinimumSocketPoolSize $MinimumSocketPoolSize)) {
            $allRows.Add($row)
        }
    }
    catch {
        $allRows.Add((New-ControlRow -ControlId 'DNS-000' -Control 'Fact collection' -Target $target -Status 'Error' -Evidence $_.Exception.Message))
    }
}

$sorted = $allRows | Sort-Object ControlId, Target
$sorted | Format-Table ControlId, Target, Status, Evidence -AutoSize

$summary = $sorted | Group-Object Status | ForEach-Object { "{0}={1}" -f $_.Name, $_.Count }
Write-Host ("Summary: {0}" -f ($summary -join ', ')) -ForegroundColor Yellow

if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
    $sorted | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "Report exported: $OutputPath" -ForegroundColor Green
}
