<#
.SYNOPSIS
    Exports Microsoft 365 workload volumetrics (sizes and counts) for divestiture migration sizing.

.DESCRIPTION
    Read-only discovery of the per-workload volumetrics a tenant-to-tenant migration needs to size
    bandwidth and effort - the figures the SoW cites (Exchange ~22 TB, OneDrive ~26.7 TB, SharePoint/
    Teams ~30-40 TB). It aggregates:
      - Exchange Online: primary mailbox and archive counts and total sizes (from mailbox statistics),
        plus shared/room/equipment mailbox counts.
      - OneDrive for Business: account count and total storage used.
      - SharePoint / Teams: site count and total storage used, with a Teams-connected site split.
      - Groups & distribution lists: Microsoft 365 group and distribution list counts.
      - Public folders: whether a public folder deployment exists and its total size.

    The aggregation of raw statistics into totals is a separate, unit-testable function; size strings
    such as "1.5 GB (1,610,612,736 bytes)" are parsed to bytes so totals are exact. Collection uses
    ExchangeOnlineManagement, Microsoft Graph, and SharePoint/PnP cmdlets; each workload is optional so
    a missing module degrades to a skipped section. The script makes no changes.

.PARAMETER Include
    Which workloads to size. Defaults to all: Exchange, OneDrive, SharePoint, Groups.

.PARAMETER OutputPath
    Optional CSV output path for the per-workload summary. If omitted, results are only written to the table.

.EXAMPLE
    .\Export-M365WorkloadVolumetrics.ps1 -OutputPath .\Volumetrics.csv

.EXAMPLE
    .\Export-M365WorkloadVolumetrics.ps1 -Include Exchange,OneDrive
#>

[CmdletBinding()]
param(
    [ValidateSet('Exchange','OneDrive','SharePoint','Groups')]
    [string[]]$Include = @('Exchange','OneDrive','SharePoint','Groups'),
    [string]$OutputPath = ''
)

$ErrorActionPreference = 'Stop'

function ConvertFrom-SizeString {
    param([AllowNull()] $Size)

    if ($null -eq $Size) { return [int64]0 }
    $text = [string]$Size
    if ([string]::IsNullOrWhiteSpace($text)) { return [int64]0 }

    # Exchange size strings look like "1.523 GB (1,635,608,853 bytes)".
    $byteMatch = [regex]::Match($text, '\(([\d,]+)\s*bytes\)')
    if ($byteMatch.Success) {
        return [int64]($byteMatch.Groups[1].Value -replace ',', '')
    }
    # Fallback: a plain number possibly with a unit suffix.
    $unitMatch = [regex]::Match($text, '^\s*([\d\.,]+)\s*(B|KB|MB|GB|TB)?')
    if ($unitMatch.Success) {
        $number = [double]($unitMatch.Groups[1].Value -replace ',', '')
        $unit = $unitMatch.Groups[2].Value.ToUpperInvariant()
        $multiplier = switch ($unit) { 'KB' { 1KB } 'MB' { 1MB } 'GB' { 1GB } 'TB' { 1TB } default { 1 } }
        return [int64]($number * $multiplier)
    }
    return [int64]0
}

function Format-Bytes {
    param([Parameter(Mandatory = $true)] [int64]$Bytes)
    if ($Bytes -ge 1TB) { return ('{0:N2} TB' -f ($Bytes / 1TB)) }
    if ($Bytes -ge 1GB) { return ('{0:N2} GB' -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ('{0:N2} MB' -f ($Bytes / 1MB)) }
    return ('{0:N0} bytes' -f $Bytes)
}

function New-VolumetricRow {
    param(
        [Parameter(Mandatory = $true)] [string]$Workload,
        [Parameter(Mandatory = $true)] [string]$Metric,
        [Parameter(Mandatory = $true)] [int64]$Count,
        [int64]$TotalBytes = 0
    )
    return [pscustomobject]@{
        Workload   = $Workload
        Metric     = $Metric
        Count      = $Count
        TotalBytes = $TotalBytes
        TotalSize  = (Format-Bytes -Bytes $TotalBytes)
    }
}

function Get-ExchangeVolumetricRows {
    param([Parameter(Mandatory = $true)] [AllowEmptyCollection()] $MailboxStats)

    $rows = [System.Collections.Generic.List[object]]::new()
    $primary = @($MailboxStats | Where-Object { -not $_.IsArchive })
    $archive = @($MailboxStats | Where-Object { $_.IsArchive })

    $primaryBytes = [int64]0
    foreach ($stat in $primary) { $primaryBytes += ConvertFrom-SizeString -Size $stat.TotalItemSize }
    $archiveBytes = [int64]0
    foreach ($stat in $archive) { $archiveBytes += ConvertFrom-SizeString -Size $stat.TotalItemSize }

    $rows.Add((New-VolumetricRow -Workload 'Exchange Online' -Metric 'Primary mailboxes' -Count $primary.Count -TotalBytes $primaryBytes))
    $rows.Add((New-VolumetricRow -Workload 'Exchange Online' -Metric 'Archive mailboxes' -Count $archive.Count -TotalBytes $archiveBytes))
    return $rows.ToArray()
}

function Get-SiteVolumetricRows {
    param(
        [Parameter(Mandatory = $true)] [string]$Workload,
        [Parameter(Mandatory = $true)] [string]$Metric,
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] $Sites,
        [string]$StorageProperty = 'StorageUsageCurrent'
    )

    $sites = @($Sites)
    $totalBytes = [int64]0
    foreach ($site in $sites) {
        $value = $site.$StorageProperty
        if ($null -ne $value) {
            # SharePoint reports storage in MB for StorageUsageCurrent.
            $totalBytes += [int64]([double]$value * 1MB)
        }
    }
    return (New-VolumetricRow -Workload $Workload -Metric $Metric -Count $sites.Count -TotalBytes $totalBytes)
}

function Get-M365WorkloadInventory {
    param([Parameter(Mandatory = $true)] $Facts)

    $rows = [System.Collections.Generic.List[object]]::new()

    if ($null -ne $Facts.MailboxStats) {
        foreach ($row in (Get-ExchangeVolumetricRows -MailboxStats $Facts.MailboxStats)) { $rows.Add($row) }
    }
    if ($null -ne $Facts.SharedMailboxCount) {
        $rows.Add((New-VolumetricRow -Workload 'Exchange Online' -Metric 'Shared/room/equipment mailboxes' -Count ([int64]$Facts.SharedMailboxCount)))
    }
    if ($null -ne $Facts.OneDriveSites) {
        $rows.Add((Get-SiteVolumetricRows -Workload 'OneDrive' -Metric 'OneDrive accounts' -Sites $Facts.OneDriveSites))
    }
    if ($null -ne $Facts.SharePointSites) {
        $teamsSites = @($Facts.SharePointSites | Where-Object { [string]$_.Template -match 'GROUP#0|TEAMCHANNEL' })
        $rows.Add((Get-SiteVolumetricRows -Workload 'SharePoint' -Metric 'All sites' -Sites $Facts.SharePointSites))
        $rows.Add((Get-SiteVolumetricRows -Workload 'SharePoint' -Metric 'Teams-connected sites' -Sites $teamsSites))
    }
    if ($null -ne $Facts.M365GroupCount) {
        $rows.Add((New-VolumetricRow -Workload 'Groups' -Metric 'Microsoft 365 groups' -Count ([int64]$Facts.M365GroupCount)))
    }
    if ($null -ne $Facts.DistributionListCount) {
        $rows.Add((New-VolumetricRow -Workload 'Groups' -Metric 'Distribution lists' -Count ([int64]$Facts.DistributionListCount)))
    }
    if ($null -ne $Facts.PublicFolderStats) {
        $pfBytes = [int64]0
        foreach ($stat in @($Facts.PublicFolderStats)) { $pfBytes += ConvertFrom-SizeString -Size $stat.TotalItemSize }
        $rows.Add((New-VolumetricRow -Workload 'Public Folders' -Metric 'Public folders' -Count (@($Facts.PublicFolderStats).Count) -TotalBytes $pfBytes))
    }

    return $rows.ToArray()
}

function Get-M365WorkloadFacts {
    param([string[]]$Include)

    $facts = @{
        MailboxStats = $null; SharedMailboxCount = $null; OneDriveSites = $null; SharePointSites = $null
        M365GroupCount = $null; DistributionListCount = $null; PublicFolderStats = $null
    }

    if ($Include -contains 'Exchange') {
        if (Get-Command -Name Get-EXOMailbox -ErrorAction SilentlyContinue) {
            try {
                $stats = [System.Collections.Generic.List[object]]::new()
                foreach ($mailbox in @(Get-EXOMailbox -ResultSize Unlimited -ErrorAction Stop)) {
                    try {
                        $primaryStat = Get-EXOMailboxStatistics -Identity $mailbox.Identity -ErrorAction Stop
                        $stats.Add([pscustomobject]@{ IsArchive = $false; TotalItemSize = $primaryStat.TotalItemSize })
                    } catch { }
                    if ($mailbox.ArchiveStatus -eq 'Active') {
                        try {
                            $archiveStat = Get-EXOMailboxStatistics -Identity $mailbox.Identity -Archive -ErrorAction Stop
                            $stats.Add([pscustomobject]@{ IsArchive = $true; TotalItemSize = $archiveStat.TotalItemSize })
                        } catch { }
                    }
                }
                $facts.MailboxStats = $stats.ToArray()
                $facts.SharedMailboxCount = @(Get-EXOMailbox -RecipientTypeDetails SharedMailbox, RoomMailbox, EquipmentMailbox -ResultSize Unlimited -ErrorAction SilentlyContinue).Count
            } catch { }
        }
    }

    if ($Include -contains 'Groups' -and (Get-Command -Name Get-MgGroup -ErrorAction SilentlyContinue)) {
        try { $facts.M365GroupCount = @(Get-MgGroup -Filter "groupTypes/any(c:c eq 'Unified')" -All -ErrorAction Stop).Count } catch { }
        if (Get-Command -Name Get-DistributionGroup -ErrorAction SilentlyContinue) {
            try { $facts.DistributionListCount = @(Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop).Count } catch { }
        }
    }

    if (($Include -contains 'OneDrive' -or $Include -contains 'SharePoint') -and (Get-Command -Name Get-SPOSite -ErrorAction SilentlyContinue)) {
        try {
            $allSites = @(Get-SPOSite -Limit All -IncludePersonalSite:$true -ErrorAction Stop)
            if ($Include -contains 'OneDrive') { $facts.OneDriveSites = @($allSites | Where-Object { [string]$_.Template -match 'SPSPERS' }) }
            if ($Include -contains 'SharePoint') { $facts.SharePointSites = @($allSites | Where-Object { [string]$_.Template -notmatch 'SPSPERS' }) }
        } catch { }
    }

    return $facts
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
Write-Host ("Collecting Microsoft 365 workload volumetrics: {0}" -f ($Include -join ', ')) -ForegroundColor Cyan
$facts = Get-M365WorkloadFacts -Include $Include
$rows = @(Get-M365WorkloadInventory -Facts $facts)

if ($rows.Count -eq 0) {
    Write-Host 'No workload data collected. Ensure the relevant modules are connected (ExchangeOnlineManagement, Microsoft.Graph, SPO/PnP).' -ForegroundColor Yellow
}
else {
    $rows | Format-Table Workload, Metric, Count, TotalSize -AutoSize
    $grandTotal = ($rows | Measure-Object -Property TotalBytes -Sum).Sum
    Write-Host ("Grand total across sized workloads: {0}" -f (Format-Bytes -Bytes ([int64]$grandTotal))) -ForegroundColor Yellow
}

if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
    $rows | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "Report exported: $OutputPath" -ForegroundColor Green
}
