<#
.SYNOPSIS
    Exports the Entra Connect / Cloud Sync configuration for divestiture discovery.

.DESCRIPTION
    Read-only discovery of the hybrid synchronization configuration the SoW calls out as a common
    failure point when re-anchoring identities into a new forest and tenant:
      - Sync service: whether the local host runs Entra Connect (ADSync) or Cloud Sync, its version,
        and whether it is in staging mode (HA).
      - Global settings: sync scheduler interval, and the source-anchor / soft-match posture
        (ms-DS-ConsistencyGuid vs objectGUID).
      - Connectors: each AD/Entra connector, its type, and partition/OU inclusion scope.
      - Sync rules: inbound/outbound rules with direction, connector, link type, and precedence.

    Facts are read from the ADSync module when present (Get-ADSyncConnector / Get-ADSyncRule /
    Get-ADSyncGlobalSettings / Get-ADSyncScheduler). When ADSync is not available the script reports
    that Entra Connect is not installed on this host and exits cleanly - it is meant to run on the
    sync server. Row shaping is separated so it is unit-testable with mock data. The script makes no
    changes.

    This complements Get-EntraDirSyncProtectionSettings.ps1 (which reads the tenant-side hard/soft
    match protection flags via Graph); it does not duplicate that check.

.PARAMETER OutputFolder
    Optional folder for the CSV outputs (Connectors.csv, SyncRules.csv, GlobalSettings.csv).
    If omitted, results are only written to the pipeline/table.

.EXAMPLE
    .\Export-EntraConnectConfiguration.ps1 -OutputFolder .\EntraConnect
#>

[CmdletBinding()]
param(
    [string]$OutputFolder = ''
)

$ErrorActionPreference = 'Stop'

function New-ConnectorRow {
    param([Parameter(Mandatory = $true)] $Connector)

    $inclusions = [System.Collections.Generic.List[string]]::new()
    foreach ($partition in @($Connector.Partitions)) {
        foreach ($container in @($partition.ConnectorPartitionScope.ContainerInclusionList | Where-Object { $_ })) {
            $inclusions.Add([string]$container)
        }
    }

    return [pscustomobject]@{
        Name              = [string]$Connector.Name
        ConnectorType     = [string]$Connector.Type
        Description       = [string]$Connector.Description
        PartitionCount    = @($Connector.Partitions).Count
        IncludedContainers = ($inclusions.ToArray() -join ';')
    }
}

function New-SyncRuleRow {
    param([Parameter(Mandatory = $true)] $Rule)
    return [pscustomobject]@{
        Name          = [string]$Rule.Name
        Direction     = [string]$Rule.Direction
        Connector     = [string]$Rule.ConnectorName
        SourceObjectType = [string]$Rule.SourceObjectType
        TargetObjectType = [string]$Rule.TargetObjectType
        LinkType      = [string]$Rule.LinkType
        Precedence    = [int]$Rule.Precedence
        Disabled      = [bool]$Rule.Disabled
    }
}

function ConvertTo-GlobalSettingsRow {
    param(
        [Parameter(Mandatory = $true)] [AllowNull()] $GlobalSettings,
        [Parameter(Mandatory = $true)] [AllowNull()] $Scheduler,
        [Parameter(Mandatory = $true)] [string]$Version
    )

    $parameters = @{}
    foreach ($parameter in @($GlobalSettings.Parameters)) {
        if ($parameter.Name) { $parameters[[string]$parameter.Name] = [string]$parameter.Value }
    }

    $sourceAnchor = ''
    foreach ($key in @('Microsoft.SourceAnchor','SourceAnchor','Microsoft.OptionalFeature.ConsistencyGuid')) {
        if ($parameters.ContainsKey($key)) { $sourceAnchor = $parameters[$key]; break }
    }

    return [pscustomobject]@{
        Version               = $Version
        StagingModeEnabled    = [bool]$Scheduler.StagingModeEnabled
        SyncCycleEnabled      = [bool]$Scheduler.SyncCycleEnabled
        AllowedSyncCycleInterval = [string]$Scheduler.AllowedSyncCycleInterval
        SourceAnchorAttribute = $sourceAnchor
        ConsistencyGuidInUse  = ($sourceAnchor -match 'consistencyGuid|ms-DS-ConsistencyGuid')
    }
}

function Get-EntraConnectFacts {
    $facts = @{
        Installed = $false; Version = ''; Connectors = @(); SyncRules = @(); GlobalSettings = $null; Scheduler = $null
    }

    if (-not (Get-Command -Name Get-ADSyncConnector -ErrorAction SilentlyContinue)) {
        try { Import-Module ADSync -ErrorAction Stop } catch { }
    }
    if (-not (Get-Command -Name Get-ADSyncConnector -ErrorAction SilentlyContinue)) {
        return $facts
    }

    $facts.Installed = $true
    try {
        $adSyncModule = Get-Module -Name ADSync
        if ($adSyncModule) { $facts.Version = [string]$adSyncModule.Version }
    } catch { }
    try { $facts.Connectors = @(Get-ADSyncConnector -ErrorAction Stop) } catch { }
    try { $facts.SyncRules = @(Get-ADSyncRule -ErrorAction Stop) } catch { }
    try { $facts.GlobalSettings = Get-ADSyncGlobalSettings -ErrorAction Stop } catch { }
    try { $facts.Scheduler = Get-ADSyncScheduler -ErrorAction Stop } catch { }

    return $facts
}

function Get-EntraConnectInventory {
    param([Parameter(Mandatory = $true)] $Facts)

    return @{
        Installed      = [bool]$Facts.Installed
        GlobalSettings = ConvertTo-GlobalSettingsRow -GlobalSettings $Facts.GlobalSettings -Scheduler $Facts.Scheduler -Version ([string]$Facts.Version)
        Connectors     = @(@($Facts.Connectors) | ForEach-Object { New-ConnectorRow -Connector $_ })
        SyncRules      = @(@($Facts.SyncRules) | ForEach-Object { New-SyncRuleRow -Rule $_ })
    }
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
Write-Host 'Collecting Entra Connect (ADSync) configuration...' -ForegroundColor Cyan
$facts = Get-EntraConnectFacts

if (-not $facts.Installed) {
    Write-Host 'Entra Connect (ADSync) is not installed on this host. Run this script on the Entra Connect sync server.' -ForegroundColor Yellow
    return
}

$inventory = Get-EntraConnectInventory -Facts $facts

Write-Host "`n=== Global settings ===" -ForegroundColor Yellow
$inventory.GlobalSettings | Format-List
Write-Host "=== Connectors ($(@($inventory.Connectors).Count)) ===" -ForegroundColor Yellow
$inventory.Connectors | Format-Table Name, ConnectorType, PartitionCount, IncludedContainers -AutoSize
Write-Host "=== Sync rules ($(@($inventory.SyncRules).Count)) ===" -ForegroundColor Yellow
$inventory.SyncRules | Format-Table Name, Direction, Connector, Precedence, Disabled -AutoSize

if (-not [string]::IsNullOrWhiteSpace($OutputFolder)) {
    if (-not (Test-Path -Path $OutputFolder)) { New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null }
    @($inventory.GlobalSettings) | Export-Csv -Path (Join-Path $OutputFolder 'GlobalSettings.csv') -NoTypeInformation -Encoding UTF8
    $inventory.Connectors | Export-Csv -Path (Join-Path $OutputFolder 'Connectors.csv') -NoTypeInformation -Encoding UTF8
    $inventory.SyncRules | Export-Csv -Path (Join-Path $OutputFolder 'SyncRules.csv') -NoTypeInformation -Encoding UTF8
    Write-Host ("Reports exported to: {0}" -f $OutputFolder) -ForegroundColor Green
}
