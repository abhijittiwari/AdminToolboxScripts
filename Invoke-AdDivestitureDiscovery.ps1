<#
.SYNOPSIS
    Runs the divestiture discovery modules and consolidates their outputs into one run folder.

.DESCRIPTION
    Orchestrates the read-only discovery scripts that gather the Discovery & Design evidence a
    tenant-to-tenant / new-forest divestiture needs, and writes each module's output plus a
    consolidated index into a single per-run folder:

      On-premises AD (always runs where the ActiveDirectory module is present):
        AdForest       Export-AdForestInventory.ps1
        AdSites        Export-AdSiteTopologyInventory.ps1
        AdTrusts       Export-AdTrustInventory.ps1
        AdOuGpo        Export-AdOuGpoInventory.ps1
        AdPopulation   Export-AdObjectPopulationInventory.ps1
        AdPrivileged   Export-AdPrivilegedIdentityInventory.ps1
        AdSidHistory   Export-AdSidHistoryInventory.ps1
        AdInfra        Export-AdInfraDependencyInventory.ps1
        EntraConnect   Export-EntraConnectConfiguration.ps1
        ExchangeOnPrem Export-ExchangeOnPremInventory.ps1

      Cloud modules (require a Graph / Exchange Online / Teams / SCC session; opt in with -IncludeCloud):
        EntraAuth      Export-EntraAuthenticationModel.ps1
        M365Volume     Export-M365WorkloadVolumetrics.ps1
        M365License    Export-M365LicenseInventory.ps1
        TeamsVoice     Export-TeamsVoiceInventory.ps1
        Purview        Export-PurviewComplianceInventory.ps1

    Each module runs into its own sub-folder (or CSV) under the run folder. A module that throws is
    recorded as an Error in the run index rather than aborting the run. The module-to-script catalogue
    and the run planning are separated so they are unit-testable. This orchestrator only reads.

.PARAMETER Module
    Which modules to run. Defaults to the on-premises AD set. Cloud modules are added via -IncludeCloud
    or by naming them explicitly.

.PARAMETER IncludeCloud
    Add the cloud modules (EntraAuth, M365Volume, M365License, TeamsVoice, Purview). They require the
    relevant connected sessions.

.PARAMETER OutputFolder
    Root folder for the run. Defaults to a timestamped folder under the current location.

.PARAMETER Credential
    Optional credential forwarded to the on-premises modules that accept one.

.PARAMETER HoneytokenNamePattern
    (Unused placeholder for parity with the hardening orchestrator; discovery modules take no honeytoken input.)

.EXAMPLE
    .\Invoke-AdDivestitureDiscovery.ps1 -OutputFolder .\DiscoveryRun

.EXAMPLE
    .\Invoke-AdDivestitureDiscovery.ps1 -IncludeCloud
#>

[CmdletBinding()]
param(
    [ValidateSet('AdForest','AdSites','AdTrusts','AdOuGpo','AdPopulation','AdPrivileged','AdSidHistory','AdInfra','EntraConnect','ExchangeOnPrem','EntraAuth','M365Volume','M365License','TeamsVoice','Purview')]
    [string[]]$Module = @('AdForest','AdSites','AdTrusts','AdOuGpo','AdPopulation','AdPrivileged','AdSidHistory','AdInfra','EntraConnect','ExchangeOnPrem'),
    [switch]$IncludeCloud,
    [string]$OutputFolder = '',
    [pscredential]$Credential
)

$ErrorActionPreference = 'Stop'

function Get-DiscoveryCatalog {
    # Ordered catalogue: module key -> script, output kind (Folder|File), and whether it needs a cloud session.
    return [ordered]@{
        AdForest       = @{ Script = 'Export-AdForestInventory.ps1';              Output = 'Folder'; Cloud = $false }
        AdSites        = @{ Script = 'Export-AdSiteTopologyInventory.ps1';         Output = 'Folder'; Cloud = $false }
        AdTrusts       = @{ Script = 'Export-AdTrustInventory.ps1';                Output = 'File';   Cloud = $false }
        AdOuGpo        = @{ Script = 'Export-AdOuGpoInventory.ps1';                Output = 'Folder'; Cloud = $false }
        AdPopulation   = @{ Script = 'Export-AdObjectPopulationInventory.ps1';     Output = 'Folder'; Cloud = $false }
        AdPrivileged   = @{ Script = 'Export-AdPrivilegedIdentityInventory.ps1';   Output = 'Folder'; Cloud = $false }
        AdSidHistory   = @{ Script = 'Export-AdSidHistoryInventory.ps1';           Output = 'File';   Cloud = $false }
        AdInfra        = @{ Script = 'Export-AdInfraDependencyInventory.ps1';      Output = 'Folder'; Cloud = $false }
        EntraConnect   = @{ Script = 'Export-EntraConnectConfiguration.ps1';       Output = 'Folder'; Cloud = $false }
        ExchangeOnPrem = @{ Script = 'Export-ExchangeOnPremInventory.ps1';         Output = 'Folder'; Cloud = $false }
        EntraAuth      = @{ Script = 'Export-EntraAuthenticationModel.ps1';        Output = 'File';   Cloud = $true }
        M365Volume     = @{ Script = 'Export-M365WorkloadVolumetrics.ps1';         Output = 'File';   Cloud = $true }
        M365License    = @{ Script = 'Export-M365LicenseInventory.ps1';            Output = 'Folder'; Cloud = $true }
        TeamsVoice     = @{ Script = 'Export-TeamsVoiceInventory.ps1';             Output = 'Folder'; Cloud = $true }
        Purview        = @{ Script = 'Export-PurviewComplianceInventory.ps1';      Output = 'Folder'; Cloud = $true }
    }
}

function Get-DiscoveryRunPlan {
    param(
        [Parameter(Mandatory = $true)] [string[]]$Module,
        [bool]$IncludeCloud = $false
    )

    $catalog = Get-DiscoveryCatalog
    $selected = [System.Collections.Generic.List[string]]::new()
    foreach ($item in $Module) { if ($selected -notcontains $item) { $selected.Add($item) } }
    if ($IncludeCloud) {
        foreach ($key in $catalog.Keys) {
            if ($catalog[$key].Cloud -and $selected -notcontains $key) { $selected.Add($key) }
        }
    }

    $plan = [System.Collections.Generic.List[object]]::new()
    foreach ($key in $catalog.Keys) {
        if ($selected -contains $key) {
            $plan.Add([pscustomobject]@{ Module = $key; Script = $catalog[$key].Script; Output = $catalog[$key].Output; Cloud = $catalog[$key].Cloud })
        }
    }
    return $plan.ToArray()
}

function Get-DiscoveryModuleArgument {
    param(
        [Parameter(Mandatory = $true)] $PlanEntry,
        [Parameter(Mandatory = $true)] [string]$RunFolder,
        [pscredential]$Credential
    )

    $arguments = @{}
    if ($PlanEntry.Output -eq 'Folder') {
        $arguments.OutputFolder = Join-Path $RunFolder $PlanEntry.Module
    }
    else {
        $arguments.OutputPath = Join-Path $RunFolder ("{0}.csv" -f $PlanEntry.Module)
    }
    # Only the on-premises modules that query AD accept -Credential.
    if ($Credential -and -not $PlanEntry.Cloud -and $PlanEntry.Module -notin @('EntraConnect','ExchangeOnPrem')) {
        $arguments.Credential = $Credential
    }
    return $arguments
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
if ([string]::IsNullOrWhiteSpace($OutputFolder)) {
    $OutputFolder = Join-Path (Get-Location) ("AdDivestitureDiscovery_{0:yyyyMMdd_HHmmss}" -f (Get-Date))
}
if (-not (Test-Path -Path $OutputFolder)) { New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null }

$plan = @(Get-DiscoveryRunPlan -Module $Module -IncludeCloud:$IncludeCloud)
Write-Host ("Running {0} discovery module(s): {1}" -f $plan.Count, (($plan | ForEach-Object Module) -join ', ')) -ForegroundColor Cyan

$index = [System.Collections.Generic.List[object]]::new()
foreach ($entry in $plan) {
    $scriptPath = Join-Path $PSScriptRoot $entry.Script
    Write-Host ("`n=== {0}: {1} ===" -f $entry.Module, $entry.Script) -ForegroundColor Cyan

    if (-not (Test-Path -Path $scriptPath)) {
        Write-Host ("  Module script not found: {0}" -f $scriptPath) -ForegroundColor Red
        $index.Add([pscustomobject]@{ Module = $entry.Module; Script = $entry.Script; Status = 'Error'; Detail = "Script not found: $scriptPath" })
        continue
    }

    $arguments = Get-DiscoveryModuleArgument -PlanEntry $entry -RunFolder $OutputFolder -Credential $Credential
    try {
        & $scriptPath @arguments
        $target = if ($arguments.ContainsKey('OutputFolder')) { $arguments.OutputFolder } else { $arguments.OutputPath }
        $index.Add([pscustomobject]@{ Module = $entry.Module; Script = $entry.Script; Status = 'Completed'; Detail = $target })
    }
    catch {
        Write-Host ("  Module failed: {0}" -f $_.Exception.Message) -ForegroundColor Red
        $index.Add([pscustomobject]@{ Module = $entry.Module; Script = $entry.Script; Status = 'Error'; Detail = $_.Exception.Message })
    }
}

$indexPath = Join-Path $OutputFolder 'DiscoveryIndex.csv'
$index | Export-Csv -Path $indexPath -NoTypeInformation -Encoding UTF8

Write-Host "`n===== Discovery run summary =====" -ForegroundColor Yellow
$index | Format-Table Module, Status, Detail -AutoSize
$completed = @($index | Where-Object Status -eq 'Completed').Count
$errors = @($index | Where-Object Status -eq 'Error').Count
Write-Host ("{0} module(s): {1} completed, {2} error(s)" -f $index.Count, $completed, $errors) -ForegroundColor Yellow
Write-Host ("Run index: {0}" -f $indexPath) -ForegroundColor Green
