<#
.SYNOPSIS
    Runs the Appendix B security hardening assessment modules and merges results into one report.

.DESCRIPTION
    Orchestrates the seven hardening assessment scripts that implement the Appendix B baseline
    controls and consolidates their per-control rows into a single CSV plus an on-screen summary:

      DC   Export-DcHardeningAssessment.ps1               (DC-001 .. DC-022)
      DJ   Export-DomainComputerHardeningAssessment.ps1   (DJ-001 .. DJ-020)
      AAD  Export-EntraIdHardeningAssessment.ps1          (AAD-001 .. AAD-015)
      DNS  Export-DnsHardeningAssessment.ps1              (DNS-001 .. DNS-010)
      ADCS Export-AdcsHardeningAssessment.ps1             (ADCS-001 .. ADCS-010)
      BR   Export-AdBackupRecoveryAssessment.ps1          (BR-001 .. BR-010)
      MON  Export-AdMonitoringAssessment.ps1              (MON-001 .. MON-010)

    Each module is invoked with -OutputPath into a per-run folder, then the individual CSVs are
    merged (a Module column is added). Modules that throw are recorded as a single Error row so
    one failing module never aborts the run. The Entra ID module (AAD) requires a Microsoft Graph
    connection and is skipped unless -IncludeEntraId is specified.

    This script only reads; every underlying module is read-only.

.PARAMETER Module
    Which modules to run. Default: DC, DJ, DNS, ADCS, BR, MON (on-premises). Add AAD via
    -IncludeEntraId or by listing it explicitly.

.PARAMETER IncludeEntraId
    Include the Entra ID (AAD) module. Requires an established Microsoft Graph session.

.PARAMETER OutputFolder
    Folder for the per-module CSVs and the merged report. Defaults to a timestamped folder
    under the current location.

.PARAMETER Credential
    Optional credential forwarded to the on-premises modules that accept one.

.PARAMETER BreakGlassUpn
    Forwarded to the Entra ID module for break-glass account verification (AAD-007).

.PARAMETER ComputersCsv
    Forwarded to the domain-joined computer module (DJ) to scope the assessed sample.

.PARAMETER HoneytokenNamePattern
    Forwarded to the monitoring module (MON) for honeytoken account verification (MON-007).

.EXAMPLE
    .\Invoke-AdSecurityHardeningAssessment.ps1 -OutputFolder .\HardeningRun

.EXAMPLE
    .\Invoke-AdSecurityHardeningAssessment.ps1 -IncludeEntraId -BreakGlassUpn bg1@contoso.com
#>

[CmdletBinding()]
param(
    [ValidateSet('DC','DJ','AAD','DNS','ADCS','BR','MON')]
    [string[]]$Module = @('DC','DJ','DNS','ADCS','BR','MON'),
    [switch]$IncludeEntraId,
    [string]$OutputFolder = '',
    [pscredential]$Credential,
    [string[]]$BreakGlassUpn = @(),
    [string]$ComputersCsv = '',
    [string]$HoneytokenNamePattern = ''
)

$ErrorActionPreference = 'Stop'

function Get-ModuleRunPlan {
    param(
        [Parameter(Mandatory = $true)] [string[]]$Module,
        [bool]$IncludeEntraId = $false
    )

    $selected = [System.Collections.Generic.List[string]]::new()
    foreach ($item in $Module) { if ($selected -notcontains $item) { $selected.Add($item) } }
    if ($IncludeEntraId -and $selected -notcontains 'AAD') { $selected.Add('AAD') }

    $catalog = [ordered]@{
        DC   = 'Export-DcHardeningAssessment.ps1'
        DJ   = 'Export-DomainComputerHardeningAssessment.ps1'
        AAD  = 'Export-EntraIdHardeningAssessment.ps1'
        DNS  = 'Export-DnsHardeningAssessment.ps1'
        ADCS = 'Export-AdcsHardeningAssessment.ps1'
        BR   = 'Export-AdBackupRecoveryAssessment.ps1'
        MON  = 'Export-AdMonitoringAssessment.ps1'
    }

    $plan = [System.Collections.Generic.List[object]]::new()
    foreach ($key in $catalog.Keys) {
        if ($selected -contains $key) {
            $plan.Add([pscustomobject]@{ Module = $key; Script = $catalog[$key] })
        }
    }
    return $plan.ToArray()
}

function Get-ModuleArgument {
    param(
        [Parameter(Mandatory = $true)] [string]$Module,
        [Parameter(Mandatory = $true)] [string]$OutputPath,
        [pscredential]$Credential,
        [string[]]$BreakGlassUpn = @(),
        [string]$ComputersCsv = '',
        [string]$HoneytokenNamePattern = ''
    )

    $arguments = @{ OutputPath = $OutputPath }
    switch ($Module) {
        'DC'   { if ($Credential) { $arguments.Credential = $Credential } }
        'DJ'   {
            if ($Credential) { $arguments.Credential = $Credential }
            if (-not [string]::IsNullOrWhiteSpace($ComputersCsv)) { $arguments.ComputersCsv = $ComputersCsv }
        }
        'AAD'  { if ($BreakGlassUpn.Count -gt 0) { $arguments.BreakGlassUpn = $BreakGlassUpn } }
        'DNS'  { if ($Credential) { $arguments.Credential = $Credential } }
        'ADCS' { }
        'BR'   { if ($Credential) { $arguments.Credential = $Credential } }
        'MON'  {
            if ($Credential) { $arguments.Credential = $Credential }
            if (-not [string]::IsNullOrWhiteSpace($HoneytokenNamePattern)) { $arguments.HoneytokenNamePattern = $HoneytokenNamePattern }
        }
    }
    return $arguments
}

function Merge-AssessmentResults {
    param([Parameter(Mandatory = $true)] [object[]]$ModuleResults)

    $merged = [System.Collections.Generic.List[object]]::new()
    foreach ($result in $ModuleResults) {
        foreach ($row in @($result.Rows)) {
            $merged.Add([pscustomobject]@{
                Module    = $result.Module
                ControlId = $row.ControlId
                Control   = $row.Control
                Target    = $row.Target
                Status    = $row.Status
                Evidence  = $row.Evidence
            })
        }
    }
    return $merged.ToArray()
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
if ([string]::IsNullOrWhiteSpace($OutputFolder)) {
    $OutputFolder = Join-Path (Get-Location) ("AdHardeningAssessment_{0:yyyyMMdd_HHmmss}" -f (Get-Date))
}
if (-not (Test-Path -Path $OutputFolder)) {
    New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
}

$plan = @(Get-ModuleRunPlan -Module $Module -IncludeEntraId:$IncludeEntraId)
Write-Host ("Running {0} assessment module(s): {1}" -f $plan.Count, (($plan | ForEach-Object Module) -join ', ')) -ForegroundColor Cyan

$moduleResults = [System.Collections.Generic.List[object]]::new()
foreach ($entry in $plan) {
    $scriptPath = Join-Path $PSScriptRoot $entry.Script
    $csvPath = Join-Path $OutputFolder ("{0}.csv" -f $entry.Module)
    Write-Host ("`n=== {0}: {1} ===" -f $entry.Module, $entry.Script) -ForegroundColor Cyan

    if (-not (Test-Path -Path $scriptPath)) {
        Write-Host ("  Module script not found: {0}" -f $scriptPath) -ForegroundColor Red
        $moduleResults.Add([pscustomobject]@{ Module = $entry.Module; Rows = @([pscustomobject]@{ ControlId = "$($entry.Module)-000"; Control = 'Module execution'; Target = $entry.Script; Status = 'Error'; Evidence = "Module script not found: $scriptPath" }) })
        continue
    }

    $arguments = Get-ModuleArgument -Module $entry.Module -OutputPath $csvPath -Credential $Credential `
        -BreakGlassUpn $BreakGlassUpn -ComputersCsv $ComputersCsv -HoneytokenNamePattern $HoneytokenNamePattern
    try {
        & $scriptPath @arguments
        $rows = if (Test-Path -Path $csvPath) { @(Import-Csv -Path $csvPath) } else { @() }
        $moduleResults.Add([pscustomobject]@{ Module = $entry.Module; Rows = $rows })
    }
    catch {
        Write-Host ("  Module failed: {0}" -f $_.Exception.Message) -ForegroundColor Red
        $moduleResults.Add([pscustomobject]@{ Module = $entry.Module; Rows = @([pscustomobject]@{ ControlId = "$($entry.Module)-000"; Control = 'Module execution'; Target = $entry.Script; Status = 'Error'; Evidence = $_.Exception.Message }) })
    }
}

$merged = @(Merge-AssessmentResults -ModuleResults $moduleResults)
$mergedPath = Join-Path $OutputFolder 'AllControls.csv'
$merged | Export-Csv -Path $mergedPath -NoTypeInformation -Encoding UTF8

Write-Host "`n===== Consolidated summary =====" -ForegroundColor Yellow
$merged | Group-Object Module | ForEach-Object {
    $statusCounts = ($_.Group | Group-Object Status | Sort-Object Name | ForEach-Object { "{0}={1}" -f $_.Name, $_.Count }) -join ', '
    Write-Host ("{0,-5} {1} control row(s): {2}" -f $_.Name, $_.Count, $statusCounts)
}
$overall = ($merged | Group-Object Status | Sort-Object Name | ForEach-Object { "{0}={1}" -f $_.Name, $_.Count }) -join ', '
Write-Host ("TOTAL {0} row(s): {1}" -f $merged.Count, $overall) -ForegroundColor Yellow
Write-Host ("`nMerged report: {0}" -f $mergedPath) -ForegroundColor Green
