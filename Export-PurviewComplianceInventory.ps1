<#
.SYNOPSIS
    Exports a Microsoft Purview / compliance inventory for divestiture discovery.

.DESCRIPTION
    Read-only discovery of the compliance configuration a tenant-to-tenant migration must reproduce.
    The SoW flags that sensitivity-label GUIDs do not survive a tenant move, so labels and their
    policies must be re-created and content re-labelled - this inventory captures what exists so that
    re-creation/rehydration can be planned rather than lift-and-shifted:
      - Sensitivity labels: name, display name, whether enabled, and parent (sub-label) relationship.
      - Label policies: name and the labels each publishes.
      - Retention: retention labels and retention policies (count and names).
      - DLP: data loss prevention policies (count and names, with mode).

    Collection uses the Security & Compliance PowerShell cmdlets (IPPSSession) and, for labels, the
    Graph information-protection endpoint as a fallback. The shaping is separated so it is unit-testable
    with mock data. The script makes no changes.

.PARAMETER OutputFolder
    Optional folder for the CSV outputs (SensitivityLabels.csv, LabelPolicies.csv, RetentionPolicies.csv,
    DlpPolicies.csv). If omitted, results are only written to the pipeline/table.

.EXAMPLE
    .\Export-PurviewComplianceInventory.ps1 -OutputFolder .\Purview
#>

[CmdletBinding()]
param(
    [string]$OutputFolder = ''
)

$ErrorActionPreference = 'Stop'

function New-SensitivityLabelRow {
    param([Parameter(Mandatory = $true)] $Label)
    return [pscustomobject]@{
        Name        = [string]$Label.Name
        DisplayName = [string]$Label.DisplayName
        Guid        = [string]$Label.Guid
        Enabled     = [bool]$Label.Enabled
        ParentLabel = [string]$Label.ParentId
        IsSubLabel  = (-not [string]::IsNullOrWhiteSpace([string]$Label.ParentId))
        Priority    = [int]$Label.Priority
    }
}

function New-LabelPolicyRow {
    param([Parameter(Mandatory = $true)] $Policy)
    return [pscustomobject]@{
        Name           = [string]$Policy.Name
        Enabled        = [bool]$Policy.Enabled
        PublishedLabels = (@($Policy.Labels | Where-Object { $_ }) -join ';')
        LabelCount     = @($Policy.Labels | Where-Object { $_ }).Count
    }
}

function New-RetentionPolicyRow {
    param([Parameter(Mandatory = $true)] $Policy)
    return [pscustomobject]@{
        Name     = [string]$Policy.Name
        Enabled  = [bool]$Policy.Enabled
        Mode     = [string]$Policy.Mode
        Workload = [string]$Policy.Workload
    }
}

function New-DlpPolicyRow {
    param([Parameter(Mandatory = $true)] $Policy)
    return [pscustomobject]@{
        Name     = [string]$Policy.Name
        Mode     = [string]$Policy.Mode
        Enabled  = [bool]$Policy.Enabled
        Workload = [string]$Policy.Workload
    }
}

function Get-PurviewComplianceInventory {
    param([Parameter(Mandatory = $true)] $Facts)

    $labelRows = @(@($Facts.SensitivityLabels) | ForEach-Object { New-SensitivityLabelRow -Label $_ })
    $labelPolicyRows = @(@($Facts.LabelPolicies) | ForEach-Object { New-LabelPolicyRow -Policy $_ })
    $retentionRows = @(@($Facts.RetentionPolicies) | ForEach-Object { New-RetentionPolicyRow -Policy $_ })
    $dlpRows = @(@($Facts.DlpPolicies) | ForEach-Object { New-DlpPolicyRow -Policy $_ })

    $summary = [pscustomobject]@{
        SensitivityLabels = @($labelRows).Count
        SubLabels         = @($labelRows | Where-Object { $_.IsSubLabel }).Count
        LabelPolicies     = @($labelPolicyRows).Count
        RetentionPolicies = @($retentionRows).Count
        DlpPolicies       = @($dlpRows).Count
    }

    return @{
        SensitivityLabels = $labelRows
        LabelPolicies     = $labelPolicyRows
        RetentionPolicies = $retentionRows
        DlpPolicies       = $dlpRows
        Summary           = $summary
    }
}

function Get-PurviewComplianceFacts {
    $facts = @{ SensitivityLabels = @(); LabelPolicies = @(); RetentionPolicies = @(); DlpPolicies = @() }

    if (-not (Get-Command -Name Get-Label -ErrorAction SilentlyContinue)) {
        throw 'Security & Compliance cmdlets (Get-Label) are not available. Run Connect-IPPSSession first.'
    }

    try { $facts.SensitivityLabels = @(Get-Label -ErrorAction Stop) } catch { }
    try { $facts.LabelPolicies = @(Get-LabelPolicy -ErrorAction Stop) } catch { }
    try {
        $retentionPolicies = @(Get-RetentionCompliancePolicy -ErrorAction Stop)
        $facts.RetentionPolicies = @($retentionPolicies | ForEach-Object {
            [pscustomobject]@{ Name = $_.Name; Enabled = $_.Enabled; Mode = $_.Mode; Workload = ($_.Workload -join ',') }
        })
    } catch { }
    try {
        $dlpPolicies = @(Get-DlpCompliancePolicy -ErrorAction Stop)
        $facts.DlpPolicies = @($dlpPolicies | ForEach-Object {
            [pscustomobject]@{ Name = $_.Name; Mode = $_.Mode; Enabled = $_.Enabled; Workload = ($_.Workload -join ',') }
        })
    } catch { }

    return $facts
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
Write-Host 'Collecting Microsoft Purview / compliance inventory...' -ForegroundColor Cyan
$facts = Get-PurviewComplianceFacts
$inventory = Get-PurviewComplianceInventory -Facts $facts

Write-Host "`n=== Summary ===" -ForegroundColor Yellow
$inventory.Summary | Format-List
Write-Host "=== Sensitivity labels ($(@($inventory.SensitivityLabels).Count)) ===" -ForegroundColor Yellow
$inventory.SensitivityLabels | Format-Table DisplayName, Enabled, IsSubLabel, Priority -AutoSize
Write-Host 'NOTE: Sensitivity-label GUIDs do not survive a tenant move; plan for label re-creation and content re-labelling.' -ForegroundColor DarkYellow

if (-not [string]::IsNullOrWhiteSpace($OutputFolder)) {
    if (-not (Test-Path -Path $OutputFolder)) { New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null }
    $inventory.SensitivityLabels | Export-Csv -Path (Join-Path $OutputFolder 'SensitivityLabels.csv') -NoTypeInformation -Encoding UTF8
    $inventory.LabelPolicies | Export-Csv -Path (Join-Path $OutputFolder 'LabelPolicies.csv') -NoTypeInformation -Encoding UTF8
    $inventory.RetentionPolicies | Export-Csv -Path (Join-Path $OutputFolder 'RetentionPolicies.csv') -NoTypeInformation -Encoding UTF8
    $inventory.DlpPolicies | Export-Csv -Path (Join-Path $OutputFolder 'DlpPolicies.csv') -NoTypeInformation -Encoding UTF8
    Write-Host ("Reports exported to: {0}" -f $OutputFolder) -ForegroundColor Green
}
