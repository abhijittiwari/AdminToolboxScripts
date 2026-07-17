<#
.SYNOPSIS
    Exports Microsoft 365 license assignments and SKU consumption for divestiture discovery.

.DESCRIPTION
    Read-only discovery of the licensing picture a tenant-to-tenant migration needs for target-tenant
    SKU mapping (a SoW review comment flags Microsoft/VMware licensing constraints):
      - Tenant SKUs: each subscribed SKU with its friendly name, enabled/consumed/warning unit counts,
        and remaining available units.
      - Per-user assignments: one row per (user, assigned SKU), including whether the assignment is
        direct or inherited from a group (group-based licensing), so the target-tenant mapping and any
        re-licensing effort is visible.

    SKU part numbers are mapped to friendly product names via a built-in lookup (extensible), falling
    back to the raw part number. Collection uses Microsoft Graph; the shaping is separated so it is
    unit-testable with mock data. The script makes no changes.

.PARAMETER OutputFolder
    Optional folder for the CSV outputs (TenantSkus.csv, UserLicenseAssignments.csv). If omitted,
    results are only written to the pipeline/table.

.EXAMPLE
    .\Export-M365LicenseInventory.ps1 -OutputFolder .\Licensing
#>

[CmdletBinding()]
param(
    [string]$OutputFolder = ''
)

$ErrorActionPreference = 'Stop'

function Get-SkuFriendlyName {
    param([Parameter(Mandatory = $true)] [string]$SkuPartNumber)

    $map = @{
        'ENTERPRISEPACK'            = 'Office 365 E3'
        'ENTERPRISEPREMIUM'         = 'Office 365 E5'
        'SPE_E3'                    = 'Microsoft 365 E3'
        'SPE_E5'                    = 'Microsoft 365 E5'
        'SPE_F1'                    = 'Microsoft 365 F3'
        'O365_BUSINESS_PREMIUM'     = 'Microsoft 365 Business Standard'
        'SPB'                       = 'Microsoft 365 Business Premium'
        'EMS'                       = 'Enterprise Mobility + Security E3'
        'EMSPREMIUM'                = 'Enterprise Mobility + Security E5'
        'AAD_PREMIUM'               = 'Entra ID P1'
        'AAD_PREMIUM_P2'            = 'Entra ID P2'
        'EXCHANGESTANDARD'          = 'Exchange Online (Plan 1)'
        'EXCHANGEENTERPRISE'        = 'Exchange Online (Plan 2)'
        'MCOEV'                     = 'Microsoft Teams Phone'
        'FLOW_FREE'                 = 'Power Automate Free'
        'POWER_BI_STANDARD'         = 'Power BI (free)'
    }
    if ($map.ContainsKey($SkuPartNumber)) { return $map[$SkuPartNumber] }
    return $SkuPartNumber
}

function New-TenantSkuRow {
    param([Parameter(Mandatory = $true)] $Sku)

    $enabled = [int](Get-ObjectPropertyValueSafe -InputObject $Sku.prepaidUnits -PropertyName 'enabled')
    $consumed = [int]$Sku.consumedUnits
    $warning = [int](Get-ObjectPropertyValueSafe -InputObject $Sku.prepaidUnits -PropertyName 'warning')
    return [pscustomobject]@{
        SkuPartNumber = [string]$Sku.skuPartNumber
        ProductName   = (Get-SkuFriendlyName -SkuPartNumber ([string]$Sku.skuPartNumber))
        SkuId         = [string]$Sku.skuId
        EnabledUnits  = $enabled
        ConsumedUnits = $consumed
        WarningUnits  = $warning
        AvailableUnits = ($enabled - $consumed)
    }
}

function Get-ObjectPropertyValueSafe {
    param([AllowNull()] $InputObject, [Parameter(Mandatory = $true)] [string]$PropertyName)
    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [hashtable]) { if ($InputObject.ContainsKey($PropertyName)) { return $InputObject[$PropertyName] } else { return $null } }
    $property = $InputObject.PSObject.Properties[$PropertyName]
    if ($property) { return $property.Value }
    return $null
}

function Get-M365LicenseInventory {
    param([Parameter(Mandatory = $true)] $Facts)

    $skuNameById = @{}
    foreach ($sku in @($Facts.Skus)) { $skuNameById[[string]$sku.skuId] = [string]$sku.skuPartNumber }

    $skuRows = @(@($Facts.Skus) | ForEach-Object { New-TenantSkuRow -Sku $_ })

    $assignmentRows = [System.Collections.Generic.List[object]]::new()
    foreach ($user in @($Facts.Users)) {
        $assignedByGroup = @{}
        foreach ($state in @($user.licenseAssignmentStates)) {
            $skuId = [string]$state.skuId
            $assignedByGroup[$skuId] = -not [string]::IsNullOrWhiteSpace([string]$state.assignedByGroup)
        }
        foreach ($license in @($user.assignedLicenses)) {
            $skuId = [string]$license.skuId
            $partNumber = if ($skuNameById.ContainsKey($skuId)) { $skuNameById[$skuId] } else { $skuId }
            $viaGroup = if ($assignedByGroup.ContainsKey($skuId)) { $assignedByGroup[$skuId] } else { $false }
            $assignmentRows.Add([pscustomobject]@{
                UserPrincipalName = [string]$user.userPrincipalName
                DisplayName       = [string]$user.displayName
                Enabled           = [bool]$user.accountEnabled
                SkuPartNumber     = $partNumber
                ProductName       = (Get-SkuFriendlyName -SkuPartNumber $partNumber)
                SkuId             = $skuId
                AssignedViaGroup  = $viaGroup
                AssignmentType    = if ($viaGroup) { 'Group' } else { 'Direct' }
            })
        }
    }

    return @{ TenantSkus = $skuRows; UserAssignments = $assignmentRows.ToArray() }
}

function Get-M365LicenseFacts {
    $facts = @{ Skus = @(); Users = @() }

    if (-not (Get-Command -Name Get-MgSubscribedSku -ErrorAction SilentlyContinue)) {
        throw 'Microsoft Graph licensing cmdlets are not available. Run Connect-MgGraph with Organization.Read.All and User.Read.All first.'
    }

    try {
        $facts.Skus = @(Get-MgSubscribedSku -All -ErrorAction Stop | ForEach-Object {
            [pscustomobject]@{ skuId = $_.SkuId; skuPartNumber = $_.SkuPartNumber; consumedUnits = $_.ConsumedUnits; prepaidUnits = $_.PrepaidUnits }
        })
    } catch { }

    try {
        $facts.Users = @(Get-MgUser -All -Property 'userPrincipalName','displayName','accountEnabled','assignedLicenses','licenseAssignmentStates' -ErrorAction Stop | Where-Object { @($_.AssignedLicenses).Count -gt 0 } | ForEach-Object {
            [pscustomobject]@{
                userPrincipalName = $_.UserPrincipalName
                displayName = $_.DisplayName
                accountEnabled = $_.AccountEnabled
                assignedLicenses = @($_.AssignedLicenses | ForEach-Object { [pscustomobject]@{ skuId = $_.SkuId } })
                licenseAssignmentStates = @($_.LicenseAssignmentStates | ForEach-Object { [pscustomobject]@{ skuId = $_.SkuId; assignedByGroup = $_.AssignedByGroup } })
            }
        })
    } catch { }

    return $facts
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
Write-Host 'Collecting Microsoft 365 license SKUs and assignments...' -ForegroundColor Cyan
$facts = Get-M365LicenseFacts
$inventory = Get-M365LicenseInventory -Facts $facts

Write-Host "`n=== Tenant SKUs ($(@($inventory.TenantSkus).Count)) ===" -ForegroundColor Yellow
$inventory.TenantSkus | Format-Table ProductName, SkuPartNumber, EnabledUnits, ConsumedUnits, AvailableUnits -AutoSize
Write-Host "=== User license assignments ($(@($inventory.UserAssignments).Count)) ===" -ForegroundColor Yellow
$groupAssigned = @($inventory.UserAssignments | Where-Object AssignedViaGroup)
Write-Host ("  {0} assignment(s), {1} group-based" -f @($inventory.UserAssignments).Count, $groupAssigned.Count) -ForegroundColor Yellow

if (-not [string]::IsNullOrWhiteSpace($OutputFolder)) {
    if (-not (Test-Path -Path $OutputFolder)) { New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null }
    $inventory.TenantSkus | Export-Csv -Path (Join-Path $OutputFolder 'TenantSkus.csv') -NoTypeInformation -Encoding UTF8
    $inventory.UserAssignments | Export-Csv -Path (Join-Path $OutputFolder 'UserLicenseAssignments.csv') -NoTypeInformation -Encoding UTF8
    Write-Host ("Reports exported to: {0}" -f $OutputFolder) -ForegroundColor Green
}
