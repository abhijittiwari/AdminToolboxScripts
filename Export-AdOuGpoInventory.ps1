<#
.SYNOPSIS
    Exports the OU structure and Group Policy inventory for divestiture discovery.

.DESCRIPTION
    Read-only discovery of the OU tree and every GPO in the domain, producing the inputs the SoW
    calls out for GPO rationalisation and modernisation:
      - OUs: distinguished name, depth, blocked-inheritance flag, and the GPOs linked to each OU.
      - GPOs: display name, GUID, enabled status, WMI filter, creation/modification dates, whether
        computer/user settings are configured, the number of links, and an "Intune-candidate" flag
        that marks GPOs whose settings are commonly deliverable through Intune (a modernisation cue,
        not a guarantee).

    GPO and OU objects are read with the GroupPolicy and ActiveDirectory modules. GPO XML reports are
    parsed to determine configured extensions and Intune-candidate categories; the parsing and row
    shaping are separated so they are unit-testable with sample XML/mock data. The script makes no
    changes.

.PARAMETER Server
    Optional Domain Controller / AD Web Services server to target for the queries.

.PARAMETER Credential
    Optional credential for the AD queries.

.PARAMETER OutputFolder
    Optional folder for the CSV outputs (OrganizationalUnits.csv, Gpos.csv). If omitted, results
    are only written to the pipeline/table.

.EXAMPLE
    .\Export-AdOuGpoInventory.ps1 -OutputFolder .\OuGpoInventory
#>

[CmdletBinding()]
param(
    [string]$Server = '',
    [pscredential]$Credential,
    [string]$OutputFolder = ''
)

$ErrorActionPreference = 'Stop'

function Get-OuDepth {
    param([Parameter(Mandatory = $true)] [string]$DistinguishedName)
    return ([regex]::Matches($DistinguishedName, '(?i)OU=')).Count
}

# Client-side extension display names that Intune can typically deliver via its own policy engine.
# Presence is a modernisation cue for the assessment, not an automatic mapping.
function Test-GpoIntuneCandidate {
    param([Parameter(Mandatory = $true)] [AllowEmptyCollection()] [string[]]$ExtensionNames)

    if ($ExtensionNames.Count -eq 0) { return $false }
    $intuneDeliverable = @(
        'Administrative Templates','Registry','Security','Internet Explorer','Windows Firewall',
        'Windows Defender','BitLocker','Certificate','Public Key','Software Restriction',
        'Application Control','Device','Preferences'
    )
    $blockers = @('Folder Redirection','Scripts','Software Installation','Disk Quota','Roaming')
    $hasBlocker = @($ExtensionNames | Where-Object { $ext = $_; @($blockers | Where-Object { $ext -match $_ }).Count -gt 0 }).Count -gt 0
    if ($hasBlocker) { return $false }
    $hasDeliverable = @($ExtensionNames | Where-Object { $ext = $_; @($intuneDeliverable | Where-Object { $ext -match $_ }).Count -gt 0 }).Count -gt 0
    return $hasDeliverable
}

function ConvertFrom-GpoReportXml {
    param([Parameter(Mandatory = $true)] [xml]$ReportXml)

    $gpo = $ReportXml.GPO
    $computerEnabled = $false
    $userEnabled = $false
    $computerExtCount = 0
    $userExtCount = 0
    $extensionNames = [System.Collections.Generic.List[string]]::new()

    if ($gpo.Computer) {
        if ($gpo.Computer.Enabled -eq 'true') { $computerEnabled = $true }
        foreach ($ext in @($gpo.Computer.ExtensionData | Where-Object { $_ })) {
            $computerExtCount++
            if ($ext.Name) { $extensionNames.Add([string]$ext.Name) }
        }
    }
    if ($gpo.User) {
        if ($gpo.User.Enabled -eq 'true') { $userEnabled = $true }
        foreach ($ext in @($gpo.User.ExtensionData | Where-Object { $_ })) {
            $userExtCount++
            if ($ext.Name) { $extensionNames.Add([string]$ext.Name) }
        }
    }

    return [pscustomobject]@{
        ComputerSettingsConfigured = ($computerEnabled -and $computerExtCount -gt 0)
        UserSettingsConfigured     = ($userEnabled -and $userExtCount -gt 0)
        ExtensionNames             = @($extensionNames | Select-Object -Unique)
    }
}

function New-GpoRow {
    param(
        [Parameter(Mandatory = $true)] $Gpo,
        [Parameter(Mandatory = $true)] $ReportInfo,
        [int]$LinkCount = 0
    )

    $extensionNames = @($ReportInfo.ExtensionNames)
    return [pscustomobject]@{
        DisplayName                = [string]$Gpo.DisplayName
        Id                         = [string]$Gpo.Id
        GpoStatus                  = [string]$Gpo.GpoStatus
        WmiFilter                  = [string]$Gpo.WmiFilter
        CreationTime               = $Gpo.CreationTime
        ModificationTime           = $Gpo.ModificationTime
        ComputerSettingsConfigured = [bool]$ReportInfo.ComputerSettingsConfigured
        UserSettingsConfigured     = [bool]$ReportInfo.UserSettingsConfigured
        LinkCount                  = $LinkCount
        IsUnlinked                 = ($LinkCount -eq 0)
        ConfiguredExtensions       = ($extensionNames -join ';')
        IntuneCandidate            = (Test-GpoIntuneCandidate -ExtensionNames $extensionNames)
    }
}

function New-OuRow {
    param(
        [Parameter(Mandatory = $true)] $Ou,
        [Parameter(Mandatory = $true)] [AllowNull()] $LinkedGpoNames
    )

    $linked = @($LinkedGpoNames | Where-Object { $_ })
    return [pscustomobject]@{
        DistinguishedName    = [string]$Ou.DistinguishedName
        Name                 = [string]$Ou.Name
        Depth                = (Get-OuDepth -DistinguishedName ([string]$Ou.DistinguishedName))
        BlockInheritance     = [bool]$Ou.BlockInheritance
        LinkedGpos           = ($linked -join ';')
        LinkedGpoCount       = $linked.Count
    }
}

function Get-AdOuGpoFacts {
    param(
        [string]$Server = '',
        [pscredential]$Credential
    )

    Import-Module ActiveDirectory -ErrorAction Stop
    Import-Module GroupPolicy -ErrorAction Stop
    $adParams = @{}
    if (-not [string]::IsNullOrWhiteSpace($Server)) { $adParams.Server = $Server }
    if ($Credential) { $adParams.Credential = $Credential }

    $domain = (Get-ADDomain @adParams).DNSRoot
    $gpoParams = @{ All = $true }
    if (-not [string]::IsNullOrWhiteSpace($Server)) { $gpoParams.Server = $Server }
    else { $gpoParams.Domain = $domain }

    $gpos = @()
    foreach ($gpo in @(Get-GPO @gpoParams)) {
        $reportInfo = [pscustomobject]@{ ComputerSettingsConfigured = $false; UserSettingsConfigured = $false; ExtensionNames = @() }
        try {
            $reportParams = @{ Guid = $gpo.Id; ReportType = 'Xml' }
            if ($gpoParams.ContainsKey('Server')) { $reportParams.Server = $gpoParams.Server }
            if ($gpoParams.ContainsKey('Domain')) { $reportParams.Domain = $gpoParams.Domain }
            [xml]$reportXml = Get-GPOReport @reportParams
            $reportInfo = ConvertFrom-GpoReportXml -ReportXml $reportXml
        } catch { }
        $gpos += [pscustomobject]@{ Gpo = $gpo; ReportInfo = $reportInfo }
    }

    # OUs and their linked GPOs (gPLink attribute references GPO GUIDs).
    $ous = @(Get-ADOrganizationalUnit -Filter * -Properties gPLink @adParams)
    $gpoNameById = @{}
    foreach ($item in $gpos) { $gpoNameById[("{$($item.Gpo.Id)}").ToLowerInvariant()] = [string]$item.Gpo.DisplayName }

    return @{
        Gpos          = $gpos
        Ous           = $ous
        GpoNameById   = $gpoNameById
    }
}

function Get-LinkedGpoGuids {
    param([AllowNull()] [string]$GpLink)

    $guids = [System.Collections.Generic.List[string]]::new()
    if ([string]::IsNullOrWhiteSpace($GpLink)) { return $guids.ToArray() }
    foreach ($match in [regex]::Matches($GpLink, '\{[0-9A-Fa-f\-]+\}')) {
        $guids.Add($match.Value.ToLowerInvariant())
    }
    return $guids.ToArray()
}

function Get-AdOuGpoInventory {
    param([Parameter(Mandatory = $true)] $Facts)

    # Count links per GPO from OU gPLinks.
    $linkCountByGuid = @{}
    foreach ($ou in @($Facts.Ous)) {
        foreach ($guid in (Get-LinkedGpoGuids -GpLink ([string]$ou.gPLink))) {
            if (-not $linkCountByGuid.ContainsKey($guid)) { $linkCountByGuid[$guid] = 0 }
            $linkCountByGuid[$guid]++
        }
    }

    $gpoRows = @(@($Facts.Gpos) | ForEach-Object {
        $guidKey = ("{$($_.Gpo.Id)}").ToLowerInvariant()
        $linkCount = if ($linkCountByGuid.ContainsKey($guidKey)) { $linkCountByGuid[$guidKey] } else { 0 }
        New-GpoRow -Gpo $_.Gpo -ReportInfo $_.ReportInfo -LinkCount $linkCount
    })

    $gpoNameById = $Facts.GpoNameById
    if (-not $gpoNameById) { $gpoNameById = @{} }
    $ouRows = @(@($Facts.Ous) | ForEach-Object {
        $linkedNames = @(Get-LinkedGpoGuids -GpLink ([string]$_.gPLink) | ForEach-Object { if ($gpoNameById.ContainsKey($_)) { $gpoNameById[$_] } else { $_ } })
        New-OuRow -Ou $_ -LinkedGpoNames $linkedNames
    })

    return @{ OrganizationalUnits = $ouRows; Gpos = $gpoRows }
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
Write-Host 'Collecting OU structure and GPO inventory...' -ForegroundColor Cyan
$facts = Get-AdOuGpoFacts -Server $Server -Credential $Credential
$inventory = Get-AdOuGpoInventory -Facts $facts

Write-Host "`n=== GPOs ($(@($inventory.Gpos).Count)) ===" -ForegroundColor Yellow
$inventory.Gpos | Format-Table DisplayName, GpoStatus, LinkCount, IsUnlinked, IntuneCandidate -AutoSize
Write-Host "=== Organizational Units ($(@($inventory.OrganizationalUnits).Count)) ===" -ForegroundColor Yellow
$inventory.OrganizationalUnits | Format-Table Name, Depth, BlockInheritance, LinkedGpoCount -AutoSize

$unlinked = @($inventory.Gpos | Where-Object IsUnlinked)
$intuneCandidates = @($inventory.Gpos | Where-Object IntuneCandidate)
Write-Host ("Summary: {0} GPO(s), {1} unlinked, {2} Intune-candidate; {3} OU(s)" -f @($inventory.Gpos).Count, $unlinked.Count, $intuneCandidates.Count, @($inventory.OrganizationalUnits).Count) -ForegroundColor Yellow

if (-not [string]::IsNullOrWhiteSpace($OutputFolder)) {
    if (-not (Test-Path -Path $OutputFolder)) { New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null }
    $inventory.OrganizationalUnits | Export-Csv -Path (Join-Path $OutputFolder 'OrganizationalUnits.csv') -NoTypeInformation -Encoding UTF8
    $inventory.Gpos | Export-Csv -Path (Join-Path $OutputFolder 'Gpos.csv') -NoTypeInformation -Encoding UTF8
    Write-Host ("Reports exported to: {0}" -f $OutputFolder) -ForegroundColor Green
}
