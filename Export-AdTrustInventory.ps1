<#
.SYNOPSIS
    Exports Active Directory domain and forest trust details for divestiture discovery.

.DESCRIPTION
    Read-only discovery of every trust in the forest, capturing the attributes that the SoW
    flags as required to design SIDHistory-preserving coexistence:
      - Source and target (partner) domain.
      - Direction (Inbound / Outbound / Bidirectional).
      - Trust type (forest / external / realm / parent-child / tree-root) and transitivity.
      - SID-filtering posture (quarantine / SIDHistory allowed) and selective authentication.

    Trust attributes are decoded from the msDS-TrustForestTrustInfo / trustAttributes bit flags on
    each trustedDomain object; the raw trustDirection / trustType / trustAttributes values are read
    with the ActiveDirectory module and shaped by a separate, unit-testable function. The script
    makes no changes.

.PARAMETER Server
    Optional Domain Controller / AD Web Services server to target for the queries.

.PARAMETER Credential
    Optional credential for the AD queries.

.PARAMETER OutputPath
    Optional CSV output path. If omitted, results are only written to the pipeline/table.

.EXAMPLE
    .\Export-AdTrustInventory.ps1 -OutputPath .\Trusts.csv
#>

[CmdletBinding()]
param(
    [string]$Server = '',
    [pscredential]$Credential,
    [string]$OutputPath = ''
)

$ErrorActionPreference = 'Stop'

function ConvertTo-TrustDirectionName {
    param([AllowNull()] $Direction)

    switch ([int]$Direction) {
        0 { 'Disabled' }
        1 { 'Inbound' }
        2 { 'Outbound' }
        3 { 'Bidirectional' }
        default { "Unknown ($Direction)" }
    }
}

function ConvertTo-TrustTypeName {
    param([AllowNull()] $TrustType)

    switch ([int]$TrustType) {
        1 { 'Downlevel (Windows NT)' }
        2 { 'Uplevel (AD Domain)' }
        3 { 'MIT (Kerberos realm)' }
        4 { 'DCE' }
        default { "Unknown ($TrustType)" }
    }
}

function ConvertTo-TrustAttributeInfo {
    param([AllowNull()] $TrustAttributes)

    $value = [int]$TrustAttributes
    $flags = [System.Collections.Generic.List[string]]::new()
    # trustAttributes bit flags (MS-ADTS).
    if ($value -band 0x1)  { $flags.Add('NonTransitive') }
    if ($value -band 0x2)  { $flags.Add('UplevelOnly') }
    if ($value -band 0x4)  { $flags.Add('QuarantinedDomain (SID filtering)') }
    if ($value -band 0x8)  { $flags.Add('ForestTransitive') }
    if ($value -band 0x10) { $flags.Add('CrossOrganization (selective auth)') }
    if ($value -band 0x20) { $flags.Add('WithinForest') }
    if ($value -band 0x40) { $flags.Add('TreatAsExternal') }
    if ($value -band 0x200) { $flags.Add('EnableTgtDelegation') }

    return [pscustomobject]@{
        Flags               = $flags.ToArray()
        IsForestTransitive  = (($value -band 0x8) -ne 0)
        IsNonTransitive     = (($value -band 0x1) -ne 0)
        SidFilteringEnabled = (($value -band 0x4) -ne 0)
        SelectiveAuth       = (($value -band 0x10) -ne 0)
        WithinForest        = (($value -band 0x20) -ne 0)
    }
}

function New-TrustRow {
    param([Parameter(Mandatory = $true)] $Trust)

    $direction = ConvertTo-TrustDirectionName -Direction $Trust.Direction
    $typeName = ConvertTo-TrustTypeName -TrustType $Trust.TrustType
    $attrInfo = ConvertTo-TrustAttributeInfo -TrustAttributes $Trust.TrustAttributes

    $trustCategory =
        if ($attrInfo.WithinForest) { 'Intra-forest' }
        elseif ($attrInfo.IsForestTransitive) { 'Forest' }
        else { 'External' }

    return [pscustomobject]@{
        SourceDomain        = [string]$Trust.SourceDomain
        PartnerDomain       = [string]$Trust.Target
        Direction           = $direction
        TrustType           = $typeName
        TrustCategory       = $trustCategory
        ForestTransitive    = $attrInfo.IsForestTransitive
        NonTransitive       = $attrInfo.IsNonTransitive
        SidFilteringEnabled = $attrInfo.SidFilteringEnabled
        SelectiveAuth       = $attrInfo.SelectiveAuth
        TrustAttributes     = [int]$Trust.TrustAttributes
        AttributeFlags      = (@($attrInfo.Flags) -join ';')
    }
}

function Get-AdTrustFacts {
    param(
        [string]$Server = '',
        [pscredential]$Credential
    )

    Import-Module ActiveDirectory -ErrorAction Stop
    $adParams = @{}
    if (-not [string]::IsNullOrWhiteSpace($Server)) { $adParams.Server = $Server }
    if ($Credential) { $adParams.Credential = $Credential }

    $forest = Get-ADForest @adParams
    $trusts = @()
    foreach ($domainName in @($forest.Domains)) {
        try {
            $trustParams = $adParams.Clone()
            $trustParams['Filter'] = '*'
            $trustParams['Server'] = $domainName
            $trustParams['Properties'] = @('Direction','TrustType','TrustAttributes','Target','Source')
            foreach ($trust in @(Get-ADTrust @trustParams)) {
                $trusts += [pscustomobject]@{
                    SourceDomain    = $domainName
                    Target          = $trust.Target
                    Direction       = $trust.Direction
                    TrustType       = $trust.TrustType
                    TrustAttributes = $trust.TrustAttributes
                }
            }
        } catch { }
    }

    return @{ Trusts = $trusts }
}

function Get-AdTrustInventory {
    param([Parameter(Mandatory = $true)] $Facts)

    return @(@($Facts.Trusts) | ForEach-Object { New-TrustRow -Trust $_ })
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
Write-Host 'Collecting AD trust inventory...' -ForegroundColor Cyan
$facts = Get-AdTrustFacts -Server $Server -Credential $Credential
$rows = @(Get-AdTrustInventory -Facts $facts)

if ($rows.Count -eq 0) {
    Write-Host 'No trusts found in the forest.' -ForegroundColor Yellow
}
else {
    $rows | Format-Table SourceDomain, PartnerDomain, Direction, TrustCategory, SidFilteringEnabled, SelectiveAuth -AutoSize
}

if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
    $rows | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "Report exported: $OutputPath" -ForegroundColor Green
}
