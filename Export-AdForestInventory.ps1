<#
.SYNOPSIS
    Exports an Active Directory forest and Domain Controller inventory for divestiture discovery.

.DESCRIPTION
    Read-only discovery of the forest topology that a tenant-to-tenant / new-forest divestiture
    needs as a design baseline:
      - Forest: name, forest/domain functional levels, schema version, UPN suffixes, domain list,
        and the forest-wide FSMO role holders (Schema Master, Domain Naming Master).
      - Domains: NetBIOS name, DNS root, domain SID, and the domain FSMO role holders
        (PDC Emulator, RID Master, Infrastructure Master).
      - Domain Controllers: hostname, site, OS/version, IsGlobalCatalog, IsReadOnly, and the
        FSMO roles each holds.

    Live data is gathered with the ActiveDirectory module; the shaping of each object into rows
    is separated so it can be unit-tested with mock data. The script makes no changes.

    Schema version numbers map to a Windows Server release (e.g. 88 = 2019, 87 = 2016); the mapping
    is reported alongside the raw objectVersion so the forest schema baseline is unambiguous.

.PARAMETER Server
    Optional Domain Controller / AD Web Services server to target for the queries.

.PARAMETER Credential
    Optional credential for the AD queries.

.PARAMETER OutputFolder
    Optional folder for the CSV outputs (Forest.csv, Domains.csv, DomainControllers.csv).
    If omitted, results are only written to the pipeline/table.

.EXAMPLE
    .\Export-AdForestInventory.ps1 -OutputFolder .\ForestInventory

.EXAMPLE
    .\Export-AdForestInventory.ps1 -Server dc01.contoso.com
#>

[CmdletBinding()]
param(
    [string]$Server = '',
    [pscredential]$Credential,
    [string]$OutputFolder = ''
)

$ErrorActionPreference = 'Stop'

function ConvertTo-SchemaVersionName {
    param([AllowNull()] $ObjectVersion)

    if ($null -eq $ObjectVersion) { return 'Unknown' }
    $map = @{
        13 = 'Windows 2000'
        30 = 'Windows Server 2003'
        31 = 'Windows Server 2003 R2'
        44 = 'Windows Server 2008'
        47 = 'Windows Server 2008 R2'
        56 = 'Windows Server 2012'
        69 = 'Windows Server 2012 R2'
        87 = 'Windows Server 2016'
        88 = 'Windows Server 2019/2022'
        91 = 'Windows Server 2025'
    }
    $key = [int]$ObjectVersion
    if ($map.ContainsKey($key)) { return $map[$key] }
    return ("objectVersion {0}" -f $key)
}

function New-ForestRow {
    param(
        [Parameter(Mandatory = $true)] $Forest,
        [Parameter(Mandatory = $true)] [AllowNull()] $SchemaObjectVersion
    )

    return [pscustomobject]@{
        ForestName             = [string]$Forest.Name
        ForestMode             = [string]$Forest.ForestMode
        RootDomain             = [string]$Forest.RootDomain
        Domains                = (@($Forest.Domains) -join ';')
        GlobalCatalogs         = (@($Forest.GlobalCatalogs).Count)
        Sites                  = (@($Forest.Sites).Count)
        UpnSuffixes            = (@($Forest.UPNSuffixes) -join ';')
        SchemaObjectVersion    = $SchemaObjectVersion
        SchemaVersionName      = (ConvertTo-SchemaVersionName -ObjectVersion $SchemaObjectVersion)
        SchemaMaster           = [string]$Forest.SchemaMaster
        DomainNamingMaster     = [string]$Forest.DomainNamingMaster
    }
}

function New-DomainRow {
    param([Parameter(Mandatory = $true)] $Domain)

    return [pscustomobject]@{
        DomainName            = [string]$Domain.DNSRoot
        NetBIOSName           = [string]$Domain.NetBIOSName
        DomainMode            = [string]$Domain.DomainMode
        DomainSID             = [string]$Domain.DomainSID
        PDCEmulator           = [string]$Domain.PDCEmulator
        RIDMaster             = [string]$Domain.RIDMaster
        InfrastructureMaster  = [string]$Domain.InfrastructureMaster
    }
}

function New-DomainControllerRow {
    param([Parameter(Mandatory = $true)] $DomainController)

    return [pscustomobject]@{
        HostName          = [string]$DomainController.HostName
        Domain            = [string]$DomainController.Domain
        Site              = [string]$DomainController.Site
        OperatingSystem   = [string]$DomainController.OperatingSystem
        OSVersion         = [string]$DomainController.OperatingSystemVersion
        IPv4Address       = [string]$DomainController.IPv4Address
        IsGlobalCatalog   = [bool]$DomainController.IsGlobalCatalog
        IsReadOnly        = [bool]$DomainController.IsReadOnly
        Enabled           = [bool]$DomainController.Enabled
        FsmoRoles         = (@($DomainController.OperationMasterRoles | ForEach-Object { [string]$_ }) -join ';')
    }
}

function Get-AdForestFacts {
    param(
        [string]$Server = '',
        [pscredential]$Credential
    )

    Import-Module ActiveDirectory -ErrorAction Stop
    $adParams = @{}
    if (-not [string]::IsNullOrWhiteSpace($Server)) { $adParams.Server = $Server }
    if ($Credential) { $adParams.Credential = $Credential }

    $forest = Get-ADForest @adParams

    $schemaObjectVersion = $null
    try {
        $rootDse = Get-ADRootDSE @adParams
        $schemaObject = Get-ADObject -Identity $rootDse.schemaNamingContext -Properties objectVersion @adParams
        $schemaObjectVersion = $schemaObject.objectVersion
    } catch { }

    $domains = @()
    foreach ($domainName in @($forest.Domains)) {
        try {
            $domainParams = $adParams.Clone()
            $domainParams['Identity'] = $domainName
            $domains += Get-ADDomain @domainParams
        } catch { }
    }

    $domainControllers = @()
    foreach ($domainName in @($forest.Domains)) {
        try {
            $dcParams = $adParams.Clone()
            $dcParams['Filter'] = '*'
            $dcParams['Server'] = $domainName
            $domainControllers += Get-ADDomainController @dcParams
        } catch { }
    }

    return @{
        Forest              = $forest
        SchemaObjectVersion = $schemaObjectVersion
        Domains             = $domains
        DomainControllers   = $domainControllers
    }
}

function Get-AdForestInventory {
    param([Parameter(Mandatory = $true)] $Facts)

    return @{
        Forest            = New-ForestRow -Forest $Facts.Forest -SchemaObjectVersion $Facts.SchemaObjectVersion
        Domains           = @(@($Facts.Domains) | ForEach-Object { New-DomainRow -Domain $_ })
        DomainControllers = @(@($Facts.DomainControllers) | ForEach-Object { New-DomainControllerRow -DomainController $_ })
    }
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
Write-Host 'Collecting AD forest, domain, and Domain Controller inventory...' -ForegroundColor Cyan
$facts = Get-AdForestFacts -Server $Server -Credential $Credential
$inventory = Get-AdForestInventory -Facts $facts

Write-Host "`n=== Forest ===" -ForegroundColor Yellow
$inventory.Forest | Format-List
Write-Host "=== Domains ($(@($inventory.Domains).Count)) ===" -ForegroundColor Yellow
$inventory.Domains | Format-Table -AutoSize
Write-Host "=== Domain Controllers ($(@($inventory.DomainControllers).Count)) ===" -ForegroundColor Yellow
$inventory.DomainControllers | Format-Table HostName, Site, OperatingSystem, IsGlobalCatalog, FsmoRoles -AutoSize

if (-not [string]::IsNullOrWhiteSpace($OutputFolder)) {
    if (-not (Test-Path -Path $OutputFolder)) { New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null }
    @($inventory.Forest) | Export-Csv -Path (Join-Path $OutputFolder 'Forest.csv') -NoTypeInformation -Encoding UTF8
    $inventory.Domains | Export-Csv -Path (Join-Path $OutputFolder 'Domains.csv') -NoTypeInformation -Encoding UTF8
    $inventory.DomainControllers | Export-Csv -Path (Join-Path $OutputFolder 'DomainControllers.csv') -NoTypeInformation -Encoding UTF8
    Write-Host ("Reports exported to: {0}" -f $OutputFolder) -ForegroundColor Green
}
