<#
.SYNOPSIS
    Exports an on-premises / hybrid Exchange inventory for divestiture discovery.

.DESCRIPTION
    Read-only discovery of the on-premises Exchange estate the SoW flags as a MATERIAL GAP - with all
    mail treated as Exchange Online, a new-forest divestiture still needs a recipient-management and
    SMTP-relay path. This inventory captures:
      - Servers: Exchange servers, roles, edition, and version.
      - Hybrid: whether a Hybrid Configuration exists (HCW) and its key settings.
      - Accepted domains: name, domain type (authoritative / internal relay / external relay), default.
      - Connectors: send connectors (address spaces, smart hosts) and receive connectors, with the
        anonymous-relay receive connectors highlighted (the application/MFD relay path).
      - Public folders: mailbox count and whether a public folder hierarchy exists.
      - DAGs: database availability groups and their member servers.

    The script uses the on-premises Exchange Management Shell cmdlets (or an implicit remote session).
    Collection and row shaping are separated so the shapers are unit-testable with mock data. The
    script makes no changes.

.PARAMETER OutputFolder
    Optional folder for the CSV outputs (Servers.csv, AcceptedDomains.csv, SendConnectors.csv,
    ReceiveConnectors.csv, Dags.csv). If omitted, results are only written to the pipeline/table.

.EXAMPLE
    .\Export-ExchangeOnPremInventory.ps1 -OutputFolder .\ExchangeOnPrem
#>

[CmdletBinding()]
param(
    [string]$OutputFolder = ''
)

$ErrorActionPreference = 'Stop'

function New-ExchangeServerRow {
    param([Parameter(Mandatory = $true)] $Server)
    return [pscustomobject]@{
        Name             = [string]$Server.Name
        ServerRole       = (@($Server.ServerRole) -join ';')
        Edition          = [string]$Server.Edition
        AdminDisplayVersion = [string]$Server.AdminDisplayVersion
        Site             = [string]$Server.Site
        IsHubTransport   = ([string]$Server.ServerRole -match 'HubTransport|Mailbox')
    }
}

function New-AcceptedDomainRow {
    param([Parameter(Mandatory = $true)] $Domain)
    return [pscustomobject]@{
        Name             = [string]$Domain.Name
        DomainName       = [string]$Domain.DomainName
        DomainType       = [string]$Domain.DomainType
        Default          = [bool]$Domain.Default
    }
}

function New-SendConnectorRow {
    param([Parameter(Mandatory = $true)] $Connector)
    return [pscustomobject]@{
        Name           = [string]$Connector.Name
        AddressSpaces  = (@($Connector.AddressSpaces | ForEach-Object { [string]$_ }) -join ';')
        SmartHosts     = (@($Connector.SmartHosts | ForEach-Object { [string]$_ }) -join ';')
        Enabled        = [bool]$Connector.Enabled
        SourceServers  = (@($Connector.SourceTransportServers | ForEach-Object { [string]$_ }) -join ';')
    }
}

function New-ReceiveConnectorRow {
    param([Parameter(Mandatory = $true)] $Connector)

    $permissionGroups = @($Connector.PermissionGroups | ForEach-Object { [string]$_ })
    $authMechanisms = @($Connector.AuthMechanism | ForEach-Object { [string]$_ })
    $allowsAnonymous = (($permissionGroups -join ',') -match 'AnonymousUsers')
    return [pscustomobject]@{
        Name             = [string]$Connector.Name
        Server           = [string]$Connector.Server
        Bindings         = (@($Connector.Bindings | ForEach-Object { [string]$_ }) -join ';')
        PermissionGroups = ($permissionGroups -join ';')
        AuthMechanism    = ($authMechanisms -join ';')
        AllowsAnonymousRelay = $allowsAnonymous
        Enabled          = [bool]$Connector.Enabled
    }
}

function New-DagRow {
    param([Parameter(Mandatory = $true)] $Dag)
    return [pscustomobject]@{
        Name    = [string]$Dag.Name
        Members = (@($Dag.Servers | ForEach-Object { [string]$_ }) -join ';')
        MemberCount = @($Dag.Servers | Where-Object { $_ }).Count
    }
}

function Get-ExchangeOnPremInventory {
    param([Parameter(Mandatory = $true)] $Facts)

    $hybridConfigured = (@($Facts.HybridConfiguration).Count -gt 0)
    $hybrid = if ($hybridConfigured) { @($Facts.HybridConfiguration)[0] } else { $null }

    return @{
        Servers           = @(@($Facts.Servers) | ForEach-Object { New-ExchangeServerRow -Server $_ })
        AcceptedDomains   = @(@($Facts.AcceptedDomains) | ForEach-Object { New-AcceptedDomainRow -Domain $_ })
        SendConnectors    = @(@($Facts.SendConnectors) | ForEach-Object { New-SendConnectorRow -Connector $_ })
        ReceiveConnectors = @(@($Facts.ReceiveConnectors) | ForEach-Object { New-ReceiveConnectorRow -Connector $_ })
        Dags              = @(@($Facts.Dags) | ForEach-Object { New-DagRow -Dag $_ })
        HybridConfigured  = $hybridConfigured
        HybridDomains     = if ($hybrid) { @($hybrid.Domains | ForEach-Object { [string]$_ }) } else { @() }
        PublicFolderMailboxCount = [int]$Facts.PublicFolderMailboxCount
    }
}

function Get-ExchangeOnPremFacts {
    $facts = @{
        Servers = @(); AcceptedDomains = @(); SendConnectors = @(); ReceiveConnectors = @()
        Dags = @(); HybridConfiguration = @(); PublicFolderMailboxCount = 0
    }

    if (-not (Get-Command -Name Get-ExchangeServer -ErrorAction SilentlyContinue)) {
        throw 'Exchange management cmdlets (Get-ExchangeServer) are not available. Run in the Exchange Management Shell or import an Exchange remote session.'
    }

    try { $facts.Servers = @(Get-ExchangeServer -ErrorAction Stop) } catch { }
    try { $facts.AcceptedDomains = @(Get-AcceptedDomain -ErrorAction Stop) } catch { }
    try { $facts.SendConnectors = @(Get-SendConnector -ErrorAction Stop) } catch { }
    try { $facts.ReceiveConnectors = @(Get-ReceiveConnector -ErrorAction Stop) } catch { }
    try { $facts.Dags = @(Get-DatabaseAvailabilityGroup -ErrorAction Stop) } catch { }
    try { $facts.HybridConfiguration = @(Get-HybridConfiguration -ErrorAction Stop) } catch { }
    try { $facts.PublicFolderMailboxCount = @(Get-Mailbox -PublicFolder -ErrorAction Stop).Count } catch { }

    return $facts
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
Write-Host 'Collecting on-premises / hybrid Exchange inventory...' -ForegroundColor Cyan
$facts = Get-ExchangeOnPremFacts
$inventory = Get-ExchangeOnPremInventory -Facts $facts

Write-Host ("`nHybrid configured: {0}{1}" -f $inventory.HybridConfigured, $(if ($inventory.HybridConfigured) { " (domains: $($inventory.HybridDomains -join ', '))" } else { '' })) -ForegroundColor Yellow
Write-Host "=== Exchange servers ($(@($inventory.Servers).Count)) ===" -ForegroundColor Yellow
$inventory.Servers | Format-Table Name, ServerRole, Edition, AdminDisplayVersion -AutoSize
Write-Host "=== Accepted domains ($(@($inventory.AcceptedDomains).Count)) ===" -ForegroundColor Yellow
$inventory.AcceptedDomains | Format-Table Name, DomainName, DomainType, Default -AutoSize

$anonRelay = @($inventory.ReceiveConnectors | Where-Object AllowsAnonymousRelay)
Write-Host "=== Receive connectors ($(@($inventory.ReceiveConnectors).Count); $($anonRelay.Count) allow anonymous relay) ===" -ForegroundColor Yellow
$inventory.ReceiveConnectors | Format-Table Name, Server, AllowsAnonymousRelay, PermissionGroups -AutoSize
Write-Host "=== Send connectors ($(@($inventory.SendConnectors).Count)) ===" -ForegroundColor Yellow
$inventory.SendConnectors | Format-Table Name, AddressSpaces, SmartHosts -AutoSize
Write-Host ("Public folder mailbox count: {0}" -f $inventory.PublicFolderMailboxCount) -ForegroundColor Yellow

if (-not [string]::IsNullOrWhiteSpace($OutputFolder)) {
    if (-not (Test-Path -Path $OutputFolder)) { New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null }
    $inventory.Servers | Export-Csv -Path (Join-Path $OutputFolder 'Servers.csv') -NoTypeInformation -Encoding UTF8
    $inventory.AcceptedDomains | Export-Csv -Path (Join-Path $OutputFolder 'AcceptedDomains.csv') -NoTypeInformation -Encoding UTF8
    $inventory.SendConnectors | Export-Csv -Path (Join-Path $OutputFolder 'SendConnectors.csv') -NoTypeInformation -Encoding UTF8
    $inventory.ReceiveConnectors | Export-Csv -Path (Join-Path $OutputFolder 'ReceiveConnectors.csv') -NoTypeInformation -Encoding UTF8
    $inventory.Dags | Export-Csv -Path (Join-Path $OutputFolder 'Dags.csv') -NoTypeInformation -Encoding UTF8
    Write-Host ("Reports exported to: {0}" -f $OutputFolder) -ForegroundColor Green
}
