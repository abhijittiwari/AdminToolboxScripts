<#
.SYNOPSIS
    Exports AD-adjacent infrastructure dependencies (DNS, DHCP, NTP, RADIUS/NPS, AD CS) for divestiture discovery.

.DESCRIPTION
    Read-only discovery of the service dependencies the SoW flags as design inputs that must be
    sequenced during a new-forest divestiture (notably PKI ahead of RADIUS/NPS):
      - DNS: zones per DNS server (type, AD-integrated, dynamic-update mode) and forwarders.
      - DHCP: authorized DHCP servers and their scopes (range, subnet mask, state).
      - NTP / time: the PDC emulator's configured time source and each DC's sync type.
      - RADIUS / NPS: NPS servers, their RADIUS clients, and connection-request/network policies.
      - AD CS: enterprise CAs and published certificate templates from the configuration partition.

    Each dependency is collected independently (a missing role/module degrades to an empty section
    rather than failing the run) and written to its own CSV. Collection and shaping are separated so
    the shapers are unit-testable with mock data. The script makes no changes.

.PARAMETER DnsServer
    DNS servers to inventory. Defaults to all Domain Controllers.

.PARAMETER DhcpServer
    DHCP servers to inventory. Defaults to the authorized DHCP servers registered in AD.

.PARAMETER NpsServer
    NPS/RADIUS servers to inventory (remote NPS query). Optional; when omitted, NPS is skipped
    unless the local host runs NPS.

.PARAMETER Include
    Which dependency sections to collect. Defaults to all: DNS, DHCP, NTP, NPS, ADCS.

.PARAMETER Credential
    Optional credential for the AD queries.

.PARAMETER OutputFolder
    Optional folder for the per-service CSV outputs. If omitted, results are only written to the table.

.EXAMPLE
    .\Export-AdInfraDependencyInventory.ps1 -OutputFolder .\InfraDependencies

.EXAMPLE
    .\Export-AdInfraDependencyInventory.ps1 -Include DNS,ADCS -OutputFolder .\InfraDependencies
#>

[CmdletBinding()]
param(
    [string[]]$DnsServer = @(),
    [string[]]$DhcpServer = @(),
    [string[]]$NpsServer = @(),
    [ValidateSet('DNS','DHCP','NTP','NPS','ADCS')]
    [string[]]$Include = @('DNS','DHCP','NTP','NPS','ADCS'),
    [pscredential]$Credential,
    [string]$OutputFolder = ''
)

$ErrorActionPreference = 'Stop'

function New-DnsZoneRow {
    param(
        [Parameter(Mandatory = $true)] [string]$ServerName,
        [Parameter(Mandatory = $true)] $Zone
    )
    return [pscustomobject]@{
        DnsServer      = $ServerName
        ZoneName       = [string]$Zone.ZoneName
        ZoneType       = [string]$Zone.ZoneType
        IsDsIntegrated = [bool]$Zone.IsDsIntegrated
        IsReverse      = [bool]$Zone.IsReverseLookupZone
        DynamicUpdate  = [string]$Zone.DynamicUpdate
    }
}

function New-DhcpScopeRow {
    param(
        [Parameter(Mandatory = $true)] [string]$ServerName,
        [Parameter(Mandatory = $true)] $Scope
    )
    return [pscustomobject]@{
        DhcpServer  = $ServerName
        ScopeId     = [string]$Scope.ScopeId
        Name        = [string]$Scope.Name
        StartRange  = [string]$Scope.StartRange
        EndRange    = [string]$Scope.EndRange
        SubnetMask  = [string]$Scope.SubnetMask
        State       = [string]$Scope.State
    }
}

function New-NpsClientRow {
    param(
        [Parameter(Mandatory = $true)] [string]$ServerName,
        [Parameter(Mandatory = $true)] $Client
    )
    return [pscustomobject]@{
        NpsServer   = $ServerName
        ClientName  = [string]$Client.Name
        Address     = [string]$Client.Address
        Enabled     = [bool]$Client.Enabled
        VendorName  = [string]$Client.VendorName
    }
}

function New-AdcsTemplateRow {
    param(
        [Parameter(Mandatory = $true)] [string]$CaName,
        [Parameter(Mandatory = $true)] [string]$TemplateName
    )
    return [pscustomobject]@{
        CertificateAuthority = $CaName
        TemplateName         = $TemplateName
    }
}

function Get-AdInfraDependencyInventory {
    param([Parameter(Mandatory = $true)] $Facts)

    $dnsZoneRows = [System.Collections.Generic.List[object]]::new()
    foreach ($server in @($Facts.DnsServers)) {
        foreach ($zone in @($server.Zones)) {
            $dnsZoneRows.Add((New-DnsZoneRow -ServerName $server.ServerName -Zone $zone))
        }
    }

    $dnsForwarderRows = [System.Collections.Generic.List[object]]::new()
    foreach ($server in @($Facts.DnsServers)) {
        foreach ($forwarder in @($server.Forwarders | Where-Object { $_ })) {
            $dnsForwarderRows.Add([pscustomobject]@{ DnsServer = $server.ServerName; Forwarder = [string]$forwarder })
        }
    }

    $dhcpScopeRows = [System.Collections.Generic.List[object]]::new()
    foreach ($server in @($Facts.DhcpServers)) {
        foreach ($scope in @($server.Scopes)) {
            $dhcpScopeRows.Add((New-DhcpScopeRow -ServerName $server.ServerName -Scope $scope))
        }
    }

    $npsClientRows = [System.Collections.Generic.List[object]]::new()
    foreach ($server in @($Facts.NpsServers)) {
        foreach ($client in @($server.Clients)) {
            $npsClientRows.Add((New-NpsClientRow -ServerName $server.ServerName -Client $client))
        }
    }

    $adcsTemplateRows = [System.Collections.Generic.List[object]]::new()
    foreach ($ca in @($Facts.CertificateAuthorities)) {
        foreach ($template in @($ca.Templates | Where-Object { $_ })) {
            $adcsTemplateRows.Add((New-AdcsTemplateRow -CaName $ca.Name -TemplateName ([string]$template)))
        }
    }

    return @{
        DnsZones       = $dnsZoneRows.ToArray()
        DnsForwarders  = $dnsForwarderRows.ToArray()
        DhcpScopes     = $dhcpScopeRows.ToArray()
        NtpSources     = @($Facts.NtpSources)
        NpsClients     = $npsClientRows.ToArray()
        AdcsTemplates  = $adcsTemplateRows.ToArray()
    }
}

function Get-AdInfraDependencyFacts {
    param(
        [string[]]$DnsServer = @(),
        [string[]]$DhcpServer = @(),
        [string[]]$NpsServer = @(),
        [string[]]$Include = @('DNS','DHCP','NTP','NPS','ADCS'),
        [pscredential]$Credential
    )

    Import-Module ActiveDirectory -ErrorAction Stop
    $adParams = @{}
    if ($Credential) { $adParams.Credential = $Credential }

    $facts = @{
        DnsServers = @(); DhcpServers = @(); NtpSources = @(); NpsServers = @(); CertificateAuthorities = @()
    }

    $domainObject = Get-ADDomain @adParams
    $dcHostNames = @(Get-ADDomainController -Filter * @adParams | ForEach-Object HostName)

    if ($Include -contains 'DNS') {
        $dnsTargets = if ($DnsServer.Count -gt 0) { $DnsServer } else { $dcHostNames }
        foreach ($target in $dnsTargets) {
            try {
                Import-Module DnsServer -ErrorAction Stop
                $zones = @(Get-DnsServerZone -ComputerName $target -ErrorAction Stop)
                $forwarders = @((Get-DnsServerForwarder -ComputerName $target -ErrorAction SilentlyContinue).IPAddress | ForEach-Object { [string]$_ })
                $facts.DnsServers += [pscustomobject]@{ ServerName = $target; Zones = $zones; Forwarders = $forwarders }
            } catch { }
        }
    }

    if ($Include -contains 'DHCP') {
        $dhcpTargets = $DhcpServer
        if ($dhcpTargets.Count -eq 0) {
            try {
                Import-Module DhcpServer -ErrorAction Stop
                $dhcpTargets = @(Get-DhcpServerInDC -ErrorAction Stop | ForEach-Object DnsName)
            } catch { $dhcpTargets = @() }
        }
        foreach ($target in $dhcpTargets) {
            try {
                Import-Module DhcpServer -ErrorAction Stop
                $scopes = @(Get-DhcpServerv4Scope -ComputerName $target -ErrorAction Stop)
                $facts.DhcpServers += [pscustomobject]@{ ServerName = $target; Scopes = $scopes }
            } catch { }
        }
    }

    if ($Include -contains 'NTP') {
        foreach ($dc in $dcHostNames) {
            try {
                $isPdc = ($dc -eq $domainObject.PDCEmulator)
                $config = Invoke-Command -ComputerName $dc -ScriptBlock {
                    $params = 'HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Parameters'
                    [pscustomobject]@{
                        Type      = (Get-ItemProperty -Path $params -Name Type -ErrorAction SilentlyContinue).Type
                        NtpServer = (Get-ItemProperty -Path $params -Name NtpServer -ErrorAction SilentlyContinue).NtpServer
                    }
                } -ErrorAction Stop
                $facts.NtpSources += [pscustomobject]@{ Server = $dc; IsPdcEmulator = $isPdc; TimeType = $config.Type; NtpServer = $config.NtpServer }
            } catch {
                $facts.NtpSources += [pscustomobject]@{ Server = $dc; IsPdcEmulator = ($dc -eq $domainObject.PDCEmulator); TimeType = 'Unknown'; NtpServer = '' }
            }
        }
    }

    if ($Include -contains 'NPS' -and $NpsServer.Count -gt 0) {
        foreach ($target in $NpsServer) {
            try {
                $clients = Invoke-Command -ComputerName $target -ScriptBlock {
                    Import-Module NPS -ErrorAction SilentlyContinue
                    @(Get-NpsRadiusClient -ErrorAction SilentlyContinue)
                } -ErrorAction Stop
                $facts.NpsServers += [pscustomobject]@{ ServerName = $target; Clients = @($clients) }
            } catch { }
        }
    }

    if ($Include -contains 'ADCS') {
        try {
            $configNc = (Get-ADRootDSE @adParams).configurationNamingContext
            $enrollmentServices = @(Get-ADObject -SearchBase "CN=Enrollment Services,CN=Public Key Services,CN=Services,$configNc" -Filter { objectClass -eq 'pKIEnrollmentService' } -Properties certificateTemplates, dNSHostName @adParams)
            foreach ($ca in $enrollmentServices) {
                $facts.CertificateAuthorities += [pscustomobject]@{ Name = [string]$ca.Name; HostName = [string]$ca.dNSHostName; Templates = @($ca.certificateTemplates) }
            }
        } catch { }
    }

    return $facts
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
Write-Host ("Collecting infrastructure dependencies: {0}" -f ($Include -join ', ')) -ForegroundColor Cyan
$facts = Get-AdInfraDependencyFacts -DnsServer $DnsServer -DhcpServer $DhcpServer -NpsServer $NpsServer -Include $Include -Credential $Credential
$inventory = Get-AdInfraDependencyInventory -Facts $facts

Write-Host "`n=== DNS zones ($(@($inventory.DnsZones).Count)) ===" -ForegroundColor Yellow
$inventory.DnsZones | Format-Table DnsServer, ZoneName, ZoneType, DynamicUpdate -AutoSize
Write-Host "=== DHCP scopes ($(@($inventory.DhcpScopes).Count)) ===" -ForegroundColor Yellow
$inventory.DhcpScopes | Format-Table DhcpServer, ScopeId, StartRange, EndRange, State -AutoSize
Write-Host "=== NTP sources ($(@($inventory.NtpSources).Count)) ===" -ForegroundColor Yellow
$inventory.NtpSources | Format-Table Server, IsPdcEmulator, TimeType, NtpServer -AutoSize
Write-Host "=== NPS RADIUS clients ($(@($inventory.NpsClients).Count)) ===" -ForegroundColor Yellow
$inventory.NpsClients | Format-Table NpsServer, ClientName, Address -AutoSize
Write-Host "=== AD CS templates ($(@($inventory.AdcsTemplates).Count)) ===" -ForegroundColor Yellow
$inventory.AdcsTemplates | Format-Table CertificateAuthority, TemplateName -AutoSize

if (-not [string]::IsNullOrWhiteSpace($OutputFolder)) {
    if (-not (Test-Path -Path $OutputFolder)) { New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null }
    $inventory.DnsZones | Export-Csv -Path (Join-Path $OutputFolder 'DnsZones.csv') -NoTypeInformation -Encoding UTF8
    $inventory.DnsForwarders | Export-Csv -Path (Join-Path $OutputFolder 'DnsForwarders.csv') -NoTypeInformation -Encoding UTF8
    $inventory.DhcpScopes | Export-Csv -Path (Join-Path $OutputFolder 'DhcpScopes.csv') -NoTypeInformation -Encoding UTF8
    $inventory.NtpSources | Export-Csv -Path (Join-Path $OutputFolder 'NtpSources.csv') -NoTypeInformation -Encoding UTF8
    $inventory.NpsClients | Export-Csv -Path (Join-Path $OutputFolder 'NpsRadiusClients.csv') -NoTypeInformation -Encoding UTF8
    $inventory.AdcsTemplates | Export-Csv -Path (Join-Path $OutputFolder 'AdcsTemplates.csv') -NoTypeInformation -Encoding UTF8
    Write-Host ("Reports exported to: {0}" -f $OutputFolder) -ForegroundColor Green
}
