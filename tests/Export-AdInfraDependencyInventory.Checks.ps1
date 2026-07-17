#Requires -Version 7
$ErrorActionPreference = 'Stop'

$script:ScriptPath = Join-Path $PSScriptRoot '..' 'Export-AdInfraDependencyInventory.ps1'
$script:Failures = [System.Collections.Generic.List[string]]::new()
$script:Passes = 0

function Get-ScriptAst {
    $tokens = $null
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:ScriptPath, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count -gt 0) {
        throw "Parse errors in $($script:ScriptPath): $($parseErrors[0].Message)"
    }
    return $ast
}

function Import-ScriptFunctions {
    $ast = Get-ScriptAst
    $functions = $ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)
    foreach ($functionAst in $functions) {
        . ([scriptblock]::Create($functionAst.Extent.Text))
    }
}

function Assert-True {
    param(
        [Parameter(Mandatory = $true)] [bool]$Condition,
        [Parameter(Mandatory = $true)] [string]$Name
    )
    if ($Condition) {
        $script:Passes++
        Write-Host "PASS $Name" -ForegroundColor Green
    }
    else {
        $script:Failures.Add($Name) | Out-Null
        Write-Host "FAIL $Name" -ForegroundColor Red
    }
}

. Import-ScriptFunctions

function Test-RequiredFunctionsExist {
    $ast = Get-ScriptAst
    $functionNames = @($ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true) | ForEach-Object Name)
    foreach ($name in @('New-DnsZoneRow','New-DhcpScopeRow','New-NpsClientRow','New-AdcsTemplateRow','Get-AdInfraDependencyInventory','Get-AdInfraDependencyFacts')) {
        Assert-True -Condition ($functionNames -contains $name) -Name "Function exists: $name"
    }
}
Test-RequiredFunctionsExist

function Test-ReadOnly {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw
    Assert-True -Condition ($scriptText -notmatch 'Set-DnsServer' -and $scriptText -notmatch 'Add-DhcpServer' -and $scriptText -notmatch 'Set-DhcpServer' -and $scriptText -notmatch 'New-NpsRadiusClient') -Name 'Script performs no infrastructure write operations'
    Assert-True -Condition ($scriptText -match 'Get-DnsServerZone' -and $scriptText -match 'Get-DhcpServerv4Scope') -Name 'DNS and DHCP queries are present'
}
Test-ReadOnly

function New-MockInfraFacts {
    return @{
        DnsServers = @(
            [pscustomobject]@{
                ServerName = 'dc01.contoso.com'
                Zones = @(
                    [pscustomobject]@{ ZoneName = 'contoso.com'; ZoneType = 'Primary'; IsDsIntegrated = $true; IsReverseLookupZone = $false; DynamicUpdate = 'Secure' },
                    [pscustomobject]@{ ZoneName = '1.10.in-addr.arpa'; ZoneType = 'Primary'; IsDsIntegrated = $true; IsReverseLookupZone = $true; DynamicUpdate = 'Secure' }
                )
                Forwarders = @('8.8.8.8','1.1.1.1')
            }
        )
        DhcpServers = @(
            [pscustomobject]@{
                ServerName = 'dhcp01.contoso.com'
                Scopes = @(
                    [pscustomobject]@{ ScopeId = '10.1.0.0'; Name = 'Perth'; StartRange = '10.1.0.10'; EndRange = '10.1.0.250'; SubnetMask = '255.255.255.0'; State = 'Active' }
                )
            }
        )
        NtpSources = @(
            [pscustomobject]@{ Server = 'dc01.contoso.com'; IsPdcEmulator = $true; TimeType = 'NTP'; NtpServer = 'time.windows.com,0x9' }
        )
        NpsServers = @(
            [pscustomobject]@{
                ServerName = 'nps01.contoso.com'
                Clients = @(
                    [pscustomobject]@{ Name = 'WLC'; Address = '10.0.5.5'; Enabled = $true; VendorName = 'RADIUS Standard' }
                )
            }
        )
        CertificateAuthorities = @(
            [pscustomobject]@{ Name = 'Contoso-Issuing-CA'; HostName = 'ca01.contoso.com'; Templates = @('User','WebServer','Machine') }
        )
    }
}

function Test-InventoryShaping {
    $inventory = Get-AdInfraDependencyInventory -Facts (New-MockInfraFacts)

    Assert-True -Condition (@($inventory.DnsZones).Count -eq 2) -Name 'DNS zone rows flattened per server'
    $reverse = $inventory.DnsZones | Where-Object ZoneName -eq '1.10.in-addr.arpa'
    Assert-True -Condition ($reverse.IsReverse -eq $true -and $reverse.DnsServer -eq 'dc01.contoso.com') -Name 'Reverse zone shaped with server name'
    Assert-True -Condition (@($inventory.DnsForwarders).Count -eq 2) -Name 'DNS forwarders flattened'

    Assert-True -Condition (@($inventory.DhcpScopes).Count -eq 1 -and $inventory.DhcpScopes[0].StartRange -eq '10.1.0.10') -Name 'DHCP scope shaped'
    Assert-True -Condition (@($inventory.NtpSources).Count -eq 1 -and $inventory.NtpSources[0].IsPdcEmulator -eq $true) -Name 'NTP source carried through'
    Assert-True -Condition (@($inventory.NpsClients).Count -eq 1 -and $inventory.NpsClients[0].Address -eq '10.0.5.5') -Name 'NPS RADIUS client flattened per server'
    Assert-True -Condition (@($inventory.AdcsTemplates).Count -eq 3 -and ($inventory.AdcsTemplates | Where-Object TemplateName -eq 'WebServer')) -Name 'AD CS templates flattened per CA'
}
Test-InventoryShaping

function Test-EmptyFactsDegradeGracefully {
    $inventory = Get-AdInfraDependencyInventory -Facts @{ DnsServers = @(); DhcpServers = @(); NtpSources = @(); NpsServers = @(); CertificateAuthorities = @() }
    Assert-True -Condition (@($inventory.DnsZones).Count -eq 0 -and @($inventory.DhcpScopes).Count -eq 0 -and @($inventory.AdcsTemplates).Count -eq 0) -Name 'Empty facts produce empty sections without error'
}
Test-EmptyFactsDegradeGracefully

if ($script:Failures.Count -gt 0) {
    Write-Host "$($script:Failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host "All $($script:Passes) checks passed." -ForegroundColor Green
exit 0
