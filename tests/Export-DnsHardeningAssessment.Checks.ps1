#Requires -Version 7
$ErrorActionPreference = 'Stop'

$script:ScriptPath = Join-Path $PSScriptRoot '..' 'Export-DnsHardeningAssessment.ps1'
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
    foreach ($name in @('New-ControlRow','Get-DnsServerFacts','Test-DnsServerControls','Get-TargetDnsServers')) {
        Assert-True -Condition ($functionNames -contains $name) -Name "Function exists: $name"
    }
}
Test-RequiredFunctionsExist

function Test-StaticScriptContent {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw
    Assert-True -Condition ($scriptText -match 'Get-DnsServerZone') -Name 'Zone enumeration is present'
    Assert-True -Condition ($scriptText -match 'Get-DnsServerScavenging') -Name 'Scavenging is checked'
    Assert-True -Condition ($scriptText -match 'Get-DnsServerGlobalQueryBlockList') -Name 'Global query block list is checked'
    Assert-True -Condition ($scriptText -match 'SocketPoolSize') -Name 'Socket pool size is checked'
    Assert-True -Condition ($scriptText -match 'LockingPercent') -Name 'Cache locking is checked'
    foreach ($controlId in 1..10 | ForEach-Object { 'DNS-{0:d3}' -f $_ }) {
        Assert-True -Condition ($scriptText -match [regex]::Escape($controlId)) -Name "Control $controlId is present"
    }
    Assert-True -Condition ($scriptText -notmatch 'Set-DnsServer' -and $scriptText -notmatch 'Remove-DnsServer' -and $scriptText -notmatch 'Add-DnsServer') -Name 'Script performs no DNS configuration changes'
}
Test-StaticScriptContent

function New-HardenedDnsFacts {
    return @{
        Zones = @(
            @{ Name = 'contoso.com'; ZoneType = 'Primary'; IsDsIntegrated = $true; IsAutoCreated = $false; IsReverseLookupZone = $false; IsSigned = $true; DynamicUpdate = 'Secure'; SecureSecondaries = 'NoTransfer'; MasterServers = @(); AgingEnabled = $true },
            @{ Name = '_msdcs.contoso.com'; ZoneType = 'Primary'; IsDsIntegrated = $true; IsAutoCreated = $false; IsReverseLookupZone = $false; IsSigned = $false; DynamicUpdate = 'Secure'; SecureSecondaries = 'NoTransfer'; MasterServers = @(); AgingEnabled = $true },
            @{ Name = 'TrustAnchors'; ZoneType = 'Primary'; IsDsIntegrated = $true; IsAutoCreated = $false; IsReverseLookupZone = $false; IsSigned = $false; DynamicUpdate = 'None'; SecureSecondaries = 'NoTransfer'; MasterServers = @(); AgingEnabled = $null },
            @{ Name = 'partner.example'; ZoneType = 'Forwarder'; IsDsIntegrated = $true; IsAutoCreated = $false; IsReverseLookupZone = $false; IsSigned = $false; DynamicUpdate = 'None'; SecureSecondaries = 'NoTransfer'; MasterServers = @('10.9.9.9'); AgingEnabled = $null }
        )
        ScavengingEnabled = $true
        ScavengingIntervalDays = 7
        Forwarders = @('10.1.1.10','10.1.1.11')
        UseRootHint = $false
        RecursionEnabled = $true
        RootHintCount = 13
        QueryLoggingEnabled = $true
        EventLogLevel = 4
        GlobalQueryBlockListEnabled = $true
        GlobalQueryBlockList = @('wpad','isatap')
        SocketPoolSize = 2500
        CacheLockingPercent = 100
    }
}

function Test-HardenedDnsServer {
    $rows = @(Test-DnsServerControls -ServerName 'DC01' -Facts (New-HardenedDnsFacts) -MinimumSocketPoolSize 2500)
    $byId = @{}
    foreach ($row in $rows) { $byId[$row.ControlId] = $row }

    Assert-True -Condition ($rows.Count -eq 10) -Name 'Ten control rows are produced per server'
    foreach ($controlId in @('DNS-001','DNS-002','DNS-003','DNS-006','DNS-007','DNS-008','DNS-009')) {
        Assert-True -Condition ($byId[$controlId].Status -eq 'Pass') -Name "Hardened DNS server passes $controlId"
    }
    Assert-True -Condition ($byId['DNS-004'].Status -eq 'Manual' -and $byId['DNS-004'].Evidence -match '10\.1\.1\.10' -and $byId['DNS-004'].Evidence -match 'partner\.example') -Name 'DNS-004 lists forwarders and conditional forwarders for review'
    Assert-True -Condition ($byId['DNS-005'].Status -eq 'Manual') -Name 'Recursion enabled yields Manual DNS-005'
    Assert-True -Condition ($byId['DNS-010'].Status -eq 'Manual') -Name 'DNS-010 zone ACL review is Manual'
    Assert-True -Condition ($byId['DNS-006'].Evidence -match 'contoso\.com') -Name 'Signed zone appears in DNS-006 evidence'
}
Test-HardenedDnsServer

function Test-WeakDnsServer {
    $facts = New-HardenedDnsFacts
    $facts.Zones = @(
        @{ Name = 'contoso.com'; ZoneType = 'Primary'; IsDsIntegrated = $false; IsAutoCreated = $false; IsReverseLookupZone = $false; IsSigned = $false; DynamicUpdate = 'NonsecureAndSecure'; SecureSecondaries = 'TransferAnyServer'; MasterServers = @(); AgingEnabled = $false }
    )
    $facts.ScavengingEnabled = $false
    $facts.RecursionEnabled = $false
    $facts.QueryLoggingEnabled = $false
    $facts.GlobalQueryBlockListEnabled = $false
    $facts.GlobalQueryBlockList = @()
    $facts.SocketPoolSize = 1000
    $facts.CacheLockingPercent = 50

    $rows = @(Test-DnsServerControls -ServerName 'DC02' -Facts $facts -MinimumSocketPoolSize 2500)
    $byId = @{}
    foreach ($row in $rows) { $byId[$row.ControlId] = $row }

    Assert-True -Condition ($byId['DNS-001'].Status -eq 'Fail' -and $byId['DNS-001'].Evidence -match 'NonsecureAndSecure') -Name 'Insecure dynamic updates fail DNS-001 with evidence'
    Assert-True -Condition ($byId['DNS-002'].Status -eq 'Fail' -and $byId['DNS-002'].Evidence -match 'contoso\.com') -Name 'Open zone transfers fail DNS-002'
    Assert-True -Condition ($byId['DNS-003'].Status -eq 'Fail') -Name 'Disabled scavenging fails DNS-003'
    Assert-True -Condition ($byId['DNS-005'].Status -eq 'Pass') -Name 'Disabled recursion passes DNS-005'
    Assert-True -Condition ($byId['DNS-007'].Status -eq 'Warning') -Name 'Disabled query logging warns DNS-007'
    Assert-True -Condition ($byId['DNS-008'].Status -eq 'Fail') -Name 'Disabled block list fails DNS-008'
    Assert-True -Condition ($byId['DNS-009'].Status -eq 'Fail' -and $byId['DNS-009'].Evidence -match 'SocketPoolSize=1000') -Name 'Weak socket pool/cache locking fails DNS-009'
}
Test-WeakDnsServer

function Test-AgingPartialWarns {
    $facts = New-HardenedDnsFacts
    $facts.Zones[1].AgingEnabled = $false
    $rows = @(Test-DnsServerControls -ServerName 'DC03' -Facts $facts -MinimumSocketPoolSize 2500)
    $agingRow = $rows | Where-Object ControlId -eq 'DNS-003'
    Assert-True -Condition ($agingRow.Status -eq 'Warning' -and $agingRow.Evidence -match '_msdcs') -Name 'Aging disabled on one zone warns DNS-003 with zone name'
}
Test-AgingPartialWarns

function Test-ExplicitServerListUsed {
    $targets = @(Get-TargetDnsServers -DnsServer @('DC01','DC02') -Credential $null)
    Assert-True -Condition ($targets.Count -eq 2 -and $targets[0] -eq 'DC01') -Name 'Explicit -DnsServer list bypasses AD discovery'
}
Test-ExplicitServerListUsed

if ($script:Failures.Count -gt 0) {
    Write-Host "$($script:Failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host "All $($script:Passes) checks passed." -ForegroundColor Green
exit 0
