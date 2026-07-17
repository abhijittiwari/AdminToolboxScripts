#Requires -Version 7
$ErrorActionPreference = 'Stop'

$script:ScriptPath = Join-Path $PSScriptRoot '..' 'Export-DcHardeningAssessment.ps1'
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
    foreach ($name in @('New-ControlRow','Get-DcFacts','Test-DcComputerControls','Test-DcDomainControls','Get-DcDomainFacts','Get-TargetDomainControllers')) {
        Assert-True -Condition ($functionNames -contains $name) -Name "Function exists: $name"
    }
}
Test-RequiredFunctionsExist

function Test-StaticScriptContent {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw
    Assert-True -Condition ($scriptText -match 'LDAPServerIntegrity') -Name 'LDAP signing registry value is checked'
    Assert-True -Condition ($scriptText -match 'LdapEnforceChannelBinding') -Name 'LDAP channel binding registry value is checked'
    Assert-True -Condition ($scriptText -match 'VulnerableChannelAllowList') -Name 'Netlogon vulnerable channel allow list is checked'
    Assert-True -Condition ($scriptText -match 'HardenedPaths') -Name 'Hardened UNC paths are checked'
    Assert-True -Condition ($scriptText -match 'dSASignature') -Name 'System-state backup marker (dSASignature) is checked'
    foreach ($controlId in 1..22 | ForEach-Object { 'DC-{0:d3}' -f $_ }) {
        Assert-True -Condition ($scriptText -match [regex]::Escape($controlId)) -Name "Control $controlId is present"
    }
    Assert-True -Condition ($scriptText -notmatch 'Set-ItemProperty' -and $scriptText -notmatch 'Set-ADUser' -and $scriptText -notmatch 'Set-Service') -Name 'Script performs no configuration changes'
}
Test-StaticScriptContent

function New-HardenedDcFacts {
    return @{
        OsCaption = 'Microsoft Windows Server 2022 Standard'
        OsBuild = '20348'
        LastHotfixDate = (Get-Date).AddDays(-10)
        SecureBoot = $true
        BitLockerProtection = 'On'
        FirewallProfiles = @(
            @{ Name = 'Domain'; Enabled = $true },
            @{ Name = 'Private'; Enabled = $true },
            @{ Name = 'Public'; Enabled = $true }
        )
        AntivirusEnabled = $true
        RealTimeProtection = $true
        AvSignatureAgeDays = 0
        EdrServicesRunning = @('Sense')
        SpoolerStatus = 'Stopped'
        SpoolerStartType = 'Disabled'
        UnnecessaryServicesRunning = @()
        InstalledRoles = @('AD-Domain-Services','DNS')
        PowerShellV2Installed = $false
        ScriptBlockLogging = 1
        ModuleLogging = 1
        Transcription = 1
        LdapServerIntegrity = 2
        LdapEnforceChannelBinding = 2
        SmbServerRequireSigning = 1
        SmbClientRequireSigning = 1
        LmCompatibilityLevel = 5
        NoLMHash = 1
        AuditReceivingNtlm = 1
        RestrictReceivingNtlm = 0
        RequireSignOrSeal = 1
        VulnerableChannelAllowList = $null
        FullSecureChannelProtection = 1
        HardenedPathSysvol = 'RequireMutualAuthentication=1, RequireIntegrity=1'
        HardenedPathNetlogon = 'RequireMutualAuthentication=1, RequireIntegrity=1'
        WefSubscriptionManager = 'Server=http://wef.contoso.com'
        SiemAgentsRunning = @('AzureMonitorAgent')
        AuditPolicy = @{
            'Credential Validation' = 'Success and Failure'
            'Kerberos Authentication Service' = 'Success and Failure'
            'Kerberos Service Ticket Operations' = 'Success and Failure'
            'Computer Account Management' = 'Success'
            'Security Group Management' = 'Success'
            'User Account Management' = 'Success and Failure'
            'Directory Service Access' = 'Failure'
            'Directory Service Changes' = 'Success'
            'Audit Policy Change' = 'Success'
            'Special Logon' = 'Success'
            'Logon' = 'Success and Failure'
            'Account Lockout' = 'Failure'
            'Sensitive Privilege Use' = 'Success and Failure'
            'Security System Extension' = 'Success'
        }
        UserRights = @{
            'SeInteractiveLogonRight' = @('*S-1-5-32-544')
            'SeRemoteInteractiveLogonRight' = @('*S-1-5-32-544')
            'SeDebugPrivilege' = @('*S-1-5-32-544')
        }
    }
}

function Test-HardenedDcPasses {
    $rows = @(Test-DcComputerControls -ComputerName 'DC01' -Facts (New-HardenedDcFacts) -PatchAgeDays 30)
    $byId = @{}
    foreach ($row in $rows) { $byId[$row.ControlId] = $row }

    Assert-True -Condition ($rows.Count -eq 18) -Name 'Hardened DC produces 18 per-computer control rows'
    foreach ($controlId in @('DC-001','DC-002','DC-003','DC-004','DC-005','DC-006','DC-007','DC-010','DC-012','DC-013','DC-014','DC-015','DC-016','DC-017','DC-018','DC-019')) {
        Assert-True -Condition ($byId[$controlId].Status -eq 'Pass') -Name "Hardened DC passes $controlId"
    }
    Assert-True -Condition ($byId['DC-011'].Status -eq 'Pass') -Name 'Hardened DC passes DC-011 (PowerShell restrictions)'
    Assert-True -Condition ($byId['DC-020'].Status -eq 'Manual') -Name 'DC-020 defaults to Manual review when no broad principals hold sensitive rights'
    Assert-True -Condition (@($rows | Where-Object { $_.Target -ne 'DC01' }).Count -eq 0) -Name 'All rows target the assessed computer'
}
Test-HardenedDcPasses

function Test-WeakDcFails {
    $facts = New-HardenedDcFacts
    $facts.OsCaption = 'Microsoft Windows Server 2012 R2 Standard'
    $facts.SecureBoot = $null
    $facts.BitLockerProtection = 'Off'
    $facts.FirewallProfiles = @(@{ Name = 'Domain'; Enabled = $false }, @{ Name = 'Public'; Enabled = $true })
    $facts.SpoolerStatus = 'Running'
    $facts.SpoolerStartType = 'Automatic'
    $facts.LdapServerIntegrity = 1
    $facts.SmbServerRequireSigning = 0
    $facts.LmCompatibilityLevel = 3
    $facts.VulnerableChannelAllowList = 'legacyapp$'
    $facts.HardenedPathSysvol = $null
    $facts.PowerShellV2Installed = $true
    $facts.InstalledRoles = @('AD-Domain-Services','DNS','Web-Server')
    $facts.UnnecessaryServicesRunning = @('RemoteRegistry')
    $facts.AuditPolicy = @{ 'Logon' = 'No Auditing' }
    $facts.UserRights = @{ 'SeRemoteInteractiveLogonRight' = @('*S-1-5-32-544','*S-1-5-11') }

    $rows = @(Test-DcComputerControls -ComputerName 'DC02' -Facts $facts -PatchAgeDays 30)
    $byId = @{}
    foreach ($row in $rows) { $byId[$row.ControlId] = $row }

    foreach ($controlId in @('DC-001','DC-003','DC-004','DC-005','DC-013','DC-014','DC-015','DC-016','DC-017','DC-018','DC-019','DC-020')) {
        Assert-True -Condition ($byId[$controlId].Status -eq 'Fail') -Name "Weak DC fails $controlId"
    }
    Assert-True -Condition ($byId['DC-010'].Status -eq 'Warning' -and $byId['DC-010'].Evidence -match 'RemoteRegistry') -Name 'Weak DC warns on unnecessary services with evidence'
    Assert-True -Condition ($byId['DC-011'].Status -eq 'Fail' -and $byId['DC-011'].Evidence -match 'v2') -Name 'PowerShell v2 presence fails DC-011'
    Assert-True -Condition ($byId['DC-012'].Status -eq 'Warning' -and $byId['DC-012'].Evidence -match 'Web-Server') -Name 'Non-AD role warns DC-012 with evidence'
    Assert-True -Condition ($byId['DC-016'].Evidence -match 'legacyapp') -Name 'Netlogon allow list account appears in evidence'
}
Test-WeakDcFails

function Test-DomainControls {
    $domainFacts = @{
        DomainName = 'contoso.com'
        KrbtgtPasswordLastSet = (Get-Date).AddDays(-400)
        BuiltinAdminName = 'Administrator'
        BuiltinAdminEnabled = $true
        BuiltinAdminPasswordAgeDays = 900
        PrivilegedGroupCounts = @{
            'Domain Admins' = 12
            'Enterprise Admins' = 1
            'Schema Admins' = 2
            'Administrators' = 5
        }
        DcBackups = @(
            @{ Server = 'DC01.contoso.com'; LastBackup = (Get-Date).AddDays(-2) },
            @{ Server = 'DC02.contoso.com'; LastBackup = $null }
        )
    }

    $rows = @(Test-DcDomainControls -DomainFacts $domainFacts -KrbtgtMaxAgeDays 180 -BackupMaxAgeDays 7 -MaxDomainAdmins 5)
    $byId = @{}
    foreach ($row in $rows) { if (-not $byId.ContainsKey($row.ControlId)) { $byId[$row.ControlId] = $row } }

    Assert-True -Condition ($byId['DC-008'].Status -eq 'Warning' -and $byId['DC-008'].Evidence -match 'enabled') -Name 'Enabled built-in Administrator warns DC-008'
    Assert-True -Condition ($byId['DC-009'].Status -eq 'Warning' -and $byId['DC-009'].Evidence -match 'Schema Admins' -and $byId['DC-009'].Evidence -match 'Domain Admins has 12') -Name 'Oversized privileged groups warn DC-009'
    Assert-True -Condition ($byId['DC-021'].Status -eq 'Fail' -and $byId['DC-021'].Evidence -match '400 days') -Name 'Stale krbtgt password fails DC-021'

    $backupRows = @($rows | Where-Object ControlId -eq 'DC-022')
    Assert-True -Condition ($backupRows.Count -eq 2) -Name 'One DC-022 row per Domain Controller'
    Assert-True -Condition (($backupRows | Where-Object Target -eq 'DC01.contoso.com').Status -eq 'Pass') -Name 'Recent backup passes DC-022'
    Assert-True -Condition (($backupRows | Where-Object Target -eq 'DC02.contoso.com').Status -eq 'Fail') -Name 'Missing backup marker fails DC-022'

    $healthyFacts = @{
        DomainName = 'contoso.com'
        KrbtgtPasswordLastSet = (Get-Date).AddDays(-30)
        BuiltinAdminName = 'Administrator'
        BuiltinAdminEnabled = $false
        BuiltinAdminPasswordAgeDays = 100
        PrivilegedGroupCounts = @{ 'Domain Admins' = 3; 'Schema Admins' = 0 }
        DcBackups = @(@{ Server = 'DC01.contoso.com'; LastBackup = (Get-Date).AddDays(-1) })
    }
    $healthyRows = @(Test-DcDomainControls -DomainFacts $healthyFacts -KrbtgtMaxAgeDays 180 -BackupMaxAgeDays 7 -MaxDomainAdmins 5)
    Assert-True -Condition (@($healthyRows | Where-Object Status -ne 'Pass').Count -eq 0) -Name 'Healthy domain facts pass all domain-level controls'
}
Test-DomainControls

if ($script:Failures.Count -gt 0) {
    Write-Host "$($script:Failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host "All $($script:Passes) checks passed." -ForegroundColor Green
exit 0
