#Requires -Version 7
$ErrorActionPreference = 'Stop'

$script:ScriptPath = Join-Path $PSScriptRoot '..' 'Export-DomainComputerHardeningAssessment.ps1'
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
    foreach ($name in @('New-ControlRow','Get-DjFacts','Test-DjComputerControls','Test-DjDomainControls','Get-DjDomainFacts','Get-TargetComputers')) {
        Assert-True -Condition ($functionNames -contains $name) -Name "Function exists: $name"
    }
}
Test-RequiredFunctionsExist

function Test-StaticScriptContent {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw
    Assert-True -Condition ($scriptText -match 'msLAPS-PasswordExpirationTime') -Name 'Windows LAPS attribute is checked'
    Assert-True -Condition ($scriptText -match 'ms-Mcs-AdmPwdExpirationTime') -Name 'Legacy LAPS attribute is checked'
    Assert-True -Condition ($scriptText -match 'ms-DS-MachineAccountQuota') -Name 'Machine account quota is checked'
    Assert-True -Condition ($scriptText -match 'AttackSurfaceReductionRules_Ids') -Name 'ASR rules are checked'
    Assert-True -Condition ($scriptText -match 'EnableSMB1Protocol') -Name 'SMBv1 is checked'
    Assert-True -Condition ($scriptText -match 'EnableMulticast') -Name 'LLMNR policy is checked'
    Assert-True -Condition ($scriptText -match 'RunAsPPL') -Name 'LSA protection is checked'
    foreach ($controlId in 1..20 | ForEach-Object { 'DJ-{0:d3}' -f $_ }) {
        Assert-True -Condition ($scriptText -match [regex]::Escape($controlId)) -Name "Control $controlId is present"
    }
    Assert-True -Condition ($scriptText -notmatch 'Set-ItemProperty' -and $scriptText -notmatch 'Set-ADComputer' -and $scriptText -notmatch 'Remove-ADComputer') -Name 'Script performs no configuration changes'
}
Test-StaticScriptContent

function New-HardenedDjFacts {
    return @{
        OsCaption = 'Microsoft Windows 11 Enterprise'
        OsBuild = '26100'
        LastHotfixDate = (Get-Date).AddDays(-12)
        BitLockerProtection = 'On'
        SecureBoot = $true
        TpmPresent = $true
        TpmReady = $true
        TpmSpecVersion = '2.0, 0, 1.59'
        FirewallProfiles = @(
            @{ Name = 'Domain'; Enabled = $true },
            @{ Name = 'Private'; Enabled = $true },
            @{ Name = 'Public'; Enabled = $true }
        )
        AntivirusEnabled = $true
        RealTimeProtection = $true
        AvSignatureAgeDays = 0
        EdrServicesRunning = @('Sense')
        AsrRulesConfigured = 16
        AsrRulesEnabled = 16
        Smb1Enabled = $false
        SmbServerSigningRequired = $true
        LlmnrPolicy = 0
        NetbiosOptions = @(2, 2)
        LmCompatibilityLevel = 5
        SmbClientSigningEnabled = 1
        LocalAdministrators = @('CONTOSO\LAPS-Admins','Administrator')
        EnableLua = 1
        ConsentPromptBehaviorAdmin = 2
        PromptOnSecureDesktop = 1
        InactivityTimeoutSecs = 900
        LegalNoticeCaption = 'Authorized use only'
        LegalNoticeText = 'This system is for authorized users only.'
        CredentialGuardRunning = $true
        RunAsPpl = 1
        ScriptBlockLogging = 1
        ModuleLogging = 1
        AppliedGpoNames = @('Default Domain Policy','CIS Windows 11 Baseline')
        NoDriveTypeAutoRun = 255
        RemovableStorageDenyAll = 1
        TimeSyncType = 'NT5DS'
    }
}

function Test-HardenedComputerPasses {
    $rows = @(Test-DjComputerControls -ComputerName 'PC01' -Facts (New-HardenedDjFacts) -PatchAgeDays 30 -BaselineGpoPattern 'CIS')
    $byId = @{}
    foreach ($row in $rows) { $byId[$row.ControlId] = $row }

    Assert-True -Condition ($rows.Count -eq 17) -Name 'Hardened computer produces 17 per-computer control rows'
    foreach ($controlId in @('DJ-001','DJ-002','DJ-004','DJ-005','DJ-006','DJ-007','DJ-008','DJ-009','DJ-010','DJ-011','DJ-012','DJ-013','DJ-014','DJ-015','DJ-016','DJ-018','DJ-019')) {
        Assert-True -Condition ($byId[$controlId].Status -eq 'Pass') -Name "Hardened computer passes $controlId"
    }
}
Test-HardenedComputerPasses

function Test-WeakComputerFails {
    $facts = New-HardenedDjFacts
    $facts.OsCaption = 'Microsoft Windows 8.1 Pro'
    $facts.BitLockerProtection = 'Off'
    $facts.TpmSpecVersion = '1.2'
    $facts.AntivirusEnabled = $false
    $facts.EdrServicesRunning = @()
    $facts.AsrRulesEnabled = 0
    $facts.Smb1Enabled = $true
    $facts.LmCompatibilityLevel = 2
    $facts.SmbServerSigningRequired = $false
    $facts.SmbClientSigningEnabled = 0
    $facts.LocalAdministrators = @('CONTOSO\Domain Users','Administrator')
    $facts.EnableLua = 0
    $facts.CredentialGuardRunning = $false
    $facts.RunAsPpl = $null
    $facts.ScriptBlockLogging = $null
    $facts.NoDriveTypeAutoRun = 145
    $facts.InactivityTimeoutSecs = 0
    $facts.TimeSyncType = 'NTP'

    $rows = @(Test-DjComputerControls -ComputerName 'PC02' -Facts $facts -PatchAgeDays 30 -BaselineGpoPattern 'CIS')
    $byId = @{}
    foreach ($row in $rows) { $byId[$row.ControlId] = $row }

    foreach ($controlId in @('DJ-001','DJ-004','DJ-005','DJ-007','DJ-008','DJ-009','DJ-010','DJ-011','DJ-012','DJ-013','DJ-014','DJ-016','DJ-018','DJ-019')) {
        Assert-True -Condition ($byId[$controlId].Status -eq 'Fail') -Name "Weak computer fails $controlId"
    }
    Assert-True -Condition ($byId['DJ-009'].Evidence -match 'SMBv1') -Name 'SMBv1 appears in DJ-009 evidence'
    Assert-True -Condition ($byId['DJ-011'].Evidence -match 'Domain Users') -Name 'Broad local admin group appears in DJ-011 evidence'
}
Test-WeakComputerFails

function Test-BaselineGpoManualWhenNoPattern {
    $rows = @(Test-DjComputerControls -ComputerName 'PC03' -Facts (New-HardenedDjFacts) -PatchAgeDays 30)
    $gpoRow = $rows | Where-Object ControlId -eq 'DJ-015'
    Assert-True -Condition ($gpoRow.Status -eq 'Manual' -and $gpoRow.Evidence -match 'CIS Windows 11 Baseline') -Name 'DJ-015 is Manual with GPO list when no pattern supplied'
}
Test-BaselineGpoManualWhenNoPattern

function Test-DomainControls {
    $weakFacts = @{
        DomainName = 'contoso.com'
        LapsSchemaPresent = $true
        EnabledComputerCount = 100
        LapsManagedCount = 40
        StaleComputerCount = 12
        StaleComputerSample = @('OLDPC1','OLDPC2')
        MachineAccountQuota = 10
    }
    $rows = @(Test-DjDomainControls -DomainFacts $weakFacts -StaleDays 90 -LapsMinimumCoveragePercent 95)
    $byId = @{}
    foreach ($row in $rows) { $byId[$row.ControlId] = $row }

    Assert-True -Condition ($rows.Count -eq 3) -Name 'Domain evaluation produces 3 rows'
    Assert-True -Condition ($byId['DJ-003'].Status -eq 'Warning' -and $byId['DJ-003'].Evidence -match '40 of 100') -Name 'Partial LAPS coverage warns DJ-003 with counts'
    Assert-True -Condition ($byId['DJ-017'].Status -eq 'Warning' -and $byId['DJ-017'].Evidence -match 'OLDPC1') -Name 'Stale computers warn DJ-017 with sample'
    Assert-True -Condition ($byId['DJ-020'].Status -eq 'Fail' -and $byId['DJ-020'].Evidence -match 'MachineAccountQuota=10') -Name 'Default machine account quota fails DJ-020'

    $healthyFacts = @{
        DomainName = 'contoso.com'
        LapsSchemaPresent = $true
        EnabledComputerCount = 100
        LapsManagedCount = 98
        StaleComputerCount = 0
        StaleComputerSample = @()
        MachineAccountQuota = 0
    }
    $healthyRows = @(Test-DjDomainControls -DomainFacts $healthyFacts -StaleDays 90 -LapsMinimumCoveragePercent 95)
    Assert-True -Condition (@($healthyRows | Where-Object Status -ne 'Pass').Count -eq 0) -Name 'Healthy domain facts pass all domain-level controls'

    $noLapsFacts = @{
        DomainName = 'contoso.com'
        LapsSchemaPresent = $false
        EnabledComputerCount = 100
        LapsManagedCount = 0
        StaleComputerCount = 0
        StaleComputerSample = @()
        MachineAccountQuota = 0
    }
    $noLapsRows = @(Test-DjDomainControls -DomainFacts $noLapsFacts -StaleDays 90 -LapsMinimumCoveragePercent 95)
    Assert-True -Condition (($noLapsRows | Where-Object ControlId -eq 'DJ-003').Status -eq 'Fail') -Name 'Missing LAPS schema fails DJ-003'
}
Test-DomainControls

function Test-TargetComputersFromCsv {
    $csvPath = Join-Path ([System.IO.Path]::GetTempPath()) ("djtargets-{0}.csv" -f [guid]::NewGuid())
    @(
        [pscustomobject]@{ ComputerName = 'PC01.contoso.com' },
        [pscustomobject]@{ ComputerName = 'PC02.contoso.com' },
        [pscustomobject]@{ ComputerName = '' }
    ) | Export-Csv -Path $csvPath -NoTypeInformation
    try {
        $targets = @(Get-TargetComputers -ComputerName @() -ComputersCsv $csvPath -SearchBase '' -SampleSize 25 -Credential $null)
        Assert-True -Condition ($targets.Count -eq 2 -and $targets[0] -eq 'PC01.contoso.com') -Name 'CSV targets are read and blank rows skipped'
    }
    finally {
        Remove-Item -Path $csvPath -ErrorAction SilentlyContinue
    }

    $explicit = @(Get-TargetComputers -ComputerName @('SRV01') -ComputersCsv '' -SearchBase '' -SampleSize 25 -Credential $null)
    Assert-True -Condition ($explicit.Count -eq 1 -and $explicit[0] -eq 'SRV01') -Name 'Explicit -ComputerName takes precedence'
}
Test-TargetComputersFromCsv

if ($script:Failures.Count -gt 0) {
    Write-Host "$($script:Failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host "All $($script:Passes) checks passed." -ForegroundColor Green
exit 0
