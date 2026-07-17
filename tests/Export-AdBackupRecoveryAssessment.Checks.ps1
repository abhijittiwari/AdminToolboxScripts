#Requires -Version 7
$ErrorActionPreference = 'Stop'

$script:ScriptPath = Join-Path $PSScriptRoot '..' 'Export-AdBackupRecoveryAssessment.ps1'
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
    foreach ($name in @('New-ControlRow','Test-BackupRecoveryControls','Get-BackupRecoveryFacts')) {
        Assert-True -Condition ($functionNames -contains $name) -Name "Function exists: $name"
    }
}
Test-RequiredFunctionsExist

function Test-StaticScriptContent {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw
    Assert-True -Condition ($scriptText -match 'dSASignature') -Name 'Backup marker (dSASignature) is read'
    foreach ($controlId in 1..10 | ForEach-Object { 'BR-{0:d3}' -f $_ }) {
        Assert-True -Condition ($scriptText -match [regex]::Escape($controlId)) -Name "Control $controlId is present"
    }
    Assert-True -Condition ($scriptText -notmatch 'Restore-ADObject' -and $scriptText -notmatch 'Remove-Item' -and $scriptText -notmatch 'Restore-Computer') -Name 'Script performs no restore or destructive actions'
}
Test-StaticScriptContent

function Test-AllBackupsFreshWithDates {
    $now = Get-Date '2026-07-17'
    $facts = @{
        DomainName = 'contoso.com'
        DcBackups = @(
            @{ Server = 'DC01.contoso.com'; LastBackup = $now.AddDays(-1) },
            @{ Server = 'DC02.contoso.com'; LastBackup = $now.AddDays(-3) }
        )
    }
    $rows = @(Test-BackupRecoveryControls -Facts $facts -BackupMaxAgeDays 7 `
        -RunbookLastUpdated $now.AddDays(-30) -LastRestoreTest $now.AddDays(-60) -Now $now)
    $byId = @{}
    foreach ($row in $rows) { $byId[$row.ControlId] = $row }

    Assert-True -Condition ($rows.Count -eq 10) -Name 'Ten BR control rows are produced'
    Assert-True -Condition ($byId['BR-001'].Status -eq 'Pass' -and $byId['BR-001'].Evidence -match 'All 2 DC') -Name 'Fresh backups on all DCs pass BR-001'
    Assert-True -Condition ($byId['BR-004'].Status -eq 'Pass') -Name 'Recent runbook date passes BR-004'
    Assert-True -Condition ($byId['BR-005'].Status -eq 'Pass') -Name 'Recent restore test passes BR-005'
    Assert-True -Condition ($byId['BR-008'].Status -eq 'Manual') -Name 'BR-008 is Manual when all backups fresh (confirm alerting)'
    foreach ($controlId in @('BR-002','BR-003','BR-006','BR-007','BR-009','BR-010')) {
        Assert-True -Condition ($byId[$controlId].Status -eq 'Manual') -Name "Procedural control $controlId is Manual"
    }
}
Test-AllBackupsFreshWithDates

function Test-StaleAndMissingBackupsFail {
    $now = Get-Date '2026-07-17'
    $facts = @{
        DomainName = 'contoso.com'
        DcBackups = @(
            @{ Server = 'DC01.contoso.com'; LastBackup = $now.AddDays(-1) },
            @{ Server = 'DC02.contoso.com'; LastBackup = $now.AddDays(-30) },
            @{ Server = 'DC03.contoso.com'; LastBackup = $null }
        )
    }
    $rows = @(Test-BackupRecoveryControls -Facts $facts -BackupMaxAgeDays 7 -Now $now)
    $byId = @{}
    foreach ($row in $rows) { $byId[$row.ControlId] = $row }

    Assert-True -Condition ($byId['BR-001'].Status -eq 'Fail' -and $byId['BR-001'].Evidence -match 'DC03' -and $byId['BR-001'].Evidence -match 'DC02') -Name 'Stale and missing backups fail BR-001 with DC names'
    Assert-True -Condition ($byId['BR-008'].Status -eq 'Warning' -and $byId['BR-008'].Evidence -match '1 of 3') -Name 'Stale backups warn BR-008 with fresh count'
    Assert-True -Condition ($byId['BR-004'].Status -eq 'Manual' -and $byId['BR-005'].Status -eq 'Manual') -Name 'Runbook and restore-test default to Manual without supplied dates'
}
Test-StaleAndMissingBackupsFail

function Test-StaleRunbookAndRestoreWarn {
    $now = Get-Date '2026-07-17'
    $facts = @{ DomainName = 'contoso.com'; DcBackups = @(@{ Server = 'DC01'; LastBackup = $now.AddDays(-1) }) }
    $rows = @(Test-BackupRecoveryControls -Facts $facts -BackupMaxAgeDays 7 `
        -RunbookLastUpdated $now.AddDays(-400) -RunbookMaxAgeDays 365 `
        -LastRestoreTest $now.AddDays(-500) -RestoreTestMaxAgeDays 365 -Now $now)
    $byId = @{}
    foreach ($row in $rows) { $byId[$row.ControlId] = $row }
    Assert-True -Condition ($byId['BR-004'].Status -eq 'Warning' -and $byId['BR-004'].Evidence -match '400 days') -Name 'Old runbook warns BR-004'
    Assert-True -Condition ($byId['BR-005'].Status -eq 'Warning' -and $byId['BR-005'].Evidence -match '500 days') -Name 'Old restore test warns BR-005'
}
Test-StaleRunbookAndRestoreWarn

if ($script:Failures.Count -gt 0) {
    Write-Host "$($script:Failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host "All $($script:Passes) checks passed." -ForegroundColor Green
exit 0
