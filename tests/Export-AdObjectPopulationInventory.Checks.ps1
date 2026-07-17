#Requires -Version 7
$ErrorActionPreference = 'Stop'

$script:ScriptPath = Join-Path $PSScriptRoot '..' 'Export-AdObjectPopulationInventory.ps1'
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
    foreach ($name in @('Test-IsServerOs','Test-IsStale','Get-UserPopulationRow','Get-GroupPopulationRow','Get-ComputerPopulationRow','Get-AdPopulationFacts','Get-AdObjectPopulationInventory')) {
        Assert-True -Condition ($functionNames -contains $name) -Name "Function exists: $name"
    }
}
Test-RequiredFunctionsExist

function Test-ReadOnly {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw
    Assert-True -Condition ($scriptText -notmatch 'Set-ADUser' -and $scriptText -notmatch 'Remove-AD' -and $scriptText -notmatch 'Disable-ADAccount' -and $scriptText -notmatch 'New-ADUser') -Name 'Script performs no AD write operations'
    Assert-True -Condition ($scriptText -match 'Get-ADUser' -and $scriptText -match 'Get-ADComputer') -Name 'User and computer queries are present'
}
Test-ReadOnly

function Test-Helpers {
    Assert-True -Condition (Test-IsServerOs -OperatingSystem 'Windows Server 2019 Datacenter') -Name 'Server OS detected'
    Assert-True -Condition (-not (Test-IsServerOs -OperatingSystem 'Windows 11 Enterprise')) -Name 'Workstation OS not a server'
    Assert-True -Condition (-not (Test-IsServerOs -OperatingSystem $null)) -Name 'Null OS is not a server'

    $threshold = (Get-Date '2026-07-17').AddDays(-90)
    $recent = (Get-Date '2026-07-10').ToFileTimeUtc()
    $old = (Get-Date '2026-01-01').ToFileTimeUtc()
    Assert-True -Condition (-not (Test-IsStale -LastLogonTimestamp $recent -Threshold $threshold)) -Name 'Recent logon is not stale'
    Assert-True -Condition (Test-IsStale -LastLogonTimestamp $old -Threshold $threshold) -Name 'Old logon is stale'
    Assert-True -Condition (Test-IsStale -LastLogonTimestamp $null -Threshold $threshold) -Name 'Never logged on counts as stale'
}
Test-Helpers

function New-MockPopulationFacts {
    $now = Get-Date '2026-07-17'
    $recent = $now.AddDays(-5).ToFileTimeUtc()
    $old = $now.AddDays(-200).ToFileTimeUtc()
    return @{
        Users = @(
            [pscustomobject]@{ Enabled = $true;  lastLogonTimestamp = $recent; PasswordNeverExpires = $false; SmartcardLogonRequired = $false },
            [pscustomobject]@{ Enabled = $true;  lastLogonTimestamp = $old;    PasswordNeverExpires = $true;  SmartcardLogonRequired = $false },
            [pscustomobject]@{ Enabled = $true;  lastLogonTimestamp = $null;   PasswordNeverExpires = $false; SmartcardLogonRequired = $true },
            [pscustomobject]@{ Enabled = $false; lastLogonTimestamp = $old;    PasswordNeverExpires = $false; SmartcardLogonRequired = $false },
            [pscustomobject]@{ Enabled = $false; lastLogonTimestamp = $null;   PasswordNeverExpires = $false; SmartcardLogonRequired = $false }
        )
        Groups = @(
            [pscustomobject]@{ GroupScope = 'Global';      GroupCategory = 'Security';     mail = $null;              member = @('cn=a') },
            [pscustomobject]@{ GroupScope = 'DomainLocal'; GroupCategory = 'Security';     mail = $null;              member = @() },
            [pscustomobject]@{ GroupScope = 'Universal';   GroupCategory = 'Distribution'; mail = 'dl@contoso.com';   member = @('cn=a','cn=b') }
        )
        Computers = @(
            [pscustomobject]@{ Enabled = $true;  lastLogonTimestamp = $recent; OperatingSystem = 'Windows Server 2022 Datacenter' },
            [pscustomobject]@{ Enabled = $true;  lastLogonTimestamp = $old;    OperatingSystem = 'Windows Server 2019 Standard' },
            [pscustomobject]@{ Enabled = $true;  lastLogonTimestamp = $recent; OperatingSystem = 'Windows 11 Enterprise' },
            [pscustomobject]@{ Enabled = $false; lastLogonTimestamp = $old;    OperatingSystem = 'Windows 10 Pro' }
        )
    }
}

function Test-PopulationCounts {
    $inventory = Get-AdObjectPopulationInventory -Facts (New-MockPopulationFacts) -StaleDays 90 -Now (Get-Date '2026-07-17')

    $u = $inventory.Users
    Assert-True -Condition ($u.TotalUsers -eq 5 -and $u.EnabledUsers -eq 3 -and $u.DisabledUsers -eq 2) -Name 'User enabled/disabled counts'
    Assert-True -Condition ($u.StaleEnabledUsers -eq 2) -Name 'Stale enabled users = old-logon + never-logged-on'
    Assert-True -Condition ($u.NeverLoggedOnUsers -eq 2 -and $u.PasswordNeverExpires -eq 1 -and $u.SmartCardRequired -eq 1) -Name 'User attribute counts'

    $g = $inventory.Groups
    Assert-True -Condition ($g.TotalGroups -eq 3 -and $g.GlobalGroups -eq 1 -and $g.UniversalGroups -eq 1 -and $g.DomainLocalGroups -eq 1) -Name 'Group scope counts'
    Assert-True -Condition ($g.SecurityGroups -eq 2 -and $g.DistributionGroups -eq 1 -and $g.MailEnabledGroups -eq 1 -and $g.EmptyGroups -eq 1) -Name 'Group category/mail/empty counts'

    $c = $inventory.Computers
    Assert-True -Condition ($c.TotalComputers -eq 4 -and $c.Servers -eq 2 -and $c.Workstations -eq 2) -Name 'Computer server/workstation split'
    Assert-True -Condition ($c.EnabledComputers -eq 3 -and $c.DisabledComputers -eq 1 -and $c.StaleEnabledComputers -eq 1) -Name 'Computer enabled/stale counts'
}
Test-PopulationCounts

function Test-EmptyPopulation {
    $inventory = Get-AdObjectPopulationInventory -Facts @{ Users = @(); Groups = @(); Computers = @() } -StaleDays 90 -Now (Get-Date '2026-07-17')
    Assert-True -Condition ($inventory.Users.TotalUsers -eq 0 -and $inventory.Groups.TotalGroups -eq 0 -and $inventory.Computers.TotalComputers -eq 0) -Name 'Empty population yields zero counts (no null-array inflation)'
}
Test-EmptyPopulation

if ($script:Failures.Count -gt 0) {
    Write-Host "$($script:Failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host "All $($script:Passes) checks passed." -ForegroundColor Green
exit 0
