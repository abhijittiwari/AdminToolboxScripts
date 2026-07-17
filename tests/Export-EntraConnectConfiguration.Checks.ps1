#Requires -Version 7
$ErrorActionPreference = 'Stop'

$script:ScriptPath = Join-Path $PSScriptRoot '..' 'Export-EntraConnectConfiguration.ps1'
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
    foreach ($name in @('New-ConnectorRow','New-SyncRuleRow','ConvertTo-GlobalSettingsRow','Get-EntraConnectFacts','Get-EntraConnectInventory')) {
        Assert-True -Condition ($functionNames -contains $name) -Name "Function exists: $name"
    }
}
Test-RequiredFunctionsExist

function Test-ReadOnly {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw
    Assert-True -Condition ($scriptText -notmatch 'Set-ADSync' -and $scriptText -notmatch 'Add-ADSync' -and $scriptText -notmatch 'Start-ADSyncSyncCycle') -Name 'Script performs no ADSync write/sync operations'
    Assert-True -Condition ($scriptText -match 'Get-ADSyncConnector' -and $scriptText -match 'Get-ADSyncRule') -Name 'ADSync connector and rule queries are present'
    Assert-True -Condition ($scriptText -match 'Get-EntraDirSyncProtectionSettings') -Name 'References the complementary dirsync-protection script (no duplication)'
}
Test-ReadOnly

function New-MockEntraConnectFacts {
    return @{
        Installed = $true
        Version = '2.1.1.0'
        Connectors = @(
            [pscustomobject]@{
                Name = 'contoso.com'; Type = 'AD'; Description = 'On-prem AD'
                Partitions = @(
                    [pscustomobject]@{ ConnectorPartitionScope = [pscustomobject]@{ ContainerInclusionList = @('OU=Users,DC=contoso,DC=com','OU=Groups,DC=contoso,DC=com') } }
                )
            },
            [pscustomobject]@{
                Name = 'contoso.onmicrosoft.com - AAD'; Type = 'Extensible2'; Description = 'Entra ID'
                Partitions = @()
            }
        )
        SyncRules = @(
            [pscustomobject]@{ Name = 'In from AD - User Join'; Direction = 'Inbound'; ConnectorName = 'contoso.com'; SourceObjectType = 'user'; TargetObjectType = 'person'; LinkType = 'Provision'; Precedence = 100; Disabled = $false },
            [pscustomobject]@{ Name = 'Out to AAD - User Join'; Direction = 'Outbound'; ConnectorName = 'contoso.onmicrosoft.com - AAD'; SourceObjectType = 'person'; TargetObjectType = 'user'; LinkType = 'Provision'; Precedence = 101; Disabled = $false }
        )
        GlobalSettings = [pscustomobject]@{
            Parameters = @(
                [pscustomobject]@{ Name = 'Microsoft.SourceAnchor'; Value = 'ms-DS-ConsistencyGuid' },
                [pscustomobject]@{ Name = 'Microsoft.OptionalFeature.GroupWriteBack'; Value = 'False' }
            )
        }
        Scheduler = [pscustomobject]@{ StagingModeEnabled = $false; SyncCycleEnabled = $true; AllowedSyncCycleInterval = '00:30:00' }
    }
}

function Test-InventoryShaping {
    $inventory = Get-EntraConnectInventory -Facts (New-MockEntraConnectFacts)

    Assert-True -Condition ($inventory.Installed -eq $true) -Name 'Installed flag carried through'
    Assert-True -Condition ($inventory.GlobalSettings.SourceAnchorAttribute -eq 'ms-DS-ConsistencyGuid' -and $inventory.GlobalSettings.ConsistencyGuidInUse -eq $true) -Name 'Source anchor / consistencyGuid detected'
    Assert-True -Condition ($inventory.GlobalSettings.StagingModeEnabled -eq $false -and $inventory.GlobalSettings.Version -eq '2.1.1.0') -Name 'Scheduler + version shaped'

    Assert-True -Condition (@($inventory.Connectors).Count -eq 2) -Name 'One row per connector'
    $adConnector = $inventory.Connectors | Where-Object Name -eq 'contoso.com'
    Assert-True -Condition ($adConnector.IncludedContainers -match 'OU=Users' -and $adConnector.PartitionCount -eq 1) -Name 'Connector OU inclusion scope captured'

    Assert-True -Condition (@($inventory.SyncRules).Count -eq 2) -Name 'One row per sync rule'
    $inbound = $inventory.SyncRules | Where-Object Direction -eq 'Inbound'
    Assert-True -Condition ($inbound.Connector -eq 'contoso.com' -and $inbound.Precedence -eq 100) -Name 'Sync rule direction/connector/precedence shaped'
}
Test-InventoryShaping

function Test-StagingModeDetected {
    $facts = New-MockEntraConnectFacts
    $facts.Scheduler = [pscustomobject]@{ StagingModeEnabled = $true; SyncCycleEnabled = $false; AllowedSyncCycleInterval = '00:30:00' }
    $inventory = Get-EntraConnectInventory -Facts $facts
    Assert-True -Condition ($inventory.GlobalSettings.StagingModeEnabled -eq $true) -Name 'Staging-mode server detected'
}
Test-StagingModeDetected

if ($script:Failures.Count -gt 0) {
    Write-Host "$($script:Failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host "All $($script:Passes) checks passed." -ForegroundColor Green
exit 0
