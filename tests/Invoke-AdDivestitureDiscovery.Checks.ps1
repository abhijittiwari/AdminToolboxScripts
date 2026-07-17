#Requires -Version 7
$ErrorActionPreference = 'Stop'

$script:ScriptPath = Join-Path $PSScriptRoot '..' 'Invoke-AdDivestitureDiscovery.ps1'
$script:RepoRoot = Join-Path $PSScriptRoot '..'
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
    foreach ($name in @('Get-DiscoveryCatalog','Get-DiscoveryRunPlan','Get-DiscoveryModuleArgument')) {
        Assert-True -Condition ($functionNames -contains $name) -Name "Function exists: $name"
    }
}
Test-RequiredFunctionsExist

function Test-CatalogedScriptsExist {
    $catalog = Get-DiscoveryCatalog
    foreach ($key in $catalog.Keys) {
        $scriptPath = Join-Path $script:RepoRoot $catalog[$key].Script
        Assert-True -Condition (Test-Path -Path $scriptPath) -Name "Cataloged script exists: $($catalog[$key].Script)"
    }
}
Test-CatalogedScriptsExist

function Test-DefaultPlanIsOnPrem {
    $plan = @(Get-DiscoveryRunPlan -Module @('AdForest','AdSites','AdTrusts','AdOuGpo','AdPopulation','AdPrivileged','AdSidHistory','AdInfra','EntraConnect','ExchangeOnPrem') -IncludeCloud:$false)
    Assert-True -Condition ($plan.Count -eq 10) -Name 'Default plan runs the ten on-premises modules'
    Assert-True -Condition (@($plan | Where-Object Cloud).Count -eq 0) -Name 'Default plan excludes cloud modules'
    Assert-True -Condition ($plan[0].Module -eq 'AdForest') -Name 'Plan follows catalog order'
}
Test-DefaultPlanIsOnPrem

function Test-IncludeCloudAddsCloudModules {
    $plan = @(Get-DiscoveryRunPlan -Module @('AdForest') -IncludeCloud:$true)
    $cloudModules = @($plan | Where-Object Cloud | ForEach-Object Module)
    foreach ($cloud in @('EntraAuth','M365Volume','M365License','TeamsVoice','Purview')) {
        Assert-True -Condition ($cloudModules -contains $cloud) -Name "IncludeCloud adds $cloud"
    }
    Assert-True -Condition (@($plan | Where-Object Module -eq 'AdForest').Count -eq 1) -Name 'Explicit on-prem module retained alongside cloud'
}
Test-IncludeCloudAddsCloudModules

function Test-DeduplicationAndOrder {
    $plan = @(Get-DiscoveryRunPlan -Module @('Purview','AdForest','Purview','AdTrusts') -IncludeCloud:$false)
    Assert-True -Condition (@($plan | Where-Object Module -eq 'Purview').Count -eq 1) -Name 'Duplicate module selections de-duplicated'
    Assert-True -Condition ($plan[0].Module -eq 'AdForest' -and $plan[1].Module -eq 'AdTrusts' -and $plan[2].Module -eq 'Purview') -Name 'Plan output follows catalog order regardless of input order'
}
Test-DeduplicationAndOrder

function Test-ModuleArguments {
    $cred = [pscredential]::new('CONTOSO\svc', (ConvertTo-SecureString 'x' -AsPlainText -Force))
    $run = '/tmp/run'

    $folderEntry = [pscustomobject]@{ Module = 'AdForest'; Script = 'Export-AdForestInventory.ps1'; Output = 'Folder'; Cloud = $false }
    $folderArgs = Get-DiscoveryModuleArgument -PlanEntry $folderEntry -RunFolder $run -Credential $cred
    Assert-True -Condition ($folderArgs.ContainsKey('OutputFolder') -and $folderArgs.OutputFolder -match 'AdForest' -and $folderArgs.ContainsKey('Credential')) -Name 'Folder-output AD module gets OutputFolder and Credential'

    $fileEntry = [pscustomobject]@{ Module = 'AdTrusts'; Script = 'Export-AdTrustInventory.ps1'; Output = 'File'; Cloud = $false }
    $fileArgs = Get-DiscoveryModuleArgument -PlanEntry $fileEntry -RunFolder $run -Credential $cred
    Assert-True -Condition ($fileArgs.ContainsKey('OutputPath') -and $fileArgs.OutputPath -match 'AdTrusts\.csv') -Name 'File-output module gets OutputPath'

    $cloudEntry = [pscustomobject]@{ Module = 'EntraAuth'; Script = 'Export-EntraAuthenticationModel.ps1'; Output = 'File'; Cloud = $true }
    $cloudArgs = Get-DiscoveryModuleArgument -PlanEntry $cloudEntry -RunFolder $run -Credential $cred
    Assert-True -Condition (-not $cloudArgs.ContainsKey('Credential')) -Name 'Cloud module does not receive -Credential'

    $ecEntry = [pscustomobject]@{ Module = 'EntraConnect'; Script = 'Export-EntraConnectConfiguration.ps1'; Output = 'Folder'; Cloud = $false }
    $ecArgs = Get-DiscoveryModuleArgument -PlanEntry $ecEntry -RunFolder $run -Credential $cred
    Assert-True -Condition (-not $ecArgs.ContainsKey('Credential')) -Name 'EntraConnect (no -Credential param) does not receive Credential'
}
Test-ModuleArguments

function Test-ReadOnly {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw
    Assert-True -Condition ($scriptText -notmatch 'Set-AD' -and $scriptText -notmatch 'Remove-AD') -Name 'Orchestrator performs no AD writes'
}
Test-ReadOnly

if ($script:Failures.Count -gt 0) {
    Write-Host "$($script:Failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host "All $($script:Passes) checks passed." -ForegroundColor Green
exit 0
