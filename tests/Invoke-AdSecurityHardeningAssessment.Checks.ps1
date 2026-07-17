#Requires -Version 7
$ErrorActionPreference = 'Stop'

$script:ScriptPath = Join-Path $PSScriptRoot '..' 'Invoke-AdSecurityHardeningAssessment.ps1'
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
    foreach ($name in @('Get-ModuleRunPlan','Get-ModuleArgument','Merge-AssessmentResults')) {
        Assert-True -Condition ($functionNames -contains $name) -Name "Function exists: $name"
    }
}
Test-RequiredFunctionsExist

function Test-ModulesReferenced {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw
    foreach ($moduleScript in @(
        'Export-DcHardeningAssessment.ps1','Export-DomainComputerHardeningAssessment.ps1',
        'Export-EntraIdHardeningAssessment.ps1','Export-DnsHardeningAssessment.ps1',
        'Export-AdcsHardeningAssessment.ps1','Export-AdBackupRecoveryAssessment.ps1',
        'Export-AdMonitoringAssessment.ps1')) {
        Assert-True -Condition ($scriptText -match [regex]::Escape($moduleScript)) -Name "Orchestrator references $moduleScript"
    }
}
Test-ModulesReferenced

function Test-DefaultPlanExcludesEntra {
    $plan = @(Get-ModuleRunPlan -Module @('DC','DJ','DNS','ADCS','BR','MON') -IncludeEntraId:$false)
    Assert-True -Condition ($plan.Count -eq 6) -Name 'Default plan runs six on-premises modules'
    Assert-True -Condition (@($plan | Where-Object Module -eq 'AAD').Count -eq 0) -Name 'Default plan excludes Entra ID'
    Assert-True -Condition ($plan[0].Module -eq 'DC' -and $plan[0].Script -eq 'Export-DcHardeningAssessment.ps1') -Name 'Plan maps module to its script'
}
Test-DefaultPlanExcludesEntra

function Test-IncludeEntraAddsAad {
    $plan = @(Get-ModuleRunPlan -Module @('DC') -IncludeEntraId:$true)
    Assert-True -Condition (@($plan | Where-Object Module -eq 'AAD').Count -eq 1) -Name 'IncludeEntraId adds the AAD module'

    $explicit = @(Get-ModuleRunPlan -Module @('AAD','AAD','DC') -IncludeEntraId:$false)
    Assert-True -Condition (@($explicit | Where-Object Module -eq 'AAD').Count -eq 1) -Name 'Duplicate module selections are de-duplicated'
}
Test-IncludeEntraAddsAad

function Test-PlanPreservesCatalogOrder {
    $plan = @(Get-ModuleRunPlan -Module @('MON','DC','ADCS') -IncludeEntraId:$false)
    Assert-True -Condition ($plan[0].Module -eq 'DC' -and $plan[1].Module -eq 'ADCS' -and $plan[2].Module -eq 'MON') -Name 'Plan follows the catalog order regardless of input order'
}
Test-PlanPreservesCatalogOrder

function Test-ModuleArguments {
    $cred = [pscredential]::new('CONTOSO\svc', (ConvertTo-SecureString 'x' -AsPlainText -Force))

    $dcArgs = Get-ModuleArgument -Module 'DC' -OutputPath 'dc.csv' -Credential $cred
    Assert-True -Condition ($dcArgs.OutputPath -eq 'dc.csv' -and $dcArgs.ContainsKey('Credential')) -Name 'DC module gets OutputPath and Credential'

    $aadArgs = Get-ModuleArgument -Module 'AAD' -OutputPath 'aad.csv' -Credential $cred -BreakGlassUpn @('bg@contoso.com')
    Assert-True -Condition ($aadArgs.ContainsKey('BreakGlassUpn') -and -not $aadArgs.ContainsKey('Credential')) -Name 'AAD module gets BreakGlassUpn but not Credential'

    $djArgs = Get-ModuleArgument -Module 'DJ' -OutputPath 'dj.csv' -ComputersCsv 'servers.csv'
    Assert-True -Condition ($djArgs.ContainsKey('ComputersCsv')) -Name 'DJ module gets ComputersCsv'

    $monArgs = Get-ModuleArgument -Module 'MON' -OutputPath 'mon.csv' -HoneytokenNamePattern 'HNY-*'
    Assert-True -Condition ($monArgs.ContainsKey('HoneytokenNamePattern')) -Name 'MON module gets HoneytokenNamePattern'

    $adcsArgs = Get-ModuleArgument -Module 'ADCS' -OutputPath 'adcs.csv' -Credential $cred
    Assert-True -Condition ($adcsArgs.Count -eq 1 -and $adcsArgs.ContainsKey('OutputPath')) -Name 'ADCS module gets only OutputPath'
}
Test-ModuleArguments

function Test-MergeAddsModuleColumn {
    $moduleResults = @(
        [pscustomobject]@{ Module = 'DC'; Rows = @(
            [pscustomobject]@{ ControlId = 'DC-001'; Control = 'c1'; Target = 'DC01'; Status = 'Pass'; Evidence = 'e1' }
        ) },
        [pscustomobject]@{ Module = 'DNS'; Rows = @(
            [pscustomobject]@{ ControlId = 'DNS-001'; Control = 'c2'; Target = 'DC01'; Status = 'Fail'; Evidence = 'e2' },
            [pscustomobject]@{ ControlId = 'DNS-002'; Control = 'c3'; Target = 'DC01'; Status = 'Pass'; Evidence = 'e3' }
        ) }
    )
    $merged = @(Merge-AssessmentResults -ModuleResults $moduleResults)
    Assert-True -Condition ($merged.Count -eq 3) -Name 'Merge flattens all module rows'
    Assert-True -Condition (($merged | Where-Object ControlId -eq 'DC-001').Module -eq 'DC') -Name 'Merged rows carry the source module'
    Assert-True -Condition (($merged[0].PSObject.Properties.Name) -contains 'Module' -and ($merged[0].PSObject.Properties.Name) -contains 'Evidence') -Name 'Merged rows include Module and Evidence columns'
}
Test-MergeAddsModuleColumn

if ($script:Failures.Count -gt 0) {
    Write-Host "$($script:Failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host "All $($script:Passes) checks passed." -ForegroundColor Green
exit 0
