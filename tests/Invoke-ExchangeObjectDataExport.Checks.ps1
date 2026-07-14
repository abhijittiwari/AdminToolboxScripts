#Requires -Version 7
$ErrorActionPreference = 'Stop'

$script:ScriptPath = Join-Path $PSScriptRoot '..' 'Invoke-ExchangeObjectDataExport.ps1'
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

# ---------------------------------------------------------------------------
# Checks
# ---------------------------------------------------------------------------

function Test-InvokeExportJob {
    $result = Invoke-ExportJob -Name 'ok job' -Action { 'ran' | Out-Null }
    Assert-True -Condition ($result -eq $true) -Name 'Invoke-ExportJob returns true when the job succeeds'

    $result = Invoke-ExportJob -Name 'failing job' -Action { throw 'boom' } 3>$null
    Assert-True -Condition ($result -eq $false) -Name 'Invoke-ExportJob returns false and continues when the job throws'
}
Test-InvokeExportJob

function Test-ScriptStaticContract {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw

    # The five scripts must be invoked in dependency order.
    $order = @(
        'Export-ExchangeObjectInventory.ps1',
        'Export-ExchangeGroupMemberships.ps1',
        'Export-ExchangeSendAsPermissions.ps1',
        'Export-ExchangeFullAccessPermissions.ps1',
        'Export-ExchangeObjectDataWorkbook.ps1'
    )
    $positions = @($order | ForEach-Object { $scriptText.IndexOf($_) })
    Assert-True -Condition (@($positions | Where-Object { $_ -lt 0 }).Count -eq 0) -Name 'Orchestrator references all five job scripts'
    $sorted = $true
    for ($i = 1; $i -lt $positions.Count; $i++) { if ($positions[$i] -le $positions[$i - 1]) { $sorted = $false } }
    Assert-True -Condition $sorted -Name 'Orchestrator runs the jobs in dependency order'

    Assert-True -Condition ($scriptText -match '\$PSScriptRoot') -Name 'Orchestrator locates job scripts next to itself'
    Assert-True -Condition ($scriptText -match 'Objects\.csv') -Name 'Orchestrator feeds Objects.csv to the downstream jobs'
    Assert-True -Condition ($scriptText -match 'ParameterSetName') -Name 'Orchestrator exposes the inventory selection parameter sets'
    Assert-True -Condition ($scriptText -notmatch 'Disconnect-ExchangeOnline' -and $scriptText -notmatch 'Disconnect-MgGraph') -Name 'Orchestrator leaves sessions open like the job scripts'
    Assert-True -Condition ($scriptText -match '\bexit 1\b') -Name 'Orchestrator exits non-zero when a job fails'
}
Test-ScriptStaticContract

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
if ($script:Failures.Count -gt 0) {
    Write-Host "$($script:Failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host "All $($script:Passes) checks passed." -ForegroundColor Green
exit 0
