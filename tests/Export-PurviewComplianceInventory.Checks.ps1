#Requires -Version 7
$ErrorActionPreference = 'Stop'

$script:ScriptPath = Join-Path $PSScriptRoot '..' 'Export-PurviewComplianceInventory.ps1'
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
    foreach ($name in @('New-SensitivityLabelRow','New-LabelPolicyRow','New-RetentionPolicyRow','New-DlpPolicyRow','Get-PurviewComplianceInventory','Get-PurviewComplianceFacts')) {
        Assert-True -Condition ($functionNames -contains $name) -Name "Function exists: $name"
    }
}
Test-RequiredFunctionsExist

function Test-ReadOnly {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw
    Assert-True -Condition ($scriptText -notmatch 'Set-Label\b' -and $scriptText -notmatch 'New-Label\b' -and $scriptText -notmatch 'Remove-Label\b' -and $scriptText -notmatch 'Set-DlpCompliancePolicy') -Name 'Script performs no compliance write operations'
    Assert-True -Condition ($scriptText -match 'Get-Label\b' -and $scriptText -match 'Get-DlpCompliancePolicy') -Name 'Label and DLP queries are present'
    Assert-True -Condition ($scriptText -match 'do not survive a tenant move') -Name 'Label GUID rehydration note is present'
}
Test-ReadOnly

function New-MockPurviewFacts {
    return @{
        SensitivityLabels = @(
            [pscustomobject]@{ Name = 'Confidential'; DisplayName = 'Confidential'; Guid = 'g1'; Enabled = $true; ParentId = ''; Priority = 5 },
            [pscustomobject]@{ Name = 'Confidential-Internal'; DisplayName = 'Internal'; Guid = 'g2'; Enabled = $true; ParentId = 'g1'; Priority = 6 },
            [pscustomobject]@{ Name = 'Public'; DisplayName = 'Public'; Guid = 'g3'; Enabled = $false; ParentId = ''; Priority = 0 }
        )
        LabelPolicies = @(
            [pscustomobject]@{ Name = 'Global Label Policy'; Enabled = $true; Labels = @('Confidential','Public') }
        )
        RetentionPolicies = @(
            [pscustomobject]@{ Name = '7-Year Retention'; Enabled = $true; Mode = 'Enforce'; Workload = 'Exchange,SharePoint' }
        )
        DlpPolicies = @(
            [pscustomobject]@{ Name = 'PII DLP'; Mode = 'Enable'; Enabled = $true; Workload = 'Exchange' },
            [pscustomobject]@{ Name = 'PCI DLP'; Mode = 'TestWithNotifications'; Enabled = $true; Workload = 'SharePoint' }
        )
    }
}

function Test-InventoryShaping {
    $inventory = Get-PurviewComplianceInventory -Facts (New-MockPurviewFacts)

    Assert-True -Condition ($inventory.Summary.SensitivityLabels -eq 3 -and $inventory.Summary.SubLabels -eq 1) -Name 'Label and sub-label counts'
    $subLabel = $inventory.SensitivityLabels | Where-Object Name -eq 'Confidential-Internal'
    Assert-True -Condition ($subLabel.IsSubLabel -eq $true -and $subLabel.ParentLabel -eq 'g1') -Name 'Sub-label parent relationship captured'
    $disabled = $inventory.SensitivityLabels | Where-Object Name -eq 'Public'
    Assert-True -Condition ($disabled.Enabled -eq $false) -Name 'Disabled label captured'

    Assert-True -Condition ($inventory.LabelPolicies[0].LabelCount -eq 2 -and $inventory.LabelPolicies[0].PublishedLabels -match 'Confidential') -Name 'Label policy published-labels shaped'
    Assert-True -Condition ($inventory.Summary.RetentionPolicies -eq 1 -and $inventory.Summary.DlpPolicies -eq 2) -Name 'Retention and DLP counts'
    $testDlp = $inventory.DlpPolicies | Where-Object Name -eq 'PCI DLP'
    Assert-True -Condition ($testDlp.Mode -eq 'TestWithNotifications') -Name 'DLP policy mode captured'
}
Test-InventoryShaping

function Test-EmptyFacts {
    $inventory = Get-PurviewComplianceInventory -Facts @{ SensitivityLabels = @(); LabelPolicies = @(); RetentionPolicies = @(); DlpPolicies = @() }
    Assert-True -Condition ($inventory.Summary.SensitivityLabels -eq 0 -and $inventory.Summary.DlpPolicies -eq 0) -Name 'Empty facts produce zeroed summary'
}
Test-EmptyFacts

if ($script:Failures.Count -gt 0) {
    Write-Host "$($script:Failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host "All $($script:Passes) checks passed." -ForegroundColor Green
exit 0
