#Requires -Version 7
$ErrorActionPreference = 'Stop'

$script:ScriptPath = Join-Path $PSScriptRoot '..' 'Export-M365LicenseInventory.ps1'
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
    foreach ($name in @('Get-SkuFriendlyName','New-TenantSkuRow','Get-ObjectPropertyValueSafe','Get-M365LicenseInventory','Get-M365LicenseFacts')) {
        Assert-True -Condition ($functionNames -contains $name) -Name "Function exists: $name"
    }
}
Test-RequiredFunctionsExist

function Test-ReadOnly {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw
    Assert-True -Condition ($scriptText -notmatch 'Set-MgUserLicense' -and $scriptText -notmatch 'Set-MgUser\b' -and $scriptText -notmatch 'Update-MgUser') -Name 'Script performs no license write operations'
    Assert-True -Condition ($scriptText -match 'Get-MgSubscribedSku' -and $scriptText -match 'Get-MgUser') -Name 'SKU and user queries are present'
}
Test-ReadOnly

function Test-FriendlyNameMapping {
    Assert-True -Condition ((Get-SkuFriendlyName -SkuPartNumber 'SPE_E5') -eq 'Microsoft 365 E5') -Name 'Known SKU mapped to friendly name'
    Assert-True -Condition ((Get-SkuFriendlyName -SkuPartNumber 'AAD_PREMIUM_P2') -eq 'Entra ID P2') -Name 'Entra P2 mapped'
    Assert-True -Condition ((Get-SkuFriendlyName -SkuPartNumber 'CUSTOM_UNKNOWN') -eq 'CUSTOM_UNKNOWN') -Name 'Unknown SKU falls back to part number'
}
Test-FriendlyNameMapping

function New-MockLicenseFacts {
    return @{
        Skus = @(
            [pscustomobject]@{ skuId = 'sku-e5'; skuPartNumber = 'SPE_E5'; consumedUnits = 80; prepaidUnits = @{ enabled = 100; warning = 0 } },
            [pscustomobject]@{ skuId = 'sku-p2'; skuPartNumber = 'AAD_PREMIUM_P2'; consumedUnits = 25; prepaidUnits = @{ enabled = 25; warning = 0 } }
        )
        Users = @(
            [pscustomobject]@{
                userPrincipalName = 'user1@contoso.com'; displayName = 'User One'; accountEnabled = $true
                assignedLicenses = @([pscustomobject]@{ skuId = 'sku-e5' }, [pscustomobject]@{ skuId = 'sku-p2' })
                licenseAssignmentStates = @(
                    [pscustomobject]@{ skuId = 'sku-e5'; assignedByGroup = $null },
                    [pscustomobject]@{ skuId = 'sku-p2'; assignedByGroup = 'group-abc' }
                )
            },
            [pscustomobject]@{
                userPrincipalName = 'user2@contoso.com'; displayName = 'User Two'; accountEnabled = $false
                assignedLicenses = @([pscustomobject]@{ skuId = 'sku-e5' })
                licenseAssignmentStates = @([pscustomobject]@{ skuId = 'sku-e5'; assignedByGroup = $null })
            }
        )
    }
}

function Test-InventoryShaping {
    $inventory = Get-M365LicenseInventory -Facts (New-MockLicenseFacts)

    Assert-True -Condition (@($inventory.TenantSkus).Count -eq 2) -Name 'One row per tenant SKU'
    $e5 = $inventory.TenantSkus | Where-Object SkuPartNumber -eq 'SPE_E5'
    Assert-True -Condition ($e5.ProductName -eq 'Microsoft 365 E5' -and $e5.AvailableUnits -eq 20 -and $e5.EnabledUnits -eq 100) -Name 'SKU available units = enabled - consumed'

    Assert-True -Condition (@($inventory.UserAssignments).Count -eq 3) -Name 'One row per (user, license)'
    $groupAssigned = $inventory.UserAssignments | Where-Object { $_.UserPrincipalName -eq 'user1@contoso.com' -and $_.SkuPartNumber -eq 'AAD_PREMIUM_P2' }
    Assert-True -Condition ($groupAssigned.AssignedViaGroup -eq $true -and $groupAssigned.AssignmentType -eq 'Group') -Name 'Group-based assignment detected'
    $directAssigned = $inventory.UserAssignments | Where-Object { $_.UserPrincipalName -eq 'user1@contoso.com' -and $_.SkuPartNumber -eq 'SPE_E5' }
    Assert-True -Condition ($directAssigned.AssignedViaGroup -eq $false -and $directAssigned.AssignmentType -eq 'Direct' -and $directAssigned.ProductName -eq 'Microsoft 365 E5') -Name 'Direct assignment with friendly name'
}
Test-InventoryShaping

if ($script:Failures.Count -gt 0) {
    Write-Host "$($script:Failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host "All $($script:Passes) checks passed." -ForegroundColor Green
exit 0
