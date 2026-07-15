#Requires -Version 7
$ErrorActionPreference = 'Stop'

$script:ScriptPath = Join-Path $PSScriptRoot '..' 'Get-EntraDirSyncProtectionSettings.ps1'
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
    foreach ($name in @('Connect-GraphIfNeeded','New-GraphRequestExecutor','Get-ObjectPropertyValue','Get-OnPremisesSynchronizationObject','ConvertTo-FeatureStatus','New-DirSyncProtectionRow','Get-DirSyncProtectionSettings')) {
        Assert-True -Condition ($functionNames -contains $name) -Name "Function exists: $name"
    }
}
Test-RequiredFunctionsExist

function Test-StaticScriptContent {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw
    Assert-True -Condition ($scriptText -match 'OnPremDirectorySynchronization\.Read\.All') -Name 'Graph scope is present'
    Assert-True -Condition ($scriptText -match '/directory/onPremisesSynchronization') -Name 'Graph onPremisesSynchronization endpoint is present'
    Assert-True -Condition ($scriptText -match 'blockCloudObjectTakeoverThroughHardMatchEnabled') -Name 'Hard-match takeover flag is present'
    Assert-True -Condition ($scriptText -match 'blockSoftMatchEnabled') -Name 'Soft-match block flag is present'
    Assert-True -Condition ($scriptText -notmatch 'Import-Module\s+MSOnline' -and $scriptText -notmatch 'Get-Msol' -and $scriptText -notmatch 'Connect-Msol') -Name 'MSOnline cmdlets are not used'
}
Test-StaticScriptContent

function Test-GetObjectPropertyValueHandlesHashtableAndObject {
    $hashValue = Get-ObjectPropertyValue -InputObject @{ features = @{ blockSoftMatchEnabled = $true } } -PropertyName 'features'
    $objectValue = Get-ObjectPropertyValue -InputObject ([pscustomobject]@{ id = 'tenant-1' }) -PropertyName 'id'

    Assert-True -Condition ($hashValue.blockSoftMatchEnabled -eq $true) -Name 'Get-ObjectPropertyValue reads hashtable property'
    Assert-True -Condition ($objectValue -eq 'tenant-1') -Name 'Get-ObjectPropertyValue reads object property'
}
Test-GetObjectPropertyValueHandlesHashtableAndObject

function Test-DirSyncProtectionSettingsRows {
    $script:RequestedUris = [System.Collections.Generic.List[string]]::new()
    $mockGraph = {
        param([string]$Method, [string]$Uri, $Body)
        $script:RequestedUris.Add($Uri) | Out-Null
        return @{
            id = 'tenant-1'
            features = @{
                blockCloudObjectTakeoverThroughHardMatchEnabled = $true
                blockSoftMatchEnabled = $false
            }
        }
    }

    $rows = @(Get-DirSyncProtectionSettings -GraphRequest $mockGraph)

    Assert-True -Condition ($script:RequestedUris[0] -eq 'https://graph.microsoft.com/v1.0/directory/onPremisesSynchronization?$select=id,features') -Name 'Graph request uses v1.0 onPremisesSynchronization select'
    Assert-True -Condition ($rows.Count -eq 2) -Name 'Two feature rows are returned'
    Assert-True -Condition ($rows[0].Setting -eq 'BlockCloudObjectTakeover' -and $rows[0].Enabled -eq $true -and $rows[0].Status -eq 'Enabled') -Name 'Hard-match takeover row is enabled'
    Assert-True -Condition ($rows[1].Setting -eq 'BlockSoftMatch' -and $rows[1].Enabled -eq $false -and $rows[1].Status -eq 'Disabled') -Name 'Soft-match row is disabled'
    Assert-True -Condition ($rows[0].RecommendedState -eq 'Enabled' -and $rows[1].RecommendedState -eq 'Enabled') -Name 'Recommended state is enabled for both flags'
}
Test-DirSyncProtectionSettingsRows

function Test-DirSyncProtectionSettingsHandlesCollectionResponse {
    $mockGraph = {
        param([string]$Method, [string]$Uri, $Body)
        return @{
            value = @(
                @{
                    id = 'tenant-collection'
                    features = @{
                        blockCloudObjectTakeoverThroughHardMatchEnabled = $false
                        blockSoftMatchEnabled = $true
                    }
                }
            )
        }
    }

    $rows = @(Get-DirSyncProtectionSettings -GraphRequest $mockGraph)

    Assert-True -Condition ($rows[0].TenantId -eq 'tenant-collection') -Name 'Collection response first item is used'
    Assert-True -Condition ($rows[0].Status -eq 'Disabled' -and $rows[1].Status -eq 'Enabled') -Name 'Collection response feature values are read'
}
Test-DirSyncProtectionSettingsHandlesCollectionResponse

if ($script:Failures.Count -gt 0) {
    Write-Host "$($script:Failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host "All $($script:Passes) checks passed." -ForegroundColor Green
exit 0
