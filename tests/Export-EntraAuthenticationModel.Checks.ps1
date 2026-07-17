#Requires -Version 7
$ErrorActionPreference = 'Stop'

$script:ScriptPath = Join-Path $PSScriptRoot '..' 'Export-EntraAuthenticationModel.ps1'
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
    foreach ($name in @('Connect-GraphIfNeeded','New-GraphRequestExecutor','Get-ObjectPropertyValue','New-DomainAuthRow','Get-AuthenticationModelSummary','Get-EntraAuthModelFacts','Get-EntraAuthenticationInventory')) {
        Assert-True -Condition ($functionNames -contains $name) -Name "Function exists: $name"
    }
}
Test-RequiredFunctionsExist

function Test-ReadOnly {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw
    Assert-True -Condition ($scriptText -notmatch "'PATCH'" -and $scriptText -notmatch "'POST'" -and $scriptText -notmatch "'DELETE'") -Name 'Script only issues GET requests'
    Assert-True -Condition ($scriptText -match '/domains' -and $scriptText -match 'authentication/agents') -Name 'Domains and PTA agent endpoints are queried'
    Assert-True -Condition ($scriptText -notmatch 'Import-Module\s+MSOnline' -and $scriptText -notmatch 'Get-Msol') -Name 'MSOnline cmdlets are not used'
}
Test-ReadOnly

function New-FederatedGraphMock {
    return {
        param([string]$Method, [string]$Uri, $Body)
        switch -Regex ($Uri) {
            '/domains' {
                return @{ value = @(
                    @{ id = 'contoso.com'; authenticationType = 'Federated'; isDefault = $true; isVerified = $true; supportedServices = @('Email','OfficeCommunicationsOnline') },
                    @{ id = 'contoso.onmicrosoft.com'; authenticationType = 'Managed'; isDefault = $false; isVerified = $true; supportedServices = @('Email') }
                ) }
            }
            'organization' { return @{ value = @(@{ onPremisesSyncEnabled = $true }) } }
            'authentication/agents' { return @{ value = @() } }
            default { throw "Unexpected URI: $Uri" }
        }
    }
}

function Test-FederatedTenant {
    $facts = Get-EntraAuthModelFacts -GraphRequest (New-FederatedGraphMock)
    $inventory = Get-EntraAuthenticationInventory -Facts $facts

    Assert-True -Condition (@($inventory.Domains).Count -eq 2) -Name 'One row per domain'
    $federated = $inventory.Domains | Where-Object DomainName -eq 'contoso.com'
    Assert-True -Condition ($federated.IsFederated -eq $true -and $federated.AuthenticationType -eq 'Federated') -Name 'Federated domain detected'
    Assert-True -Condition ($inventory.Summary.AuthenticationModel -match 'Mixed' -and $inventory.Summary.FederatedDomainCount -eq 1) -Name 'Mixed federated+managed model summarised'
    Assert-True -Condition ($inventory.Summary.OnPremisesSyncEnabled -eq $true) -Name 'On-prem sync flag captured'
}
Test-FederatedTenant

function Test-PtaTenant {
    $mock = {
        param([string]$Method, [string]$Uri, $Body)
        switch -Regex ($Uri) {
            '/domains' { return @{ value = @(@{ id = 'contoso.com'; authenticationType = 'Managed'; isDefault = $true; isVerified = $true; supportedServices = @('Email') }) } }
            'organization' { return @{ value = @(@{ onPremisesSyncEnabled = $true }) } }
            'authentication/agents' { return @{ value = @(@{ id = 'agent1' }, @{ id = 'agent2' }) } }
            default { throw "Unexpected URI: $Uri" }
        }
    }
    $inventory = Get-EntraAuthenticationInventory -Facts (Get-EntraAuthModelFacts -GraphRequest $mock)
    Assert-True -Condition ($inventory.Summary.AuthenticationModel -match 'Pass-through' -and $inventory.Summary.PtaAgentCount -eq 2) -Name 'PTA agents drive Pass-through Authentication model'
}
Test-PtaTenant

function Test-PhsTenant {
    $mock = {
        param([string]$Method, [string]$Uri, $Body)
        switch -Regex ($Uri) {
            '/domains' { return @{ value = @(@{ id = 'contoso.com'; authenticationType = 'Managed'; isDefault = $true; isVerified = $true; supportedServices = @('Email') }) } }
            'organization' { return @{ value = @(@{ onPremisesSyncEnabled = $true }) } }
            'authentication/agents' { return @{ value = @() } }
            default { throw "Unexpected URI: $Uri" }
        }
    }
    $inventory = Get-EntraAuthenticationInventory -Facts (Get-EntraAuthModelFacts -GraphRequest $mock)
    Assert-True -Condition ($inventory.Summary.AuthenticationModel -match 'Password Hash Sync') -Name 'Sync enabled with no PTA agents implies PHS'
}
Test-PhsTenant

if ($script:Failures.Count -gt 0) {
    Write-Host "$($script:Failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host "All $($script:Passes) checks passed." -ForegroundColor Green
exit 0
