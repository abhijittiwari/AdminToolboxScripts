#Requires -Version 7
$ErrorActionPreference = 'Stop'

$script:ScriptPath = Join-Path $PSScriptRoot '..' 'Export-AdForestInventory.ps1'
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
    foreach ($name in @('ConvertTo-SchemaVersionName','New-ForestRow','New-DomainRow','New-DomainControllerRow','Get-AdForestFacts','Get-AdForestInventory')) {
        Assert-True -Condition ($functionNames -contains $name) -Name "Function exists: $name"
    }
}
Test-RequiredFunctionsExist

function Test-ReadOnly {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw
    Assert-True -Condition ($scriptText -notmatch 'Set-AD' -and $scriptText -notmatch 'Remove-AD' -and $scriptText -notmatch 'New-ADObject' -and $scriptText -notmatch 'Move-AD') -Name 'Script performs no AD write operations'
    Assert-True -Condition ($scriptText -match 'Get-ADForest') -Name 'Forest query is present'
    Assert-True -Condition ($scriptText -match 'OperationMasterRoles') -Name 'FSMO roles are collected'
}
Test-ReadOnly

function Test-SchemaVersionMapping {
    Assert-True -Condition ((ConvertTo-SchemaVersionName -ObjectVersion 88) -eq 'Windows Server 2019/2022') -Name 'Schema 88 maps to 2019/2022'
    Assert-True -Condition ((ConvertTo-SchemaVersionName -ObjectVersion 87) -eq 'Windows Server 2016') -Name 'Schema 87 maps to 2016'
    Assert-True -Condition ((ConvertTo-SchemaVersionName -ObjectVersion $null) -eq 'Unknown') -Name 'Null schema version returns Unknown'
    Assert-True -Condition ((ConvertTo-SchemaVersionName -ObjectVersion 999) -match '999') -Name 'Unmapped schema version echoes the raw number'
}
Test-SchemaVersionMapping

function New-MockForestFacts {
    return @{
        SchemaObjectVersion = 88
        Forest = [pscustomobject]@{
            Name = 'contoso.com'
            ForestMode = 'Windows2016Forest'
            RootDomain = 'contoso.com'
            Domains = @('contoso.com','emea.contoso.com')
            GlobalCatalogs = @('dc01','dc02','dc03')
            Sites = @('Azure-AE','Site-Perth')
            UPNSuffixes = @('south32.net','contoso.io')
            SchemaMaster = 'dc01.contoso.com'
            DomainNamingMaster = 'dc01.contoso.com'
        }
        Domains = @(
            [pscustomobject]@{ DNSRoot = 'contoso.com'; NetBIOSName = 'CONTOSO'; DomainMode = 'Windows2016Domain'; DomainSID = 'S-1-5-21-1'; PDCEmulator = 'dc01.contoso.com'; RIDMaster = 'dc01.contoso.com'; InfrastructureMaster = 'dc02.contoso.com' }
        )
        DomainControllers = @(
            [pscustomobject]@{ HostName = 'dc01.contoso.com'; Domain = 'contoso.com'; Site = 'Azure-AE'; OperatingSystem = 'Windows Server 2019 Datacenter'; OperatingSystemVersion = '10.0 (17763)'; IPv4Address = '10.0.0.1'; IsGlobalCatalog = $true; IsReadOnly = $false; Enabled = $true; OperationMasterRoles = @('SchemaMaster','PDCEmulator','RIDMaster') },
            [pscustomobject]@{ HostName = 'rodc01.emea.contoso.com'; Domain = 'emea.contoso.com'; Site = 'Site-Perth'; OperatingSystem = 'Windows Server 2016 Standard'; OperatingSystemVersion = '10.0 (14393)'; IPv4Address = '10.1.0.1'; IsGlobalCatalog = $true; IsReadOnly = $true; Enabled = $true; OperationMasterRoles = @() }
        )
    }
}

function Test-InventoryShaping {
    $inventory = Get-AdForestInventory -Facts (New-MockForestFacts)

    Assert-True -Condition ($inventory.Forest.ForestName -eq 'contoso.com') -Name 'Forest row carries the forest name'
    Assert-True -Condition ($inventory.Forest.SchemaVersionName -eq 'Windows Server 2019/2022') -Name 'Forest row resolves the schema version name'
    Assert-True -Condition ($inventory.Forest.UpnSuffixes -eq 'south32.net;contoso.io') -Name 'UPN suffixes are semicolon-joined'
    Assert-True -Condition ($inventory.Forest.GlobalCatalogs -eq 3) -Name 'Global catalog count is captured'

    Assert-True -Condition (@($inventory.Domains).Count -eq 1 -and $inventory.Domains[0].NetBIOSName -eq 'CONTOSO') -Name 'Domain row shaping'
    Assert-True -Condition ($inventory.Domains[0].PDCEmulator -eq 'dc01.contoso.com') -Name 'Domain FSMO holders are captured'

    Assert-True -Condition (@($inventory.DomainControllers).Count -eq 2) -Name 'Two DC rows are produced'
    $dc01 = $inventory.DomainControllers | Where-Object HostName -eq 'dc01.contoso.com'
    Assert-True -Condition ($dc01.FsmoRoles -eq 'SchemaMaster;PDCEmulator;RIDMaster') -Name 'DC FSMO roles are semicolon-joined'
    $rodc = $inventory.DomainControllers | Where-Object HostName -eq 'rodc01.emea.contoso.com'
    Assert-True -Condition ($rodc.IsReadOnly -eq $true -and $rodc.FsmoRoles -eq '') -Name 'RODC with no FSMO roles shapes correctly'
}
Test-InventoryShaping

if ($script:Failures.Count -gt 0) {
    Write-Host "$($script:Failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host "All $($script:Passes) checks passed." -ForegroundColor Green
exit 0
