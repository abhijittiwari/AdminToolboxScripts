#Requires -Version 7
$ErrorActionPreference = 'Stop'

$script:ScriptPath = Join-Path $PSScriptRoot '..' 'Get-CrossTenantAdminRoleAssignments.ps1'
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
    foreach ($name in @('Test-RequiredCsvColumns','Join-ReportValues','Add-UserRole','Resolve-RolePrincipal','Get-TenantRoleIndex','Get-UserReportRow','Invoke-TenantExport')) {
        Assert-True -Condition ($functionNames -contains $name) -Name "Function exists: $name"
    }
}
Test-RequiredFunctionsExist

function Test-CsvValidation {
    $valid = @([pscustomobject]@{ WAEUPN = 'a@wae.com'; Prefix = 'adma'; FortescueUPN = 'a@fortescue.com' })
    Test-RequiredCsvColumns -Rows $valid -RequiredColumns @('WAEUPN','Prefix','FortescueUPN')
    Assert-True -Condition $true -Name 'CSV validation accepts required columns'

    $threw = $false
    try {
        Test-RequiredCsvColumns -Rows @([pscustomobject]@{ WAEUPN = 'a@wae.com'; Prefix = 'adma' }) -RequiredColumns @('WAEUPN','Prefix','FortescueUPN')
    }
    catch {
        $threw = $_.Exception.Message -match 'FortescueUPN'
    }
    Assert-True -Condition $threw -Name 'CSV validation reports missing FortescueUPN'
}
Test-CsvValidation

function Test-JoinReportValues {
    $joined = Join-ReportValues -Values @('Role B', '', $null, 'Role A', 'Role A')
    Assert-True -Condition ($joined -eq 'Role A; Role B') -Name 'Join-ReportValues sorts, deduplicates, and skips blanks'
}
Test-JoinReportValues

function Test-AddUserRole {
    $roleMap = @{}
    Add-UserRole -RoleMap $roleMap -UserId 'u1' -RoleName 'Global Reader (Active)'
    Add-UserRole -RoleMap $roleMap -UserId 'u1' -RoleName 'Global Reader (Active)'
    Assert-True -Condition ($roleMap['u1'].Count -eq 1 -and $roleMap['u1'][0] -eq 'Global Reader (Active)') -Name 'Add-UserRole deduplicates user roles'
}
Test-AddUserRole

function Test-StaticScriptContent {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw
    foreach ($column in @('Environment','Prefix','InputUPN','UserPrincipalName','DisplayName','ImmutableId','ActiveDirectoryRoles','PimEligibleDirectoryRoles','AllDirectoryRoles','LicensesAssigned','LookupStatus','Error')) {
        Assert-True -Condition ($scriptText -match [regex]::Escape($column)) -Name "Output column present: $column"
    }
    Assert-True -Condition ($scriptText -match 'Directory\.Read\.All' -and $scriptText -match 'RoleManagement\.Read\.Directory') -Name 'Required Graph scopes present'
    Assert-True -Condition ($scriptText -match 'Get-MgRoleManagementDirectoryRoleAssignment' -and $scriptText -match 'Get-MgRoleManagementDirectoryRoleEligibilityScheduleInstance') -Name 'Active and PIM role cmdlets present'
    Assert-True -Condition ($scriptText -notmatch 'Connect-ExchangeOnline' -and $scriptText -notmatch 'Get-ManagementRoleAssignment') -Name 'Exchange Online RBAC is out of scope'
}
Test-StaticScriptContent

if ($script:Failures.Count -gt 0) {
    Write-Host "$($script:Failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host "All $($script:Passes) checks passed." -ForegroundColor Green
exit 0
