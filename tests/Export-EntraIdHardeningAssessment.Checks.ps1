#Requires -Version 7
$ErrorActionPreference = 'Stop'

$script:ScriptPath = Join-Path $PSScriptRoot '..' 'Export-EntraIdHardeningAssessment.ps1'
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
    foreach ($name in @('Connect-GraphIfNeeded','New-GraphRequestExecutor','Get-ObjectPropertyValue','Get-GraphCollection','Get-EntraFacts','Get-EnabledCaPolicies','Test-CaPolicyProperty','Test-EntraControls')) {
        Assert-True -Condition ($functionNames -contains $name) -Name "Function exists: $name"
    }
}
Test-RequiredFunctionsExist

function Test-StaticScriptContent {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw
    Assert-True -Condition ($scriptText -match '/identity/conditionalAccess/policies') -Name 'Conditional Access endpoint is queried'
    Assert-True -Condition ($scriptText -match 'identitySecurityDefaultsEnforcementPolicy') -Name 'Security Defaults policy is queried'
    Assert-True -Condition ($scriptText -match 'authorizationPolicy') -Name 'Authorization policy is queried'
    Assert-True -Condition ($scriptText -match 'roleEligibilityScheduleInstances') -Name 'PIM eligibility endpoint is queried'
    Assert-True -Condition ($scriptText -match '62e90394-69f5-4237-9190-012177145e10') -Name 'Global Administrator role id is present'
    foreach ($controlId in 1..15 | ForEach-Object { 'AAD-{0:d3}' -f $_ }) {
        Assert-True -Condition ($scriptText -match [regex]::Escape($controlId)) -Name "Control $controlId is present"
    }
    Assert-True -Condition ($scriptText -notmatch "'PATCH'" -and $scriptText -notmatch "'POST'" -and $scriptText -notmatch "'DELETE'") -Name 'Script only issues GET requests'
    Assert-True -Condition ($scriptText -notmatch 'Import-Module\s+MSOnline' -and $scriptText -notmatch 'Get-Msol') -Name 'MSOnline cmdlets are not used'
}
Test-StaticScriptContent

function New-HardenedCaPolicies {
    $breakGlassId = 'bg-id-1'
    return @(
        @{
            displayName = 'Require MFA for all users'
            state = 'enabled'
            conditions = @{ users = @{ includeUsers = @('All'); excludeUsers = @($breakGlassId); includeRoles = @() }; clientAppTypes = @('all') }
            grantControls = @{ builtInControls = @('mfa') }
        },
        @{
            displayName = 'Phishing-resistant MFA for admins'
            state = 'enabled'
            conditions = @{ users = @{ includeUsers = @(); excludeUsers = @(); includeRoles = @('62e90394-69f5-4237-9190-012177145e10') }; clientAppTypes = @('all') }
            grantControls = @{ authenticationStrength = @{ id = '00000000-0000-0000-0000-000000000004'; displayName = 'Phishing-resistant MFA' } }
        },
        @{
            displayName = 'Block legacy authentication'
            state = 'enabled'
            conditions = @{ users = @{ includeUsers = @('All'); excludeUsers = @($breakGlassId); includeRoles = @() }; clientAppTypes = @('exchangeActiveSync','other') }
            grantControls = @{ builtInControls = @('block') }
        },
        @{
            displayName = 'Admin portals require compliant device'
            state = 'enabled'
            conditions = @{ users = @{ includeUsers = @('All'); excludeUsers = @($breakGlassId); includeRoles = @() }; applications = @{ includeApplications = @('MicrosoftAdminPortals') }; clientAppTypes = @('all') }
            grantControls = @{ builtInControls = @('compliantDevice') }
        },
        @{
            displayName = 'Block high sign-in risk'
            state = 'enabled'
            conditions = @{ users = @{ includeUsers = @('All'); excludeUsers = @($breakGlassId); includeRoles = @() }; signInRiskLevels = @('high','medium'); clientAppTypes = @('all') }
            grantControls = @{ builtInControls = @('mfa') }
        },
        @{
            displayName = 'Remediate high user risk'
            state = 'enabled'
            conditions = @{ users = @{ includeUsers = @('All'); excludeUsers = @($breakGlassId); includeRoles = @() }; userRiskLevels = @('high'); clientAppTypes = @('all') }
            grantControls = @{ builtInControls = @('passwordChange') }
        },
        @{
            displayName = 'Disabled draft policy'
            state = 'disabled'
            conditions = @{ users = @{ includeUsers = @('All'); excludeUsers = @(); includeRoles = @() }; clientAppTypes = @('all') }
            grantControls = @{ builtInControls = @('block') }
        }
    )
}

function New-HardenedGraphMock {
    return {
        param([string]$Method, [string]$Uri, $Body)
        switch -Regex ($Uri) {
            '/organization' { return @{ value = @(@{ id = 'tenant-1'; displayName = 'Contoso' }) } }
            '/identity/conditionalAccess/policies' { return @{ value = New-HardenedCaPolicies } }
            'identitySecurityDefaultsEnforcementPolicy' { return @{ isEnabled = $false } }
            'authorizationPolicy' { return @{ allowInvitesFrom = 'adminsAndGuestInviters'; defaultUserRolePermissions = @{ permissionGrantPoliciesAssigned = @() } } }
            'roleEligibilityScheduleInstances' { return @{ value = @(@{ id = 'e1' }, @{ id = 'e2' }, @{ id = 'e3' }) } }
            'roleAssignments' { return @{ value = @(@{ id = 'a1' }, @{ id = 'a2' }) } }
            'roleManagementPolicyAssignments' {
                return @{ value = @(@{ policy = @{ rules = @(
                    @{ id = 'Approval_EndUser_Assignment'; setting = @{ isApprovalRequired = $true } },
                    @{ id = 'Enablement_EndUser_Assignment'; enabledRules = @('MultiFactorAuthentication','Justification') }
                ) } }) }
            }
            'accessReviews/definitions' { return @{ value = @(@{ id = 'review-1' }) } }
            'groupSettings' {
                return @{ value = @(@{ displayName = 'Password Rule Settings'; values = @(
                    @{ name = 'EnableBannedPasswordCheck'; value = 'True' },
                    @{ name = 'BannedPasswordCheckOnPremisesMode'; value = 'Enforce' },
                    @{ name = 'LockoutThreshold'; value = '10' }
                ) }) }
            }
            '/users/' { return @{ id = 'bg-id-1'; userPrincipalName = 'bg1@contoso.com'; accountEnabled = $true } }
            default { throw "Unexpected URI: $Uri" }
        }
    }
}

function Test-HardenedTenantPasses {
    $facts = Get-EntraFacts -GraphRequest (New-HardenedGraphMock) -BreakGlassUpn @('bg1@contoso.com')
    $rows = @(Test-EntraControls -Facts $facts -BreakGlassUpn @('bg1@contoso.com') -MaxGlobalAdmins 5)
    $byId = @{}
    foreach ($row in $rows) { $byId[$row.ControlId] = $row }

    Assert-True -Condition ($rows.Count -eq 15) -Name 'Hardened tenant produces 15 control rows'
    foreach ($controlId in @('AAD-001','AAD-002','AAD-003','AAD-004','AAD-006','AAD-007','AAD-008','AAD-009','AAD-010','AAD-012','AAD-013','AAD-014','AAD-015')) {
        Assert-True -Condition ($byId[$controlId].Status -eq 'Pass') -Name "Hardened tenant passes $controlId"
    }
    Assert-True -Condition ($byId['AAD-005'].Status -eq 'Manual') -Name 'AAD-005 per-user MFA is Manual'
    Assert-True -Condition ($byId['AAD-011'].Status -eq 'Manual') -Name 'AAD-011 log retention is Manual'
    Assert-True -Condition ($byId['AAD-001'].Evidence -match 'Require MFA for all users') -Name 'AAD-001 evidence names the policy'
    Assert-True -Condition (@($rows | Where-Object Target -ne 'Contoso').Count -eq 0) -Name 'Rows target the tenant display name'
}
Test-HardenedTenantPasses

function Test-WeakTenantFails {
    $weakMock = {
        param([string]$Method, [string]$Uri, $Body)
        switch -Regex ($Uri) {
            '/organization' { return @{ value = @(@{ id = 'tenant-2'; displayName = 'Fabrikam' }) } }
            '/identity/conditionalAccess/policies' { return @{ value = @() } }
            'identitySecurityDefaultsEnforcementPolicy' { return @{ isEnabled = $true } }
            'authorizationPolicy' { return @{ allowInvitesFrom = 'everyone'; defaultUserRolePermissions = @{ permissionGrantPoliciesAssigned = @('ManagePermissionGrantsForSelf.microsoft-user-default-legacy') } } }
            'roleEligibilityScheduleInstances' { return @{ value = @() } }
            'roleAssignments' { return @{ value = @(1..12 | ForEach-Object { @{ id = "a$_" } }) } }
            'roleManagementPolicyAssignments' { return @{ value = @() } }
            'accessReviews/definitions' { return @{ value = @() } }
            'groupSettings' { return @{ value = @() } }
            default { throw "Unexpected URI: $Uri" }
        }
    }

    $facts = Get-EntraFacts -GraphRequest $weakMock
    $rows = @(Test-EntraControls -Facts $facts -MaxGlobalAdmins 5)
    $byId = @{}
    foreach ($row in $rows) { $byId[$row.ControlId] = $row }

    Assert-True -Condition ($byId['AAD-001'].Status -eq 'Warning') -Name 'Security Defaults only MFA warns AAD-001'
    Assert-True -Condition ($byId['AAD-002'].Status -eq 'Fail') -Name 'Missing phishing-resistant policy fails AAD-002'
    Assert-True -Condition ($byId['AAD-003'].Status -eq 'Warning') -Name 'Legacy auth via Security Defaults warns AAD-003'
    Assert-True -Condition ($byId['AAD-004'].Status -eq 'Warning') -Name 'Security Defaults enabled warns AAD-004'
    Assert-True -Condition ($byId['AAD-006'].Status -eq 'Fail') -Name 'No PIM eligibility fails AAD-006'
    Assert-True -Condition ($byId['AAD-007'].Status -eq 'Manual') -Name 'AAD-007 is Manual without -BreakGlassUpn'
    Assert-True -Condition ($byId['AAD-008'].Status -eq 'Warning' -and $byId['AAD-008'].Evidence -match '12 active') -Name 'Twelve Global Admins warn AAD-008'
    Assert-True -Condition ($byId['AAD-009'].Status -eq 'Fail' -and $byId['AAD-009'].Evidence -match 'everyone') -Name 'Open guest invitations fail AAD-009'
    Assert-True -Condition ($byId['AAD-010'].Status -eq 'Fail' -and $byId['AAD-010'].Evidence -match 'legacy') -Name 'Legacy consent policy fails AAD-010'
    Assert-True -Condition ($byId['AAD-013'].Status -eq 'Fail') -Name 'Missing admin portal device policy fails AAD-013'
    Assert-True -Condition ($byId['AAD-014'].Status -eq 'Fail') -Name 'Missing risk policies fail AAD-014'
    Assert-True -Condition ($byId['AAD-015'].Status -eq 'Fail') -Name 'No access reviews fail AAD-015'
}
Test-WeakTenantFails

function Test-BreakGlassExclusionDetected {
    $mock = New-HardenedGraphMock
    $facts = Get-EntraFacts -GraphRequest $mock -BreakGlassUpn @('bg1@contoso.com')
    # Remove the break-glass exclusion from the all-user MFA policy.
    $mfaPolicy = @($facts.CaPolicies) | Where-Object { $_.displayName -eq 'Require MFA for all users' }
    $mfaPolicy.conditions.users.excludeUsers = @()

    $rows = @(Test-EntraControls -Facts $facts -BreakGlassUpn @('bg1@contoso.com') -MaxGlobalAdmins 5)
    $breakGlassRow = $rows | Where-Object ControlId -eq 'AAD-007'
    Assert-True -Condition ($breakGlassRow.Status -eq 'Fail' -and $breakGlassRow.Evidence -match 'Require MFA for all users') -Name 'Missing CA exclusion fails AAD-007 and names the policy'
}
Test-BreakGlassExclusionDetected

function Test-GraphCollectionPagination {
    $script:PageCount = 0
    $pagedMock = {
        param([string]$Method, [string]$Uri, $Body)
        $script:PageCount++
        if ($Uri -match 'skiptoken') {
            return @{ value = @(@{ id = 'p2' }) }
        }
        return @{ value = @(@{ id = 'p1' }); '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/things?$skiptoken=abc' }
    }
    $items = @(Get-GraphCollection -GraphRequest $pagedMock -Uri 'https://graph.microsoft.com/v1.0/things')
    Assert-True -Condition ($items.Count -eq 2 -and $script:PageCount -eq 2) -Name 'Get-GraphCollection follows @odata.nextLink'
}
Test-GraphCollectionPagination

if ($script:Failures.Count -gt 0) {
    Write-Host "$($script:Failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host "All $($script:Passes) checks passed." -ForegroundColor Green
exit 0
