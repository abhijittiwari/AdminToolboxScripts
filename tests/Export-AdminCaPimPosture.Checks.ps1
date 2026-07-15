#Requires -Version 7
$ErrorActionPreference = 'Stop'

$script:ScriptPath = Join-Path $PSScriptRoot '..' 'Export-AdminCaPimPosture.ps1'
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

function New-TestCaPolicy {
    param(
        [string]$Id = 'policy-1',
        [string]$DisplayName = 'Test Policy',
        [string]$State = 'enabled',
        [string[]]$IncludeUsers = @(),
        [string[]]$ExcludeUsers = @(),
        [string[]]$IncludeGroups = @(),
        [string[]]$ExcludeGroups = @(),
        [string[]]$IncludeRoles = @(),
        [string[]]$ExcludeRoles = @(),
        [string[]]$BuiltInControls = @(),
        [string]$AuthenticationStrengthName = '',
        [string]$AuthenticationStrengthId = '',
        [string[]]$IncludeApplications = @('All')
    )

    $authenticationStrength = $null
    if (-not [string]::IsNullOrWhiteSpace($AuthenticationStrengthId)) {
        $authenticationStrength = [pscustomobject]@{ Id = $AuthenticationStrengthId; DisplayName = $AuthenticationStrengthName }
    }

    return [pscustomobject]@{
        Id            = $Id
        DisplayName   = $DisplayName
        State         = $State
        Conditions    = [pscustomobject]@{
            Users        = [pscustomobject]@{
                IncludeUsers  = $IncludeUsers
                ExcludeUsers  = $ExcludeUsers
                IncludeGroups = $IncludeGroups
                ExcludeGroups = $ExcludeGroups
                IncludeRoles  = $IncludeRoles
                ExcludeRoles  = $ExcludeRoles
            }
            Applications = [pscustomobject]@{ IncludeApplications = $IncludeApplications }
        }
        GrantControls = [pscustomobject]@{
            BuiltInControls        = $BuiltInControls
            AuthenticationStrength = $authenticationStrength
        }
    }
}

function Test-RequiredFunctionsExist {
    $ast = Get-ScriptAst
    $functionNames = @($ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true) | ForEach-Object Name)
    foreach ($name in @('Join-ReportValues','Get-ObjectPropertyValue','ConvertTo-StringArray','Test-AnyValueIntersection','ConvertTo-CaUsersCondition','Test-CaPolicyUserScope','Get-CaPolicyGrantControlSummary','Test-CaPolicyRequiresMfa','Get-CaPolicyIncludedApplications','New-AdminCaPolicyRow','New-AdminPostureRow','Add-UserRoleEntry','New-RoleEntry','Resolve-RolePrincipalEntries','Get-AdminRoleIndex','Get-UserTransitiveGroupIds','New-PostureSummaryMetrics','Test-RequiredCsvColumns')) {
        Assert-True -Condition ($functionNames -contains $name) -Name "Function exists: $name"
    }
}
Test-RequiredFunctionsExist

function Test-StaticScriptContent {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw
    Assert-True -Condition ($scriptText -match 'Policy\.Read\.All' -and $scriptText -match 'RoleManagement\.Read\.Directory' -and $scriptText -match 'Directory\.Read\.All') -Name 'Graph scopes are present'
    Assert-True -Condition ($scriptText -match 'Get-MgIdentityConditionalAccessPolicy') -Name 'Conditional Access policies are read'
    Assert-True -Condition ($scriptText -match 'Get-MgRoleManagementDirectoryRoleEligibilityScheduleInstance') -Name 'PIM eligibility schedules are read'
    Assert-True -Condition ($scriptText -notmatch 'Import-Module\s+MSOnline' -and $scriptText -notmatch 'Get-Msol' -and $scriptText -notmatch 'Connect-Msol') -Name 'MSOnline cmdlets are not used'
    Assert-True -Condition ($scriptText -notmatch 'New-MgIdentityConditionalAccessPolicy' -and $scriptText -notmatch 'Update-Mg' -and $scriptText -notmatch 'Remove-Mg') -Name 'Script is read-only'
    Assert-True -Condition ($scriptText -match 'ImportExcel' -and $scriptText -match 'Export-Excel') -Name 'Workbook is built with ImportExcel'
    Assert-True -Condition ($scriptText -match "WorksheetName 'Summary'" -and $scriptText -match "WorksheetName 'AdminPosture'" -and $scriptText -match "WorksheetName 'CaPolicyMap'") -Name 'Workbook contains Summary, AdminPosture, and CaPolicyMap sheets'
    Assert-True -Condition ($scriptText -notmatch 'Export-Csv') -Name 'Output is the combined workbook, not CSVs'
}
Test-StaticScriptContent

function Test-ConvertToCaUsersCondition {
    $policy = New-TestCaPolicy -IncludeUsers @('All') -ExcludeUsers @('user-9') -IncludeRoles @('role-1')
    $condition = ConvertTo-CaUsersCondition -Policy $policy

    $hashtablePolicy = @{ Conditions = @{ Users = @{ IncludeGroups = @('group-1') } } }
    $hashtableCondition = ConvertTo-CaUsersCondition -Policy $hashtablePolicy

    Assert-True -Condition ($condition.IncludeUsers -contains 'All' -and $condition.ExcludeUsers -contains 'user-9' -and $condition.IncludeRoles -contains 'role-1') -Name 'Users condition is extracted from object policy'
    Assert-True -Condition ($condition.IncludeGroups.Count -eq 0) -Name 'Missing condition arrays default to empty'
    Assert-True -Condition ($hashtableCondition.IncludeGroups -contains 'group-1') -Name 'Users condition is extracted from hashtable policy'
}
Test-ConvertToCaUsersCondition

function Test-CaPolicyUserScopeStatuses {
    $allUsers = ConvertTo-CaUsersCondition -Policy (New-TestCaPolicy -IncludeUsers @('All'))
    $directUser = ConvertTo-CaUsersCondition -Policy (New-TestCaPolicy -IncludeUsers @('user-1'))
    $byGroup = ConvertTo-CaUsersCondition -Policy (New-TestCaPolicy -IncludeGroups @('group-1'))
    $byRole = ConvertTo-CaUsersCondition -Policy (New-TestCaPolicy -IncludeRoles @('role-ga'))
    $excludedUser = ConvertTo-CaUsersCondition -Policy (New-TestCaPolicy -IncludeUsers @('All') -ExcludeUsers @('user-1'))
    $excludedGroup = ConvertTo-CaUsersCondition -Policy (New-TestCaPolicy -IncludeUsers @('All') -ExcludeGroups @('group-1'))
    $unrelated = ConvertTo-CaUsersCondition -Policy (New-TestCaPolicy -IncludeUsers @('user-2'))

    $groupIds = @('group-1')
    $activeRoles = @('role-ga')
    $eligibleRoles = @('role-exchange')

    Assert-True -Condition ((Test-CaPolicyUserScope -UsersCondition $allUsers -UserId 'user-1' -GroupIds @() -ActiveRoleTemplateIds @() -EligibleRoleTemplateIds @()) -eq 'Applies') -Name 'All-users policy applies'
    Assert-True -Condition ((Test-CaPolicyUserScope -UsersCondition $directUser -UserId 'user-1' -GroupIds @() -ActiveRoleTemplateIds @() -EligibleRoleTemplateIds @()) -eq 'Applies') -Name 'Direct user inclusion applies'
    Assert-True -Condition ((Test-CaPolicyUserScope -UsersCondition $byGroup -UserId 'user-1' -GroupIds $groupIds -ActiveRoleTemplateIds @() -EligibleRoleTemplateIds @()) -eq 'Applies') -Name 'Group inclusion applies'
    Assert-True -Condition ((Test-CaPolicyUserScope -UsersCondition $byRole -UserId 'user-1' -GroupIds @() -ActiveRoleTemplateIds $activeRoles -EligibleRoleTemplateIds @()) -eq 'Applies') -Name 'Active role inclusion applies'
    Assert-True -Condition ((Test-CaPolicyUserScope -UsersCondition $excludedUser -UserId 'user-1' -GroupIds @() -ActiveRoleTemplateIds @() -EligibleRoleTemplateIds @()) -eq 'Excluded') -Name 'User exclusion wins over all-users inclusion'
    Assert-True -Condition ((Test-CaPolicyUserScope -UsersCondition $excludedGroup -UserId 'user-1' -GroupIds $groupIds -ActiveRoleTemplateIds @() -EligibleRoleTemplateIds @()) -eq 'Excluded') -Name 'Group exclusion wins over all-users inclusion'
    Assert-True -Condition ((Test-CaPolicyUserScope -UsersCondition $unrelated -UserId 'user-1' -GroupIds @() -ActiveRoleTemplateIds @() -EligibleRoleTemplateIds @()) -eq 'NotInScope') -Name 'Unrelated policy is not in scope'
}
Test-CaPolicyUserScopeStatuses

function Test-CaPolicyUserScopePimActivation {
    $roleScoped = ConvertTo-CaUsersCondition -Policy (New-TestCaPolicy -IncludeRoles @('role-exchange'))

    $status = Test-CaPolicyUserScope -UsersCondition $roleScoped -UserId 'user-1' -GroupIds @() -ActiveRoleTemplateIds @('role-ga') -EligibleRoleTemplateIds @('role-exchange')

    Assert-True -Condition ($status -eq 'AppliesOnPimActivation') -Name 'Eligible-only role maps to AppliesOnPimActivation'
}
Test-CaPolicyUserScopePimActivation

function Test-GrantControlSummaryAndMfa {
    $mfaPolicy = New-TestCaPolicy -BuiltInControls @('mfa', 'compliantDevice')
    $strengthPolicy = New-TestCaPolicy -AuthenticationStrengthId 'strength-1' -AuthenticationStrengthName 'Phishing-resistant MFA'
    $blockPolicy = New-TestCaPolicy -BuiltInControls @('block')
    $noGrantPolicy = [pscustomobject]@{ Id = 'p'; GrantControls = $null }

    Assert-True -Condition ((Get-CaPolicyGrantControlSummary -Policy $mfaPolicy) -eq 'compliantDevice; mfa') -Name 'Grant control summary joins built-in controls'
    Assert-True -Condition ((Get-CaPolicyGrantControlSummary -Policy $strengthPolicy) -eq 'AuthenticationStrength: Phishing-resistant MFA') -Name 'Grant control summary includes authentication strength'
    Assert-True -Condition (Test-CaPolicyRequiresMfa -Policy $mfaPolicy) -Name 'mfa built-in control counts as requiring MFA'
    Assert-True -Condition (Test-CaPolicyRequiresMfa -Policy $strengthPolicy) -Name 'Authentication strength counts as requiring MFA'
    Assert-True -Condition (-not (Test-CaPolicyRequiresMfa -Policy $blockPolicy)) -Name 'Block-only policy does not count as requiring MFA'
    Assert-True -Condition (-not (Test-CaPolicyRequiresMfa -Policy $noGrantPolicy)) -Name 'Missing grant controls do not count as requiring MFA'
}
Test-GrantControlSummaryAndMfa

function Test-AdminCaPolicyRow {
    $user = [pscustomobject]@{ Id = 'user-1'; UserPrincipalName = 'admin@contoso.com'; DisplayName = 'Admin One'; AccountEnabled = $true }
    $policy = New-TestCaPolicy -Id 'policy-1' -DisplayName 'Require MFA for admins' -BuiltInControls @('mfa')

    $row = New-AdminCaPolicyRow -User $user -Policy $policy -ApplicabilityStatus 'Applies'

    Assert-True -Condition ($row.AdminUserPrincipalName -eq 'admin@contoso.com' -and $row.PolicyName -eq 'Require MFA for admins' -and $row.PolicyState -eq 'enabled') -Name 'Policy row carries admin and policy details'
    Assert-True -Condition ($row.ApplicabilityStatus -eq 'Applies' -and $row.RequiresMfa -eq $true -and $row.IncludedApplications -eq 'All') -Name 'Policy row carries applicability, MFA flag, and apps'
}
Test-AdminCaPolicyRow

function Test-AdminPostureRow {
    $user = [pscustomobject]@{ Id = 'user-1'; UserPrincipalName = 'admin@contoso.com'; DisplayName = 'Admin One'; AccountEnabled = $true }
    $activeEntries = @((New-RoleEntry -RoleName 'Global Administrator' -RoleTemplateId 'role-ga'))
    $eligibleEntries = @((New-RoleEntry -RoleName 'Exchange Administrator' -RoleTemplateId 'role-exchange' -ViaGroup 'PIM Group'))

    $policyRows = @(
        (New-AdminCaPolicyRow -User $user -Policy (New-TestCaPolicy -Id 'p1' -DisplayName 'Enabled MFA policy' -State 'enabled' -BuiltInControls @('mfa')) -ApplicabilityStatus 'Applies'),
        (New-AdminCaPolicyRow -User $user -Policy (New-TestCaPolicy -Id 'p2' -DisplayName 'Report-only policy' -State 'enabledForReportingButNotEnforced') -ApplicabilityStatus 'Applies'),
        (New-AdminCaPolicyRow -User $user -Policy (New-TestCaPolicy -Id 'p3' -DisplayName 'Activation policy') -ApplicabilityStatus 'AppliesOnPimActivation'),
        (New-AdminCaPolicyRow -User $user -Policy (New-TestCaPolicy -Id 'p4' -DisplayName 'Excluded policy') -ApplicabilityStatus 'Excluded')
    )

    $row = New-AdminPostureRow -User $user -ActiveRoleEntries $activeEntries -EligibleRoleEntries $eligibleEntries -PolicyRows $policyRows
    $emptyRow = New-AdminPostureRow -User $user -ActiveRoleEntries @() -EligibleRoleEntries @() -PolicyRows @()

    Assert-True -Condition ($row.ActiveDirectoryRoles -eq 'Global Administrator') -Name 'Posture row lists active roles'
    Assert-True -Condition ($row.PimEligibleDirectoryRoles -eq 'Exchange Administrator via group: PIM Group') -Name 'Posture row lists PIM-eligible roles with group attribution'
    Assert-True -Condition ($row.CaEnabledPoliciesApplying -eq 'Enabled MFA policy') -Name 'Posture row lists enabled applying policies'
    Assert-True -Condition ($row.CaReportOnlyPoliciesApplying -eq 'Report-only policy') -Name 'Posture row lists report-only applying policies'
    Assert-True -Condition ($row.CaPoliciesApplyOnPimActivation -eq 'Activation policy') -Name 'Posture row lists activation-only policies'
    Assert-True -Condition ($row.CaPoliciesExcludedFrom -eq 'Excluded policy') -Name 'Posture row lists explicit exclusions'
    Assert-True -Condition ($row.MfaRequiredByEnabledCa -eq $true) -Name 'MFA flag is set when an enabled applying policy requires MFA'
    Assert-True -Condition ($emptyRow.MfaRequiredByEnabledCa -eq $false -and $emptyRow.CaEnabledPoliciesApplying -eq '') -Name 'Empty posture row reports no CA coverage'
}
Test-AdminPostureRow

function Test-RoleEntryLabels {
    $plain = New-RoleEntry -RoleName 'Global Reader' -RoleTemplateId 'role-gr'
    $scoped = New-RoleEntry -RoleName 'User Administrator' -RoleTemplateId 'role-ua' -DirectoryScopeId '/administrativeUnits/au-1'

    Assert-True -Condition ($plain.DisplayLabel -eq 'Global Reader') -Name 'Tenant-wide role label has no scope suffix'
    Assert-True -Condition ($scoped.DisplayLabel -eq 'User Administrator (scope: /administrativeUnits/au-1)') -Name 'AU-scoped role label includes the scope'
}
Test-RoleEntryLabels

function Test-PostureSummaryMetrics {
    $user = [pscustomobject]@{ Id = 'user-1'; UserPrincipalName = 'admin@contoso.com'; DisplayName = 'Admin One'; AccountEnabled = $true }
    $coveredRow = New-AdminPostureRow -User $user -ActiveRoleEntries @() -EligibleRoleEntries @((New-RoleEntry -RoleName 'Exchange Administrator' -RoleTemplateId 'role-exchange')) -PolicyRows @(
        (New-AdminCaPolicyRow -User $user -Policy (New-TestCaPolicy -DisplayName 'MFA policy' -BuiltInControls @('mfa')) -ApplicabilityStatus 'Applies')
    )
    $uncoveredRow = New-AdminPostureRow -User $user -ActiveRoleEntries @() -EligibleRoleEntries @() -PolicyRows @()
    $errorRow = New-AdminPostureRow -User $user -ActiveRoleEntries @() -EligibleRoleEntries @() -PolicyRows @() -ErrorText 'User lookup failed'

    $metrics = @(New-PostureSummaryMetrics -PostureRows @($coveredRow, $uncoveredRow, $errorRow) -CaPolicyCount 5)
    $metricMap = @{}
    foreach ($metric in $metrics) { $metricMap[$metric.Metric] = $metric.Count }

    Assert-True -Condition ($metricMap['InScopeAdmins'] -eq 3 -and $metricMap['ConditionalAccessPolicies'] -eq 5) -Name 'Metrics report admin and policy counts'
    Assert-True -Condition ($metricMap['AdminsWithPimEligibleRoles'] -eq 1) -Name 'Metrics count admins with PIM-eligible roles'
    Assert-True -Condition ($metricMap['AdminsWithNoEnabledCaPolicyApplying'] -eq 1) -Name 'Error rows are not counted as CA coverage gaps'
    Assert-True -Condition ($metricMap['AdminsWithoutMfaRequiringEnabledCa'] -eq 1) -Name 'Metrics count admins without MFA-requiring CA'
    Assert-True -Condition ($metricMap['AdminsWithLookupErrors'] -eq 1) -Name 'Metrics count lookup errors'
}
Test-PostureSummaryMetrics

function Test-AddUserRoleEntryDeduplicates {
    $roleMap = @{}
    $entry = New-RoleEntry -RoleName 'Global Administrator' -RoleTemplateId 'role-ga'

    Add-UserRoleEntry -RoleMap $roleMap -UserId 'user-1' -Entry $entry
    Add-UserRoleEntry -RoleMap $roleMap -UserId 'user-1' -Entry $entry

    Assert-True -Condition ($roleMap['user-1'].Count -eq 1) -Name 'Duplicate role entries are not added'
}
Test-AddUserRoleEntryDeduplicates

if ($script:Failures.Count -gt 0) {
    Write-Host "$($script:Failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host "All $($script:Passes) checks passed." -ForegroundColor Green
exit 0
