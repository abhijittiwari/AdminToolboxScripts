<#
.SYNOPSIS
    Maps Conditional Access policy applicability and PIM-eligible directory roles
    for in-scope Entra ID admins.

.DESCRIPTION
    Finds in-scope admins (users with active or PIM-eligible Entra ID directory
    role assignments, including assignments through role-assignable groups), or
    takes an explicit list of admin UPNs from a CSV. For each admin the script
    evaluates every Conditional Access policy's user-scope conditions - included
    and excluded users, groups, and directory roles - and reports whether the
    policy applies, is explicitly excluded, or only applies once a PIM-eligible
    role is activated.

    The script is read-only. It writes one Excel workbook with three worksheets:
    - Summary: posture metrics (in-scope admins, PIM coverage, CA gaps).
    - AdminPosture: one row per admin with active roles, PIM-eligible roles,
      applying enabled/report-only CA policies, exclusions, and whether any
      enabled applying policy requires MFA.
    - CaPolicyMap: one row per admin per relevant CA policy, including the
      policy state, applicability status, grant controls, and included apps.

    CA applicability is evaluated against user-scope conditions only. Other
    policy conditions (applications, platforms, locations, client apps, risk)
    are reported for context but not evaluated.

.PARAMETER InputCsv
    Optional CSV path with a UserPrincipalName column that defines the in-scope
    admins. When omitted, admins are discovered from active and PIM-eligible
    directory role assignments.

.PARAMETER OutputPath
    Excel workbook output path. Defaults to .\AdminCaPimPosture_yyyyMMdd_HHmmss.xlsx.

.EXAMPLE
    .\Export-AdminCaPimPosture.ps1

.EXAMPLE
    .\Export-AdminCaPimPosture.ps1 -InputCsv .\InScopeAdmins.csv

.NOTES
    Requires: ImportExcel (Install-Module ImportExcel -Scope CurrentUser). Excel
    itself is not required.
#>

[CmdletBinding()]
param(
    [string]$InputCsv,
    [string]$OutputPath = ".\AdminCaPimPosture_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
)

$ErrorActionPreference = 'Stop'

if (-not (Get-Module -ListAvailable ImportExcel)) {
    throw 'The ImportExcel module is required to build the workbook. Install it with: Install-Module ImportExcel -Scope CurrentUser'
}
Import-Module ImportExcel -ErrorAction Stop

function Join-ReportValues {
    param([AllowEmptyCollection()] [object[]]$Values)

    $cleanValues = @($Values | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | ForEach-Object { ([string]$_).Trim() } | Sort-Object -Unique)
    return ($cleanValues -join '; ')
}

function Get-ObjectPropertyValue {
    param(
        $InputObject,
        [Parameter(Mandatory = $true)] [string]$PropertyName
    )

    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [hashtable]) { return $InputObject[$PropertyName] }

    $property = $InputObject.PSObject.Properties[$PropertyName]
    if ($null -eq $property) { return $null }
    return $property.Value
}

function ConvertTo-StringArray {
    param($Value)

    return @(@($Value) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | ForEach-Object { [string]$_ })
}

function Test-AnyValueIntersection {
    param(
        [AllowEmptyCollection()] [string[]]$First,
        [AllowEmptyCollection()] [string[]]$Second
    )

    foreach ($value in @($First)) {
        if (@($Second) -contains $value) { return $true }
    }
    return $false
}

function ConvertTo-CaUsersCondition {
    param([Parameter(Mandatory = $true)] $Policy)

    $conditions = Get-ObjectPropertyValue -InputObject $Policy -PropertyName 'Conditions'
    if ($null -eq $conditions) { $conditions = Get-ObjectPropertyValue -InputObject $Policy -PropertyName 'conditions' }
    $users = Get-ObjectPropertyValue -InputObject $conditions -PropertyName 'Users'
    if ($null -eq $users) { $users = Get-ObjectPropertyValue -InputObject $conditions -PropertyName 'users' }

    return [pscustomobject]@{
        IncludeUsers  = ConvertTo-StringArray -Value (Get-ObjectPropertyValue -InputObject $users -PropertyName 'IncludeUsers')
        ExcludeUsers  = ConvertTo-StringArray -Value (Get-ObjectPropertyValue -InputObject $users -PropertyName 'ExcludeUsers')
        IncludeGroups = ConvertTo-StringArray -Value (Get-ObjectPropertyValue -InputObject $users -PropertyName 'IncludeGroups')
        ExcludeGroups = ConvertTo-StringArray -Value (Get-ObjectPropertyValue -InputObject $users -PropertyName 'ExcludeGroups')
        IncludeRoles  = ConvertTo-StringArray -Value (Get-ObjectPropertyValue -InputObject $users -PropertyName 'IncludeRoles')
        ExcludeRoles  = ConvertTo-StringArray -Value (Get-ObjectPropertyValue -InputObject $users -PropertyName 'ExcludeRoles')
    }
}

function Test-CaPolicyUserScope {
    param(
        [Parameter(Mandatory = $true)] $UsersCondition,
        [Parameter(Mandatory = $true)] [string]$UserId,
        [AllowEmptyCollection()] [string[]]$GroupIds,
        [AllowEmptyCollection()] [string[]]$ActiveRoleTemplateIds,
        [AllowEmptyCollection()] [string[]]$EligibleRoleTemplateIds
    )

    $isExcluded = ($UsersCondition.ExcludeUsers -contains $UserId) -or
        (Test-AnyValueIntersection -First $UsersCondition.ExcludeGroups -Second $GroupIds) -or
        (Test-AnyValueIntersection -First $UsersCondition.ExcludeRoles -Second $ActiveRoleTemplateIds)

    $isIncluded = ($UsersCondition.IncludeUsers -contains 'All') -or
        ($UsersCondition.IncludeUsers -contains $UserId) -or
        (Test-AnyValueIntersection -First $UsersCondition.IncludeGroups -Second $GroupIds) -or
        (Test-AnyValueIntersection -First $UsersCondition.IncludeRoles -Second $ActiveRoleTemplateIds)

    $isIncludedOnActivation = (-not $isIncluded) -and
        (Test-AnyValueIntersection -First $UsersCondition.IncludeRoles -Second $EligibleRoleTemplateIds)

    if ($isExcluded -and ($isIncluded -or $isIncludedOnActivation)) { return 'Excluded' }
    if ($isIncluded) { return 'Applies' }
    if ($isIncludedOnActivation) { return 'AppliesOnPimActivation' }
    return 'NotInScope'
}

function Get-CaPolicyGrantControlSummary {
    param([Parameter(Mandatory = $true)] $Policy)

    $grantControls = Get-ObjectPropertyValue -InputObject $Policy -PropertyName 'GrantControls'
    if ($null -eq $grantControls) { return '' }

    $parts = [System.Collections.Generic.List[string]]::new()
    foreach ($control in (ConvertTo-StringArray -Value (Get-ObjectPropertyValue -InputObject $grantControls -PropertyName 'BuiltInControls'))) {
        $parts.Add($control) | Out-Null
    }

    $authenticationStrength = Get-ObjectPropertyValue -InputObject $grantControls -PropertyName 'AuthenticationStrength'
    $strengthName = [string](Get-ObjectPropertyValue -InputObject $authenticationStrength -PropertyName 'DisplayName')
    if (-not [string]::IsNullOrWhiteSpace($strengthName)) {
        $parts.Add("AuthenticationStrength: $strengthName") | Out-Null
    }

    return Join-ReportValues -Values $parts.ToArray()
}

function Test-CaPolicyRequiresMfa {
    param([Parameter(Mandatory = $true)] $Policy)

    $grantControls = Get-ObjectPropertyValue -InputObject $Policy -PropertyName 'GrantControls'
    if ($null -eq $grantControls) { return $false }

    $builtInControls = ConvertTo-StringArray -Value (Get-ObjectPropertyValue -InputObject $grantControls -PropertyName 'BuiltInControls')
    if ($builtInControls -contains 'mfa') { return $true }

    $authenticationStrength = Get-ObjectPropertyValue -InputObject $grantControls -PropertyName 'AuthenticationStrength'
    $strengthId = [string](Get-ObjectPropertyValue -InputObject $authenticationStrength -PropertyName 'Id')
    return (-not [string]::IsNullOrWhiteSpace($strengthId))
}

function Get-CaPolicyIncludedApplications {
    param([Parameter(Mandatory = $true)] $Policy)

    $conditions = Get-ObjectPropertyValue -InputObject $Policy -PropertyName 'Conditions'
    $applications = Get-ObjectPropertyValue -InputObject $conditions -PropertyName 'Applications'
    return Join-ReportValues -Values (ConvertTo-StringArray -Value (Get-ObjectPropertyValue -InputObject $applications -PropertyName 'IncludeApplications'))
}

function New-AdminCaPolicyRow {
    param(
        [Parameter(Mandatory = $true)] $User,
        [Parameter(Mandatory = $true)] $Policy,
        [Parameter(Mandatory = $true)] [string]$ApplicabilityStatus
    )

    return [pscustomobject]@{
        AdminUserPrincipalName = [string]$User.UserPrincipalName
        AdminDisplayName       = [string]$User.DisplayName
        PolicyName             = [string](Get-ObjectPropertyValue -InputObject $Policy -PropertyName 'DisplayName')
        PolicyId               = [string](Get-ObjectPropertyValue -InputObject $Policy -PropertyName 'Id')
        PolicyState            = [string](Get-ObjectPropertyValue -InputObject $Policy -PropertyName 'State')
        ApplicabilityStatus    = $ApplicabilityStatus
        RequiresMfa            = Test-CaPolicyRequiresMfa -Policy $Policy
        GrantControls          = Get-CaPolicyGrantControlSummary -Policy $Policy
        IncludedApplications   = Get-CaPolicyIncludedApplications -Policy $Policy
    }
}

function New-AdminPostureRow {
    param(
        [Parameter(Mandatory = $true)] $User,
        [AllowEmptyCollection()] [object[]]$ActiveRoleEntries,
        [AllowEmptyCollection()] [object[]]$EligibleRoleEntries,
        [AllowEmptyCollection()] [object[]]$PolicyRows,
        [AllowEmptyString()] [string]$ErrorText = ''
    )

    $applyingRows = @($PolicyRows | Where-Object ApplicabilityStatus -eq 'Applies')
    $enabledApplying = @($applyingRows | Where-Object PolicyState -eq 'enabled')

    return [pscustomobject]@{
        UserPrincipalName              = [string]$User.UserPrincipalName
        DisplayName                    = [string]$User.DisplayName
        UserId                         = [string]$User.Id
        AccountEnabled                 = $User.AccountEnabled
        ActiveDirectoryRoles           = Join-ReportValues -Values @($ActiveRoleEntries | ForEach-Object DisplayLabel)
        PimEligibleDirectoryRoles      = Join-ReportValues -Values @($EligibleRoleEntries | ForEach-Object DisplayLabel)
        CaEnabledPoliciesApplying      = Join-ReportValues -Values @($enabledApplying | ForEach-Object PolicyName)
        CaReportOnlyPoliciesApplying   = Join-ReportValues -Values @($applyingRows | Where-Object PolicyState -eq 'enabledForReportingButNotEnforced' | ForEach-Object PolicyName)
        CaPoliciesApplyOnPimActivation = Join-ReportValues -Values @($PolicyRows | Where-Object ApplicabilityStatus -eq 'AppliesOnPimActivation' | ForEach-Object PolicyName)
        CaPoliciesExcludedFrom         = Join-ReportValues -Values @($PolicyRows | Where-Object ApplicabilityStatus -eq 'Excluded' | ForEach-Object PolicyName)
        MfaRequiredByEnabledCa         = (@($enabledApplying | Where-Object RequiresMfa -eq $true).Count -gt 0)
        Error                          = $ErrorText
    }
}

function Add-UserRoleEntry {
    param(
        [Parameter(Mandatory = $true)] [hashtable]$RoleMap,
        [Parameter(Mandatory = $true)] [string]$UserId,
        [Parameter(Mandatory = $true)] $Entry
    )

    if (-not $RoleMap.ContainsKey($UserId)) {
        $RoleMap[$UserId] = [System.Collections.Generic.List[object]]::new()
    }
    if (-not @($RoleMap[$UserId] | Where-Object { $_.DisplayLabel -eq $Entry.DisplayLabel }).Count) {
        $RoleMap[$UserId].Add($Entry) | Out-Null
    }
}

function New-RoleEntry {
    param(
        [Parameter(Mandatory = $true)] [string]$RoleName,
        [Parameter(Mandatory = $true)] [string]$RoleTemplateId,
        [AllowEmptyString()] [string]$DirectoryScopeId = '/',
        [AllowEmptyString()] [string]$ViaGroup = ''
    )

    $label = $RoleName
    if (-not [string]::IsNullOrWhiteSpace($ViaGroup)) { $label = "$label via group: $ViaGroup" }
    if (-not [string]::IsNullOrWhiteSpace($DirectoryScopeId) -and $DirectoryScopeId -ne '/') { $label = "$label (scope: $DirectoryScopeId)" }

    return [pscustomobject]@{
        RoleName       = $RoleName
        RoleTemplateId = $RoleTemplateId
        DisplayLabel   = $label
    }
}

function Resolve-RolePrincipalEntries {
    param(
        [Parameter(Mandatory = $true)] [string]$PrincipalId,
        [Parameter(Mandatory = $true)] [string]$RoleName,
        [Parameter(Mandatory = $true)] [string]$RoleTemplateId,
        [AllowEmptyString()] [string]$DirectoryScopeId,
        [Parameter(Mandatory = $true)] [hashtable]$RoleMap
    )

    $principal = Get-MgDirectoryObjectById -Ids $PrincipalId -ErrorAction SilentlyContinue
    if (-not $principal) { return }

    switch ($principal.AdditionalProperties.'@odata.type') {
        '#microsoft.graph.user' {
            Add-UserRoleEntry -RoleMap $RoleMap -UserId $PrincipalId -Entry (New-RoleEntry -RoleName $RoleName -RoleTemplateId $RoleTemplateId -DirectoryScopeId $DirectoryScopeId)
        }
        '#microsoft.graph.group' {
            $groupName = $principal.AdditionalProperties.displayName
            try {
                Get-MgGroupTransitiveMember -GroupId $PrincipalId -All -ErrorAction Stop |
                    Where-Object { $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.user' } |
                    ForEach-Object {
                        Add-UserRoleEntry -RoleMap $RoleMap -UserId $_.Id -Entry (New-RoleEntry -RoleName $RoleName -RoleTemplateId $RoleTemplateId -DirectoryScopeId $DirectoryScopeId -ViaGroup $groupName)
                    }
            }
            catch {
                Write-Warning "Could not expand role group '$groupName' ($PrincipalId): $($_.Exception.Message)"
            }
        }
    }
}

function Get-AdminRoleIndex {
    $roleDefinitions = @{}
    Get-MgRoleManagementDirectoryRoleDefinition -All -ErrorAction Stop | ForEach-Object {
        $roleDefinitions[$_.Id] = $_.DisplayName
    }

    $activeRoles = @{}
    $eligibleRoles = @{}

    Write-Host 'Collecting active Entra ID directory role assignments...' -ForegroundColor Cyan
    foreach ($assignment in @(Get-MgRoleManagementDirectoryRoleAssignment -All -ErrorAction Stop)) {
        $roleName = $roleDefinitions[$assignment.RoleDefinitionId]
        if (-not [string]::IsNullOrWhiteSpace($roleName)) {
            Resolve-RolePrincipalEntries -PrincipalId $assignment.PrincipalId -RoleName $roleName -RoleTemplateId $assignment.RoleDefinitionId -DirectoryScopeId ([string]$assignment.DirectoryScopeId) -RoleMap $activeRoles
        }
    }

    Write-Host 'Collecting PIM-eligible Entra ID directory role assignments...' -ForegroundColor Cyan
    try {
        foreach ($assignment in @(Get-MgRoleManagementDirectoryRoleEligibilityScheduleInstance -All -ErrorAction Stop)) {
            $roleName = $roleDefinitions[$assignment.RoleDefinitionId]
            if (-not [string]::IsNullOrWhiteSpace($roleName)) {
                Resolve-RolePrincipalEntries -PrincipalId $assignment.PrincipalId -RoleName $roleName -RoleTemplateId $assignment.RoleDefinitionId -DirectoryScopeId ([string]$assignment.DirectoryScopeId) -RoleMap $eligibleRoles
            }
        }
    }
    catch {
        Write-Warning "Could not read PIM eligibility schedules: $($_.Exception.Message)"
    }

    return @{ Active = $activeRoles; Eligible = $eligibleRoles }
}

function Get-UserTransitiveGroupIds {
    param([Parameter(Mandatory = $true)] [string]$UserId)

    return @(Get-MgUserTransitiveMemberOf -UserId $UserId -All -ErrorAction Stop |
        Where-Object { $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group' } |
        ForEach-Object Id)
}

function New-PostureSummaryMetrics {
    param(
        [AllowEmptyCollection()] [object[]]$PostureRows,
        [Parameter(Mandatory = $true)] [int]$CaPolicyCount
    )

    $validRows = @($PostureRows | Where-Object { [string]::IsNullOrWhiteSpace($_.Error) })

    return @(
        [pscustomobject]@{ Metric = 'InScopeAdmins'; Count = @($PostureRows).Count }
        [pscustomobject]@{ Metric = 'ConditionalAccessPolicies'; Count = $CaPolicyCount }
        [pscustomobject]@{ Metric = 'AdminsWithPimEligibleRoles'; Count = @($PostureRows | Where-Object { -not [string]::IsNullOrWhiteSpace($_.PimEligibleDirectoryRoles) }).Count }
        [pscustomobject]@{ Metric = 'AdminsWithNoEnabledCaPolicyApplying'; Count = @($validRows | Where-Object { [string]::IsNullOrWhiteSpace($_.CaEnabledPoliciesApplying) }).Count }
        [pscustomobject]@{ Metric = 'AdminsWithoutMfaRequiringEnabledCa'; Count = @($validRows | Where-Object { $_.MfaRequiredByEnabledCa -ne $true }).Count }
        [pscustomobject]@{ Metric = 'AdminsExcludedFromAnyCaPolicy'; Count = @($PostureRows | Where-Object { -not [string]::IsNullOrWhiteSpace($_.CaPoliciesExcludedFrom) }).Count }
        [pscustomobject]@{ Metric = 'AdminsWithLookupErrors'; Count = @($PostureRows | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Error) }).Count }
    )
}

function Test-RequiredCsvColumns {
    param(
        [Parameter(Mandatory = $true)] [object[]]$Rows,
        [Parameter(Mandatory = $true)] [string[]]$RequiredColumns
    )

    if ($Rows.Count -eq 0) {
        throw 'Input CSV contains no data rows.'
    }

    $actualColumns = @($Rows[0].PSObject.Properties.Name)
    $missingColumns = @($RequiredColumns | Where-Object { $actualColumns -notcontains $_ })
    if ($missingColumns.Count -gt 0) {
        throw "Input CSV is missing required column(s): $($missingColumns -join ', ')"
    }
}

Write-Host 'Connecting to Microsoft Graph...' -ForegroundColor Cyan
Connect-MgGraph -Scopes @('Policy.Read.All', 'RoleManagement.Read.Directory', 'Directory.Read.All') -NoWelcome

$roleIndex = Get-AdminRoleIndex

$adminUserIds = [System.Collections.Generic.List[string]]::new()
if (-not [string]::IsNullOrWhiteSpace($InputCsv)) {
    if (-not (Test-Path -Path $InputCsv -PathType Leaf)) {
        throw "Input CSV not found: $InputCsv"
    }
    $inputRows = @(Import-Csv -Path $InputCsv)
    Test-RequiredCsvColumns -Rows $inputRows -RequiredColumns @('UserPrincipalName')

    Write-Host "Resolving $($inputRows.Count) in-scope admins from $InputCsv..." -ForegroundColor Cyan
    foreach ($row in $inputRows) {
        $upn = ([string]$row.UserPrincipalName).Trim()
        if ([string]::IsNullOrWhiteSpace($upn)) { continue }
        try {
            $resolved = Get-MgUser -UserId $upn -Property 'id' -ErrorAction Stop
            if (-not $adminUserIds.Contains($resolved.Id)) { $adminUserIds.Add($resolved.Id) | Out-Null }
        }
        catch {
            Write-Warning "Could not resolve in-scope admin '$upn': $($_.Exception.Message)"
        }
    }
}
else {
    foreach ($userId in @($roleIndex.Active.Keys) + @($roleIndex.Eligible.Keys)) {
        if (-not $adminUserIds.Contains($userId)) { $adminUserIds.Add($userId) | Out-Null }
    }
}

if ($adminUserIds.Count -eq 0) {
    throw 'No in-scope admins were found.'
}
Write-Host "In-scope admins: $($adminUserIds.Count)" -ForegroundColor Cyan

Write-Host 'Collecting Conditional Access policies...' -ForegroundColor Cyan
$caPolicies = @(Get-MgIdentityConditionalAccessPolicy -All -ErrorAction Stop)
Write-Host "Conditional Access policies: $($caPolicies.Count)" -ForegroundColor Cyan

$policyConditions = @{}
foreach ($policy in $caPolicies) {
    $policyConditions[[string]$policy.Id] = ConvertTo-CaUsersCondition -Policy $policy
}

$postureRows = [System.Collections.Generic.List[object]]::new()
$policyMapRows = [System.Collections.Generic.List[object]]::new()
$count = 0

foreach ($adminUserId in $adminUserIds) {
    $count++

    $activeEntries = if ($roleIndex.Active.ContainsKey($adminUserId)) { @($roleIndex.Active[$adminUserId]) } else { @() }
    $eligibleEntries = if ($roleIndex.Eligible.ContainsKey($adminUserId)) { @($roleIndex.Eligible[$adminUserId]) } else { @() }

    try {
        $user = Get-MgUser -UserId $adminUserId -Property 'id,displayName,userPrincipalName,accountEnabled' -ErrorAction Stop
    }
    catch {
        $postureRows.Add((New-AdminPostureRow -User ([pscustomobject]@{ Id = $adminUserId; UserPrincipalName = ''; DisplayName = ''; AccountEnabled = $null }) -ActiveRoleEntries $activeEntries -EligibleRoleEntries $eligibleEntries -PolicyRows @() -ErrorText "User lookup failed: $($_.Exception.Message)")) | Out-Null
        continue
    }

    Write-Progress -Activity 'Mapping Conditional Access and PIM posture' `
        -Status ("{0} of {1} - {2}" -f $count, $adminUserIds.Count, $user.UserPrincipalName) `
        -PercentComplete (($count / $adminUserIds.Count) * 100)

    $errorText = ''
    $groupIds = @()
    try {
        $groupIds = Get-UserTransitiveGroupIds -UserId $adminUserId
    }
    catch {
        $errorText = "Group membership lookup failed: $($_.Exception.Message)"
    }

    $activeTemplateIds = @($activeEntries | ForEach-Object RoleTemplateId | Sort-Object -Unique)
    $eligibleTemplateIds = @($eligibleEntries | ForEach-Object RoleTemplateId | Sort-Object -Unique)

    $userPolicyRows = [System.Collections.Generic.List[object]]::new()
    foreach ($policy in $caPolicies) {
        $status = Test-CaPolicyUserScope -UsersCondition $policyConditions[[string]$policy.Id] -UserId $adminUserId `
            -GroupIds $groupIds -ActiveRoleTemplateIds $activeTemplateIds -EligibleRoleTemplateIds $eligibleTemplateIds
        if ($status -eq 'NotInScope') { continue }
        $userPolicyRows.Add((New-AdminCaPolicyRow -User $user -Policy $policy -ApplicabilityStatus $status)) | Out-Null
    }

    $policyMapRows.AddRange($userPolicyRows)
    $postureRows.Add((New-AdminPostureRow -User $user -ActiveRoleEntries $activeEntries -EligibleRoleEntries $eligibleEntries -PolicyRows $userPolicyRows.ToArray() -ErrorText $errorText)) | Out-Null
}
Write-Progress -Activity 'Mapping Conditional Access and PIM posture' -Completed

$summaryMetrics = New-PostureSummaryMetrics -PostureRows $postureRows.ToArray() -CaPolicyCount $caPolicies.Count

# Export-Excel appends sheets to an existing workbook; remove it so stale sheets never survive.
if (Test-Path -Path $OutputPath) {
    Remove-Item -Path $OutputPath -Force
}

$summaryMetrics | Export-Excel -Path $OutputPath -WorksheetName 'Summary' -AutoFilter -FreezeTopRow -BoldTopRow

$postureSheetRows = @($postureRows | Sort-Object UserPrincipalName)
if ($postureSheetRows.Count -eq 0) {
    $postureSheetRows = @([pscustomobject]@{ Note = 'No in-scope admins were found.' })
}
$postureSheetRows | Export-Excel -Path $OutputPath -WorksheetName 'AdminPosture' -AutoFilter -FreezeTopRow -BoldTopRow

$policyMapSheetRows = @($policyMapRows | Sort-Object AdminUserPrincipalName, PolicyName)
if ($policyMapSheetRows.Count -eq 0) {
    $policyMapSheetRows = @([pscustomobject]@{ Note = 'No Conditional Access policies were relevant to the in-scope admins.' })
}
$policyMapSheetRows | Export-Excel -Path $OutputPath -WorksheetName 'CaPolicyMap' -AutoFilter -FreezeTopRow -BoldTopRow

$noEnabledCa = @($summaryMetrics | Where-Object Metric -eq 'AdminsWithNoEnabledCaPolicyApplying')[0].Count
$noMfaCa = @($summaryMetrics | Where-Object Metric -eq 'AdminsWithoutMfaRequiringEnabledCa')[0].Count
$withEligible = @($summaryMetrics | Where-Object Metric -eq 'AdminsWithPimEligibleRoles')[0].Count

Write-Host ''
Write-Host "Workbook exported: $OutputPath (AdminPosture: $($postureRows.Count) rows, CaPolicyMap: $($policyMapRows.Count) rows)" -ForegroundColor Green
Write-Host "Admins with PIM-eligible roles: $withEligible" -ForegroundColor Green
Write-Host "Admins with no enabled CA policy applying: $noEnabledCa" -ForegroundColor $(if ($noEnabledCa -gt 0) { 'Yellow' } else { 'Green' })
Write-Host "Admins without an MFA-requiring enabled CA policy: $noMfaCa" -ForegroundColor $(if ($noMfaCa -gt 0) { 'Yellow' } else { 'Green' })
