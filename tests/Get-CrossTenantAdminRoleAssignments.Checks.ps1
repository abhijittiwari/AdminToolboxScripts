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

$script:GraphStub = @{}

function Reset-GraphStub {
    $script:GraphStub = @{
        Users             = @{}
        UserFailures      = @{}
        UserSearchResults = @{}
        LicenseFailures   = @{}
        LicenseDetails    = @{}
        RoleDefinitions   = @()
        ActiveAssignments = @()
        EligibleAssignments = @()
        DirectoryObjects  = @{}
        GroupMembers      = @{}
        GroupFailures     = @{}
    }
}

function New-StubErrorRecord {
    param(
        [Parameter(Mandatory = $true)] [string]$Message,
        [Parameter(Mandatory = $true)] [string]$ErrorId
    )

    $exception = [System.Exception]::new($Message)
    return [System.Management.Automation.ErrorRecord]::new(
        $exception,
        $ErrorId,
        [System.Management.Automation.ErrorCategory]::NotSpecified,
        $null
    )
}

function Get-MgUser {
    param(
        [string]$UserId,
        [string]$Filter,
        [string[]]$Property,
        [switch]$All,
        [System.Management.Automation.ActionPreference]$ErrorAction
    )

    if (-not [string]::IsNullOrWhiteSpace($Filter)) {
        if (-not $script:GraphStub.UserSearchResults.ContainsKey($Filter)) {
            return @()
        }
        return @($script:GraphStub.UserSearchResults[$Filter])
    }

    if ($script:GraphStub.UserFailures.ContainsKey($UserId)) {
        throw $script:GraphStub.UserFailures[$UserId]
    }

    return $script:GraphStub.Users[$UserId]
}

function Get-MgUserLicenseDetail {
    param(
        [string]$UserId,
        [System.Management.Automation.ActionPreference]$ErrorAction
    )

    if ($script:GraphStub.LicenseFailures.ContainsKey($UserId)) {
        throw $script:GraphStub.LicenseFailures[$UserId]
    }

    return @($script:GraphStub.LicenseDetails[$UserId])
}

function Get-MgRoleManagementDirectoryRoleDefinition {
    param([switch]$All)
    return @($script:GraphStub.RoleDefinitions)
}

function Get-MgRoleManagementDirectoryRoleAssignment {
    param([switch]$All)
    return @($script:GraphStub.ActiveAssignments)
}

function Get-MgRoleManagementDirectoryRoleEligibilityScheduleInstance {
    param([switch]$All)
    return @($script:GraphStub.EligibleAssignments)
}

function Get-MgDirectoryObjectById {
    param(
        [string]$Ids,
        [System.Management.Automation.ActionPreference]$ErrorAction
    )
    return $script:GraphStub.DirectoryObjects[$Ids]
}

function Get-MgGroupTransitiveMember {
    param(
        [string]$GroupId,
        [switch]$All,
        [System.Management.Automation.ActionPreference]$ErrorAction
    )

    if ($script:GraphStub.GroupFailures.ContainsKey($GroupId)) {
        if ($ErrorAction -eq [System.Management.Automation.ActionPreference]::Stop) {
            throw $script:GraphStub.GroupFailures[$GroupId]
        }
        return @()
    }

    return @($script:GraphStub.GroupMembers[$GroupId])
}

function Test-RequiredFunctionsExist {
    $ast = Get-ScriptAst
    $functionNames = @($ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true) | ForEach-Object Name)
    foreach ($name in @('Test-RequiredCsvColumns','Join-ReportValues','Add-UserRole','Test-GraphNotFoundError','ConvertTo-AdminlessDisplayName','Find-MgUserByDisplayName','Get-DisplayNameFallbackCandidates','Resolve-RolePrincipal','Get-TenantRoleIndex','Get-UserReportRow','Invoke-TenantExport')) {
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

function Test-ConvertToAdminlessDisplayName {
    $cleanName = ConvertTo-AdminlessDisplayName -DisplayName '  George   Merry (Admin)  '
    Assert-True -Condition ($cleanName -eq 'George Merry') -Name 'ConvertTo-AdminlessDisplayName removes Admin marker and normalizes spaces'
}
Test-ConvertToAdminlessDisplayName

function Test-DisplayNameFallbackCandidatesPreferAdminPrefix {
    $candidates = @(Get-DisplayNameFallbackCandidates -DisplayName 'Amandeep Bhatoa (Admin)')
    Assert-True -Condition ($candidates.Count -eq 2 -and $candidates[0] -eq 'Admin Amandeep Bhatoa' -and $candidates[1] -eq 'Amandeep Bhatoa') -Name 'Display name fallback candidates prefer Admin-prefixed account'
}
Test-DisplayNameFallbackCandidatesPreferAdminPrefix

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

function New-EmptyRoleIndex {
    return [pscustomobject]@{
        ActiveRoles         = @{}
        EligibleRoles       = @{}
        PimWarning          = ''
        EnvironmentWarnings = ''
    }
}

function Test-UserLookupStatusDistinguishesNotFoundAndUnexpectedErrors {
    Reset-GraphStub
    $script:GraphStub.UserFailures['missing@example.com'] = New-StubErrorRecord -Message "Resource 'missing@example.com' does not exist or one of its queried reference-property objects are not present." -ErrorId 'Request_ResourceNotFound'
    $script:GraphStub.UserFailures['throttle@example.com'] = New-StubErrorRecord -Message 'Too many requests' -ErrorId 'TooManyRequests'

    $missingRow = Get-UserReportRow -Environment 'WAE' -Prefix 'adma' -InputUPN 'missing@example.com' -RoleIndex (New-EmptyRoleIndex)
    $throttleRow = Get-UserReportRow -Environment 'WAE' -Prefix 'adma' -InputUPN 'throttle@example.com' -RoleIndex (New-EmptyRoleIndex)

    Assert-True -Condition ($missingRow.LookupStatus -eq 'NotFound') -Name 'User lookup missing user remains NotFound'
    Assert-True -Condition ($throttleRow.LookupStatus -eq 'Error') -Name 'Unexpected user lookup failure is Error'
}
Test-UserLookupStatusDistinguishesNotFoundAndUnexpectedErrors

function Test-LicenseFailureSetsLookupStatusError {
    Reset-GraphStub
    $script:GraphStub.Users['licensed@example.com'] = [pscustomobject]@{
        Id                    = 'user-1'
        UserPrincipalName     = 'licensed@example.com'
        DisplayName           = 'Licensed User'
        OnPremisesImmutableId = 'immutable-1'
    }
    $script:GraphStub.LicenseFailures['user-1'] = New-StubErrorRecord -Message 'License endpoint unavailable' -ErrorId 'ServiceUnavailable'

    $row = Get-UserReportRow -Environment 'WAE' -Prefix 'adma' -InputUPN 'licensed@example.com' -RoleIndex (New-EmptyRoleIndex)

    Assert-True -Condition ($row.LookupStatus -eq 'Error') -Name 'License lookup failure sets LookupStatus Error'
    Assert-True -Condition ($row.Error -match 'License lookup failed') -Name 'License lookup failure is reported in Error'
}
Test-LicenseFailureSetsLookupStatusError

function Test-PimWarningPersistsWhenUserLookupFails {
    Reset-GraphStub
    $script:GraphStub.UserFailures['missing-pim@example.com'] = New-StubErrorRecord -Message 'Not found' -ErrorId 'Request_ResourceNotFound'
    $roleIndex = New-EmptyRoleIndex
    $roleIndex.PimWarning = 'Could not read PIM eligibility schedules: denied'

    $row = Get-UserReportRow -Environment 'WAE' -Prefix 'adma' -InputUPN 'missing-pim@example.com' -RoleIndex $roleIndex

    Assert-True -Condition ($row.Error -match [regex]::Escape($roleIndex.PimWarning)) -Name 'PIM warning remains in Error when user lookup fails'
}
Test-PimWarningPersistsWhenUserLookupFails

function Test-BlankUpnReturnsErrorRow {
    Reset-GraphStub
    $row = $null
    $threw = $false
    try {
        $row = Get-UserReportRow -Environment 'WAE' -Prefix 'adma' -InputUPN '' -RoleIndex (New-EmptyRoleIndex)
    }
    catch {
        $threw = $true
    }

    Assert-True -Condition (-not $threw) -Name 'Blank UPN does not abort row generation'
    Assert-True -Condition ($null -ne $row -and $row.LookupStatus -eq 'Error' -and $row.Error -match 'blank') -Name 'Blank UPN returns an Error row'
}
Test-BlankUpnReturnsErrorRow

function Test-DisplayNameFallbackFindsFortescueUser {
    Reset-GraphStub
    $script:GraphStub.UserFailures['admgmerry@wae.com'] = New-StubErrorRecord -Message "Resource 'admgmerry@wae.com' does not exist." -ErrorId 'Request_ResourceNotFound'
    $script:GraphStub.UserSearchResults["displayName eq 'George Merry'"] = @([pscustomobject]@{
        Id                    = 'fortescue-user-1'
        UserPrincipalName     = 'george.merry.a@fortescue.com'
        DisplayName           = 'George Merry'
        OnPremisesImmutableId = 'immutable-fortescue-1'
    })
    $script:GraphStub.LicenseDetails['fortescue-user-1'] = @([pscustomobject]@{ SkuPartNumber = 'M365_E5' })

    $roleIndex = New-EmptyRoleIndex
    $roleIndex.EligibleRoles['fortescue-user-1'] = [System.Collections.Generic.List[string]]::new()
    $roleIndex.EligibleRoles['fortescue-user-1'].Add('Global Administrator (PIM Eligible)') | Out-Null

    $row = Get-UserReportRow -Environment 'Fortescue' -Prefix 'admgmerry' -InputUPN 'admgmerry@wae.com' -RoleIndex $roleIndex -DisplayNameHint 'George Merry (Admin)'

    Assert-True -Condition ($row.LookupStatus -eq 'FoundByDisplayName') -Name 'Display name fallback reports FoundByDisplayName'
    Assert-True -Condition ($row.UserPrincipalName -eq 'george.merry.a@fortescue.com') -Name 'Display name fallback returns Fortescue UPN'
    Assert-True -Condition ($row.PimEligibleDirectoryRoles -eq 'Global Administrator (PIM Eligible)') -Name 'PIM eligible roles are included for display-name match'
}
Test-DisplayNameFallbackFindsFortescueUser

function Test-DisplayNameFallbackPrefersAdminAccountWhenBothExist {
    Reset-GraphStub
    $script:GraphStub.UserFailures['admabhatoa@wae.com'] = New-StubErrorRecord -Message "Resource 'admabhatoa@wae.com' does not exist." -ErrorId 'Request_ResourceNotFound'
    $script:GraphStub.UserSearchResults["displayName eq 'Admin Amandeep Bhatoa'"] = @([pscustomobject]@{
        Id                    = 'admin-user-1'
        UserPrincipalName     = 'admin.amandeep.bhatoa@fortescue.com'
        DisplayName           = 'Admin Amandeep Bhatoa'
        OnPremisesImmutableId = 'immutable-admin-1'
    })
    $script:GraphStub.UserSearchResults["displayName eq 'Amandeep Bhatoa'"] = @([pscustomobject]@{
        Id                    = 'standard-user-1'
        UserPrincipalName     = 'amandeep.bhatoa@fortescue.com'
        DisplayName           = 'Amandeep Bhatoa'
        OnPremisesImmutableId = 'immutable-standard-1'
    })

    $row = Get-UserReportRow -Environment 'Fortescue' -Prefix 'admabhatoa' -InputUPN 'admabhatoa@wae.com' -RoleIndex (New-EmptyRoleIndex) -DisplayNameHint 'Amandeep Bhatoa (Admin)'

    Assert-True -Condition ($row.UserPrincipalName -eq 'admin.amandeep.bhatoa@fortescue.com') -Name 'Display name fallback prefers Fortescue Admin-prefixed account'
    Assert-True -Condition ($row.DisplayName -eq 'Admin Amandeep Bhatoa') -Name 'Display name fallback returns Admin-prefixed display name'
}
Test-DisplayNameFallbackPrefersAdminAccountWhenBothExist

function Test-DisplayNameFallbackRejectsAmbiguousMatches {
    Reset-GraphStub
    $script:GraphStub.UserFailures['admgmerry@wae.com'] = New-StubErrorRecord -Message "Resource 'admgmerry@wae.com' does not exist." -ErrorId 'Request_ResourceNotFound'
    $script:GraphStub.UserSearchResults["displayName eq 'George Merry'"] = @(
        [pscustomobject]@{ Id = 'user-1'; UserPrincipalName = 'george.merry.a@fortescue.com'; DisplayName = 'George Merry'; OnPremisesImmutableId = 'immutable-1' },
        [pscustomobject]@{ Id = 'user-2'; UserPrincipalName = 'george.merry.b@fortescue.com'; DisplayName = 'George Merry'; OnPremisesImmutableId = 'immutable-2' }
    )

    $row = Get-UserReportRow -Environment 'Fortescue' -Prefix 'admgmerry' -InputUPN 'admgmerry@wae.com' -RoleIndex (New-EmptyRoleIndex) -DisplayNameHint 'George Merry (Admin)'

    Assert-True -Condition ($row.LookupStatus -eq 'Error') -Name 'Ambiguous display name fallback is an Error'
    Assert-True -Condition ($row.Error -match 'matched multiple users') -Name 'Ambiguous display name fallback is reported'
}
Test-DisplayNameFallbackRejectsAmbiguousMatches

function Test-RoleIndexIncludesDirectPimEligibleRoles {
    Reset-GraphStub
    $script:GraphStub.RoleDefinitions = @([pscustomobject]@{ Id = 'role-eligible'; DisplayName = 'Privileged Role Administrator' })
    $script:GraphStub.EligibleAssignments = @([pscustomobject]@{ RoleDefinitionId = 'role-eligible'; PrincipalId = 'user-eligible' })
    $script:GraphStub.DirectoryObjects['user-eligible'] = [pscustomobject]@{
        Id = 'user-eligible'
        AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.user' }
    }

    $roleIndex = Get-TenantRoleIndex

    Assert-True -Condition ($roleIndex.EligibleRoles.ContainsKey('user-eligible')) -Name 'Role index captures direct PIM eligible user assignment'
    Assert-True -Condition ($roleIndex.EligibleRoles['user-eligible'][0] -eq 'Privileged Role Administrator (PIM Eligible)') -Name 'Role index labels PIM eligible user assignment'
}
Test-RoleIndexIncludesDirectPimEligibleRoles

function Test-GroupExpansionWarningIsCapturedAndPropagated {
    Reset-GraphStub
    $script:GraphStub.RoleDefinitions = @([pscustomobject]@{ Id = 'role-1'; DisplayName = 'Global Reader' })
    $script:GraphStub.ActiveAssignments = @([pscustomobject]@{ RoleDefinitionId = 'role-1'; PrincipalId = 'group-1' })
    $script:GraphStub.DirectoryObjects['group-1'] = [pscustomobject]@{
        Id = 'group-1'
        AdditionalProperties = @{
            '@odata.type' = '#microsoft.graph.group'
            displayName   = 'Admins'
        }
    }
    $script:GraphStub.GroupFailures['group-1'] = New-StubErrorRecord -Message 'Group members denied' -ErrorId 'Authorization_RequestDenied'
    $script:GraphStub.Users['admin@example.com'] = [pscustomobject]@{
        Id                    = 'user-2'
        UserPrincipalName     = 'admin@example.com'
        DisplayName           = 'Admin User'
        OnPremisesImmutableId = 'immutable-2'
    }

    $roleIndex = Get-TenantRoleIndex
    $row = Get-UserReportRow -Environment 'WAE' -Prefix 'adma' -InputUPN 'admin@example.com' -RoleIndex $roleIndex

    Assert-True -Condition ($roleIndex.EnvironmentWarnings -match 'Group role expansion failed') -Name 'Group expansion warning is captured in role index'
    Assert-True -Condition ($row.Error -match 'Group role expansion failed') -Name 'Group expansion warning is included in row Error'
}
Test-GroupExpansionWarningIsCapturedAndPropagated

if ($script:Failures.Count -gt 0) {
    Write-Host "$($script:Failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host "All $($script:Passes) checks passed." -ForegroundColor Green
exit 0
