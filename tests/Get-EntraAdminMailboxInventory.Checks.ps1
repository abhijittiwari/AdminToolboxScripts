#Requires -Version 7
$ErrorActionPreference = 'Stop'

$script:ScriptPath = Join-Path $PSScriptRoot '..' 'Get-EntraAdminMailboxInventory.ps1'
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
$script:ExchangeStub = @{}

function Reset-Stubs {
    $script:GraphStub = @{
        RoleDefinitions     = @()
        ActiveAssignments   = @()
        EligibleAssignments = @()
        DirectoryObjects    = @{}
        GroupMembers        = @{}
        Users               = @{}
        LicenseDetails      = @{}
    }
    $script:ExchangeStub = @{
        Mailboxes = @{}
        Failures  = @{}
    }
}

function Get-MgRoleManagementDirectoryRoleDefinition {
    param([switch]$All, [System.Management.Automation.ActionPreference]$ErrorAction)
    return @($script:GraphStub.RoleDefinitions)
}

function Get-MgRoleManagementDirectoryRoleAssignment {
    param([switch]$All, [System.Management.Automation.ActionPreference]$ErrorAction)
    return @($script:GraphStub.ActiveAssignments)
}

function Get-MgRoleManagementDirectoryRoleEligibilityScheduleInstance {
    param([switch]$All, [System.Management.Automation.ActionPreference]$ErrorAction)
    return @($script:GraphStub.EligibleAssignments)
}

function Get-MgDirectoryObjectById {
    param([string]$Ids, [System.Management.Automation.ActionPreference]$ErrorAction)
    return $script:GraphStub.DirectoryObjects[$Ids]
}

function Get-MgGroupTransitiveMember {
    param([string]$GroupId, [switch]$All, [System.Management.Automation.ActionPreference]$ErrorAction)
    return @($script:GraphStub.GroupMembers[$GroupId])
}

function Get-MgUser {
    param([string]$UserId, [string[]]$Property, [System.Management.Automation.ActionPreference]$ErrorAction)
    return $script:GraphStub.Users[$UserId]
}

function Get-MgUserLicenseDetail {
    param([string]$UserId, [System.Management.Automation.ActionPreference]$ErrorAction)
    return @($script:GraphStub.LicenseDetails[$UserId])
}

function Get-EXOMailbox {
    param(
        [string]$Identity,
        [string[]]$Properties,
        [System.Management.Automation.ActionPreference]$ErrorAction
    )

    if ($script:ExchangeStub.Failures.ContainsKey($Identity)) {
        throw $script:ExchangeStub.Failures[$Identity]
    }
    if (-not $script:ExchangeStub.Mailboxes.ContainsKey($Identity)) {
        throw "The operation couldn't be performed because object '$Identity' couldn't be found."
    }
    return $script:ExchangeStub.Mailboxes[$Identity]
}

function Test-RequiredFunctionsExist {
    $ast = Get-ScriptAst
    $functionNames = @($ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true) | ForEach-Object Name)
    foreach ($name in @('Join-ReportValues','Add-UserRole','Resolve-RolePrincipal','Get-AdminUserRoleMap','Get-AdminMailboxReportRow','Test-RequiredCommand','Connect-Services')) {
        Assert-True -Condition ($functionNames -contains $name) -Name "Function exists: $name"
    }
}
Test-RequiredFunctionsExist

function Test-StaticScriptContract {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw
    foreach ($column in @('DisplayName','UserPrincipalName','HasMailbox','ExchangeGuid','LicenseName','MailboxType')) {
        Assert-True -Condition ($scriptText -match [regex]::Escape($column)) -Name "Output column present: $column"
    }
    Assert-True -Condition ($scriptText -match 'Get-MgRoleManagementDirectoryRoleAssignment' -and $scriptText -match 'Get-MgRoleManagementDirectoryRoleEligibilityScheduleInstance') -Name 'Active and PIM role cmdlets present'
    Assert-True -Condition ($scriptText -match 'Get-MgGroupTransitiveMember' -and $scriptText -match 'Get-MgDirectoryObjectById') -Name 'Role group expansion cmdlets present'
    Assert-True -Condition ($scriptText -match 'Get-EXOMailbox' -and $scriptText -match 'ExchangeGuid' -and $scriptText -match 'RecipientTypeDetails') -Name 'Exchange mailbox fields are queried'
    Assert-True -Condition ($scriptText -match 'Directory\.Read\.All' -and $scriptText -match 'RoleManagement\.Read\.Directory' -and $scriptText -match 'User\.Read\.All') -Name 'Required Graph scopes present'
    Assert-True -Condition ($scriptText -match 'Test-RequiredCommand' -and $scriptText -match 'Install Microsoft.Graph and ExchangeOnlineManagement') -Name 'Missing module commands fail with actionable guidance'
}
Test-StaticScriptContract

function Test-JoinReportValues {
    $joined = Join-ReportValues -Values @('SPE_E5', '', $null, 'EMS', 'SPE_E5')
    Assert-True -Condition ($joined -eq 'EMS; SPE_E5') -Name 'Join-ReportValues sorts, deduplicates, and skips blanks'
}
Test-JoinReportValues

function Test-AdminRoleMapIncludesActivePimAndGroupUsers {
    Reset-Stubs
    $script:GraphStub.RoleDefinitions = @(
        [pscustomobject]@{ Id = 'role-1'; DisplayName = 'Global Administrator' }
        [pscustomobject]@{ Id = 'role-2'; DisplayName = 'Exchange Administrator' }
    )
    $script:GraphStub.ActiveAssignments = @([pscustomobject]@{ RoleDefinitionId = 'role-1'; PrincipalId = 'user-active' })
    $script:GraphStub.EligibleAssignments = @([pscustomobject]@{ RoleDefinitionId = 'role-2'; PrincipalId = 'group-pim' })
    $script:GraphStub.DirectoryObjects['user-active'] = [pscustomobject]@{ AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.user' } }
    $script:GraphStub.DirectoryObjects['group-pim'] = [pscustomobject]@{ AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.group'; displayName = 'PIM Admin Group' } }
    $script:GraphStub.GroupMembers['group-pim'] = @(
        [pscustomobject]@{ Id = 'user-pim'; AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.user' } }
        [pscustomobject]@{ Id = 'spn-1'; AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.servicePrincipal' } }
    )

    $roleMap = Get-AdminUserRoleMap

    Assert-True -Condition ($roleMap.ContainsKey('user-active') -and $roleMap['user-active'][0] -eq 'Global Administrator (Active)') -Name 'Admin role map includes active user assignment'
    Assert-True -Condition ($roleMap.ContainsKey('user-pim') -and $roleMap['user-pim'][0] -match 'Exchange Administrator \(PIM Eligible\) via group') -Name 'Admin role map includes PIM group user assignment'
    Assert-True -Condition (-not $roleMap.ContainsKey('spn-1')) -Name 'Admin role map excludes non-user group members'
}
Test-AdminRoleMapIncludesActivePimAndGroupUsers

function Test-AdminMailboxReportRowWithMailbox {
    Reset-Stubs
    $script:GraphStub.Users['user-1'] = [pscustomobject]@{ DisplayName = 'Admin User'; UserPrincipalName = 'admin@example.com' }
    $script:GraphStub.LicenseDetails['user-1'] = @(
        [pscustomobject]@{ SkuPartNumber = 'SPE_E5' }
        [pscustomobject]@{ SkuPartNumber = 'EMS' }
    )
    $script:ExchangeStub.Mailboxes['admin@example.com'] = [pscustomobject]@{ ExchangeGuid = '11111111-1111-1111-1111-111111111111'; RecipientTypeDetails = 'UserMailbox' }

    $row = Get-AdminMailboxReportRow -UserId 'user-1'

    Assert-True -Condition ($row.DisplayName -eq 'Admin User' -and $row.UserPrincipalName -eq 'admin@example.com') -Name 'Report row includes user identity fields'
    Assert-True -Condition ($row.HasMailbox -eq $true -and $row.ExchangeGuid -eq '11111111-1111-1111-1111-111111111111' -and $row.MailboxType -eq 'UserMailbox') -Name 'Report row includes mailbox details'
    Assert-True -Condition ($row.LicenseName -eq 'EMS; SPE_E5') -Name 'Report row includes joined license names'
}
Test-AdminMailboxReportRowWithMailbox

function Test-AdminMailboxReportRowWithoutMailbox {
    Reset-Stubs
    $script:GraphStub.Users['user-2'] = [pscustomobject]@{ DisplayName = 'No Mailbox'; UserPrincipalName = 'nomailbox@example.com' }

    $row = Get-AdminMailboxReportRow -UserId 'user-2'

    Assert-True -Condition ($row.HasMailbox -eq $false -and $row.ExchangeGuid -eq '' -and $row.MailboxType -eq '') -Name 'Report row handles admins without mailboxes'
}
Test-AdminMailboxReportRowWithoutMailbox

if ($script:Failures.Count -gt 0) {
    Write-Host "$($script:Failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host "All $($script:Passes) checks passed." -ForegroundColor Green
exit 0
