<#
.SYNOPSIS
    Exports Entra ID admin users with mailbox and license details.

.DESCRIPTION
    Finds users with active or PIM-eligible Entra ID directory role assignments,
    including users assigned through role-assignable groups, then checks whether
    each admin has an Exchange Online mailbox.

    The script is read-only. It writes a CSV with exactly these primary fields:
    DisplayName, UserPrincipalName, HasMailbox, ExchangeGuid, LicenseName, and
    MailboxType.

.PARAMETER OutputPath
    CSV output path. Defaults to .\EntraAdminMailboxInventory_yyyyMMdd_HHmmss.csv.

.EXAMPLE
    .\Get-EntraAdminMailboxInventory.ps1
#>

[CmdletBinding()]
param(
    [string]$OutputPath = ".\EntraAdminMailboxInventory_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

$ErrorActionPreference = 'Stop'

function Join-ReportValues {
    param([AllowEmptyCollection()] [object[]]$Values)

    $cleanValues = @($Values | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | ForEach-Object { ([string]$_).Trim() } | Sort-Object -Unique)
    return ($cleanValues -join '; ')
}

function Add-UserRole {
    param(
        [Parameter(Mandatory = $true)] [hashtable]$RoleMap,
        [Parameter(Mandatory = $true)] [string]$UserId,
        [Parameter(Mandatory = $true)] [string]$RoleName
    )

    if (-not $RoleMap.ContainsKey($UserId)) {
        $RoleMap[$UserId] = [System.Collections.Generic.List[string]]::new()
    }
    if (-not $RoleMap[$UserId].Contains($RoleName)) {
        $RoleMap[$UserId].Add($RoleName) | Out-Null
    }
}

function Resolve-RolePrincipal {
    param(
        [Parameter(Mandatory = $true)] [string]$PrincipalId,
        [Parameter(Mandatory = $true)] [string]$RoleName,
        [Parameter(Mandatory = $true)] [hashtable]$RoleMap
    )

    $principal = Get-MgDirectoryObjectById -Ids $PrincipalId -ErrorAction SilentlyContinue
    if (-not $principal) { return }

    switch ($principal.AdditionalProperties.'@odata.type') {
        '#microsoft.graph.user' {
            Add-UserRole -RoleMap $RoleMap -UserId $PrincipalId -RoleName $RoleName
        }
        '#microsoft.graph.group' {
            $groupName = $principal.AdditionalProperties.displayName
            try {
                Get-MgGroupTransitiveMember -GroupId $PrincipalId -All -ErrorAction Stop |
                    Where-Object { $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.user' } |
                    ForEach-Object {
                        Add-UserRole -RoleMap $RoleMap -UserId $_.Id -RoleName "$RoleName via group: $groupName"
                    }
            }
            catch {
                Write-Warning "Could not expand role group '$groupName' ($PrincipalId): $($_.Exception.Message)"
            }
        }
    }
}

function Get-AdminUserRoleMap {
    $roleDefinitions = @{}
    Get-MgRoleManagementDirectoryRoleDefinition -All -ErrorAction Stop | ForEach-Object {
        $roleDefinitions[$_.Id] = $_.DisplayName
    }

    $roleMap = @{}

    Write-Host 'Collecting active Entra ID directory role assignments...' -ForegroundColor Cyan
    $activeAssignments = @(Get-MgRoleManagementDirectoryRoleAssignment -All -ErrorAction Stop)
    foreach ($assignment in $activeAssignments) {
        $roleName = $roleDefinitions[$assignment.RoleDefinitionId]
        if (-not [string]::IsNullOrWhiteSpace($roleName)) {
            Resolve-RolePrincipal -PrincipalId $assignment.PrincipalId -RoleName "$roleName (Active)" -RoleMap $roleMap
        }
    }

    Write-Host 'Collecting PIM-eligible Entra ID directory role assignments...' -ForegroundColor Cyan
    try {
        $eligibleAssignments = @(Get-MgRoleManagementDirectoryRoleEligibilityScheduleInstance -All -ErrorAction Stop)
        foreach ($assignment in $eligibleAssignments) {
            $roleName = $roleDefinitions[$assignment.RoleDefinitionId]
            if (-not [string]::IsNullOrWhiteSpace($roleName)) {
                Resolve-RolePrincipal -PrincipalId $assignment.PrincipalId -RoleName "$roleName (PIM Eligible)" -RoleMap $roleMap
            }
        }
    }
    catch {
        Write-Warning "Could not read PIM eligibility schedules: $($_.Exception.Message)"
    }

    return $roleMap
}

function Get-AdminMailboxReportRow {
    param([Parameter(Mandatory = $true)] [string]$UserId)

    $user = Get-MgUser -UserId $UserId -Property 'id,displayName,userPrincipalName' -ErrorAction Stop
    $licenseName = ''
    try {
        $licenseName = Join-ReportValues -Values @(Get-MgUserLicenseDetail -UserId $UserId -ErrorAction Stop | Select-Object -ExpandProperty SkuPartNumber)
    }
    catch {
        Write-Warning "License lookup failed for $($user.UserPrincipalName): $($_.Exception.Message)"
    }

    $mailbox = $null
    try {
        $mailbox = Get-EXOMailbox -Identity $user.UserPrincipalName -Properties ExchangeGuid,RecipientTypeDetails -ErrorAction Stop
    }
    catch {
        if ($_.Exception.Message -notmatch 'couldn''t be found|not found|ManagementObjectNotFoundException|RecipientNotFound') {
            Write-Warning "Mailbox lookup failed for $($user.UserPrincipalName): $($_.Exception.Message)"
        }
    }

    return [pscustomobject]@{
        DisplayName       = [string]$user.DisplayName
        UserPrincipalName = [string]$user.UserPrincipalName
        HasMailbox        = [bool]($null -ne $mailbox)
        ExchangeGuid      = if ($mailbox) { [string]$mailbox.ExchangeGuid } else { '' }
        LicenseName       = $licenseName
        MailboxType       = if ($mailbox) { [string]$mailbox.RecipientTypeDetails } else { '' }
    }
}

function Test-RequiredCommand {
    param([Parameter(Mandatory = $true)] [string[]]$CommandName)

    $missingCommands = @($CommandName | Where-Object { -not (Get-Command $_ -ErrorAction SilentlyContinue) })
    if ($missingCommands.Count -gt 0) {
        throw "Missing required command(s): $($missingCommands -join ', '). Install Microsoft.Graph and ExchangeOnlineManagement modules."
    }
}

function Connect-Services {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module ExchangeOnlineManagement -ErrorAction Stop

    Test-RequiredCommand -CommandName @(
        'Get-MgRoleManagementDirectoryRoleDefinition',
        'Get-MgRoleManagementDirectoryRoleAssignment',
        'Get-MgRoleManagementDirectoryRoleEligibilityScheduleInstance',
        'Get-MgDirectoryObjectById',
        'Get-MgGroupTransitiveMember',
        'Get-MgUser',
        'Get-MgUserLicenseDetail',
        'Get-EXOMailbox'
    )

    if (-not (Get-MgContext -ErrorAction SilentlyContinue)) {
        Connect-MgGraph -Scopes @('Directory.Read.All', 'RoleManagement.Read.Directory', 'User.Read.All') -NoWelcome
    }

    if (@(Get-ConnectionInformation -ErrorAction SilentlyContinue).Count -eq 0) {
        Connect-ExchangeOnline -ShowBanner:$false
    }
}

Connect-Services

$roleMap = Get-AdminUserRoleMap
$rows = [System.Collections.Generic.List[object]]::new()
$count = 0
foreach ($userId in @($roleMap.Keys | Sort-Object)) {
    $count++
    Write-Progress -Activity 'Processing admin mailbox inventory' -Status "$count of $($roleMap.Count)" -PercentComplete (($count / [Math]::Max($roleMap.Count, 1)) * 100)
    $rows.Add((Get-AdminMailboxReportRow -UserId $userId)) | Out-Null
}
Write-Progress -Activity 'Processing admin mailbox inventory' -Completed

$columns = 'DisplayName','UserPrincipalName','HasMailbox','ExchangeGuid','LicenseName','MailboxType'
$rows | Sort-Object UserPrincipalName | Select-Object $columns | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

Write-Host "Report exported: $OutputPath ($($rows.Count) rows)" -ForegroundColor Green
