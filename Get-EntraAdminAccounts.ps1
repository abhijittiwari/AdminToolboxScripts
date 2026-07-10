<#
.SYNOPSIS
    Reports all admin accounts in Entra ID (active + PIM-eligible role assignments).

.DESCRIPTION
    For every user holding a directory role (directly, via PIM eligibility, or via a
    role-assignable group), records:
      UPN, ImmutableId, Roles, Licenses, MFA status/methods, Mailbox,
      Group memberships (comma-separated), Last sign-in, Created date.

.REQUIREMENTS
    Modules : Microsoft.Graph (Install-Module Microsoft.Graph)
              ExchangeOnlineManagement (optional, for mailbox details)
    Roles   : Global Reader (or equivalent) recommended.
    Scopes  : Directory.Read.All, RoleManagement.Read.Directory,
              AuditLog.Read.All, UserAuthenticationMethod.Read.All

.PARAMETER IncludeMailbox
    Connects to Exchange Online and records mailbox type (UserMailbox, SharedMailbox,
    etc.). Without this switch, falls back to the user's Mail attribute.

.PARAMETER OutputPath
    CSV output path. Defaults to .\EntraAdminAccounts_<timestamp>.csv

.EXAMPLE
    .\Get-EntraAdminAccounts.ps1 -IncludeMailbox
#>

[CmdletBinding()]
param(
    [switch]$IncludeMailbox,
    [string]$OutputPath = ".\EntraAdminAccounts_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Overall progress bar (Id 0). Sub-steps use Id 1 with -ParentId 0.
# ---------------------------------------------------------------------------
$script:stepNum   = 0
$script:stepTotal = 6
if ($IncludeMailbox) { $script:stepTotal = 7 }

function Show-Step {
    param([string]$Status)
    $script:stepNum++
    Write-Progress -Id 0 -Activity 'Entra ID Admin Account Report' `
        -Status ("Step {0} of {1}: {2}" -f $script:stepNum, $script:stepTotal, $Status) `
        -PercentComplete (($script:stepNum / $script:stepTotal) * 100)
}

# ---------------------------------------------------------------------------
# 1. Connect
# ---------------------------------------------------------------------------
Show-Step 'Connecting to Microsoft Graph'
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes @(
    'Directory.Read.All',
    'RoleManagement.Read.Directory',
    'AuditLog.Read.All',
    'UserAuthenticationMethod.Read.All'
) -NoWelcome

$exoConnected = $false
if ($IncludeMailbox) {
    Show-Step 'Connecting to Exchange Online'
    try {
        Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        Connect-ExchangeOnline -ShowBanner:$false
        $exoConnected = $true
    }
    catch {
        Write-Warning "Exchange Online connection failed ($_). Falling back to Mail attribute."
    }
}

# ---------------------------------------------------------------------------
# 2. Role definitions lookup
# ---------------------------------------------------------------------------
Show-Step 'Loading role definitions'
Write-Host "Loading role definitions..." -ForegroundColor Cyan
$roleDefs = @{}
Get-MgRoleManagementDirectoryRoleDefinition -All | ForEach-Object {
    $roleDefs[$_.Id] = $_.DisplayName
}

# ---------------------------------------------------------------------------
# 3. Collect role assignments: active + PIM eligible
#    $userRoles : userId -> [System.Collections.Generic.List[string]] of role names
# ---------------------------------------------------------------------------
$userRoles = @{}

function Add-UserRole {
    param([string]$UserId, [string]$RoleName)
    if (-not $userRoles.ContainsKey($UserId)) {
        $userRoles[$UserId] = @()
    }
    if ($userRoles[$UserId] -notcontains $RoleName) {
        $userRoles[$UserId] = @($userRoles[$UserId]) + $RoleName
    }
}

function Resolve-Principal {
    param([string]$PrincipalId, [string]$RoleName)
    # Principal can be a user, a (role-assignable) group, or a service principal.
    $obj = Get-MgDirectoryObjectById -Ids $PrincipalId -ErrorAction SilentlyContinue
    if (-not $obj) { return }
    switch ($obj.AdditionalProperties.'@odata.type') {
        '#microsoft.graph.user' {
            Add-UserRole -UserId $PrincipalId -RoleName $RoleName
        }
        '#microsoft.graph.group' {
            $grpName = $obj.AdditionalProperties.displayName
            Get-MgGroupTransitiveMember -GroupId $PrincipalId -All |
                Where-Object { $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.user' } |
                ForEach-Object {
                    Add-UserRole -UserId $_.Id -RoleName "$RoleName (via group: $grpName)"
                }
        }
        default { } # service principals excluded from user report
    }
}

Show-Step 'Collecting active role assignments'
Write-Host "Collecting active role assignments..." -ForegroundColor Cyan
$activeAssignments = @(Get-MgRoleManagementDirectoryRoleAssignment -All)
$n = 0
foreach ($assignment in $activeAssignments) {
    $n++
    $roleName = $roleDefs[$assignment.RoleDefinitionId]
    Write-Progress -Id 1 -ParentId 0 -Activity 'Active role assignments' `
        -Status ("{0} of {1} - {2}" -f $n, $activeAssignments.Count, $roleName) `
        -PercentComplete (($n / $activeAssignments.Count) * 100)
    if ($roleName) { Resolve-Principal -PrincipalId $assignment.PrincipalId -RoleName "$roleName (Active)" }
}
Write-Progress -Id 1 -Activity 'Active role assignments' -Completed

Show-Step 'Collecting PIM eligible role assignments'
Write-Host "Collecting PIM eligible role assignments..." -ForegroundColor Cyan
try {
    $pimEligible = @(Get-MgRoleManagementDirectoryRoleEligibilityScheduleInstance -All)
    $n = 0
    foreach ($eligibility in $pimEligible) {
        $n++
        $roleName = $roleDefs[$eligibility.RoleDefinitionId]
        Write-Progress -Id 1 -ParentId 0 -Activity 'PIM eligible assignments' `
            -Status ("{0} of {1} - {2}" -f $n, $pimEligible.Count, $roleName) `
            -PercentComplete (($n / $pimEligible.Count) * 100)
        if ($roleName) { Resolve-Principal -PrincipalId $eligibility.PrincipalId -RoleName "$roleName (PIM Eligible)" }
    }
    Write-Progress -Id 1 -Activity 'PIM eligible assignments' -Completed
}
catch {
    Write-Warning "Could not read PIM eligibility schedules (requires Entra ID P2 / permissions): $_"
}

Write-Host ("Found {0} unique admin user(s)." -f $userRoles.Count) -ForegroundColor Green

# ---------------------------------------------------------------------------
# 4. Build report per user
# ---------------------------------------------------------------------------
Show-Step 'Processing admin accounts (users, licenses, MFA, mailbox, groups)'
$report = @()
$i = 0

foreach ($userId in $userRoles.Keys) {
    $i++
    try {
        $user = Get-MgUser -UserId $userId -Property `
            'id,displayName,userPrincipalName,onPremisesImmutableId,createdDateTime,signInActivity,mail,accountEnabled'
    }
    catch {
        Write-Warning "Could not retrieve user $userId : $_"
        continue
    }

    Write-Progress -Id 1 -ParentId 0 -Activity "Processing admin accounts" `
        -Status ("{0} of {1} - {2}" -f $i, $userRoles.Count, $user.UserPrincipalName) `
        -PercentComplete (($i / $userRoles.Count) * 100)

    # --- Licenses ---
    $licenses = ''
    try {
        $licenses = (Get-MgUserLicenseDetail -UserId $userId -ErrorAction SilentlyContinue |
            Select-Object -ExpandProperty SkuPartNumber) -join ', '
    } catch { }

    # --- MFA (registration report) ---
    $mfaRegistered = 'Unknown'; $mfaMethods = ''; $mfaDefault = ''
    try {
        $reg = Get-MgReportAuthenticationMethodUserRegistrationDetail `
            -UserRegistrationDetailsId $userId -ErrorAction Stop
        if ($reg) {
            $mfaRegistered = $reg.IsMfaRegistered
            $mfaMethods    = ($reg.MethodsRegistered -join ', ')
            $mfaDefault    = $reg.DefaultMfaMethod
        }
    }
    catch {
        Write-Warning "Could not retrieve MFA registration details for $($user.UserPrincipalName): $($_.Exception.Message)"
    }

    # --- Mailbox ---
    if ($exoConnected) {
        $mbx = Get-EXOMailbox -Identity $user.UserPrincipalName -ErrorAction SilentlyContinue
        $mailbox = if ($mbx) { "$($mbx.RecipientTypeDetails) ($($mbx.PrimarySmtpAddress))" } else { 'No mailbox' }
    }
    else {
        $mailbox = if ($user.Mail) { $user.Mail } else { 'No mail attribute' }
    }

    # --- Group membership (names, comma-separated) ---
    $groups = ''
    try {
        $groups = (Get-MgUserMemberOf -UserId $userId -All -ErrorAction SilentlyContinue |
            Where-Object { $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group' } |
            ForEach-Object { $_.AdditionalProperties.displayName }) -join ', '
    } catch { }

    # --- Sign-in activity ---
    $lastInteractive    = $user.SignInActivity.LastSignInDateTime
    $lastNonInteractive = $user.SignInActivity.LastNonInteractiveSignInDateTime
    $lastSignIn = @($lastInteractive, $lastNonInteractive) |
        Where-Object { $_ } | Sort-Object -Descending | Select-Object -First 1

    # New-Object PSObject -Property is allowed in Constrained Language Mode
    # (unlike a [PSCustomObject] hashtable cast).
    $report += New-Object -TypeName PSObject -Property @{
        DisplayName              = $user.DisplayName
        UPN                      = $user.UserPrincipalName
        ImmutableId              = $user.OnPremisesImmutableId
        AccountEnabled           = $user.AccountEnabled
        Roles                    = ($userRoles[$userId] -join '; ')
        Licenses                 = $licenses
        MfaRegistered            = $mfaRegistered
        MfaMethods               = $mfaMethods
        MfaDefaultMethod         = $mfaDefault
        Mailbox                  = $mailbox
        GroupMembership          = $groups
        LastSignIn               = $lastSignIn
        LastInteractiveSignIn    = $lastInteractive
        LastNonInteractiveSignIn = $lastNonInteractive
        CreatedDate              = $user.CreatedDateTime
    }
}

Write-Progress -Id 1 -Activity "Processing admin accounts" -Completed

# ---------------------------------------------------------------------------
# 5. Output
# ---------------------------------------------------------------------------
Show-Step 'Exporting report'
# Explicit column order (New-Object -Property does not preserve hashtable order)
$columns = 'DisplayName','UPN','ImmutableId','AccountEnabled','Roles','Licenses',
           'MfaRegistered','MfaMethods','MfaDefaultMethod','Mailbox','GroupMembership',
           'LastSignIn','LastInteractiveSignIn','LastNonInteractiveSignIn','CreatedDate'

$report | Sort-Object UPN | Select-Object $columns |
    Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host "Report exported: $OutputPath ($($report.Count) accounts)" -ForegroundColor Green

$report | Sort-Object UPN | Format-Table UPN, Roles, MfaRegistered, LastSignIn -AutoSize

if ($exoConnected) { Disconnect-ExchangeOnline -Confirm:$false }
Disconnect-MgGraph | Out-Null
Write-Progress -Id 0 -Activity 'Entra ID Admin Account Report' -Completed
