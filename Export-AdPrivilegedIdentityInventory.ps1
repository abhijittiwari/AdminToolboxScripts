<#
.SYNOPSIS
    Exports privileged, service, and non-human identity details for divestiture discovery.

.DESCRIPTION
    Read-only discovery of the privileged and non-human identity plane the SoW calls out for the
    least-privilege / tooling-privilege design:
      - Privileged group members: recursive membership of the tier-0 groups (Domain/Enterprise/
        Schema Admins, Administrators, Account/Server/Print/Backup Operators, and more), one row per
        (group, member) with the member's enabled state and last logon.
      - adminCount=1 objects: accounts currently or historically in a protected group (AdminSDHolder
        stamp) - useful for spotting orphaned privilege.
      - Service / non-human accounts: users carrying a servicePrincipalName (SPN), plus Group Managed
        Service Accounts (gMSA), with a Kerberos-delegation flag (unconstrained/constrained/RBCD) so
        the risky delegation cases surface.

    Objects are read with the ActiveDirectory module; the classification/shaping is separated so it is
    unit-testable with mock data. The script makes no changes.

.PARAMETER PrivilegedGroup
    Privileged group names to expand. Defaults to the standard tier-0 / operator groups.

.PARAMETER StaleDays
    lastLogonTimestamp age in days after which a member is flagged stale. Default 90.

.PARAMETER Server
    Optional Domain Controller / AD Web Services server to target for the queries.

.PARAMETER Credential
    Optional credential for the AD queries.

.PARAMETER OutputFolder
    Optional folder for the CSV outputs (PrivilegedGroupMembers.csv, AdminCountObjects.csv,
    ServiceAccounts.csv). If omitted, results are only written to the pipeline/table.

.EXAMPLE
    .\Export-AdPrivilegedIdentityInventory.ps1 -OutputFolder .\PrivilegedIdentities
#>

[CmdletBinding()]
param(
    [string[]]$PrivilegedGroup = @('Domain Admins','Enterprise Admins','Schema Admins','Administrators','Account Operators','Server Operators','Print Operators','Backup Operators','Group Policy Creator Owners','DnsAdmins','Cert Publishers'),
    [int]$StaleDays = 90,
    [string]$Server = '',
    [pscredential]$Credential,
    [string]$OutputFolder = ''
)

$ErrorActionPreference = 'Stop'

function ConvertTo-DelegationType {
    param(
        [AllowNull()] $TrustedForDelegation,
        [AllowNull()] $TrustedToAuthForDelegation,
        [AllowNull()] $AllowedToDelegateTo,
        [AllowNull()] $PrincipalsAllowedToDelegateToAccount
    )

    if ([bool]$TrustedForDelegation) { return 'Unconstrained' }
    if ([bool]$TrustedToAuthForDelegation -or @($AllowedToDelegateTo).Where({ $_ }).Count -gt 0) { return 'Constrained' }
    if (@($PrincipalsAllowedToDelegateToAccount).Where({ $_ }).Count -gt 0) { return 'ResourceBased' }
    return 'None'
}

function New-PrivilegedMemberRow {
    param(
        [Parameter(Mandatory = $true)] [string]$GroupName,
        [Parameter(Mandatory = $true)] $Member,
        [Parameter(Mandatory = $true)] [datetime]$StaleThreshold
    )

    $lastLogon = $null
    $isStale = $true
    if ($null -ne $Member.lastLogonTimestamp) {
        try {
            $lastLogon = [datetime]::FromFileTimeUtc([int64]$Member.lastLogonTimestamp)
            $isStale = ($lastLogon -lt $StaleThreshold)
        } catch { }
    }

    return [pscustomobject]@{
        GroupName        = $GroupName
        MemberName       = [string]$Member.SamAccountName
        MemberClass      = [string]$Member.objectClass
        DistinguishedName = [string]$Member.DistinguishedName
        Enabled          = [bool]$Member.Enabled
        LastLogon        = $lastLogon
        StaleMember      = $isStale
    }
}

function New-ServiceAccountRow {
    param([Parameter(Mandatory = $true)] $Account)

    $spns = @($Account.servicePrincipalName)
    return [pscustomobject]@{
        SamAccountName   = [string]$Account.SamAccountName
        ObjectClass      = [string]$Account.objectClass
        Enabled          = [bool]$Account.Enabled
        SpnCount         = @($spns | Where-Object { $_ }).Count
        ServicePrincipalNames = (@($spns | Where-Object { $_ }) -join ';')
        DelegationType   = (ConvertTo-DelegationType -TrustedForDelegation $Account.TrustedForDelegation -TrustedToAuthForDelegation $Account.TrustedToAuthForDelegation -AllowedToDelegateTo $Account.'msDS-AllowedToDelegateTo' -PrincipalsAllowedToDelegateToAccount $Account.PrincipalsAllowedToDelegateToAccount)
        PasswordNeverExpires = [bool]$Account.PasswordNeverExpires
    }
}

function New-AdminCountRow {
    param([Parameter(Mandatory = $true)] $Object)

    return [pscustomobject]@{
        SamAccountName    = [string]$Object.SamAccountName
        ObjectClass       = [string]$Object.objectClass
        Enabled           = [bool]$Object.Enabled
        DistinguishedName = [string]$Object.DistinguishedName
    }
}

function Get-AdPrivilegedIdentityFacts {
    param(
        [string[]]$PrivilegedGroup,
        [string]$Server = '',
        [pscredential]$Credential
    )

    Import-Module ActiveDirectory -ErrorAction Stop
    $adParams = @{}
    if (-not [string]::IsNullOrWhiteSpace($Server)) { $adParams.Server = $Server }
    if ($Credential) { $adParams.Credential = $Credential }

    $memberProps = @('SamAccountName','Enabled','lastLogonTimestamp','objectClass','DistinguishedName')

    $groupMembers = @()
    foreach ($groupName in $PrivilegedGroup) {
        try {
            $members = @(Get-ADGroupMember -Identity $groupName -Recursive @adParams)
            foreach ($member in $members) {
                $detail = $null
                try { $detail = Get-ADObject -Identity $member.DistinguishedName -Properties $memberProps @adParams } catch { }
                if (-not $detail) { $detail = $member }
                $groupMembers += [pscustomobject]@{ GroupName = $groupName; Member = $detail }
            }
        } catch { }
    }

    $adminCountObjects = @()
    try {
        $adminCountObjects = @(Get-ADObject -LDAPFilter '(&(adminCount=1)(|(objectClass=user)(objectClass=group)))' -Properties SamAccountName, Enabled, objectClass @adParams)
    } catch { }

    $serviceAccounts = @()
    try {
        $spnProps = @('SamAccountName','Enabled','servicePrincipalName','objectClass','TrustedForDelegation','TrustedToAuthForDelegation','msDS-AllowedToDelegateTo','PrincipalsAllowedToDelegateToAccount','PasswordNeverExpires')
        $serviceAccounts += @(Get-ADUser -LDAPFilter '(servicePrincipalName=*)' -Properties $spnProps @adParams)
        $serviceAccounts += @(Get-ADServiceAccount -Filter * -Properties $spnProps @adParams)
    } catch { }

    return @{
        GroupMembers      = $groupMembers
        AdminCountObjects = $adminCountObjects
        ServiceAccounts   = $serviceAccounts
    }
}

function Get-AdPrivilegedIdentityInventory {
    param(
        [Parameter(Mandatory = $true)] $Facts,
        [int]$StaleDays = 90,
        [nullable[datetime]]$Now
    )

    if (-not $Now) { $Now = Get-Date }
    $staleThreshold = $Now.AddDays(-$StaleDays)

    return @{
        PrivilegedGroupMembers = @(@($Facts.GroupMembers) | ForEach-Object { New-PrivilegedMemberRow -GroupName $_.GroupName -Member $_.Member -StaleThreshold $staleThreshold })
        AdminCountObjects      = @(@($Facts.AdminCountObjects) | ForEach-Object { New-AdminCountRow -Object $_ })
        ServiceAccounts        = @(@($Facts.ServiceAccounts) | ForEach-Object { New-ServiceAccountRow -Account $_ })
    }
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
Write-Host 'Collecting privileged, service, and non-human identity inventory...' -ForegroundColor Cyan
$facts = Get-AdPrivilegedIdentityFacts -PrivilegedGroup $PrivilegedGroup -Server $Server -Credential $Credential
$inventory = Get-AdPrivilegedIdentityInventory -Facts $facts -StaleDays $StaleDays

Write-Host "`n=== Privileged group members ($(@($inventory.PrivilegedGroupMembers).Count)) ===" -ForegroundColor Yellow
$inventory.PrivilegedGroupMembers | Format-Table GroupName, MemberName, Enabled, StaleMember -AutoSize
Write-Host "=== Service / non-human accounts ($(@($inventory.ServiceAccounts).Count)) ===" -ForegroundColor Yellow
$inventory.ServiceAccounts | Format-Table SamAccountName, SpnCount, DelegationType, Enabled -AutoSize
Write-Host "=== adminCount=1 objects ($(@($inventory.AdminCountObjects).Count)) ===" -ForegroundColor Yellow
$inventory.AdminCountObjects | Format-Table SamAccountName, ObjectClass, Enabled -AutoSize

$riskyDelegation = @($inventory.ServiceAccounts | Where-Object { $_.DelegationType -in @('Unconstrained','Constrained','ResourceBased') })
Write-Host ("Summary: {0} privileged member row(s), {1} service account(s) ({2} with delegation), {3} adminCount object(s)" -f @($inventory.PrivilegedGroupMembers).Count, @($inventory.ServiceAccounts).Count, $riskyDelegation.Count, @($inventory.AdminCountObjects).Count) -ForegroundColor Yellow

if (-not [string]::IsNullOrWhiteSpace($OutputFolder)) {
    if (-not (Test-Path -Path $OutputFolder)) { New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null }
    $inventory.PrivilegedGroupMembers | Export-Csv -Path (Join-Path $OutputFolder 'PrivilegedGroupMembers.csv') -NoTypeInformation -Encoding UTF8
    $inventory.AdminCountObjects | Export-Csv -Path (Join-Path $OutputFolder 'AdminCountObjects.csv') -NoTypeInformation -Encoding UTF8
    $inventory.ServiceAccounts | Export-Csv -Path (Join-Path $OutputFolder 'ServiceAccounts.csv') -NoTypeInformation -Encoding UTF8
    Write-Host ("Reports exported to: {0}" -f $OutputFolder) -ForegroundColor Green
}
