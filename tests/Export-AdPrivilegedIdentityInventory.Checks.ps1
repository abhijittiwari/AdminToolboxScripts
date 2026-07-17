#Requires -Version 7
$ErrorActionPreference = 'Stop'

$script:ScriptPath = Join-Path $PSScriptRoot '..' 'Export-AdPrivilegedIdentityInventory.ps1'
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
    foreach ($name in @('ConvertTo-DelegationType','New-PrivilegedMemberRow','New-ServiceAccountRow','New-AdminCountRow','Get-AdPrivilegedIdentityFacts','Get-AdPrivilegedIdentityInventory')) {
        Assert-True -Condition ($functionNames -contains $name) -Name "Function exists: $name"
    }
}
Test-RequiredFunctionsExist

function Test-ReadOnly {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw
    Assert-True -Condition ($scriptText -notmatch 'Set-ADUser' -and $scriptText -notmatch 'Remove-ADGroupMember' -and $scriptText -notmatch 'Add-ADGroupMember' -and $scriptText -notmatch 'Set-ADAccountControl') -Name 'Script performs no AD write operations'
    Assert-True -Condition ($scriptText -match 'Get-ADGroupMember' -and $scriptText -match 'adminCount=1') -Name 'Group membership and adminCount queries are present'
}
Test-ReadOnly

function Test-DelegationClassification {
    Assert-True -Condition ((ConvertTo-DelegationType -TrustedForDelegation $true -TrustedToAuthForDelegation $false -AllowedToDelegateTo @() -PrincipalsAllowedToDelegateToAccount @()) -eq 'Unconstrained') -Name 'Unconstrained delegation classified'
    Assert-True -Condition ((ConvertTo-DelegationType -TrustedForDelegation $false -TrustedToAuthForDelegation $false -AllowedToDelegateTo @('HOST/sql01') -PrincipalsAllowedToDelegateToAccount @()) -eq 'Constrained') -Name 'Constrained (msDS-AllowedToDelegateTo) classified'
    Assert-True -Condition ((ConvertTo-DelegationType -TrustedForDelegation $false -TrustedToAuthForDelegation $false -AllowedToDelegateTo @() -PrincipalsAllowedToDelegateToAccount @('CN=web01')) -eq 'ResourceBased') -Name 'Resource-based delegation classified'
    Assert-True -Condition ((ConvertTo-DelegationType -TrustedForDelegation $false -TrustedToAuthForDelegation $false -AllowedToDelegateTo @() -PrincipalsAllowedToDelegateToAccount @()) -eq 'None') -Name 'No delegation classified'
}
Test-DelegationClassification

function New-MockPrivilegedFacts {
    $now = Get-Date '2026-07-17'
    $recent = $now.AddDays(-5).ToFileTimeUtc()
    $old = $now.AddDays(-300).ToFileTimeUtc()
    return @{
        GroupMembers = @(
            [pscustomobject]@{ GroupName = 'Domain Admins'; Member = [pscustomobject]@{ SamAccountName = 'admin1'; Enabled = $true;  lastLogonTimestamp = $recent; objectClass = 'user'; DistinguishedName = 'CN=admin1,DC=contoso,DC=com' } },
            [pscustomobject]@{ GroupName = 'Domain Admins'; Member = [pscustomobject]@{ SamAccountName = 'staleadmin'; Enabled = $true; lastLogonTimestamp = $old; objectClass = 'user'; DistinguishedName = 'CN=staleadmin,DC=contoso,DC=com' } },
            [pscustomobject]@{ GroupName = 'Enterprise Admins'; Member = [pscustomobject]@{ SamAccountName = 'ea1'; Enabled = $false; lastLogonTimestamp = $null; objectClass = 'user'; DistinguishedName = 'CN=ea1,DC=contoso,DC=com' } }
        )
        AdminCountObjects = @(
            [pscustomobject]@{ SamAccountName = 'orphaned'; Enabled = $true; objectClass = 'user'; DistinguishedName = 'CN=orphaned,DC=contoso,DC=com' }
        )
        ServiceAccounts = @(
            [pscustomobject]@{ SamAccountName = 'svc-sql'; Enabled = $true; objectClass = 'user'; servicePrincipalName = @('MSSQLSvc/sql01:1433','MSSQLSvc/sql01.contoso.com:1433'); TrustedForDelegation = $true; TrustedToAuthForDelegation = $false; 'msDS-AllowedToDelegateTo' = @(); PrincipalsAllowedToDelegateToAccount = @(); PasswordNeverExpires = $true },
            [pscustomobject]@{ SamAccountName = 'gmsa-web$'; Enabled = $true; objectClass = 'msDS-GroupManagedServiceAccount'; servicePrincipalName = @('HTTP/web01'); TrustedForDelegation = $false; TrustedToAuthForDelegation = $false; 'msDS-AllowedToDelegateTo' = @(); PrincipalsAllowedToDelegateToAccount = @(); PasswordNeverExpires = $false }
        )
    }
}

function Test-InventoryShaping {
    $inventory = Get-AdPrivilegedIdentityInventory -Facts (New-MockPrivilegedFacts) -StaleDays 90 -Now (Get-Date '2026-07-17')

    Assert-True -Condition (@($inventory.PrivilegedGroupMembers).Count -eq 3) -Name 'One row per privileged (group, member)'
    $stale = $inventory.PrivilegedGroupMembers | Where-Object MemberName -eq 'staleadmin'
    Assert-True -Condition ($stale.StaleMember -eq $true -and $stale.GroupName -eq 'Domain Admins') -Name 'Stale privileged member flagged'
    $active = $inventory.PrivilegedGroupMembers | Where-Object MemberName -eq 'admin1'
    Assert-True -Condition ($active.StaleMember -eq $false) -Name 'Recently used privileged member not stale'
    $disabledEa = $inventory.PrivilegedGroupMembers | Where-Object MemberName -eq 'ea1'
    Assert-True -Condition ($disabledEa.Enabled -eq $false -and $disabledEa.StaleMember -eq $true) -Name 'Disabled never-logged-on member flagged stale'

    Assert-True -Condition (@($inventory.AdminCountObjects).Count -eq 1 -and $inventory.AdminCountObjects[0].SamAccountName -eq 'orphaned') -Name 'adminCount object shaped'

    $svcSql = $inventory.ServiceAccounts | Where-Object SamAccountName -eq 'svc-sql'
    Assert-True -Condition ($svcSql.SpnCount -eq 2 -and $svcSql.DelegationType -eq 'Unconstrained') -Name 'SPN account: SPN count + unconstrained delegation'
    $gmsa = $inventory.ServiceAccounts | Where-Object SamAccountName -eq 'gmsa-web$'
    Assert-True -Condition ($gmsa.ObjectClass -eq 'msDS-GroupManagedServiceAccount' -and $gmsa.DelegationType -eq 'None') -Name 'gMSA classified with no delegation'
}
Test-InventoryShaping

if ($script:Failures.Count -gt 0) {
    Write-Host "$($script:Failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host "All $($script:Passes) checks passed." -ForegroundColor Green
exit 0
