#Requires -Version 7
$ErrorActionPreference = 'Stop'

$script:ScriptPath = Join-Path $PSScriptRoot '..' 'Export-AdSidHistoryInventory.ps1'
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
    foreach ($name in @('Get-DomainSidFromSid','New-SidHistoryRow','Get-AdSidHistoryFacts','Get-AdSidHistoryInventory')) {
        Assert-True -Condition ($functionNames -contains $name) -Name "Function exists: $name"
    }
}
Test-RequiredFunctionsExist

function Test-ReadOnly {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw
    Assert-True -Condition ($scriptText -notmatch 'Set-AD' -and $scriptText -notmatch 'Remove-AD') -Name 'Script performs no AD write operations'
    Assert-True -Condition ($scriptText -match 'sIDHistory=\*') -Name 'SIDHistory LDAP filter is present'
}
Test-ReadOnly

function Test-DomainSidExtraction {
    Assert-True -Condition ((Get-DomainSidFromSid -Sid 'S-1-5-21-1111111111-2222222222-3333333333-1105') -eq 'S-1-5-21-1111111111-2222222222-3333333333') -Name 'Domain SID = account SID minus RID'
    Assert-True -Condition ((Get-DomainSidFromSid -Sid 'S-1-5-32-544') -eq 'S-1-5-32-544') -Name 'Well-known SID returned unchanged'
}
Test-DomainSidExtraction

function Test-SidHistoryShaping {
    $facts = @{
        Objects = @(
            [pscustomobject]@{
                SamAccountName = 'user1'; Enabled = $true; objectClass = 'user'; DistinguishedName = 'CN=user1,DC=contoso,DC=com'
                SIDHistory = @('S-1-5-21-1111111111-2222222222-3333333333-1105','S-1-5-21-1111111111-2222222222-3333333333-1106')
            },
            [pscustomobject]@{
                SamAccountName = 'grp1'; Enabled = $true; objectClass = 'group'; DistinguishedName = 'CN=grp1,DC=contoso,DC=com'
                SIDHistory = @('S-1-5-21-1111111111-2222222222-3333333333-1200','S-1-5-21-9999999999-8888888888-7777777777-1300')
            }
        )
    }
    $rows = @(Get-AdSidHistoryInventory -Facts $facts)
    Assert-True -Condition ($rows.Count -eq 2) -Name 'One row per object with SIDHistory'

    $user1 = $rows | Where-Object SamAccountName -eq 'user1'
    Assert-True -Condition ($user1.SidHistoryCount -eq 2 -and $user1.SourceDomainCount -eq 1) -Name 'Two SIDs from one source domain'

    $grp1 = $rows | Where-Object SamAccountName -eq 'grp1'
    Assert-True -Condition ($grp1.SidHistoryCount -eq 2 -and $grp1.SourceDomainCount -eq 2) -Name 'SIDs from two distinct source domains counted'
    Assert-True -Condition ($grp1.ObjectClass -eq 'group') -Name 'Object class captured'
}
Test-SidHistoryShaping

if ($script:Failures.Count -gt 0) {
    Write-Host "$($script:Failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host "All $($script:Passes) checks passed." -ForegroundColor Green
exit 0
