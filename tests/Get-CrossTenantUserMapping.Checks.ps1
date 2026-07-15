#Requires -Version 7
$ErrorActionPreference = 'Stop'

$script:ScriptPath = Join-Path $PSScriptRoot '..' 'Get-CrossTenantUserMapping.ps1'
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
$script:MatchAttributeNames = @('ExtensionAttribute13', 'ExtensionAttribute14')

function New-TestTargetUser {
    param(
        [string]$Id,
        [string]$Upn,
        [string]$DisplayName = 'Test User',
        [string]$Attribute13 = '',
        [string]$Attribute14 = ''
    )

    return [pscustomobject]@{
        Id                            = $Id
        UserPrincipalName             = $Upn
        DisplayName                   = $DisplayName
        AccountEnabled                = $true
        OnPremisesSyncEnabled         = $true
        OnPremisesExtensionAttributes = [pscustomobject]@{
            ExtensionAttribute13 = $Attribute13
            ExtensionAttribute14 = $Attribute14
        }
    }
}

function Test-RequiredFunctionsExist {
    $ast = Get-ScriptAst
    $functionNames = @($ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true) | ForEach-Object Name)
    foreach ($name in @('ConvertTo-UpnKey','Join-ReportValues','Get-UserExtensionAttributeValue','New-TargetAttributeIndex','Find-TargetUserMatches','New-UserMappingRow','Connect-TenantGraph','Get-TenantUsers')) {
        Assert-True -Condition ($functionNames -contains $name) -Name "Function exists: $name"
    }
}
Test-RequiredFunctionsExist

function Test-StaticScriptContent {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw
    Assert-True -Condition ($scriptText -match 'User\.Read\.All') -Name 'Graph scope is present'
    Assert-True -Condition ($scriptText -match 'onPremisesExtensionAttributes') -Name 'Extension attributes property is requested'
    Assert-True -Condition ($scriptText -match 'ExtensionAttribute13' -and $scriptText -match 'ExtensionAttribute14') -Name 'Extension attributes 13 and 14 are used'
    Assert-True -Condition ($scriptText -match 'Disconnect-MgGraph') -Name 'Graph sessions are disconnected between tenants'
    Assert-True -Condition ($scriptText -notmatch 'Import-Module\s+MSOnline' -and $scriptText -notmatch 'Get-Msol' -and $scriptText -notmatch 'Connect-Msol') -Name 'MSOnline cmdlets are not used'
}
Test-StaticScriptContent

function Test-ConvertToUpnKey {
    Assert-True -Condition ((ConvertTo-UpnKey -Value '  User@Contoso.COM ') -eq 'user@contoso.com') -Name 'ConvertTo-UpnKey trims and lowercases'
    Assert-True -Condition ((ConvertTo-UpnKey -Value '') -eq '') -Name 'ConvertTo-UpnKey handles empty input'
}
Test-ConvertToUpnKey

function Test-GetUserExtensionAttributeValue {
    $objectUser = New-TestTargetUser -Id 't1' -Upn 't1@fabrikam.com' -Attribute13 'user@contoso.com'
    $hashtableUser = [pscustomobject]@{ OnPremisesExtensionAttributes = @{ ExtensionAttribute14 = 'other@contoso.com' } }
    $nullUser = [pscustomobject]@{ OnPremisesExtensionAttributes = $null }

    Assert-True -Condition ((Get-UserExtensionAttributeValue -User $objectUser -AttributeName 'ExtensionAttribute13') -eq 'user@contoso.com') -Name 'Extension attribute is read from object'
    Assert-True -Condition ((Get-UserExtensionAttributeValue -User $hashtableUser -AttributeName 'ExtensionAttribute14') -eq 'other@contoso.com') -Name 'Extension attribute is read from hashtable'
    Assert-True -Condition ((Get-UserExtensionAttributeValue -User $nullUser -AttributeName 'ExtensionAttribute13') -eq '') -Name 'Null extension attributes return empty string'
}
Test-GetUserExtensionAttributeValue

function Test-NewTargetAttributeIndex {
    $targetUsers = @(
        (New-TestTargetUser -Id 't1' -Upn 't1@fabrikam.com' -Attribute13 'User.One@Contoso.com'),
        (New-TestTargetUser -Id 't2' -Upn 't2@fabrikam.com' -Attribute14 'user.two@contoso.com'),
        (New-TestTargetUser -Id 't3' -Upn 't3@fabrikam.com')
    )

    $index = New-TargetAttributeIndex -TargetUsers $targetUsers

    Assert-True -Condition ($index.ContainsKey('user.one@contoso.com')) -Name 'Index key is normalized to lowercase'
    Assert-True -Condition ($index.ContainsKey('user.two@contoso.com')) -Name 'ExtensionAttribute14 values are indexed'
    Assert-True -Condition ($index.Count -eq 2) -Name 'Blank attribute values are not indexed'
    Assert-True -Condition ($index['user.one@contoso.com'][0].Attribute -eq 'ExtensionAttribute13') -Name 'Index entry records the matched attribute'
}
Test-NewTargetAttributeIndex

function Test-FindTargetUserMatchesExact {
    $targetUsers = @(
        (New-TestTargetUser -Id 't1' -Upn 't1@fabrikam.com' -Attribute13 'user.one@contoso.com'),
        (New-TestTargetUser -Id 't2' -Upn 't2@fabrikam.com' -Attribute14 'user.two@contoso.com')
    )
    $index = New-TargetAttributeIndex -TargetUsers $targetUsers

    $matchResults13 = @(Find-TargetUserMatches -SourceUserPrincipalName 'User.One@Contoso.com' -AttributeIndex $index -TargetUsers $targetUsers)
    $matchResults14 = @(Find-TargetUserMatches -SourceUserPrincipalName 'user.two@contoso.com' -AttributeIndex $index -TargetUsers $targetUsers)
    $noMatches = @(Find-TargetUserMatches -SourceUserPrincipalName 'missing@contoso.com' -AttributeIndex $index -TargetUsers $targetUsers)

    Assert-True -Condition ($matchResults13.Count -eq 1 -and $matchResults13[0].User.Id -eq 't1' -and $matchResults13[0].MatchType -eq 'Exact') -Name 'Exact match on ExtensionAttribute13 is case-insensitive'
    Assert-True -Condition ($matchResults14.Count -eq 1 -and $matchResults14[0].User.Id -eq 't2' -and $matchResults14[0].MatchedAttributes -eq 'ExtensionAttribute14') -Name 'Exact match on ExtensionAttribute14 is found'
    Assert-True -Condition ($noMatches.Count -eq 0) -Name 'Unmatched UPN returns no matches'
}
Test-FindTargetUserMatchesExact

function Test-FindTargetUserMatchesCollapsesDuplicateUser {
    $targetUsers = @(
        (New-TestTargetUser -Id 't1' -Upn 't1@fabrikam.com' -Attribute13 'user.one@contoso.com' -Attribute14 'user.one@contoso.com')
    )
    $index = New-TargetAttributeIndex -TargetUsers $targetUsers

    $matchResults = @(Find-TargetUserMatches -SourceUserPrincipalName 'user.one@contoso.com' -AttributeIndex $index -TargetUsers $targetUsers)

    Assert-True -Condition ($matchResults.Count -eq 1) -Name 'UPN in both attributes yields a single match'
    Assert-True -Condition ($matchResults[0].MatchedAttributes -eq 'ExtensionAttribute13; ExtensionAttribute14') -Name 'Collapsed match lists both attributes'
}
Test-FindTargetUserMatchesCollapsesDuplicateUser

function Test-FindTargetUserMatchesContainsFallback {
    $targetUsers = @(
        (New-TestTargetUser -Id 't1' -Upn 't1@fabrikam.com' -Attribute13 'SMTP:user.one@contoso.com;source-upn')
    )
    $index = New-TargetAttributeIndex -TargetUsers $targetUsers

    $matchResults = @(Find-TargetUserMatches -SourceUserPrincipalName 'user.one@contoso.com' -AttributeIndex $index -TargetUsers $targetUsers)

    Assert-True -Condition ($matchResults.Count -eq 1 -and $matchResults[0].User.Id -eq 't1') -Name 'Contains fallback finds UPN embedded in attribute value'
    Assert-True -Condition ($matchResults[0].MatchType -eq 'Contains') -Name 'Contains fallback is labeled as Contains match'
}
Test-FindTargetUserMatchesContainsFallback

function Test-NewUserMappingRowStatuses {
    $sourceUser = [pscustomobject]@{
        Id                = 's1'
        UserPrincipalName = 'user.one@contoso.com'
        DisplayName       = 'User One'
        AccountEnabled    = $true
    }
    $targetUsers = @(
        (New-TestTargetUser -Id 't1' -Upn 't1@fabrikam.com' -DisplayName 'User One (Fabrikam)' -Attribute13 'user.one@contoso.com'),
        (New-TestTargetUser -Id 't2' -Upn 't2@fabrikam.com' -DisplayName 'User One Dup' -Attribute14 'user.one@contoso.com')
    )
    $index = New-TargetAttributeIndex -TargetUsers $targetUsers

    $singleMatch = @(Find-TargetUserMatches -SourceUserPrincipalName 'user.one@contoso.com' -AttributeIndex (New-TargetAttributeIndex -TargetUsers @($targetUsers[0])) -TargetUsers @($targetUsers[0]))
    $matchedRow = New-UserMappingRow -SourceUser $sourceUser -MatchResults $singleMatch
    $ambiguousRow = New-UserMappingRow -SourceUser $sourceUser -MatchResults @(Find-TargetUserMatches -SourceUserPrincipalName 'user.one@contoso.com' -AttributeIndex $index -TargetUsers $targetUsers)
    $notFoundRow = New-UserMappingRow -SourceUser $sourceUser -MatchResults @()

    Assert-True -Condition ($matchedRow.MatchStatus -eq 'Matched' -and $matchedRow.TargetUserPrincipalName -eq 't1@fabrikam.com' -and $matchedRow.TargetUserId -eq 't1') -Name 'Single match row has Matched status and target details'
    Assert-True -Condition ($matchedRow.SourceUserPrincipalName -eq 'user.one@contoso.com' -and $matchedRow.SourceUserId -eq 's1') -Name 'Row carries source user details'
    Assert-True -Condition ($matchedRow.MatchedAttribute -eq 'ExtensionAttribute13' -and $matchedRow.TargetExtensionAttribute13 -eq 'user.one@contoso.com') -Name 'Row reports matched attribute and its value'
    Assert-True -Condition ($ambiguousRow.MatchStatus -eq 'Ambiguous' -and $ambiguousRow.TargetUserId -eq 't1; t2') -Name 'Multiple matches produce Ambiguous status with joined targets'
    Assert-True -Condition ($notFoundRow.MatchStatus -eq 'NotFound' -and $notFoundRow.TargetUserPrincipalName -eq '') -Name 'No matches produce NotFound status with empty target fields'
}
Test-NewUserMappingRowStatuses

if ($script:Failures.Count -gt 0) {
    Write-Host "$($script:Failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host "All $($script:Passes) checks passed." -ForegroundColor Green
exit 0
