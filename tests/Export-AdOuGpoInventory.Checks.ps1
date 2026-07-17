#Requires -Version 7
$ErrorActionPreference = 'Stop'

$script:ScriptPath = Join-Path $PSScriptRoot '..' 'Export-AdOuGpoInventory.ps1'
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
    foreach ($name in @('Get-OuDepth','Test-GpoIntuneCandidate','ConvertFrom-GpoReportXml','New-GpoRow','New-OuRow','Get-LinkedGpoGuids','Get-AdOuGpoFacts','Get-AdOuGpoInventory')) {
        Assert-True -Condition ($functionNames -contains $name) -Name "Function exists: $name"
    }
}
Test-RequiredFunctionsExist

function Test-ReadOnly {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw
    Assert-True -Condition ($scriptText -notmatch 'Set-GPO\b' -and $scriptText -notmatch 'New-GPO\b' -and $scriptText -notmatch 'Remove-GPO\b' -and $scriptText -notmatch 'New-GPLink\b' -and $scriptText -notmatch 'Set-ADOrganizationalUnit\b') -Name 'Script performs no GPO/OU write operations'
    Assert-True -Condition ($scriptText -match 'Get-GPO' -and $scriptText -match 'Get-GPOReport') -Name 'GPO and GPO report queries are present'
}
Test-ReadOnly

function Test-OuDepthAndGuidParsing {
    Assert-True -Condition ((Get-OuDepth -DistinguishedName 'OU=Servers,OU=Perth,DC=contoso,DC=com') -eq 2) -Name 'OU depth counts OU components'
    Assert-True -Condition ((Get-OuDepth -DistinguishedName 'DC=contoso,DC=com') -eq 0) -Name 'Domain root has depth 0'
    $guids = @(Get-LinkedGpoGuids -GpLink '[LDAP://cn={31B2F340-016D-11D2-945F-00C04FB984F9},cn=policies,cn=system,DC=contoso,DC=com;0][LDAP://cn={AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE},cn=policies;0]')
    Assert-True -Condition ($guids.Count -eq 2 -and $guids[0] -eq '{31b2f340-016d-11d2-945f-00c04fb984f9}') -Name 'gPLink GUIDs parsed and lowercased'
    Assert-True -Condition ((Get-LinkedGpoGuids -GpLink '').Count -eq 0) -Name 'Empty gPLink yields no GUIDs'
}
Test-OuDepthAndGuidParsing

function Test-IntuneCandidateHeuristic {
    Assert-True -Condition (Test-GpoIntuneCandidate -ExtensionNames @('Administrative Templates','Registry')) -Name 'Admin templates/registry flagged as Intune candidate'
    Assert-True -Condition (-not (Test-GpoIntuneCandidate -ExtensionNames @('Folder Redirection'))) -Name 'Folder redirection is not an Intune candidate'
    Assert-True -Condition (-not (Test-GpoIntuneCandidate -ExtensionNames @('Administrative Templates','Scripts'))) -Name 'Scripts blocker overrides deliverable extensions'
    Assert-True -Condition (-not (Test-GpoIntuneCandidate -ExtensionNames @())) -Name 'No configured extensions is not an Intune candidate'
}
Test-IntuneCandidateHeuristic

function Test-GpoReportParsing {
    $xml = [xml]@'
<GPO xmlns="http://www.microsoft.com/GroupPolicy/Settings">
  <Computer>
    <Enabled>true</Enabled>
    <ExtensionData><Name>Registry</Name></ExtensionData>
    <ExtensionData><Name>Security</Name></ExtensionData>
  </Computer>
  <User>
    <Enabled>true</Enabled>
  </User>
</GPO>
'@
    $info = ConvertFrom-GpoReportXml -ReportXml $xml
    Assert-True -Condition ($info.ComputerSettingsConfigured -eq $true) -Name 'Computer settings detected as configured'
    Assert-True -Condition ($info.UserSettingsConfigured -eq $false) -Name 'User section with no extensions is not configured'
    Assert-True -Condition (@($info.ExtensionNames).Count -eq 2 -and ($info.ExtensionNames -contains 'Registry')) -Name 'Extension names extracted from report'
}
Test-GpoReportParsing

function Test-InventoryShaping {
    $guid1 = '4d5e6f7a-1111-2222-3333-444455556666'
    $guid2 = 'aaaabbbb-cccc-dddd-eeee-ffff00001111'
    $facts = @{
        Gpos = @(
            [pscustomobject]@{
                Gpo = [pscustomobject]@{ DisplayName = 'Baseline'; Id = $guid1; GpoStatus = 'AllSettingsEnabled'; WmiFilter = $null; CreationTime = (Get-Date); ModificationTime = (Get-Date) }
                ReportInfo = [pscustomobject]@{ ComputerSettingsConfigured = $true; UserSettingsConfigured = $false; ExtensionNames = @('Administrative Templates') }
            },
            [pscustomobject]@{
                Gpo = [pscustomobject]@{ DisplayName = 'Legacy-FolderRedirect'; Id = $guid2; GpoStatus = 'UserSettingsDisabled'; WmiFilter = $null; CreationTime = (Get-Date); ModificationTime = (Get-Date) }
                ReportInfo = [pscustomobject]@{ ComputerSettingsConfigured = $false; UserSettingsConfigured = $true; ExtensionNames = @('Folder Redirection') }
            }
        )
        Ous = @(
            [pscustomobject]@{ DistinguishedName = 'OU=Workstations,DC=contoso,DC=com'; Name = 'Workstations'; BlockInheritance = $false; gPLink = "[LDAP://cn={$guid1},cn=policies,cn=system,DC=contoso,DC=com;0]" },
            [pscustomobject]@{ DistinguishedName = 'OU=Empty,DC=contoso,DC=com'; Name = 'Empty'; BlockInheritance = $true; gPLink = $null }
        )
        GpoNameById = @{ "{$guid1}" = 'Baseline'; "{$guid2}" = 'Legacy-FolderRedirect' }
    }

    $inventory = Get-AdOuGpoInventory -Facts $facts
    $byName = @{}
    foreach ($g in $inventory.Gpos) { $byName[$g.DisplayName] = $g }

    Assert-True -Condition ($byName['Baseline'].LinkCount -eq 1 -and $byName['Baseline'].IsUnlinked -eq $false -and $byName['Baseline'].IntuneCandidate -eq $true) -Name 'Linked admin-template GPO: link counted, Intune candidate'
    Assert-True -Condition ($byName['Legacy-FolderRedirect'].LinkCount -eq 0 -and $byName['Legacy-FolderRedirect'].IsUnlinked -eq $true -and $byName['Legacy-FolderRedirect'].IntuneCandidate -eq $false) -Name 'Unlinked folder-redirect GPO: unlinked, not Intune candidate'

    $workstationsOu = $inventory.OrganizationalUnits | Where-Object Name -eq 'Workstations'
    Assert-True -Condition ($workstationsOu.LinkedGpoCount -eq 1 -and $workstationsOu.LinkedGpos -eq 'Baseline' -and $workstationsOu.Depth -eq 1) -Name 'OU resolves linked GPO name, count, and depth'
    $emptyOu = $inventory.OrganizationalUnits | Where-Object Name -eq 'Empty'
    Assert-True -Condition ($emptyOu.LinkedGpoCount -eq 0 -and $emptyOu.BlockInheritance -eq $true) -Name 'OU with no links and blocked inheritance shapes correctly'
}
Test-InventoryShaping

if ($script:Failures.Count -gt 0) {
    Write-Host "$($script:Failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host "All $($script:Passes) checks passed." -ForegroundColor Green
exit 0
