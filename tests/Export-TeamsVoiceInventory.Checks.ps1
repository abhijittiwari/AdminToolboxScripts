#Requires -Version 7
$ErrorActionPreference = 'Stop'

$script:ScriptPath = Join-Path $PSScriptRoot '..' 'Export-TeamsVoiceInventory.ps1'
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
    foreach ($name in @('New-TeamRow','Get-PhoneNumberTypeSummary','Get-TeamsVoiceInventory','Get-TeamsVoiceFacts')) {
        Assert-True -Condition ($functionNames -contains $name) -Name "Function exists: $name"
    }
}
Test-RequiredFunctionsExist

function Test-ReadOnly {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw
    Assert-True -Condition ($scriptText -notmatch 'Set-CsPhoneNumberAssignment' -and $scriptText -notmatch 'Remove-Team\b' -and $scriptText -notmatch 'New-Team\b' -and $scriptText -notmatch 'Set-CsUser') -Name 'Script performs no Teams/Voice write operations'
    Assert-True -Condition ($scriptText -match 'Get-Team\b' -and $scriptText -match 'Get-CsPhoneNumberAssignment') -Name 'Teams and phone-number queries are present'
    Assert-True -Condition ($scriptText -match 'excluded pending vendor selection') -Name 'SoW Voice-scope tension is noted'
}
Test-ReadOnly

function New-MockTeamsVoiceFacts {
    return @{
        Teams = @(
            [pscustomobject]@{ DisplayName = 'Engineering'; GroupId = 'g1'; Visibility = 'Private'; Archived = $false; StandardChannelCount = 5; PrivateChannelCount = 2; SharedChannelCount = 1 },
            [pscustomobject]@{ DisplayName = 'All Company'; GroupId = 'g2'; Visibility = 'Public'; Archived = $false; StandardChannelCount = 3; PrivateChannelCount = 0; SharedChannelCount = 0 },
            [pscustomobject]@{ DisplayName = 'Old Project'; GroupId = 'g3'; Visibility = 'Private'; Archived = $true; StandardChannelCount = 1; PrivateChannelCount = 0; SharedChannelCount = 0 }
        )
        PhoneNumbers = @(
            [pscustomobject]@{ NumberType = 'CallingPlan'; AssignedUpn = 'user1@contoso.com' },
            [pscustomobject]@{ NumberType = 'CallingPlan'; AssignedUpn = $null },
            [pscustomobject]@{ NumberType = 'DirectRouting'; AssignedUpn = 'user2@contoso.com' }
        )
        AutoAttendantCount = 4
        CallQueueCount = 7
        VoiceRoutingPolicyCount = 3
        DialPlanCount = 2
    }
}

function Test-InventoryShaping {
    $inventory = Get-TeamsVoiceInventory -Facts (New-MockTeamsVoiceFacts)

    $ts = $inventory.TeamsSummary
    Assert-True -Condition ($ts.TotalTeams -eq 3 -and $ts.PublicTeams -eq 1 -and $ts.PrivateTeams -eq 2 -and $ts.ArchivedTeams -eq 1) -Name 'Team public/private/archived counts'
    Assert-True -Condition ($ts.TotalStandardChannels -eq 9 -and $ts.TotalPrivateChannels -eq 2 -and $ts.TotalSharedChannels -eq 1) -Name 'Channel totals incl. private/shared split'

    $vs = $inventory.VoiceSummary
    Assert-True -Condition ($vs.TotalPhoneNumbers -eq 3 -and $vs.AutoAttendants -eq 4 -and $vs.CallQueues -eq 7) -Name 'Voice summary counts'

    $callingPlan = $inventory.PhoneNumberTypes | Where-Object NumberType -eq 'CallingPlan'
    Assert-True -Condition ($callingPlan.TotalNumbers -eq 2 -and $callingPlan.AssignedToUser -eq 1 -and $callingPlan.Unassigned -eq 1) -Name 'Phone numbers grouped by type with assignment split'
}
Test-InventoryShaping

function Test-EmptyFacts {
    $inventory = Get-TeamsVoiceInventory -Facts @{ Teams = @(); PhoneNumbers = @(); AutoAttendantCount = 0; CallQueueCount = 0; VoiceRoutingPolicyCount = 0; DialPlanCount = 0 }
    Assert-True -Condition ($inventory.TeamsSummary.TotalTeams -eq 0 -and @($inventory.PhoneNumberTypes).Count -eq 0) -Name 'Empty facts produce zeroed summary and no number rows'
}
Test-EmptyFacts

if ($script:Failures.Count -gt 0) {
    Write-Host "$($script:Failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host "All $($script:Passes) checks passed." -ForegroundColor Green
exit 0
