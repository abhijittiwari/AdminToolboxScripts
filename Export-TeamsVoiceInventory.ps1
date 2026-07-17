<#
.SYNOPSIS
    Exports a Microsoft Teams and Teams Voice inventory for divestiture discovery.

.DESCRIPTION
    Read-only discovery of the Teams estate for migration scoping. The SoW flags cross-tenant Teams as
    the hardest workload and notes a scope tension (Voice design is in scope, but Voice migration/testing
    is excluded pending vendor selection) - this inventory quantifies both so the decision can be made:
      - Teams: team count, a public/private split, and channel counts including the private/shared
        channel split (private and shared channels are the parts that migrate worst).
      - Voice - numbers: assigned phone numbers by type (Calling Plan / Operator Connect / Direct
        Routing) and user vs service assignment.
      - Voice - routing: auto attendants and call queues (the resource-account-backed workflows).
      - Voice - policies: voice-routing / dial-plan policy counts.

    NOTE: Per the SoW, Voice migration/testing may be descoped pending the network squad's vendor
    selection; this script inventories Voice for design only and makes no changes.

    Collection uses the MicrosoftTeams module; the shaping is separated so it is unit-testable with mock
    data. The script makes no changes.

.PARAMETER Include
    Which sections to collect. Defaults to Teams, Voice.

.PARAMETER OutputFolder
    Optional folder for the CSV outputs (Teams.csv, PhoneNumbers.csv, VoiceApps.csv). If omitted,
    results are only written to the pipeline/table.

.EXAMPLE
    .\Export-TeamsVoiceInventory.ps1 -OutputFolder .\TeamsVoice
#>

[CmdletBinding()]
param(
    [ValidateSet('Teams','Voice')]
    [string[]]$Include = @('Teams','Voice'),
    [string]$OutputFolder = ''
)

$ErrorActionPreference = 'Stop'

function New-TeamRow {
    param([Parameter(Mandatory = $true)] $Team)
    return [pscustomobject]@{
        DisplayName        = [string]$Team.DisplayName
        GroupId            = [string]$Team.GroupId
        Visibility         = [string]$Team.Visibility
        Archived           = [bool]$Team.Archived
        StandardChannels   = [int]$Team.StandardChannelCount
        PrivateChannels    = [int]$Team.PrivateChannelCount
        SharedChannels     = [int]$Team.SharedChannelCount
    }
}

function Get-PhoneNumberTypeSummary {
    param([Parameter(Mandatory = $true)] [AllowEmptyCollection()] $PhoneNumbers)

    $rows = [System.Collections.Generic.List[object]]::new()
    $groups = @($PhoneNumbers) | Group-Object -Property { [string]$_.NumberType }
    foreach ($group in $groups) {
        $assignedToUser = @($group.Group | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.AssignedUpn) })
        $rows.Add([pscustomobject]@{
            NumberType      = if ([string]::IsNullOrWhiteSpace($group.Name)) { 'Unknown' } else { $group.Name }
            TotalNumbers    = @($group.Group).Count
            AssignedToUser  = $assignedToUser.Count
            Unassigned      = (@($group.Group).Count - $assignedToUser.Count)
        })
    }
    return $rows.ToArray()
}

function Get-TeamsVoiceInventory {
    param([Parameter(Mandatory = $true)] $Facts)

    $teamRows = @(@($Facts.Teams) | ForEach-Object { New-TeamRow -Team $_ })

    $teamsSummary = [pscustomobject]@{
        TotalTeams        = @($teamRows).Count
        PublicTeams       = @($teamRows | Where-Object { $_.Visibility -eq 'Public' }).Count
        PrivateTeams      = @($teamRows | Where-Object { $_.Visibility -eq 'Private' }).Count
        ArchivedTeams     = @($teamRows | Where-Object { $_.Archived }).Count
        TotalStandardChannels = (@($teamRows | Measure-Object -Property StandardChannels -Sum).Sum)
        TotalPrivateChannels  = (@($teamRows | Measure-Object -Property PrivateChannels -Sum).Sum)
        TotalSharedChannels   = (@($teamRows | Measure-Object -Property SharedChannels -Sum).Sum)
    }

    $voiceSummary = [pscustomobject]@{
        TotalPhoneNumbers = @($Facts.PhoneNumbers).Count
        AutoAttendants    = [int]$Facts.AutoAttendantCount
        CallQueues        = [int]$Facts.CallQueueCount
        VoiceRoutingPolicies = [int]$Facts.VoiceRoutingPolicyCount
        DialPlans         = [int]$Facts.DialPlanCount
    }

    return @{
        Teams          = $teamRows
        TeamsSummary   = $teamsSummary
        PhoneNumberTypes = (Get-PhoneNumberTypeSummary -PhoneNumbers @($Facts.PhoneNumbers))
        VoiceSummary   = $voiceSummary
    }
}

function Get-TeamsVoiceFacts {
    param([string[]]$Include)

    $facts = @{
        Teams = @(); PhoneNumbers = @(); AutoAttendantCount = 0; CallQueueCount = 0
        VoiceRoutingPolicyCount = 0; DialPlanCount = 0
    }

    if (-not (Get-Command -Name Get-Team -ErrorAction SilentlyContinue)) {
        throw 'The MicrosoftTeams module is not connected. Run Connect-MicrosoftTeams first.'
    }

    if ($Include -contains 'Teams') {
        try {
            foreach ($team in @(Get-Team -ErrorAction Stop)) {
                $standard = 0; $private = 0; $shared = 0
                try { $standard = @(Get-TeamChannel -GroupId $team.GroupId -ErrorAction Stop).Count } catch { }
                try { $private = @(Get-TeamChannel -GroupId $team.GroupId -MembershipType Private -ErrorAction SilentlyContinue).Count } catch { }
                try { $shared = @(Get-TeamChannel -GroupId $team.GroupId -MembershipType Shared -ErrorAction SilentlyContinue).Count } catch { }
                $facts.Teams += [pscustomobject]@{
                    DisplayName = $team.DisplayName; GroupId = $team.GroupId; Visibility = $team.Visibility; Archived = $team.Archived
                    StandardChannelCount = $standard; PrivateChannelCount = $private; SharedChannelCount = $shared
                }
            }
        } catch { }
    }

    if ($Include -contains 'Voice') {
        try {
            $facts.PhoneNumbers = @(Get-CsPhoneNumberAssignment -ErrorAction Stop | ForEach-Object {
                [pscustomobject]@{ NumberType = $_.NumberType; AssignedUpn = $_.AssignedPstnTargetId }
            })
        } catch { }
        try { $facts.AutoAttendantCount = @(Get-CsAutoAttendant -ErrorAction Stop).Count } catch { }
        try { $facts.CallQueueCount = @(Get-CsCallQueue -ErrorAction Stop).Count } catch { }
        try { $facts.VoiceRoutingPolicyCount = @(Get-CsOnlineVoiceRoutingPolicy -ErrorAction Stop).Count } catch { }
        try { $facts.DialPlanCount = @(Get-CsTenantDialPlan -ErrorAction Stop).Count } catch { }
    }

    return $facts
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
Write-Host ("Collecting Microsoft Teams / Voice inventory: {0}" -f ($Include -join ', ')) -ForegroundColor Cyan
$facts = Get-TeamsVoiceFacts -Include $Include
$inventory = Get-TeamsVoiceInventory -Facts $facts

Write-Host "`n=== Teams summary ===" -ForegroundColor Yellow
$inventory.TeamsSummary | Format-List
Write-Host "=== Voice summary ===" -ForegroundColor Yellow
$inventory.VoiceSummary | Format-List
Write-Host "=== Phone numbers by type ===" -ForegroundColor Yellow
$inventory.PhoneNumberTypes | Format-Table -AutoSize
Write-Host 'NOTE: Per the SoW, Voice migration/testing may be excluded pending vendor selection; this inventory is for design scoping only.' -ForegroundColor DarkYellow

if (-not [string]::IsNullOrWhiteSpace($OutputFolder)) {
    if (-not (Test-Path -Path $OutputFolder)) { New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null }
    $inventory.Teams | Export-Csv -Path (Join-Path $OutputFolder 'Teams.csv') -NoTypeInformation -Encoding UTF8
    $inventory.PhoneNumberTypes | Export-Csv -Path (Join-Path $OutputFolder 'PhoneNumbers.csv') -NoTypeInformation -Encoding UTF8
    @($inventory.VoiceSummary) | Export-Csv -Path (Join-Path $OutputFolder 'VoiceApps.csv') -NoTypeInformation -Encoding UTF8
    Write-Host ("Reports exported to: {0}" -f $OutputFolder) -ForegroundColor Green
}
