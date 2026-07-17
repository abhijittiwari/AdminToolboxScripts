<#
.SYNOPSIS
    Exports Active Directory user, group, and computer population volumetrics for divestiture discovery.

.DESCRIPTION
    Read-only discovery of the object populations a divestiture must scope and license, producing the
    volumetrics the SoW cites (and the disabled-account scope question it flags):
      - Users: total, enabled vs disabled, stale (no recent logon), never-logged-on, password-never-
        expires, and smart-card-required counts.
      - Groups: total, by scope (Global/DomainLocal/Universal) and category (Security/Distribution),
        mail-enabled, and empty groups.
      - Computers: total, enabled vs disabled, stale, and a server-vs-workstation split derived from
        the operatingSystem attribute (the ~470 servers to re-domain vs. ~1,860 devices to domain-join).

    Objects are read once with the ActiveDirectory module (lastLogonTimestamp for staleness); the
    counting logic is a separate, unit-testable function fed from the raw object lists. The script
    makes no changes.

.PARAMETER SearchBase
    Optional OU distinguished name to scope the counts.

.PARAMETER StaleDays
    lastLogonTimestamp age in days after which an enabled object counts as stale. Default 90.

.PARAMETER Server
    Optional Domain Controller / AD Web Services server to target for the queries.

.PARAMETER Credential
    Optional credential for the AD queries.

.PARAMETER OutputFolder
    Optional folder for the CSV outputs (UserPopulation.csv, GroupPopulation.csv, ComputerPopulation.csv).
    If omitted, results are only written to the pipeline/table.

.EXAMPLE
    .\Export-AdObjectPopulationInventory.ps1 -OutputFolder .\Population -StaleDays 90
#>

[CmdletBinding()]
param(
    [string]$SearchBase = '',
    [int]$StaleDays = 90,
    [string]$Server = '',
    [pscredential]$Credential,
    [string]$OutputFolder = ''
)

$ErrorActionPreference = 'Stop'

function Test-IsServerOs {
    param([AllowNull()] [string]$OperatingSystem)
    if ([string]::IsNullOrWhiteSpace($OperatingSystem)) { return $false }
    return ($OperatingSystem -match 'Server')
}

function Test-IsStale {
    param(
        [AllowNull()] $LastLogonTimestamp,
        [Parameter(Mandatory = $true)] [datetime]$Threshold
    )
    if ($null -eq $LastLogonTimestamp) { return $true }  # never logged on counts as stale
    try {
        $lastLogon = [datetime]::FromFileTimeUtc([int64]$LastLogonTimestamp)
    } catch {
        return $true
    }
    return ($lastLogon -lt $Threshold)
}

function Get-UserPopulationRow {
    param(
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] $Users,
        [Parameter(Mandatory = $true)] [datetime]$StaleThreshold
    )

    $users = @($Users)
    $enabled = @($users | Where-Object { $_.Enabled })
    $staleEnabled = @($enabled | Where-Object { Test-IsStale -LastLogonTimestamp $_.lastLogonTimestamp -Threshold $StaleThreshold })
    return [pscustomobject]@{
        TotalUsers            = $users.Count
        EnabledUsers          = $enabled.Count
        DisabledUsers         = @($users | Where-Object { -not $_.Enabled }).Count
        StaleEnabledUsers     = $staleEnabled.Count
        NeverLoggedOnUsers    = @($users | Where-Object { $null -eq $_.lastLogonTimestamp }).Count
        PasswordNeverExpires  = @($users | Where-Object { $_.PasswordNeverExpires }).Count
        SmartCardRequired     = @($users | Where-Object { $_.SmartcardLogonRequired }).Count
    }
}

function Get-GroupPopulationRow {
    param([Parameter(Mandatory = $true)] [AllowEmptyCollection()] $Groups)

    $groups = @($Groups)
    return [pscustomobject]@{
        TotalGroups        = $groups.Count
        GlobalGroups       = @($groups | Where-Object { $_.GroupScope -eq 'Global' }).Count
        DomainLocalGroups  = @($groups | Where-Object { $_.GroupScope -eq 'DomainLocal' }).Count
        UniversalGroups    = @($groups | Where-Object { $_.GroupScope -eq 'Universal' }).Count
        SecurityGroups     = @($groups | Where-Object { $_.GroupCategory -eq 'Security' }).Count
        DistributionGroups = @($groups | Where-Object { $_.GroupCategory -eq 'Distribution' }).Count
        MailEnabledGroups  = @($groups | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.mail) }).Count
        EmptyGroups        = @($groups | Where-Object { @($_.member).Count -eq 0 }).Count
    }
}

function Get-ComputerPopulationRow {
    param(
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] $Computers,
        [Parameter(Mandatory = $true)] [datetime]$StaleThreshold
    )

    $computers = @($Computers)
    $enabled = @($computers | Where-Object { $_.Enabled })
    $servers = @($computers | Where-Object { Test-IsServerOs -OperatingSystem ([string]$_.OperatingSystem) })
    return [pscustomobject]@{
        TotalComputers        = $computers.Count
        EnabledComputers      = $enabled.Count
        DisabledComputers     = @($computers | Where-Object { -not $_.Enabled }).Count
        Servers               = $servers.Count
        Workstations          = ($computers.Count - $servers.Count)
        StaleEnabledComputers = @($enabled | Where-Object { Test-IsStale -LastLogonTimestamp $_.lastLogonTimestamp -Threshold $StaleThreshold }).Count
    }
}

function Get-AdPopulationFacts {
    param(
        [string]$SearchBase = '',
        [string]$Server = '',
        [pscredential]$Credential
    )

    Import-Module ActiveDirectory -ErrorAction Stop
    $adParams = @{}
    if (-not [string]::IsNullOrWhiteSpace($Server)) { $adParams.Server = $Server }
    if ($Credential) { $adParams.Credential = $Credential }
    if (-not [string]::IsNullOrWhiteSpace($SearchBase)) { $adParams.SearchBase = $SearchBase }

    $users = @(Get-ADUser -Filter * -Properties Enabled, lastLogonTimestamp, PasswordNeverExpires, SmartcardLogonRequired @adParams)
    $groups = @(Get-ADGroup -Filter * -Properties GroupScope, GroupCategory, mail, member @adParams)
    $computers = @(Get-ADComputer -Filter * -Properties Enabled, lastLogonTimestamp, OperatingSystem @adParams)

    return @{ Users = $users; Groups = $groups; Computers = $computers }
}

function Get-AdObjectPopulationInventory {
    param(
        [Parameter(Mandatory = $true)] $Facts,
        [int]$StaleDays = 90,
        [nullable[datetime]]$Now
    )

    if (-not $Now) { $Now = Get-Date }
    $staleThreshold = $Now.AddDays(-$StaleDays)

    return @{
        Users     = Get-UserPopulationRow -Users $Facts.Users -StaleThreshold $staleThreshold
        Groups    = Get-GroupPopulationRow -Groups $Facts.Groups
        Computers = Get-ComputerPopulationRow -Computers $Facts.Computers -StaleThreshold $staleThreshold
    }
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
Write-Host 'Collecting AD user/group/computer population volumetrics...' -ForegroundColor Cyan
$facts = Get-AdPopulationFacts -SearchBase $SearchBase -Server $Server -Credential $Credential
$inventory = Get-AdObjectPopulationInventory -Facts $facts -StaleDays $StaleDays

Write-Host "`n=== Users ===" -ForegroundColor Yellow
$inventory.Users | Format-List
Write-Host "=== Groups ===" -ForegroundColor Yellow
$inventory.Groups | Format-List
Write-Host "=== Computers ===" -ForegroundColor Yellow
$inventory.Computers | Format-List

if (-not [string]::IsNullOrWhiteSpace($OutputFolder)) {
    if (-not (Test-Path -Path $OutputFolder)) { New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null }
    @($inventory.Users) | Export-Csv -Path (Join-Path $OutputFolder 'UserPopulation.csv') -NoTypeInformation -Encoding UTF8
    @($inventory.Groups) | Export-Csv -Path (Join-Path $OutputFolder 'GroupPopulation.csv') -NoTypeInformation -Encoding UTF8
    @($inventory.Computers) | Export-Csv -Path (Join-Path $OutputFolder 'ComputerPopulation.csv') -NoTypeInformation -Encoding UTF8
    Write-Host ("Reports exported to: {0}" -f $OutputFolder) -ForegroundColor Green
}
