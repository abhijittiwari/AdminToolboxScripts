<#
.SYNOPSIS
    Exports Active Directory objects that carry SIDHistory for divestiture discovery.

.DESCRIPTION
    Read-only discovery of every user, group, and computer whose sIDHistory attribute is populated.
    A new-forest divestiture that preserves resource access through SIDHistory (with the trust's SID
    filtering relaxed during coexistence) needs a precise inventory of what SIDHistory already exists
    and which foreign domains those historical SIDs come from, so it can be reproduced or cleaned up.

    For each object the script reports the count of historical SIDs and the distinct source domain SIDs
    they belong to (the domain SID is the historical SID minus its trailing RID), which shows how many
    prior domains contributed SIDHistory. Objects are read with the ActiveDirectory module; parsing and
    shaping are separated so they are unit-testable with mock data. The script makes no changes.

.PARAMETER Server
    Optional Domain Controller / AD Web Services server to target for the queries.

.PARAMETER Credential
    Optional credential for the AD queries.

.PARAMETER OutputPath
    Optional CSV output path. If omitted, results are only written to the pipeline/table.

.EXAMPLE
    .\Export-AdSidHistoryInventory.ps1 -OutputPath .\SidHistory.csv
#>

[CmdletBinding()]
param(
    [string]$Server = '',
    [pscredential]$Credential,
    [string]$OutputPath = ''
)

$ErrorActionPreference = 'Stop'

function Get-DomainSidFromSid {
    param([Parameter(Mandatory = $true)] [string]$Sid)

    # A domain SID is the account SID with the final RID sub-authority removed.
    if ($Sid -notmatch '^S-1-5-21-') { return $Sid }
    $parts = $Sid -split '-'
    if ($parts.Count -le 1) { return $Sid }
    return ($parts[0..($parts.Count - 2)] -join '-')
}

function New-SidHistoryRow {
    param([Parameter(Mandatory = $true)] $Object)

    $sids = @($Object.SIDHistory | ForEach-Object { [string]$_ } | Where-Object { $_ })
    $sourceDomainSids = @($sids | ForEach-Object { Get-DomainSidFromSid -Sid $_ } | Select-Object -Unique)

    return [pscustomobject]@{
        SamAccountName    = [string]$Object.SamAccountName
        ObjectClass       = [string]$Object.objectClass
        Enabled           = [bool]$Object.Enabled
        DistinguishedName = [string]$Object.DistinguishedName
        SidHistoryCount   = $sids.Count
        SourceDomainCount = $sourceDomainSids.Count
        SourceDomainSids  = ($sourceDomainSids -join ';')
        SidHistory        = ($sids -join ';')
    }
}

function Get-AdSidHistoryFacts {
    param(
        [string]$Server = '',
        [pscredential]$Credential
    )

    Import-Module ActiveDirectory -ErrorAction Stop
    $adParams = @{}
    if (-not [string]::IsNullOrWhiteSpace($Server)) { $adParams.Server = $Server }
    if ($Credential) { $adParams.Credential = $Credential }

    # Any object class with sIDHistory populated (users, groups, computers).
    $objects = @(Get-ADObject -LDAPFilter '(sIDHistory=*)' -Properties SamAccountName, Enabled, objectClass, SIDHistory @adParams)
    return @{ Objects = $objects }
}

function Get-AdSidHistoryInventory {
    param([Parameter(Mandatory = $true)] $Facts)
    return @(@($Facts.Objects) | ForEach-Object { New-SidHistoryRow -Object $_ })
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
Write-Host 'Collecting objects with SIDHistory...' -ForegroundColor Cyan
$facts = Get-AdSidHistoryFacts -Server $Server -Credential $Credential
$rows = @(Get-AdSidHistoryInventory -Facts $facts)

if ($rows.Count -eq 0) {
    Write-Host 'No objects with SIDHistory found.' -ForegroundColor Yellow
}
else {
    $rows | Format-Table SamAccountName, ObjectClass, Enabled, SidHistoryCount, SourceDomainCount -AutoSize
    $distinctDomains = @($rows | ForEach-Object { $_.SourceDomainSids -split ';' } | Where-Object { $_ } | Select-Object -Unique)
    Write-Host ("Summary: {0} object(s) with SIDHistory across {1} distinct source domain(s)" -f $rows.Count, $distinctDomains.Count) -ForegroundColor Yellow
}

if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
    $rows | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "Report exported: $OutputPath" -ForegroundColor Green
}
