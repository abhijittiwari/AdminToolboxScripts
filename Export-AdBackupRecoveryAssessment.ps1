<#
.SYNOPSIS
    Assesses Active Directory backup and forest-recovery controls BR-001 to BR-010 (Appendix B baseline).

.DESCRIPTION
    Assesses the AD backup and recovery posture against the security hardening baseline:
    system-state/AD backup coverage of all Domain Controllers, backup encryption and data
    residency, offline/immutable copies, the forest-recovery runbook, restore testing,
    retention vs RPO/RTO, isolated recovery environment, backup success/failure alerting,
    restricted backup access, and authoritative/non-authoritative restore validation.

    BR-001 and BR-008 are evaluated from AD replication metadata: each Domain Controller's
    dsaSignature attribute records when the directory database was last backed up, so the
    script can confirm coverage and recency across every DC. The remaining controls are
    procedural and are reported as Manual with the specific evidence to gather. Runbook and
    restore-test dates can be supplied so those controls are evaluated against your policy
    interval instead of left for manual review.

    Each control is emitted as one row with Status: Pass / Fail / Warning / Manual / Error.
    The script makes no changes to any system.

.PARAMETER Server
    Optional Domain Controller for AD queries.

.PARAMETER Credential
    Optional credential for AD queries.

.PARAMETER BackupMaxAgeDays
    Maximum age in days of a Domain Controller's last backup before BR-001 flags it. Default 7.

.PARAMETER RunbookLastUpdated
    Optional date the forest-recovery runbook was last reviewed/updated (BR-004).

.PARAMETER RunbookMaxAgeDays
    Maximum age in days of the runbook review before BR-004 warns. Default 365.

.PARAMETER LastRestoreTest
    Optional date of the most recent restore/DR test (BR-005).

.PARAMETER RestoreTestMaxAgeDays
    Maximum age in days of the last restore test before BR-005 warns. Default 365.

.PARAMETER OutputPath
    Optional CSV output path. If omitted, results are only written to the pipeline/table.

.EXAMPLE
    .\Export-AdBackupRecoveryAssessment.ps1 -OutputPath .\BackupRecovery.csv

.EXAMPLE
    .\Export-AdBackupRecoveryAssessment.ps1 -LastRestoreTest '2026-05-01' -RunbookLastUpdated '2026-06-15'
#>

[CmdletBinding()]
param(
    [string]$Server = '',
    [pscredential]$Credential,
    [int]$BackupMaxAgeDays = 7,
    [nullable[datetime]]$RunbookLastUpdated,
    [int]$RunbookMaxAgeDays = 365,
    [nullable[datetime]]$LastRestoreTest,
    [int]$RestoreTestMaxAgeDays = 365,
    [string]$OutputPath = ''
)

$ErrorActionPreference = 'Stop'

function New-ControlRow {
    param(
        [Parameter(Mandatory = $true)] [string]$ControlId,
        [Parameter(Mandatory = $true)] [string]$Control,
        [Parameter(Mandatory = $true)] [string]$Target,
        [Parameter(Mandatory = $true)] [ValidateSet('Pass','Fail','Warning','Manual','Error')] [string]$Status,
        [Parameter(Mandatory = $false)] [string]$Evidence = ''
    )

    return [pscustomobject]@{
        ControlId = $ControlId
        Control   = $Control
        Target    = $Target
        Status    = $Status
        Evidence  = $Evidence
    }
}

function Test-BackupRecoveryControls {
    param(
        [Parameter(Mandatory = $true)] $Facts,
        [int]$BackupMaxAgeDays = 7,
        [nullable[datetime]]$RunbookLastUpdated,
        [int]$RunbookMaxAgeDays = 365,
        [nullable[datetime]]$LastRestoreTest,
        [int]$RestoreTestMaxAgeDays = 365,
        [nullable[datetime]]$Now
    )

    if (-not $Now) { $Now = Get-Date }
    $rows = [System.Collections.Generic.List[object]]::new()
    $domain = [string]$Facts.DomainName

    $backups = @($Facts.DcBackups)
    $freshBackups = [System.Collections.Generic.List[object]]::new()
    $staleBackups = [System.Collections.Generic.List[object]]::new()
    $missingBackups = [System.Collections.Generic.List[object]]::new()
    foreach ($backup in $backups) {
        if (-not $backup.LastBackup) { $missingBackups.Add($backup); continue }
        $ageDays = [int]($Now - [datetime]$backup.LastBackup).TotalDays
        if ($ageDays -le $BackupMaxAgeDays) { $freshBackups.Add($backup) } else { $staleBackups.Add($backup) }
    }

    # BR-001 backups cover all DCs
    if ($backups.Count -eq 0) {
        $rows.Add((New-ControlRow -ControlId 'BR-001' -Control 'System-state / AD backups shall cover all Domain Controllers' -Target $domain -Status 'Error' -Evidence 'No Domain Controllers enumerated'))
    }
    elseif ($missingBackups.Count -eq 0 -and $staleBackups.Count -eq 0) {
        $rows.Add((New-ControlRow -ControlId 'BR-001' -Control 'System-state / AD backups shall cover all Domain Controllers' -Target $domain -Status 'Pass' -Evidence ("All {0} DC(s) have a system-state backup within {1} day(s)" -f $backups.Count, $BackupMaxAgeDays)))
    }
    else {
        $issues = [System.Collections.Generic.List[string]]::new()
        if ($missingBackups.Count -gt 0) { $issues.Add(("no backup marker: {0}" -f (($missingBackups | ForEach-Object Server) -join ', '))) }
        if ($staleBackups.Count -gt 0) { $issues.Add(("stale (> {0} days): {1}" -f $BackupMaxAgeDays, (($staleBackups | ForEach-Object Server) -join ', '))) }
        $rows.Add((New-ControlRow -ControlId 'BR-001' -Control 'System-state / AD backups shall cover all Domain Controllers' -Target $domain -Status 'Fail' -Evidence ($issues -join '; ')))
    }

    # BR-002 encryption + data residency
    $rows.Add((New-ControlRow -ControlId 'BR-002' -Control 'Backups shall be encrypted and stored within approved regions (data residency)' -Target $domain -Status 'Manual' -Evidence 'Confirm backup-at-rest encryption and that backup storage stays within approved regions in the backup solution configuration'))

    # BR-003 offline/immutable copy
    $rows.Add((New-ControlRow -ControlId 'BR-003' -Control 'At least one offline or immutable backup copy shall be maintained (ransomware resilience)' -Target $domain -Status 'Manual' -Evidence 'Verify an offline or immutable (WORM / air-gapped) backup copy exists for forest recovery'))

    # BR-004 forest-recovery runbook
    if ($RunbookLastUpdated) {
        $ageDays = [int]($Now - [datetime]$RunbookLastUpdated).TotalDays
        $status = if ($ageDays -le $RunbookMaxAgeDays) { 'Pass' } else { 'Warning' }
        $rows.Add((New-ControlRow -ControlId 'BR-004' -Control 'A documented AD forest-recovery runbook shall be maintained' -Target $domain -Status $status -Evidence ("Runbook last updated {0:yyyy-MM-dd} ({1} days ago; threshold {2})" -f [datetime]$RunbookLastUpdated, $ageDays, $RunbookMaxAgeDays)))
    }
    else {
        $rows.Add((New-ControlRow -ControlId 'BR-004' -Control 'A documented AD forest-recovery runbook shall be maintained' -Target $domain -Status 'Manual' -Evidence 'Confirm a current documented forest-recovery runbook exists; pass -RunbookLastUpdated to evaluate its currency'))
    }

    # BR-005 restore testing
    if ($LastRestoreTest) {
        $ageDays = [int]($Now - [datetime]$LastRestoreTest).TotalDays
        $status = if ($ageDays -le $RestoreTestMaxAgeDays) { 'Pass' } else { 'Warning' }
        $rows.Add((New-ControlRow -ControlId 'BR-005' -Control 'Restore procedures shall be tested at defined intervals' -Target $domain -Status $status -Evidence ("Last restore test {0:yyyy-MM-dd} ({1} days ago; threshold {2})" -f [datetime]$LastRestoreTest, $ageDays, $RestoreTestMaxAgeDays)))
    }
    else {
        $rows.Add((New-ControlRow -ControlId 'BR-005' -Control 'Restore procedures shall be tested at defined intervals' -Target $domain -Status 'Manual' -Evidence 'Confirm restore/DR testing occurs at defined intervals; pass -LastRestoreTest to evaluate its currency'))
    }

    # BR-006 retention vs RPO/RTO
    $rows.Add((New-ControlRow -ControlId 'BR-006' -Control 'Backup retention shall meet policy and recovery objectives (RPO/RTO)' -Target $domain -Status 'Manual' -Evidence 'Compare configured backup retention and schedule against documented RPO/RTO targets'))

    # BR-007 isolated recovery environment
    $rows.Add((New-ControlRow -ControlId 'BR-007' -Control 'An isolated recovery environment shall be considered for forest recovery' -Target $domain -Status 'Manual' -Evidence 'Confirm whether an isolated/clean-room recovery environment is defined for forest recovery'))

    # BR-008 backup success/failure alerting
    if ($backups.Count -eq 0) {
        $rows.Add((New-ControlRow -ControlId 'BR-008' -Control 'Backup success and failure shall be monitored and alerted' -Target $domain -Status 'Error' -Evidence 'No Domain Controllers enumerated'))
    }
    elseif ($staleBackups.Count -gt 0 -or $missingBackups.Count -gt 0) {
        $rows.Add((New-ControlRow -ControlId 'BR-008' -Control 'Backup success and failure shall be monitored and alerted' -Target $domain -Status 'Warning' -Evidence ("{0} of {1} DC(s) have a fresh backup; stale/missing backups indicate alerting may not be acted upon - confirm backup job alerting" -f $freshBackups.Count, $backups.Count)))
    }
    else {
        $rows.Add((New-ControlRow -ControlId 'BR-008' -Control 'Backup success and failure shall be monitored and alerted' -Target $domain -Status 'Manual' -Evidence ("All {0} DC(s) show recent backups; confirm backup success/failure alerts are configured and routed to operations" -f $backups.Count)))
    }

    # BR-009 restricted access to backup systems
    $rows.Add((New-ControlRow -ControlId 'BR-009' -Control 'Access to backup systems and media shall be restricted and audited' -Target $domain -Status 'Manual' -Evidence 'Review access controls and audit logging on the backup infrastructure and media'))

    # BR-010 restore capability validated
    $rows.Add((New-ControlRow -ControlId 'BR-010' -Control 'Authoritative and non-authoritative restore capability shall be validated' -Target $domain -Status 'Manual' -Evidence 'Confirm both authoritative and non-authoritative restore procedures have been validated (typically during BR-005 testing)'))

    return $rows.ToArray()
}

function Get-BackupRecoveryFacts {
    param(
        [string]$Server = '',
        [pscredential]$Credential
    )

    Import-Module ActiveDirectory -ErrorAction Stop
    $adParams = @{}
    if ($Credential) { $adParams.Credential = $Credential }
    if (-not [string]::IsNullOrWhiteSpace($Server)) { $adParams.Server = $Server }

    $domainObject = Get-ADDomain @adParams
    $facts = @{ DomainName = $domainObject.DNSRoot; DcBackups = @() }

    foreach ($dc in @(Get-ADDomainController -Filter * @adParams)) {
        $lastBackup = $null
        try {
            $metadata = Get-ADReplicationAttributeMetadata -Object $domainObject.DistinguishedName -Server $dc.HostName -ErrorAction Stop |
                Where-Object { $_.AttributeName -eq 'dSASignature' }
            if ($metadata) { $lastBackup = $metadata.LastOriginatingChangeTime }
        } catch { }
        $facts.DcBackups += @{ Server = $dc.HostName; LastBackup = $lastBackup }
    }

    return $facts
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
Write-Host 'Collecting AD backup markers (dsaSignature) from all Domain Controllers...' -ForegroundColor Cyan
$facts = Get-BackupRecoveryFacts -Server $Server -Credential $Credential

$rows = @(Test-BackupRecoveryControls -Facts $facts -BackupMaxAgeDays $BackupMaxAgeDays `
    -RunbookLastUpdated $RunbookLastUpdated -RunbookMaxAgeDays $RunbookMaxAgeDays `
    -LastRestoreTest $LastRestoreTest -RestoreTestMaxAgeDays $RestoreTestMaxAgeDays)

$sorted = $rows | Sort-Object ControlId, Target
$sorted | Format-Table ControlId, Target, Status, Evidence -AutoSize

$summary = $sorted | Group-Object Status | ForEach-Object { "{0}={1}" -f $_.Name, $_.Count }
Write-Host ("Summary: {0}" -f ($summary -join ', ')) -ForegroundColor Yellow

if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
    $sorted | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "Report exported: $OutputPath" -ForegroundColor Green
}
