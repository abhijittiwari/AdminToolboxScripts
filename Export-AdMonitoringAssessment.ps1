<#
.SYNOPSIS
    Assesses Active Directory monitoring and detection controls MON-001 to MON-010 (Appendix B baseline).

.DESCRIPTION
    Assesses the AD monitoring and detection posture against the security hardening baseline:
    security log forwarding to a SIEM, advanced audit policy coverage, alerting on privileged
    group changes, detection of suspicious authentication (Kerberoasting, AS-REP roasting,
    password spray), log retention, time synchronization for log correlation, honeytoken/canary
    accounts, directory-service health monitoring, legacy authentication monitoring, and periodic
    detection use-case review.

    Technically verifiable controls are collected per Domain Controller with Invoke-Command
    (WinRM): advanced audit subcategory settings, WEF/SIEM agent presence, NTLM auditing, and
    time-source configuration. Honeytoken accounts are matched from Active Directory by name
    pattern. SIEM detection-content controls are procedural and reported as Manual with the
    specific evidence to gather.

    Each control is emitted as one row (per DC where relevant) with Status:
      Pass / Fail / Warning / Manual / Error
    The script makes no changes to any system.

.PARAMETER Server
    One or more Domain Controllers to assess. Defaults to all DCs in the domain.

.PARAMETER Credential
    Optional credential for AD queries and remoting.

.PARAMETER HoneytokenNamePattern
    Wildcard pattern matched against sAMAccountName to locate deployed honeytoken/canary
    accounts (MON-007), e.g. 'svc-legacy*' or 'HNY-*'. When omitted, MON-007 is reported as Manual.

.PARAMETER OutputPath
    Optional CSV output path. If omitted, results are only written to the pipeline/table.

.EXAMPLE
    .\Export-AdMonitoringAssessment.ps1 -OutputPath .\Monitoring.csv

.EXAMPLE
    .\Export-AdMonitoringAssessment.ps1 -HoneytokenNamePattern 'HNY-*'
#>

[CmdletBinding()]
param(
    [string[]]$Server = @(),
    [pscredential]$Credential,
    [string]$HoneytokenNamePattern = '',
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

$script:MonFactScriptBlock = {
    $facts = @{}

    function Get-RegValue {
        param([string]$Path, [string]$Name)
        try { (Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop).$Name } catch { $null }
    }

    $facts.AuditPolicy = @{}
    $auditCsv = auditpol /get /category:* /r 2>$null
    if ($auditCsv) {
        try {
            foreach ($line in ($auditCsv | Where-Object { $_ } | ConvertFrom-Csv)) {
                $facts.AuditPolicy[$line.Subcategory] = $line.'Inclusion Setting'
            }
        } catch { }
    }

    $facts.WefSubscriptionManager = Get-RegValue 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\EventLog\EventForwarding\SubscriptionManager' '1'
    $siemServiceNames = @('HealthService','AzureMonitorAgent','MMAExtensionHeartbeatService','SplunkForwarder','nxlog','winlogbeat','ADAuditPlusPro')
    $facts.SiemAgentsRunning = @(Get-Service -ErrorAction SilentlyContinue |
        Where-Object { $siemServiceNames -contains $_.Name -and $_.Status -eq 'Running' } |
        ForEach-Object Name)

    $lsaMsv = 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa\MSV1_0'
    $facts.AuditReceivingNtlm = Get-RegValue $lsaMsv 'AuditReceivingNTLMTraffic'
    $facts.AuditNtlmInDomain = Get-RegValue 'HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon\Parameters' 'AuditNTLMInDomain'

    $facts.TimeSyncType = Get-RegValue 'HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Parameters' 'Type'
    $facts.TimeNtpServer = Get-RegValue 'HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Parameters' 'NtpServer'

    return $facts
}

function Get-MonDcFacts {
    param(
        [Parameter(Mandatory = $true)] [string]$ComputerName,
        [pscredential]$Credential
    )

    $invokeParams = @{
        ComputerName = $ComputerName
        ScriptBlock  = $script:MonFactScriptBlock
        ErrorAction  = 'Stop'
    }
    if ($Credential) { $invokeParams.Credential = $Credential }
    return Invoke-Command @invokeParams
}

function Test-MonDcControls {
    param(
        [Parameter(Mandatory = $true)] [string]$ServerName,
        [Parameter(Mandatory = $true)] $Facts,
        [bool]$IsPdcEmulator = $false
    )

    $rows = [System.Collections.Generic.List[object]]::new()
    $add = {
        param($Id, $Control, $Status, $Evidence)
        $rows.Add((New-ControlRow -ControlId $Id -Control $Control -Target $ServerName -Status $Status -Evidence $Evidence))
    }

    $auditPolicy = $Facts.AuditPolicy
    $isAudited = {
        param([string]$Subcategory, [string]$RequiredMode = 'any')
        $setting = [string]$auditPolicy[$Subcategory]
        if ([string]::IsNullOrWhiteSpace($setting) -or $setting -match 'No Auditing') { return $false }
        switch ($RequiredMode) {
            'Success' { return ($setting -match 'Success') }
            'Failure' { return ($setting -match 'Failure') }
            default   { return $true }
        }
    }

    $forwarding = @(@($Facts.SiemAgentsRunning) + @($(if ($Facts.WefSubscriptionManager) { 'WEF subscription' })) | Where-Object { $_ })

    # MON-001 log forwarding to SIEM
    if ($forwarding.Count -gt 0) {
        & $add 'MON-001' 'Security event logs from DCs and key servers shall be forwarded to a SIEM' 'Pass' ("Log collection detected: {0}; confirm coverage extends to key member servers" -f ($forwarding -join ', '))
    }
    else {
        & $add 'MON-001' 'Security event logs from DCs and key servers shall be forwarded to a SIEM' 'Warning' 'No WEF subscription or known SIEM/log agent detected; verify security log forwarding'
    }

    # MON-002 advanced audit policy coverage
    $requiredSubcategories = @(
        'Credential Validation','Kerberos Authentication Service','Kerberos Service Ticket Operations',
        'Computer Account Management','Security Group Management','User Account Management',
        'Directory Service Access','Directory Service Changes','Audit Policy Change',
        'Special Logon','Logon','Logoff','Account Lockout','Sensitive Privilege Use','Security System Extension'
    )
    if (-not $auditPolicy -or $auditPolicy.Count -eq 0) {
        & $add 'MON-002' 'Advanced audit policy shall be configured to capture security-relevant events' 'Error' 'auditpol output unavailable or not parseable'
    }
    else {
        $notAudited = @($requiredSubcategories | Where-Object { -not (& $isAudited $_) })
        if ($notAudited.Count -eq 0) {
            & $add 'MON-002' 'Advanced audit policy shall be configured to capture security-relevant events' 'Pass' 'All required audit subcategories are enabled'
        }
        else {
            & $add 'MON-002' 'Advanced audit policy shall be configured to capture security-relevant events' 'Fail' ("Subcategories not audited: {0}" -f ($notAudited -join ', '))
        }
    }

    # MON-003 privileged group change alerting (audit prerequisite)
    if (& $isAudited 'Security Group Management' 'Success') {
        & $add 'MON-003' 'Changes to privileged groups (Domain/Enterprise/Schema Admins) shall be alerted' 'Manual' 'Security Group Management auditing is enabled; confirm the SIEM alerts on changes to Domain/Enterprise/Schema Admins (events 4728/4732/4756)'
    }
    else {
        & $add 'MON-003' 'Changes to privileged groups (Domain/Enterprise/Schema Admins) shall be alerted' 'Fail' 'Security Group Management (Success) auditing is not enabled - privileged group change events will not be generated'
    }

    # MON-006 time synchronization
    $timeType = [string]$Facts.TimeSyncType
    if ($IsPdcEmulator) {
        if ($timeType -eq 'NTP' -and -not [string]::IsNullOrWhiteSpace([string]$Facts.TimeNtpServer)) {
            & $add 'MON-006' 'Time synchronization shall be enforced to ensure reliable log correlation' 'Pass' ("PDC emulator syncs from external NTP: {0}" -f $Facts.TimeNtpServer)
        }
        else {
            & $add 'MON-006' 'Time synchronization shall be enforced to ensure reliable log correlation' 'Warning' ("PDC emulator time type is '{0}' (expected NTP with an authoritative external source)" -f $timeType)
        }
    }
    else {
        if ($timeType -eq 'NT5DS') {
            & $add 'MON-006' 'Time synchronization shall be enforced to ensure reliable log correlation' 'Pass' 'Syncs from the AD time hierarchy (NT5DS)'
        }
        else {
            & $add 'MON-006' 'Time synchronization shall be enforced to ensure reliable log correlation' 'Warning' ("Time type is '{0}' (expected NT5DS for a non-PDC DC)" -f $timeType)
        }
    }

    # MON-009 legacy authentication monitoring
    $ntlmAuditOn = ($Facts.AuditReceivingNtlm -ge 1 -or $Facts.AuditNtlmInDomain -ge 1)
    if ($ntlmAuditOn) {
        & $add 'MON-009' 'Legacy authentication usage (LDAP/NTLM) shall be monitored to support phase-out' 'Pass' ("NTLM auditing enabled (AuditReceivingNTLMTraffic={0}, AuditNTLMInDomain={1}); also confirm LDAP signing/binding events (2886-2889) are monitored" -f $Facts.AuditReceivingNtlm, $Facts.AuditNtlmInDomain)
    }
    else {
        & $add 'MON-009' 'Legacy authentication usage (LDAP/NTLM) shall be monitored to support phase-out' 'Warning' 'NTLM auditing not enabled; enable NTLM audit policy and monitor LDAP signing/binding events (2886-2889) to support phase-out'
    }

    return $rows.ToArray()
}

function Test-MonDomainControls {
    param(
        [Parameter(Mandatory = $true)] $Facts,
        [string]$HoneytokenNamePattern = ''
    )

    $rows = [System.Collections.Generic.List[object]]::new()
    $domain = [string]$Facts.DomainName

    # MON-004 suspicious authentication detection
    $rows.Add((New-ControlRow -ControlId 'MON-004' -Control 'Suspicious authentication (Kerberoasting, AS-REP roasting, password spray) shall be detected' -Target $domain -Status 'Manual' -Evidence 'Confirm SIEM detection content exists for Kerberoasting (4769 RC4/encryption downgrade), AS-REP roasting (4768 preauth-not-required), and password spray (distributed 4771/4625)'))

    # MON-005 log retention
    $rows.Add((New-ControlRow -ControlId 'MON-005' -Control 'Log retention shall meet policy and investigative requirements' -Target $domain -Status 'Manual' -Evidence 'Confirm SIEM/log retention meets policy and investigative requirements (commonly 90+ days hot, 1 year cold)'))

    # MON-007 honeytoken accounts
    if ([string]::IsNullOrWhiteSpace($HoneytokenNamePattern)) {
        $rows.Add((New-ControlRow -ControlId 'MON-007' -Control 'Honeytoken / canary accounts shall be deployed and monitored' -Target $domain -Status 'Manual' -Evidence 'Confirm honeytoken/canary accounts are deployed and alert on any use; pass -HoneytokenNamePattern to verify their existence'))
    }
    elseif (@($Facts.HoneytokenAccounts).Count -gt 0) {
        $rows.Add((New-ControlRow -ControlId 'MON-007' -Control 'Honeytoken / canary accounts shall be deployed and monitored' -Target $domain -Status 'Pass' -Evidence ("Honeytoken account(s) matching '{0}': {1}; confirm any authentication attempt triggers an alert" -f $HoneytokenNamePattern, (@($Facts.HoneytokenAccounts) -join ', '))))
    }
    else {
        $rows.Add((New-ControlRow -ControlId 'MON-007' -Control 'Honeytoken / canary accounts shall be deployed and monitored' -Target $domain -Status 'Fail' -Evidence ("No accounts match the honeytoken pattern '{0}'" -f $HoneytokenNamePattern)))
    }

    # MON-008 directory-service health monitoring
    $rows.Add((New-ControlRow -ControlId 'MON-008' -Control 'Domain Controller and directory-service health shall be continuously monitored' -Target $domain -Status 'Manual' -Evidence 'Confirm continuous monitoring of DC availability, replication, and directory-service health (e.g. via the monitoring platform or Microsoft Defender for Identity)'))

    # MON-010 detection use-case review
    $rows.Add((New-ControlRow -ControlId 'MON-010' -Control 'Detection use-cases shall be reviewed and updated periodically' -Target $domain -Status 'Manual' -Evidence 'Confirm a periodic review process keeps detection use-cases current with the threat landscape'))

    return $rows.ToArray()
}

function Get-MonDomainFacts {
    param(
        [string[]]$Server,
        [pscredential]$Credential,
        [string]$HoneytokenNamePattern = ''
    )

    Import-Module ActiveDirectory -ErrorAction Stop
    $adParams = @{}
    if ($Credential) { $adParams.Credential = $Credential }

    $domainObject = Get-ADDomain @adParams
    $facts = @{
        DomainName      = $domainObject.DNSRoot
        PdcEmulator     = $domainObject.PDCEmulator
        HoneytokenAccounts = @()
    }

    if ($Server.Count -gt 0) {
        $facts.DomainControllers = @($Server | ForEach-Object { @{ HostName = $_; IsPdcEmulator = ($_ -eq $domainObject.PDCEmulator) } })
    }
    else {
        $facts.DomainControllers = @(Get-ADDomainController -Filter * @adParams | ForEach-Object { @{ HostName = $_.HostName; IsPdcEmulator = [bool]$_.OperationMasterRoles -and ($_.OperationMasterRoles -contains 'PDCEmulator') } })
    }

    if (-not [string]::IsNullOrWhiteSpace($HoneytokenNamePattern)) {
        try {
            $facts.HoneytokenAccounts = @(Get-ADUser -Filter { SamAccountName -like $HoneytokenNamePattern } @adParams | ForEach-Object SamAccountName)
        } catch { }
    }

    return $facts
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
$allRows = [System.Collections.Generic.List[object]]::new()

Write-Host 'Collecting domain-level monitoring facts...' -ForegroundColor Cyan
$domainFacts = Get-MonDomainFacts -Server $Server -Credential $Credential -HoneytokenNamePattern $HoneytokenNamePattern
foreach ($row in (Test-MonDomainControls -Facts $domainFacts -HoneytokenNamePattern $HoneytokenNamePattern)) {
    $allRows.Add($row)
}

$targets = @($domainFacts.DomainControllers)
Write-Host ("Assessing {0} Domain Controller(s)" -f $targets.Count) -ForegroundColor Cyan

foreach ($target in $targets) {
    Write-Host ("  Collecting monitoring facts from {0}..." -f $target.HostName) -ForegroundColor Cyan
    try {
        $facts = Get-MonDcFacts -ComputerName $target.HostName -Credential $Credential
        foreach ($row in (Test-MonDcControls -ServerName $target.HostName -Facts $facts -IsPdcEmulator ([bool]$target.IsPdcEmulator))) {
            $allRows.Add($row)
        }
    }
    catch {
        $allRows.Add((New-ControlRow -ControlId 'MON-000' -Control 'Fact collection' -Target $target.HostName -Status 'Error' -Evidence $_.Exception.Message))
    }
}

$sorted = $allRows | Sort-Object ControlId, Target
$sorted | Format-Table ControlId, Target, Status, Evidence -AutoSize

$summary = $sorted | Group-Object Status | ForEach-Object { "{0}={1}" -f $_.Name, $_.Count }
Write-Host ("Summary: {0}" -f ($summary -join ', ')) -ForegroundColor Yellow

if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
    $sorted | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "Report exported: $OutputPath" -ForegroundColor Green
}
