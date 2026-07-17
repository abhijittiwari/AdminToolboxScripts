<#
.SYNOPSIS
    Assesses domain-joined computer security hardening controls DJ-001 to DJ-020 (Appendix B baseline).

.DESCRIPTION
    Read-only assessment of domain-joined member servers and workstations against the
    security hardening baseline: supported/patched OS, Windows LAPS coverage, BitLocker,
    Secure Boot and TPM 2.0, firewall, Defender AV/EDR, ASR rules, legacy protocols
    (SMBv1/LLMNR/NBT-NS), NTLM and SMB signing, local administrator membership, UAC,
    Credential Guard and LSA protection, PowerShell logging, security baseline GPOs,
    removable media/AutoRun, stale computer accounts, screen lock and logon banner,
    time synchronization, and domain-join restrictions.

    Domain-wide controls (DJ-003 LAPS coverage, DJ-017 stale accounts, DJ-020 machine
    account quota) are evaluated once from Active Directory. Per-computer controls are
    collected remotely with Invoke-Command (WinRM) from the targeted sample of computers.
    Each control is emitted as one row per target with Status:
      Pass / Fail / Warning / Manual / Error

    The script makes no changes to any system.

.PARAMETER ComputerName
    Explicit list of computers to assess remotely. When omitted, enabled computer
    accounts are discovered from Active Directory (optionally scoped by -SearchBase)
    and sampled up to -SampleSize.

.PARAMETER ComputersCsv
    Optional CSV with a ComputerName column (fallback: Name) listing computers to assess.

.PARAMETER SearchBase
    Optional OU distinguished name to scope AD computer discovery.

.PARAMETER SampleSize
    Maximum number of discovered computers to assess remotely. Default 25.

.PARAMETER StaleDays
    lastLogonTimestamp age in days after which an enabled computer account counts as stale. Default 90.

.PARAMETER PatchAgeDays
    Maximum age in days of the most recent installed hotfix before DJ-002 is flagged. Default 30.

.PARAMETER LapsMinimumCoveragePercent
    Minimum percentage of enabled computers with a LAPS-managed password before DJ-003 passes. Default 95.

.PARAMETER BaselineGpoPattern
    Regex matched against applied GPO names to verify the security baseline GPO (DJ-015).
    When omitted, DJ-015 reports the applied GPO list for manual review.

.PARAMETER Credential
    Optional credential for AD queries and remoting.

.PARAMETER OutputPath
    Optional CSV output path. If omitted, results are only written to the pipeline/table.

.EXAMPLE
    .\Export-DomainComputerHardeningAssessment.ps1 -OutputPath .\DjHardening.csv

.EXAMPLE
    .\Export-DomainComputerHardeningAssessment.ps1 -ComputersCsv .\servers.csv -BaselineGpoPattern 'CIS|Security Baseline'
#>

[CmdletBinding()]
param(
    [string[]]$ComputerName = @(),
    [string]$ComputersCsv = '',
    [string]$SearchBase = '',
    [int]$SampleSize = 25,
    [int]$StaleDays = 90,
    [int]$PatchAgeDays = 30,
    [int]$LapsMinimumCoveragePercent = 95,
    [string]$BaselineGpoPattern = '',
    [pscredential]$Credential,
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

# Remote fact collection. Runs on each computer under Windows PowerShell 5.1.
$script:DjFactScriptBlock = {
    $facts = @{}

    function Get-RegValue {
        param([string]$Path, [string]$Name)
        try { (Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop).$Name } catch { $null }
    }

    $os = Get-CimInstance -ClassName Win32_OperatingSystem
    $facts.OsCaption = $os.Caption
    $facts.OsBuild   = $os.BuildNumber

    $hotfixes = @(Get-HotFix -ErrorAction SilentlyContinue | Where-Object InstalledOn)
    $facts.LastHotfixDate = if ($hotfixes.Count -gt 0) { ($hotfixes | Sort-Object InstalledOn -Descending)[0].InstalledOn } else { $null }

    try {
        $blv = Get-BitLockerVolume -MountPoint $env:SystemDrive -ErrorAction Stop
        $facts.BitLockerProtection = [string]$blv.ProtectionStatus
    } catch { $facts.BitLockerProtection = $null }

    try { $facts.SecureBoot = Confirm-SecureBootUEFI } catch { $facts.SecureBoot = $null }
    try {
        $tpm = Get-Tpm -ErrorAction Stop
        $facts.TpmPresent = [bool]$tpm.TpmPresent
        $facts.TpmReady   = [bool]$tpm.TpmReady
        $tpmInfo = Get-CimInstance -Namespace 'root/cimv2/Security/MicrosoftTpm' -ClassName Win32_Tpm -ErrorAction SilentlyContinue
        $facts.TpmSpecVersion = if ($tpmInfo) { [string]$tpmInfo.SpecVersion } else { $null }
    } catch {
        $facts.TpmPresent = $null
        $facts.TpmReady = $null
        $facts.TpmSpecVersion = $null
    }

    $facts.FirewallProfiles = @(Get-NetFirewallProfile -ErrorAction SilentlyContinue |
        ForEach-Object { @{ Name = $_.Name; Enabled = [bool]$_.Enabled } })

    try {
        $mp = Get-MpComputerStatus -ErrorAction Stop
        $facts.AntivirusEnabled = [bool]$mp.AntivirusEnabled
        $facts.RealTimeProtection = [bool]$mp.RealTimeProtectionEnabled
        $facts.AvSignatureAgeDays = [int]$mp.AntivirusSignatureAge
    } catch {
        $facts.AntivirusEnabled = $null
        $facts.RealTimeProtection = $null
        $facts.AvSignatureAgeDays = $null
    }

    $edrServiceNames = @('Sense','CSFalconService','CbDefense','SentinelAgent','xagt','CylanceSvc','mfeens','TaniumClient')
    $facts.EdrServicesRunning = @(Get-Service -ErrorAction SilentlyContinue |
        Where-Object { $edrServiceNames -contains $_.Name -and $_.Status -eq 'Running' } |
        ForEach-Object Name)

    try {
        $mpPreference = Get-MpPreference -ErrorAction Stop
        $asrIds = @($mpPreference.AttackSurfaceReductionRules_Ids)
        $asrActions = @($mpPreference.AttackSurfaceReductionRules_Actions)
        $enabledCount = 0
        for ($i = 0; $i -lt $asrIds.Count; $i++) {
            if ($i -lt $asrActions.Count -and $asrActions[$i] -in @(1, 6)) { $enabledCount++ }  # 1=Block, 6=Warn
        }
        $facts.AsrRulesConfigured = $asrIds.Count
        $facts.AsrRulesEnabled = $enabledCount
    } catch {
        $facts.AsrRulesConfigured = $null
        $facts.AsrRulesEnabled = $null
    }

    try {
        $smbServer = Get-SmbServerConfiguration -ErrorAction Stop
        $facts.Smb1Enabled = [bool]$smbServer.EnableSMB1Protocol
        $facts.SmbServerSigningRequired = [bool]$smbServer.RequireSecuritySignature
    } catch {
        $facts.Smb1Enabled = $null
        $facts.SmbServerSigningRequired = $null
    }
    $facts.LlmnrPolicy = Get-RegValue 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\DNSClient' 'EnableMulticast'
    $netbiosSettings = @()
    foreach ($interfaceKey in (Get-ChildItem 'HKLM:\SYSTEM\CurrentControlSet\Services\NetBT\Parameters\Interfaces' -ErrorAction SilentlyContinue)) {
        $value = Get-RegValue $interfaceKey.PSPath 'NetbiosOptions'
        if ($null -ne $value) { $netbiosSettings += [int]$value }
    }
    $facts.NetbiosOptions = $netbiosSettings

    $facts.LmCompatibilityLevel = Get-RegValue 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa' 'LmCompatibilityLevel'
    $facts.SmbClientSigningEnabled = Get-RegValue 'HKLM:\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Parameters' 'EnableSecuritySignature'

    $facts.LocalAdministrators = @()
    try {
        $facts.LocalAdministrators = @(Get-LocalGroupMember -Group 'Administrators' -ErrorAction Stop | ForEach-Object Name)
    } catch {
        try {
            $group = [ADSI]'WinNT://./Administrators,group'
            $facts.LocalAdministrators = @($group.Invoke('Members') | ForEach-Object {
                $_.GetType().InvokeMember('AdsPath', 'GetProperty', $null, $_, $null) -replace '^WinNT://', '' -replace '/', '\'
            })
        } catch { }
    }

    $policySystem = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System'
    $facts.EnableLua = Get-RegValue $policySystem 'EnableLUA'
    $facts.ConsentPromptBehaviorAdmin = Get-RegValue $policySystem 'ConsentPromptBehaviorAdmin'
    $facts.PromptOnSecureDesktop = Get-RegValue $policySystem 'PromptOnSecureDesktop'
    $facts.InactivityTimeoutSecs = Get-RegValue $policySystem 'InactivityTimeoutSecs'
    $facts.LegalNoticeCaption = Get-RegValue $policySystem 'legalnoticecaption'
    $facts.LegalNoticeText = Get-RegValue $policySystem 'legalnoticetext'

    try {
        $deviceGuard = Get-CimInstance -Namespace 'root\Microsoft\Windows\DeviceGuard' -ClassName Win32_DeviceGuard -ErrorAction Stop
        $facts.CredentialGuardRunning = (@($deviceGuard.SecurityServicesRunning) -contains 1)
    } catch { $facts.CredentialGuardRunning = $null }
    $facts.RunAsPpl = Get-RegValue 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa' 'RunAsPPL'

    $facts.ScriptBlockLogging = Get-RegValue 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging' 'EnableScriptBlockLogging'
    $facts.ModuleLogging      = Get-RegValue 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\PowerShell\ModuleLogging' 'EnableModuleLogging'

    $facts.AppliedGpoNames = @()
    foreach ($historyKey in (Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\History' -ErrorAction SilentlyContinue)) {
        foreach ($gpoKey in (Get-ChildItem $historyKey.PSPath -ErrorAction SilentlyContinue)) {
            $displayName = Get-RegValue $gpoKey.PSPath 'DisplayName'
            if ($displayName -and $facts.AppliedGpoNames -notcontains $displayName) { $facts.AppliedGpoNames += $displayName }
        }
    }

    $facts.NoDriveTypeAutoRun = Get-RegValue 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer' 'NoDriveTypeAutoRun'
    $facts.RemovableStorageDenyAll = Get-RegValue 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\RemovableStorageDevices' 'Deny_All'

    $facts.TimeSyncType = Get-RegValue 'HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Parameters' 'Type'

    return $facts
}

function Get-DjFacts {
    param(
        [Parameter(Mandatory = $true)] [string]$ComputerName,
        [pscredential]$Credential
    )

    $invokeParams = @{
        ComputerName = $ComputerName
        ScriptBlock  = $script:DjFactScriptBlock
        ErrorAction  = 'Stop'
    }
    if ($Credential) { $invokeParams.Credential = $Credential }
    return Invoke-Command @invokeParams
}

function Test-DjComputerControls {
    param(
        [Parameter(Mandatory = $true)] [string]$ComputerName,
        [Parameter(Mandatory = $true)] $Facts,
        [int]$PatchAgeDays = 30,
        [string]$BaselineGpoPattern = ''
    )

    $rows = [System.Collections.Generic.List[object]]::new()
    $add = {
        param($Id, $Control, $Status, $Evidence)
        $rows.Add((New-ControlRow -ControlId $Id -Control $Control -Target $ComputerName -Status $Status -Evidence $Evidence))
    }

    # DJ-001 supported OS
    $osCaption = [string]$Facts.OsCaption
    if ($osCaption -match 'Windows 11|Server (2016|2019|2022|2025)') {
        $status = if ($osCaption -match 'Server 2016') { 'Warning' } else { 'Pass' }
        & $add 'DJ-001' 'Domain-joined computers shall run supported, currently patched operating system versions' $status ("{0} build {1}" -f $osCaption, $Facts.OsBuild)
    }
    elseif ($osCaption -match 'Windows 10') {
        & $add 'DJ-001' 'Domain-joined computers shall run supported, currently patched operating system versions' 'Warning' ("{0} build {1} - Windows 10 end of support Oct 2025; verify ESU coverage" -f $osCaption, $Facts.OsBuild)
    }
    else {
        & $add 'DJ-001' 'Domain-joined computers shall run supported, currently patched operating system versions' 'Fail' ("Unsupported or unknown OS: {0}" -f $osCaption)
    }

    # DJ-002 patch SLA
    if ($Facts.LastHotfixDate) {
        $ageDays = [int]((Get-Date) - [datetime]$Facts.LastHotfixDate).TotalDays
        $status = if ($ageDays -le $PatchAgeDays) { 'Pass' } else { 'Warning' }
        & $add 'DJ-002' 'OS and third-party security patches shall be applied within defined SLAs' $status ("Most recent hotfix installed {0:yyyy-MM-dd} ({1} days ago); third-party patch state requires patch tooling verification" -f [datetime]$Facts.LastHotfixDate, $ageDays)
    }
    else {
        & $add 'DJ-002' 'OS and third-party security patches shall be applied within defined SLAs' 'Manual' 'No hotfix install date available; verify via patch management tooling'
    }

    # DJ-004 BitLocker
    switch ([string]$Facts.BitLockerProtection) {
        'On'    { & $add 'DJ-004' 'BitLocker full-disk encryption shall be enabled on all endpoints' 'Pass' 'System drive protection is On' }
        'Off'   { & $add 'DJ-004' 'BitLocker full-disk encryption shall be enabled on all endpoints' 'Fail' 'System drive protection is Off' }
        default { & $add 'DJ-004' 'BitLocker full-disk encryption shall be enabled on all endpoints' 'Fail' 'BitLocker status unavailable' }
    }

    # DJ-005 Secure Boot + TPM 2.0
    $tpm2 = ([string]$Facts.TpmSpecVersion -match '2\.0')
    if ($Facts.SecureBoot -eq $true -and $Facts.TpmPresent -eq $true -and $tpm2) {
        & $add 'DJ-005' 'Secure Boot and TPM 2.0 shall be enabled' 'Pass' ("Secure Boot on; TPM present and ready={0}; spec {1}" -f $Facts.TpmReady, $Facts.TpmSpecVersion)
    }
    else {
        & $add 'DJ-005' 'Secure Boot and TPM 2.0 shall be enabled' 'Fail' ("SecureBoot={0}, TpmPresent={1}, TpmSpecVersion={2}" -f $Facts.SecureBoot, $Facts.TpmPresent, $Facts.TpmSpecVersion)
    }

    # DJ-006 firewall on all profiles
    $profiles = @($Facts.FirewallProfiles)
    $disabledProfiles = @($profiles | Where-Object { -not $_.Enabled } | ForEach-Object Name)
    if ($profiles.Count -eq 0) { & $add 'DJ-006' 'Windows Firewall shall be enabled on all network profiles' 'Error' 'Firewall profile state could not be read' }
    elseif ($disabledProfiles.Count -eq 0) { & $add 'DJ-006' 'Windows Firewall shall be enabled on all network profiles' 'Pass' 'All firewall profiles enabled' }
    else { & $add 'DJ-006' 'Windows Firewall shall be enabled on all network profiles' 'Fail' ("Disabled profiles: {0}" -f ($disabledProfiles -join ', ')) }

    # DJ-007 Defender AV + EDR
    $edr = @($Facts.EdrServicesRunning)
    if ($Facts.AntivirusEnabled -eq $true -and $edr.Count -gt 0) {
        & $add 'DJ-007' 'Microsoft Defender Antivirus and EDR shall be deployed, active, and reporting' 'Pass' ("AV enabled (real-time {0}); EDR: {1}; verify reporting in the security console" -f $Facts.RealTimeProtection, ($edr -join ', '))
    }
    elseif ($Facts.AntivirusEnabled -eq $true) {
        & $add 'DJ-007' 'Microsoft Defender Antivirus and EDR shall be deployed, active, and reporting' 'Warning' 'AV enabled but no known EDR service detected'
    }
    else {
        & $add 'DJ-007' 'Microsoft Defender Antivirus and EDR shall be deployed, active, and reporting' 'Fail' ("AntivirusEnabled={0}, EDR services: {1}" -f $Facts.AntivirusEnabled, ($edr -join ', '))
    }

    # DJ-008 ASR rules
    if ($null -eq $Facts.AsrRulesConfigured) {
        & $add 'DJ-008' 'Attack Surface Reduction (ASR) rules shall be enabled' 'Error' 'Defender preferences could not be read'
    }
    elseif ([int]$Facts.AsrRulesEnabled -gt 0) {
        & $add 'DJ-008' 'Attack Surface Reduction (ASR) rules shall be enabled' 'Pass' ("{0} of {1} configured ASR rules in block/warn mode" -f $Facts.AsrRulesEnabled, $Facts.AsrRulesConfigured)
    }
    else {
        & $add 'DJ-008' 'Attack Surface Reduction (ASR) rules shall be enabled' 'Fail' ("No ASR rules in block/warn mode ({0} configured)" -f $Facts.AsrRulesConfigured)
    }

    # DJ-009 legacy protocols
    $legacyIssues = [System.Collections.Generic.List[string]]::new()
    if ($Facts.Smb1Enabled -eq $true) { $legacyIssues.Add('SMBv1 enabled') }
    if ($Facts.LlmnrPolicy -ne 0) { $legacyIssues.Add(("LLMNR not disabled by policy (EnableMulticast={0})" -f $Facts.LlmnrPolicy)) }
    $netbiosEnabled = @(@($Facts.NetbiosOptions) | Where-Object { $_ -ne 2 })
    if (@($Facts.NetbiosOptions).Count -gt 0 -and $netbiosEnabled.Count -gt 0) { $legacyIssues.Add(("NetBIOS over TCP/IP not disabled on {0} interface(s)" -f $netbiosEnabled.Count)) }
    if ($legacyIssues.Count -eq 0) {
        & $add 'DJ-009' 'Legacy protocols (SMBv1, LLMNR, NBT-NS) shall be disabled' 'Pass' 'SMBv1 off, LLMNR disabled by policy, NetBIOS disabled on all interfaces'
    }
    else {
        $status = if ($Facts.Smb1Enabled -eq $true) { 'Fail' } else { 'Warning' }
        & $add 'DJ-009' 'Legacy protocols (SMBv1, LLMNR, NBT-NS) shall be disabled' $status ($legacyIssues -join '; ')
    }

    # DJ-010 NTLMv1/LM disabled, SMB signing
    if ($Facts.LmCompatibilityLevel -ge 5 -and ($Facts.SmbServerSigningRequired -eq $true -or $Facts.SmbClientSigningEnabled -eq 1)) {
        & $add 'DJ-010' 'NTLMv1/LM authentication shall be disabled and SMB signing enabled' 'Pass' ("LmCompatibilityLevel={0}; SMB server signing required={1}" -f $Facts.LmCompatibilityLevel, $Facts.SmbServerSigningRequired)
    }
    else {
        & $add 'DJ-010' 'NTLMv1/LM authentication shall be disabled and SMB signing enabled' 'Fail' ("LmCompatibilityLevel={0}, SMB server signing required={1}, SMB client signing enabled={2}" -f $Facts.LmCompatibilityLevel, $Facts.SmbServerSigningRequired, $Facts.SmbClientSigningEnabled)
    }

    # DJ-011 least privilege on local Administrators
    $localAdmins = @($Facts.LocalAdministrators)
    $broadMembers = @($localAdmins | Where-Object { $_ -match '\\(Domain Users|Users|Everyone|Authenticated Users)$' })
    if ($localAdmins.Count -eq 0) {
        & $add 'DJ-011' 'Standard users shall not hold local administrator rights (least privilege)' 'Error' 'Local Administrators membership could not be read'
    }
    elseif ($broadMembers.Count -gt 0) {
        & $add 'DJ-011' 'Standard users shall not hold local administrator rights (least privilege)' 'Fail' ("Broad groups in local Administrators: {0}" -f ($broadMembers -join ', '))
    }
    else {
        $status = if ($localAdmins.Count -le 3) { 'Pass' } else { 'Warning' }
        & $add 'DJ-011' 'Standard users shall not hold local administrator rights (least privilege)' $status ("Members: {0}" -f ($localAdmins -join ', '))
    }

    # DJ-012 UAC
    if ($Facts.EnableLua -eq 1 -and $Facts.ConsentPromptBehaviorAdmin -ge 2 -and $Facts.PromptOnSecureDesktop -ne 0) {
        & $add 'DJ-012' 'User Account Control (UAC) shall be enabled at recommended settings' 'Pass' ("EnableLUA=1, ConsentPromptBehaviorAdmin={0}, secure desktop on" -f $Facts.ConsentPromptBehaviorAdmin)
    }
    else {
        & $add 'DJ-012' 'User Account Control (UAC) shall be enabled at recommended settings' 'Fail' ("EnableLUA={0}, ConsentPromptBehaviorAdmin={1}, PromptOnSecureDesktop={2}" -f $Facts.EnableLua, $Facts.ConsentPromptBehaviorAdmin, $Facts.PromptOnSecureDesktop)
    }

    # DJ-013 Credential Guard + LSA protection
    $lsaProtected = ($Facts.RunAsPpl -ge 1)
    if ($Facts.CredentialGuardRunning -eq $true -and $lsaProtected) {
        & $add 'DJ-013' 'Credential Guard and LSA protection shall be enabled where supported' 'Pass' ("Credential Guard running; RunAsPPL={0}" -f $Facts.RunAsPpl)
    }
    elseif ($lsaProtected) {
        & $add 'DJ-013' 'Credential Guard and LSA protection shall be enabled where supported' 'Warning' ("LSA protection on (RunAsPPL={0}) but Credential Guard not running (may be unsupported hardware)" -f $Facts.RunAsPpl)
    }
    else {
        & $add 'DJ-013' 'Credential Guard and LSA protection shall be enabled where supported' 'Fail' ("CredentialGuardRunning={0}, RunAsPPL={1}" -f $Facts.CredentialGuardRunning, $Facts.RunAsPpl)
    }

    # DJ-014 PowerShell logging
    if ($Facts.ScriptBlockLogging -eq 1 -and $Facts.ModuleLogging -eq 1) {
        & $add 'DJ-014' 'PowerShell script-block and module logging shall be enabled' 'Pass' 'Script block and module logging enabled by policy'
    }
    else {
        & $add 'DJ-014' 'PowerShell script-block and module logging shall be enabled' 'Fail' ("EnableScriptBlockLogging={0}, EnableModuleLogging={1}" -f $Facts.ScriptBlockLogging, $Facts.ModuleLogging)
    }

    # DJ-015 baseline GPOs
    $gpoNames = @($Facts.AppliedGpoNames)
    if (-not [string]::IsNullOrWhiteSpace($BaselineGpoPattern)) {
        $matched = @($gpoNames | Where-Object { $_ -match $BaselineGpoPattern })
        if ($matched.Count -gt 0) {
            & $add 'DJ-015' 'CIS/Microsoft security baseline GPOs shall be applied to all domain-joined computers' 'Pass' ("Baseline GPO(s) applied: {0}" -f ($matched -join ', '))
        }
        else {
            & $add 'DJ-015' 'CIS/Microsoft security baseline GPOs shall be applied to all domain-joined computers' 'Fail' ("No applied GPO matches '{0}'. Applied: {1}" -f $BaselineGpoPattern, ($gpoNames -join ', '))
        }
    }
    else {
        & $add 'DJ-015' 'CIS/Microsoft security baseline GPOs shall be applied to all domain-joined computers' 'Manual' ("Applied GPOs: {0}. Provide -BaselineGpoPattern to verify automatically" -f ($gpoNames -join ', '))
    }

    # DJ-016 removable media / AutoRun
    if ($Facts.NoDriveTypeAutoRun -eq 255) {
        $storageNote = if ($Facts.RemovableStorageDenyAll -eq 1) { 'removable storage denied by policy' } else { 'removable storage not blocked by policy (verify DLP/endpoint control)' }
        $status = if ($Facts.RemovableStorageDenyAll -eq 1) { 'Pass' } else { 'Warning' }
        & $add 'DJ-016' 'Removable media and AutoRun/AutoPlay shall be controlled' $status ("AutoRun disabled for all drive types; {0}" -f $storageNote)
    }
    else {
        & $add 'DJ-016' 'Removable media and AutoRun/AutoPlay shall be controlled' 'Fail' ("NoDriveTypeAutoRun={0} (expected 255)" -f $Facts.NoDriveTypeAutoRun)
    }

    # DJ-018 screen lock + banner
    $timeout = $Facts.InactivityTimeoutSecs
    $bannerConfigured = (-not [string]::IsNullOrWhiteSpace([string]$Facts.LegalNoticeCaption)) -and (-not [string]::IsNullOrWhiteSpace([string]$Facts.LegalNoticeText))
    if ($timeout -ge 1 -and $timeout -le 900 -and $bannerConfigured) {
        & $add 'DJ-018' 'Screen-lock inactivity timeout and logon banner shall be configured' 'Pass' ("InactivityTimeoutSecs={0}; logon banner configured" -f $timeout)
    }
    else {
        & $add 'DJ-018' 'Screen-lock inactivity timeout and logon banner shall be configured' 'Fail' ("InactivityTimeoutSecs={0} (expected 1-900); banner configured={1}" -f $timeout, $bannerConfigured)
    }

    # DJ-019 time hierarchy
    if ([string]$Facts.TimeSyncType -eq 'NT5DS') {
        & $add 'DJ-019' 'Domain-joined computers shall synchronize time from the AD time hierarchy' 'Pass' 'W32Time type NT5DS (domain hierarchy)'
    }
    else {
        & $add 'DJ-019' 'Domain-joined computers shall synchronize time from the AD time hierarchy' 'Fail' ("W32Time type is '{0}' (expected NT5DS)" -f $Facts.TimeSyncType)
    }

    return $rows.ToArray()
}

function Test-DjDomainControls {
    param(
        [Parameter(Mandatory = $true)] $DomainFacts,
        [int]$StaleDays = 90,
        [int]$LapsMinimumCoveragePercent = 95
    )

    $rows = [System.Collections.Generic.List[object]]::new()
    $domain = [string]$DomainFacts.DomainName

    # DJ-003 LAPS coverage
    if ($DomainFacts.LapsSchemaPresent) {
        $total = [int]$DomainFacts.EnabledComputerCount
        $managed = [int]$DomainFacts.LapsManagedCount
        $coverage = if ($total -gt 0) { [math]::Round(100.0 * $managed / $total, 1) } else { 0 }
        $status = if ($coverage -ge $LapsMinimumCoveragePercent) { 'Pass' } elseif ($coverage -gt 0) { 'Warning' } else { 'Fail' }
        $rows.Add((New-ControlRow -ControlId 'DJ-003' -Control 'Local Administrator passwords shall be unique per device and centrally managed (Windows LAPS)' -Target $domain -Status $status -Evidence ("{0} of {1} enabled computers ({2}%) have a LAPS-managed password (threshold {3}%)" -f $managed, $total, $coverage, $LapsMinimumCoveragePercent)))
    }
    else {
        $rows.Add((New-ControlRow -ControlId 'DJ-003' -Control 'Local Administrator passwords shall be unique per device and centrally managed (Windows LAPS)' -Target $domain -Status 'Fail' -Evidence 'Neither Windows LAPS nor legacy LAPS schema attributes found in the directory'))
    }

    # DJ-017 stale computer accounts
    $staleCount = [int]$DomainFacts.StaleComputerCount
    if ($staleCount -eq 0) {
        $rows.Add((New-ControlRow -ControlId 'DJ-017' -Control 'Stale computer accounts shall be identified and disabled or removed' -Target $domain -Status 'Pass' -Evidence ("No enabled computer accounts inactive more than {0} days" -f $StaleDays)))
    }
    else {
        $rows.Add((New-ControlRow -ControlId 'DJ-017' -Control 'Stale computer accounts shall be identified and disabled or removed' -Target $domain -Status 'Warning' -Evidence ("{0} enabled computer account(s) inactive more than {1} days (sample: {2})" -f $staleCount, $StaleDays, (@($DomainFacts.StaleComputerSample) -join ', '))))
    }

    # DJ-020 domain join restriction
    $quota = $DomainFacts.MachineAccountQuota
    if ($quota -eq 0) {
        $rows.Add((New-ControlRow -ControlId 'DJ-020' -Control 'Domain join shall be restricted to authorized administrators and computer accounts monitored' -Target $domain -Status 'Pass' -Evidence 'ms-DS-MachineAccountQuota=0; verify computer-account creation alerting in the SIEM'))
    }
    else {
        $rows.Add((New-ControlRow -ControlId 'DJ-020' -Control 'Domain join shall be restricted to authorized administrators and computer accounts monitored' -Target $domain -Status 'Fail' -Evidence ("ms-DS-MachineAccountQuota={0}; any authenticated user can join {0} computer(s) to the domain" -f $quota)))
    }

    return $rows.ToArray()
}

function Get-DjDomainFacts {
    param(
        [pscredential]$Credential,
        [int]$StaleDays = 90
    )

    Import-Module ActiveDirectory -ErrorAction Stop
    $adParams = @{}
    if ($Credential) { $adParams.Credential = $Credential }

    $domainObject = Get-ADDomain @adParams
    $facts = @{ DomainName = $domainObject.DNSRoot }

    $domainRoot = Get-ADObject -Identity $domainObject.DistinguishedName -Properties 'ms-DS-MachineAccountQuota' @adParams
    $facts.MachineAccountQuota = $domainRoot.'ms-DS-MachineAccountQuota'

    $enabledComputers = @(Get-ADComputer -Filter 'Enabled -eq $true' -Properties lastLogonTimestamp @adParams)
    $facts.EnabledComputerCount = $enabledComputers.Count

    $staleThreshold = (Get-Date).AddDays(-$StaleDays).ToFileTimeUtc()
    $staleComputers = @($enabledComputers | Where-Object { $_.lastLogonTimestamp -and $_.lastLogonTimestamp -lt $staleThreshold })
    $facts.StaleComputerCount = $staleComputers.Count
    $facts.StaleComputerSample = @($staleComputers | Select-Object -First 10 | ForEach-Object Name)

    $facts.LapsSchemaPresent = $false
    $facts.LapsManagedCount = 0
    foreach ($lapsAttribute in @('msLAPS-PasswordExpirationTime','ms-Mcs-AdmPwdExpirationTime')) {
        try {
            $managed = @(Get-ADComputer -Filter ('Enabled -eq $true -and {0} -like "*"' -f $lapsAttribute) @adParams)
            $facts.LapsSchemaPresent = $true
            if ($managed.Count -gt $facts.LapsManagedCount) { $facts.LapsManagedCount = $managed.Count }
        } catch { }
    }

    return $facts
}

function Get-TargetComputers {
    param(
        [string[]]$ComputerName,
        [string]$ComputersCsv,
        [string]$SearchBase,
        [int]$SampleSize,
        [pscredential]$Credential
    )

    if ($ComputerName.Count -gt 0) { return $ComputerName }

    if (-not [string]::IsNullOrWhiteSpace($ComputersCsv)) {
        $csvRows = @(Import-Csv -Path $ComputersCsv)
        if ($csvRows.Count -eq 0) { throw "No rows found in $ComputersCsv" }
        $column = if ($csvRows[0].PSObject.Properties['ComputerName']) { 'ComputerName' }
                  elseif ($csvRows[0].PSObject.Properties['Name']) { 'Name' }
                  else { throw "CSV $ComputersCsv must contain a ComputerName or Name column" }
        return @($csvRows | ForEach-Object { $_.$column } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }

    Import-Module ActiveDirectory -ErrorAction Stop
    $adParams = @{ Filter = 'Enabled -eq $true' }
    if ($Credential) { $adParams.Credential = $Credential }
    if (-not [string]::IsNullOrWhiteSpace($SearchBase)) { $adParams.SearchBase = $SearchBase }
    $computers = @(Get-ADComputer @adParams | ForEach-Object { if ($_.DNSHostName) { $_.DNSHostName } else { $_.Name } })
    return @($computers | Select-Object -First $SampleSize)
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
$allRows = [System.Collections.Generic.List[object]]::new()

Write-Host 'Collecting domain-level facts (LAPS coverage, stale accounts, machine account quota)...' -ForegroundColor Cyan
$domainFacts = Get-DjDomainFacts -Credential $Credential -StaleDays $StaleDays
foreach ($row in (Test-DjDomainControls -DomainFacts $domainFacts -StaleDays $StaleDays -LapsMinimumCoveragePercent $LapsMinimumCoveragePercent)) {
    $allRows.Add($row)
}

$targets = @(Get-TargetComputers -ComputerName $ComputerName -ComputersCsv $ComputersCsv -SearchBase $SearchBase -SampleSize $SampleSize -Credential $Credential)
Write-Host ("Assessing {0} computer(s) remotely" -f $targets.Count) -ForegroundColor Cyan

foreach ($target in $targets) {
    Write-Host ("  Collecting facts from {0}..." -f $target) -ForegroundColor Cyan
    try {
        $facts = Get-DjFacts -ComputerName $target -Credential $Credential
        foreach ($row in (Test-DjComputerControls -ComputerName $target -Facts $facts -PatchAgeDays $PatchAgeDays -BaselineGpoPattern $BaselineGpoPattern)) {
            $allRows.Add($row)
        }
    }
    catch {
        $allRows.Add((New-ControlRow -ControlId 'DJ-000' -Control 'Fact collection' -Target $target -Status 'Error' -Evidence $_.Exception.Message))
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
