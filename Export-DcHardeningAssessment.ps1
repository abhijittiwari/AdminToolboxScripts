<#
.SYNOPSIS
    Assesses Domain Controller security hardening controls DC-001 to DC-022 (Appendix B baseline).

.DESCRIPTION
    Read-only assessment of Domain Controllers against the security hardening baseline:
    supported OS and patch currency, Secure Boot, BitLocker, firewall, anti-malware/EDR,
    service/role footprint, PowerShell restrictions, LDAP signing and channel binding,
    SMB signing, NTLM restrictions, Netlogon secure channel, hardened UNC paths,
    Print Spooler, advanced audit policy and log forwarding, user rights assignments,
    krbtgt password age, and system-state backup currency.

    Per-DC facts are collected remotely with Invoke-Command (WinRM); domain-level facts
    (krbtgt, privileged groups, backup metadata) are collected once via the ActiveDirectory
    module. Each control is emitted as one row per target with Status:
      Pass / Fail / Warning / Manual / Error
    Controls that cannot be fully verified programmatically (e.g. SIEM onboarding,
    restore testing) are reported as Manual with the evidence that was collected.

    The script makes no changes to any system.

.PARAMETER Server
    One or more Domain Controller names to assess. Defaults to all DCs in the domain.

.PARAMETER Credential
    Optional credential for AD queries and remoting.

.PARAMETER PatchAgeDays
    Maximum age in days of the most recent installed hotfix before DC-002 is flagged. Default 30.

.PARAMETER KrbtgtMaxAgeDays
    Maximum krbtgt password age in days before DC-021 fails. Default 180.

.PARAMETER BackupMaxAgeDays
    Maximum system-state backup age in days before DC-022 fails. Default 7.

.PARAMETER MaxDomainAdmins
    Domain Admins member count above which DC-009 warns. Default 5.

.PARAMETER OutputPath
    Optional CSV output path. If omitted, results are only written to the pipeline/table.

.EXAMPLE
    .\Export-DcHardeningAssessment.ps1 -OutputPath .\DcHardening.csv

.EXAMPLE
    .\Export-DcHardeningAssessment.ps1 -Server DC01,DC02 -KrbtgtMaxAgeDays 90
#>

[CmdletBinding()]
param(
    [string[]]$Server = @(),
    [pscredential]$Credential,
    [int]$PatchAgeDays = 30,
    [int]$KrbtgtMaxAgeDays = 180,
    [int]$BackupMaxAgeDays = 7,
    [int]$MaxDomainAdmins = 5,
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

# Remote fact collection. Runs on each DC under Windows PowerShell 5.1; must stay
# self-contained (no outer-scope references except $using: parameters).
$script:DcFactScriptBlock = {
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

    try { $facts.SecureBoot = Confirm-SecureBootUEFI } catch { $facts.SecureBoot = $null }

    try {
        $blv = Get-BitLockerVolume -MountPoint $env:SystemDrive -ErrorAction Stop
        $facts.BitLockerProtection = [string]$blv.ProtectionStatus
    } catch { $facts.BitLockerProtection = $null }

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

    $spooler = Get-Service -Name Spooler -ErrorAction SilentlyContinue
    $facts.SpoolerStatus    = if ($spooler) { [string]$spooler.Status } else { 'NotInstalled' }
    $facts.SpoolerStartType = if ($spooler) { [string]$spooler.StartType } else { 'NotInstalled' }

    $unnecessaryServiceNames = @('Fax','Browser','SSDPSRV','upnphost','RemoteRegistry','SharedAccess','WMPNetworkSvc','XblAuthManager','icssvc')
    $facts.UnnecessaryServicesRunning = @(Get-Service -ErrorAction SilentlyContinue |
        Where-Object { $unnecessaryServiceNames -contains $_.Name -and $_.Status -eq 'Running' } |
        ForEach-Object Name)

    $facts.InstalledRoles = @(Get-WindowsFeature -ErrorAction SilentlyContinue |
        Where-Object { $_.Installed -and $_.FeatureType -eq 'Role' } | ForEach-Object Name)
    $ps2 = Get-WindowsFeature -Name PowerShell-V2 -ErrorAction SilentlyContinue
    $facts.PowerShellV2Installed = if ($ps2) { [bool]$ps2.Installed } else { $false }

    $facts.ScriptBlockLogging = Get-RegValue 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging' 'EnableScriptBlockLogging'
    $facts.ModuleLogging      = Get-RegValue 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\PowerShell\ModuleLogging' 'EnableModuleLogging'
    $facts.Transcription      = Get-RegValue 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\PowerShell\Transcription' 'EnableTranscripting'

    $ntds = 'HKLM:\SYSTEM\CurrentControlSet\Services\NTDS\Parameters'
    $facts.LdapServerIntegrity       = Get-RegValue $ntds 'LDAPServerIntegrity'
    $facts.LdapEnforceChannelBinding = Get-RegValue $ntds 'LdapEnforceChannelBinding'

    $facts.SmbServerRequireSigning = Get-RegValue 'HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters' 'RequireSecuritySignature'
    $facts.SmbClientRequireSigning = Get-RegValue 'HKLM:\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Parameters' 'RequireSecuritySignature'

    $lsa = 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa'
    $facts.LmCompatibilityLevel = Get-RegValue $lsa 'LmCompatibilityLevel'
    $facts.NoLMHash             = Get-RegValue $lsa 'NoLMHash'
    $facts.AuditReceivingNtlm   = Get-RegValue "$lsa\MSV1_0" 'AuditReceivingNTLMTraffic'
    $facts.RestrictReceivingNtlm = Get-RegValue "$lsa\MSV1_0" 'RestrictReceivingNTLMTraffic'

    $netlogon = 'HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon\Parameters'
    $facts.RequireSignOrSeal            = Get-RegValue $netlogon 'RequireSignOrSeal'
    $facts.VulnerableChannelAllowList   = Get-RegValue $netlogon 'VulnerableChannelAllowList'
    $facts.FullSecureChannelProtection  = Get-RegValue $netlogon 'FullSecureChannelProtection'

    $hardenedPaths = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\NetworkProvider\HardenedPaths'
    $facts.HardenedPathSysvol   = Get-RegValue $hardenedPaths '\\*\SYSVOL'
    $facts.HardenedPathNetlogon = Get-RegValue $hardenedPaths '\\*\NETLOGON'

    $facts.WefSubscriptionManager = Get-RegValue 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\EventLog\EventForwarding\SubscriptionManager' '1'
    $siemServiceNames = @('HealthService','AzureMonitorAgent','SplunkForwarder','nxlog','winlogbeat')
    $facts.SiemAgentsRunning = @(Get-Service -ErrorAction SilentlyContinue |
        Where-Object { $siemServiceNames -contains $_.Name -and $_.Status -eq 'Running' } |
        ForEach-Object Name)

    $facts.AuditPolicy = @{}
    $auditCsv = auditpol /get /category:* /r 2>$null
    if ($auditCsv) {
        try {
            foreach ($line in ($auditCsv | Where-Object { $_ } | ConvertFrom-Csv)) {
                $facts.AuditPolicy[$line.Subcategory] = $line.'Inclusion Setting'
            }
        } catch { }
    }

    $facts.UserRights = @{}
    $seceditFile = Join-Path $env:TEMP ("dcassess-{0}.inf" -f [guid]::NewGuid())
    try {
        secedit /export /cfg $seceditFile /areas USER_RIGHTS /quiet | Out-Null
        foreach ($line in (Get-Content -Path $seceditFile -ErrorAction SilentlyContinue)) {
            if ($line -match '^(Se\S+)\s*=\s*(.+)$') {
                $facts.UserRights[$Matches[1]] = ($Matches[2] -split ',' | ForEach-Object { $_.Trim() })
            }
        }
    } finally {
        Remove-Item -Path $seceditFile -ErrorAction SilentlyContinue
    }

    return $facts
}

function Get-DcFacts {
    param(
        [Parameter(Mandatory = $true)] [string]$ComputerName,
        [pscredential]$Credential
    )

    $invokeParams = @{
        ComputerName = $ComputerName
        ScriptBlock  = $script:DcFactScriptBlock
        ErrorAction  = 'Stop'
    }
    if ($Credential) { $invokeParams.Credential = $Credential }
    return Invoke-Command @invokeParams
}

function Test-DcComputerControls {
    param(
        [Parameter(Mandatory = $true)] [string]$ComputerName,
        [Parameter(Mandatory = $true)] $Facts,
        [int]$PatchAgeDays = 30
    )

    $rows = [System.Collections.Generic.List[object]]::new()
    $add = {
        param($Id, $Control, $Status, $Evidence)
        $rows.Add((New-ControlRow -ControlId $Id -Control $Control -Target $ComputerName -Status $Status -Evidence $Evidence))
    }

    # DC-001 supported OS
    $osCaption = [string]$Facts.OsCaption
    $supportedPattern = 'Server (2016|2019|2022|2025)'
    if ($osCaption -match $supportedPattern) {
        $status = if ($osCaption -match 'Server 2016') { 'Warning' } else { 'Pass' }
        $note = if ($status -eq 'Warning') { ' (2016 approaching end of support; plan upgrade)' } else { '' }
        & $add 'DC-001' 'Domain Controllers shall run supported Windows Server versions' $status ("{0} build {1}{2}" -f $osCaption, $Facts.OsBuild, $note)
    }
    else {
        & $add 'DC-001' 'Domain Controllers shall run supported Windows Server versions' 'Fail' ("Unsupported or unknown OS: {0}" -f $osCaption)
    }

    # DC-002 patch currency (best effort; cross-check with patch management tooling)
    if ($Facts.LastHotfixDate) {
        $ageDays = [int]((Get-Date) - [datetime]$Facts.LastHotfixDate).TotalDays
        $status = if ($ageDays -le $PatchAgeDays) { 'Pass' } else { 'Warning' }
        & $add 'DC-002' 'Latest security patches shall be installed' $status ("Most recent hotfix installed {0:yyyy-MM-dd} ({1} days ago); cross-check missing updates in patch management tooling" -f [datetime]$Facts.LastHotfixDate, $ageDays)
    }
    else {
        & $add 'DC-002' 'Latest security patches shall be installed' 'Manual' 'No hotfix install date available; verify patch level via patch management tooling'
    }

    # DC-003 Secure Boot
    if ($Facts.SecureBoot -eq $true) { & $add 'DC-003' 'Secure Boot shall be enabled' 'Pass' 'Confirm-SecureBootUEFI returned True' }
    elseif ($Facts.SecureBoot -eq $false) { & $add 'DC-003' 'Secure Boot shall be enabled' 'Fail' 'UEFI present but Secure Boot disabled' }
    else { & $add 'DC-003' 'Secure Boot shall be enabled' 'Fail' 'Secure Boot state unavailable (legacy BIOS firmware or query failed)' }

    # DC-004 BitLocker
    switch ([string]$Facts.BitLockerProtection) {
        'On'    { & $add 'DC-004' 'BitLocker shall be enabled' 'Pass' 'System drive protection is On' }
        'Off'   { & $add 'DC-004' 'BitLocker shall be enabled' 'Fail' 'System drive protection is Off' }
        default { & $add 'DC-004' 'BitLocker shall be enabled' 'Fail' 'BitLocker status unavailable (feature likely not installed)' }
    }

    # DC-005 firewall on all profiles
    $profiles = @($Facts.FirewallProfiles)
    $disabledProfiles = @($profiles | Where-Object { -not $_.Enabled } | ForEach-Object Name)
    if ($profiles.Count -eq 0) { & $add 'DC-005' 'Windows Firewall shall be enabled' 'Error' 'Firewall profile state could not be read' }
    elseif ($disabledProfiles.Count -eq 0) { & $add 'DC-005' 'Windows Firewall shall be enabled' 'Pass' 'All firewall profiles enabled' }
    else { & $add 'DC-005' 'Windows Firewall shall be enabled' 'Fail' ("Disabled profiles: {0}" -f ($disabledProfiles -join ', ')) }

    # DC-006 anti-malware
    if ($Facts.AntivirusEnabled -eq $true) {
        $status = if ($Facts.RealTimeProtection -eq $true) { 'Pass' } else { 'Warning' }
        & $add 'DC-006' 'Anti-malware protection shall be installed' $status ("Defender AV enabled; real-time protection: {0}; signature age {1} day(s)" -f $Facts.RealTimeProtection, $Facts.AvSignatureAgeDays)
    }
    elseif (@($Facts.EdrServicesRunning).Count -gt 0) {
        & $add 'DC-006' 'Anti-malware protection shall be installed' 'Warning' ("Defender AV not active; third-party agent running: {0} - verify AV coverage" -f (@($Facts.EdrServicesRunning) -join ', '))
    }
    else {
        & $add 'DC-006' 'Anti-malware protection shall be installed' 'Fail' 'No active anti-malware detected'
    }

    # DC-007 EDR
    if (@($Facts.EdrServicesRunning).Count -gt 0) {
        & $add 'DC-007' 'EDR shall be deployed' 'Pass' ("EDR service(s) running: {0}" -f (@($Facts.EdrServicesRunning) -join ', '))
    }
    else {
        & $add 'DC-007' 'EDR shall be deployed' 'Manual' 'No known EDR service detected; verify EDR onboarding in the security console'
    }

    # DC-010 unnecessary services
    $running = @($Facts.UnnecessaryServicesRunning)
    if ($running.Count -eq 0) { & $add 'DC-010' 'Unnecessary services shall be disabled' 'Pass' 'No known-unnecessary services running' }
    else { & $add 'DC-010' 'Unnecessary services shall be disabled' 'Warning' ("Running services to review: {0}" -f ($running -join ', ')) }

    # DC-011 PowerShell restrictions
    $psIssues = [System.Collections.Generic.List[string]]::new()
    if ($Facts.PowerShellV2Installed) { $psIssues.Add('PowerShell v2 engine installed (downgrade attack surface)') }
    if ($Facts.ScriptBlockLogging -ne 1) { $psIssues.Add('script block logging not enabled') }
    if ($Facts.ModuleLogging -ne 1) { $psIssues.Add('module logging not enabled') }
    if ($psIssues.Count -eq 0) {
        & $add 'DC-011' 'PowerShell restrictions shall be implemented' 'Pass' ("Script block + module logging enabled, PSv2 removed; transcription: {0}" -f ($(if ($Facts.Transcription -eq 1) { 'enabled' } else { 'disabled' })))
    }
    else {
        $status = if ($Facts.PowerShellV2Installed) { 'Fail' } else { 'Warning' }
        & $add 'DC-011' 'PowerShell restrictions shall be implemented' $status ($psIssues -join '; ')
    }

    # DC-012 non-AD workloads
    $allowedRoles = @('AD-Domain-Services','DNS','FileAndStorage-Services','FS-FileServer','Storage-Services','Windows-Defender')
    $extraRoles = @($Facts.InstalledRoles | Where-Object { $allowedRoles -notcontains $_ })
    if ($extraRoles.Count -eq 0) { & $add 'DC-012' 'Domain Controllers shall not be used for non-AD workloads' 'Pass' ("Installed roles: {0}" -f (@($Facts.InstalledRoles) -join ', ')) }
    else { & $add 'DC-012' 'Domain Controllers shall not be used for non-AD workloads' 'Warning' ("Non-AD roles installed: {0}" -f ($extraRoles -join ', ')) }

    # DC-013 LDAP signing + channel binding
    if ($Facts.LdapServerIntegrity -eq 2 -and $Facts.LdapEnforceChannelBinding -eq 2) {
        & $add 'DC-013' 'LDAP server signing and LDAP channel binding shall be enforced' 'Pass' 'LDAPServerIntegrity=2, LdapEnforceChannelBinding=2'
    }
    else {
        & $add 'DC-013' 'LDAP server signing and LDAP channel binding shall be enforced' 'Fail' ("LDAPServerIntegrity={0}, LdapEnforceChannelBinding={1} (both should be 2)" -f $Facts.LdapServerIntegrity, $Facts.LdapEnforceChannelBinding)
    }

    # DC-014 SMB signing
    if ($Facts.SmbServerRequireSigning -eq 1 -and $Facts.SmbClientRequireSigning -eq 1) {
        & $add 'DC-014' 'SMB signing shall be required for all SMB communications' 'Pass' 'Server and client RequireSecuritySignature=1'
    }
    else {
        & $add 'DC-014' 'SMB signing shall be required for all SMB communications' 'Fail' ("Server RequireSecuritySignature={0}, client RequireSecuritySignature={1} (both should be 1)" -f $Facts.SmbServerRequireSigning, $Facts.SmbClientRequireSigning)
    }

    # DC-015 NTLMv1/LM disabled, NTLM minimized
    $ntlmEvidence = "LmCompatibilityLevel={0}, NoLMHash={1}, AuditReceivingNTLMTraffic={2}, RestrictReceivingNTLMTraffic={3}" -f $Facts.LmCompatibilityLevel, $Facts.NoLMHash, $Facts.AuditReceivingNtlm, $Facts.RestrictReceivingNtlm
    if ($Facts.LmCompatibilityLevel -ge 5 -and $Facts.NoLMHash -ne 0) {
        $status = if ($Facts.AuditReceivingNtlm -ge 1 -or $Facts.RestrictReceivingNtlm -ge 1) { 'Pass' } else { 'Warning' }
        $note = if ($status -eq 'Warning') { '; NTLM usage audit/restriction not configured' } else { '' }
        & $add 'DC-015' 'NTLMv1 and LM authentication shall be disabled and overall NTLM usage minimized' $status ($ntlmEvidence + $note)
    }
    else {
        & $add 'DC-015' 'NTLMv1 and LM authentication shall be disabled and overall NTLM usage minimized' 'Fail' ($ntlmEvidence + ' (LmCompatibilityLevel should be 5 and NoLMHash 1)')
    }

    # DC-016 Netlogon secure channel
    $allowList = [string]$Facts.VulnerableChannelAllowList
    if (-not [string]::IsNullOrWhiteSpace($allowList)) {
        & $add 'DC-016' 'Vulnerable Netlogon secure channel connections shall be blocked' 'Fail' ("VulnerableChannelAllowList exempts accounts: {0}" -f $allowList)
    }
    elseif ($Facts.RequireSignOrSeal -eq 0) {
        & $add 'DC-016' 'Vulnerable Netlogon secure channel connections shall be blocked' 'Fail' 'RequireSignOrSeal=0'
    }
    else {
        & $add 'DC-016' 'Vulnerable Netlogon secure channel connections shall be blocked' 'Pass' 'Enforcement default active; no vulnerable-channel allow list'
    }

    # DC-017 hardened UNC paths
    $sysvolOk = ([string]$Facts.HardenedPathSysvol -match 'RequireMutualAuthentication=1' -and [string]$Facts.HardenedPathSysvol -match 'RequireIntegrity=1')
    $netlogonOk = ([string]$Facts.HardenedPathNetlogon -match 'RequireMutualAuthentication=1' -and [string]$Facts.HardenedPathNetlogon -match 'RequireIntegrity=1')
    if ($sysvolOk -and $netlogonOk) {
        & $add 'DC-017' 'Hardened UNC paths shall be enforced for SYSVOL and NETLOGON shares' 'Pass' 'Mutual authentication and integrity required for \\*\SYSVOL and \\*\NETLOGON'
    }
    else {
        & $add 'DC-017' 'Hardened UNC paths shall be enforced for SYSVOL and NETLOGON shares' 'Fail' ("SYSVOL='{0}', NETLOGON='{1}'" -f $Facts.HardenedPathSysvol, $Facts.HardenedPathNetlogon)
    }

    # DC-018 Print Spooler
    if ($Facts.SpoolerStatus -eq 'NotInstalled' -or ($Facts.SpoolerStatus -ne 'Running' -and $Facts.SpoolerStartType -eq 'Disabled')) {
        & $add 'DC-018' 'The Print Spooler service shall be disabled on Domain Controllers' 'Pass' ("Spooler: {0} (start type {1})" -f $Facts.SpoolerStatus, $Facts.SpoolerStartType)
    }
    else {
        & $add 'DC-018' 'The Print Spooler service shall be disabled on Domain Controllers' 'Fail' ("Spooler: {0} (start type {1})" -f $Facts.SpoolerStatus, $Facts.SpoolerStartType)
    }

    # DC-019 advanced audit policy + central collection
    $requiredSubcategories = @(
        'Credential Validation','Kerberos Authentication Service','Kerberos Service Ticket Operations',
        'Computer Account Management','Security Group Management','User Account Management',
        'Directory Service Access','Directory Service Changes','Audit Policy Change',
        'Special Logon','Logon','Account Lockout','Sensitive Privilege Use','Security System Extension'
    )
    $auditPolicy = $Facts.AuditPolicy
    if (-not $auditPolicy -or $auditPolicy.Count -eq 0) {
        & $add 'DC-019' 'Advanced audit policy shall be configured and security logs centrally collected (SIEM)' 'Error' 'auditpol output unavailable or not parseable'
    }
    else {
        $notAudited = @($requiredSubcategories | Where-Object {
            $setting = [string]$auditPolicy[$_]
            [string]::IsNullOrWhiteSpace($setting) -or $setting -match 'No Auditing'
        })
        $forwarding = @(@($Facts.SiemAgentsRunning) + @($(if ($Facts.WefSubscriptionManager) { 'WEF subscription' })) | Where-Object { $_ })
        $forwardingEvidence = if ($forwarding.Count -gt 0) { "log collection: $($forwarding -join ', ')" } else { 'no WEF subscription or known SIEM agent detected' }
        if ($notAudited.Count -eq 0 -and $forwarding.Count -gt 0) {
            & $add 'DC-019' 'Advanced audit policy shall be configured and security logs centrally collected (SIEM)' 'Pass' ("All key subcategories audited; {0}" -f $forwardingEvidence)
        }
        elseif ($notAudited.Count -eq 0) {
            & $add 'DC-019' 'Advanced audit policy shall be configured and security logs centrally collected (SIEM)' 'Manual' ("Audit policy OK; {0} - confirm SIEM onboarding" -f $forwardingEvidence)
        }
        else {
            & $add 'DC-019' 'Advanced audit policy shall be configured and security logs centrally collected (SIEM)' 'Fail' ("Subcategories not audited: {0}; {1}" -f ($notAudited -join ', '), $forwardingEvidence)
        }
    }

    # DC-020 user rights assignments
    $userRights = $Facts.UserRights
    if (-not $userRights -or $userRights.Count -eq 0) {
        & $add 'DC-020' 'User rights assignments on Domain Controllers shall be restricted to authorized administrative groups' 'Error' 'secedit user-rights export unavailable'
    }
    else {
        # Broad principals that must never hold interactive/remote/debug rights on a DC.
        $broadSids = @('*S-1-1-0','*S-1-5-32-545','*S-1-5-11','*S-1-5-32-546')  # Everyone, Users, Authenticated Users, Guests
        $sensitiveRights = @('SeInteractiveLogonRight','SeRemoteInteractiveLogonRight','SeDebugPrivilege','SeTcbPrivilege','SeTakeOwnershipPrivilege','SeRestorePrivilege','SeBackupPrivilege')
        $violations = [System.Collections.Generic.List[string]]::new()
        foreach ($right in $sensitiveRights) {
            if ($userRights.ContainsKey($right)) {
                $broad = @($userRights[$right] | Where-Object { $broadSids -contains $_ })
                if ($broad.Count -gt 0) { $violations.Add(("{0}: {1}" -f $right, ($broad -join ','))) }
            }
        }
        if ($violations.Count -gt 0) {
            & $add 'DC-020' 'User rights assignments on Domain Controllers shall be restricted to authorized administrative groups' 'Fail' ("Broad principals hold sensitive rights: {0}" -f ($violations -join '; '))
        }
        else {
            & $add 'DC-020' 'User rights assignments on Domain Controllers shall be restricted to authorized administrative groups' 'Manual' ("No broad principals on sensitive rights; review full assignment list against the tiering model ({0} rights exported)" -f $userRights.Count)
        }
    }

    return $rows.ToArray()
}

function Test-DcDomainControls {
    param(
        [Parameter(Mandatory = $true)] $DomainFacts,
        [int]$KrbtgtMaxAgeDays = 180,
        [int]$BackupMaxAgeDays = 7,
        [int]$MaxDomainAdmins = 5
    )

    $rows = [System.Collections.Generic.List[object]]::new()
    $domain = [string]$DomainFacts.DomainName

    # DC-008 built-in Administrator (RID 500)
    $builtinIssues = [System.Collections.Generic.List[string]]::new()
    if ($DomainFacts.BuiltinAdminEnabled) { $builtinIssues.Add('account enabled') }
    if ($DomainFacts.BuiltinAdminPasswordAgeDays -gt 365) { $builtinIssues.Add(("password {0} days old" -f $DomainFacts.BuiltinAdminPasswordAgeDays)) }
    if ($builtinIssues.Count -eq 0) {
        $rows.Add((New-ControlRow -ControlId 'DC-008' -Control 'Local Administrator accounts shall be restricted' -Target $domain -Status 'Pass' -Evidence ("Built-in Administrator ({0}) disabled and password rotated" -f $DomainFacts.BuiltinAdminName)))
    }
    else {
        $rows.Add((New-ControlRow -ControlId 'DC-008' -Control 'Local Administrator accounts shall be restricted' -Target $domain -Status 'Warning' -Evidence ("Built-in Administrator ({0}): {1}" -f $DomainFacts.BuiltinAdminName, ($builtinIssues -join '; '))))
    }

    # DC-009 administrative access restricted
    $groupCounts = $DomainFacts.PrivilegedGroupCounts
    $emptyExpected = @('Schema Admins','Account Operators','Server Operators','Print Operators')
    $issues = [System.Collections.Generic.List[string]]::new()
    foreach ($groupName in $emptyExpected) {
        if ($groupCounts.ContainsKey($groupName) -and [int]$groupCounts[$groupName] -gt 0) {
            $issues.Add(("{0} has {1} member(s), expected 0" -f $groupName, $groupCounts[$groupName]))
        }
    }
    if ($groupCounts.ContainsKey('Domain Admins') -and [int]$groupCounts['Domain Admins'] -gt $MaxDomainAdmins) {
        $issues.Add(("Domain Admins has {0} members (threshold {1})" -f $groupCounts['Domain Admins'], $MaxDomainAdmins))
    }
    $countsEvidence = ($groupCounts.Keys | Sort-Object | ForEach-Object { "{0}={1}" -f $_, $groupCounts[$_] }) -join ', '
    if ($issues.Count -eq 0) {
        $rows.Add((New-ControlRow -ControlId 'DC-009' -Control 'Administrative access shall be restricted' -Target $domain -Status 'Pass' -Evidence $countsEvidence))
    }
    else {
        $rows.Add((New-ControlRow -ControlId 'DC-009' -Control 'Administrative access shall be restricted' -Target $domain -Status 'Warning' -Evidence (($issues -join '; ') + ". Membership: $countsEvidence")))
    }

    # DC-021 krbtgt password age
    if ($DomainFacts.KrbtgtPasswordLastSet) {
        $krbtgtAge = [int]((Get-Date) - [datetime]$DomainFacts.KrbtgtPasswordLastSet).TotalDays
        $status = if ($krbtgtAge -le $KrbtgtMaxAgeDays) { 'Pass' } else { 'Fail' }
        $rows.Add((New-ControlRow -ControlId 'DC-021' -Control 'The krbtgt account password shall be rotated at regular intervals' -Target $domain -Status $status -Evidence ("krbtgt password last set {0:yyyy-MM-dd} ({1} days ago; threshold {2})" -f [datetime]$DomainFacts.KrbtgtPasswordLastSet, $krbtgtAge, $KrbtgtMaxAgeDays)))
    }
    else {
        $rows.Add((New-ControlRow -ControlId 'DC-021' -Control 'The krbtgt account password shall be rotated at regular intervals' -Target $domain -Status 'Error' -Evidence 'krbtgt PasswordLastSet could not be read'))
    }

    # DC-022 system-state backups per DC
    foreach ($backup in @($DomainFacts.DcBackups)) {
        if ($backup.LastBackup) {
            $backupAge = [int]((Get-Date) - [datetime]$backup.LastBackup).TotalDays
            $status = if ($backupAge -le $BackupMaxAgeDays) { 'Pass' } else { 'Fail' }
            $rows.Add((New-ControlRow -ControlId 'DC-022' -Control 'System-state backups shall be performed regularly and restore procedures tested' -Target $backup.Server -Status $status -Evidence ("dsaSignature backup marker {0:yyyy-MM-dd} ({1} days ago; threshold {2}); restore testing requires manual verification" -f [datetime]$backup.LastBackup, $backupAge, $BackupMaxAgeDays)))
        }
        else {
            $rows.Add((New-ControlRow -ControlId 'DC-022' -Control 'System-state backups shall be performed regularly and restore procedures tested' -Target $backup.Server -Status 'Fail' -Evidence 'No system-state backup marker (dsaSignature) found'))
        }
    }

    return $rows.ToArray()
}

function Get-DcDomainFacts {
    param([pscredential]$Credential)

    Import-Module ActiveDirectory -ErrorAction Stop
    $adParams = @{}
    if ($Credential) { $adParams.Credential = $Credential }

    $domainObject = Get-ADDomain @adParams
    $facts = @{ DomainName = $domainObject.DNSRoot }

    $krbtgt = Get-ADUser -Identity 'krbtgt' -Properties PasswordLastSet @adParams
    $facts.KrbtgtPasswordLastSet = $krbtgt.PasswordLastSet

    $builtinAdmin = Get-ADUser -Identity ("{0}-500" -f $domainObject.DomainSID.Value) -Properties Enabled, PasswordLastSet @adParams
    $facts.BuiltinAdminName = $builtinAdmin.SamAccountName
    $facts.BuiltinAdminEnabled = [bool]$builtinAdmin.Enabled
    $facts.BuiltinAdminPasswordAgeDays = if ($builtinAdmin.PasswordLastSet) { [int]((Get-Date) - $builtinAdmin.PasswordLastSet).TotalDays } else { $null }

    $facts.PrivilegedGroupCounts = @{}
    foreach ($groupName in @('Domain Admins','Enterprise Admins','Schema Admins','Administrators','Account Operators','Server Operators','Print Operators','Backup Operators')) {
        try {
            $members = @(Get-ADGroupMember -Identity $groupName -ErrorAction Stop @adParams)
            $facts.PrivilegedGroupCounts[$groupName] = $members.Count
        } catch { }
    }

    $facts.DcBackups = @()
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

function Get-TargetDomainControllers {
    param(
        [string[]]$Server,
        [pscredential]$Credential
    )

    if ($Server.Count -gt 0) { return $Server }
    Import-Module ActiveDirectory -ErrorAction Stop
    $adParams = @{}
    if ($Credential) { $adParams.Credential = $Credential }
    return @(Get-ADDomainController -Filter * @adParams | ForEach-Object HostName)
}

$allRows = [System.Collections.Generic.List[object]]::new()

Write-Host 'Collecting domain-level facts (krbtgt, privileged groups, backup markers)...' -ForegroundColor Cyan
$domainFacts = Get-DcDomainFacts -Credential $Credential
foreach ($row in (Test-DcDomainControls -DomainFacts $domainFacts -KrbtgtMaxAgeDays $KrbtgtMaxAgeDays -BackupMaxAgeDays $BackupMaxAgeDays -MaxDomainAdmins $MaxDomainAdmins)) {
    $allRows.Add($row)
}

$targets = @(Get-TargetDomainControllers -Server $Server -Credential $Credential)
Write-Host ("Assessing {0} Domain Controller(s): {1}" -f $targets.Count, ($targets -join ', ')) -ForegroundColor Cyan

foreach ($target in $targets) {
    Write-Host ("  Collecting facts from {0}..." -f $target) -ForegroundColor Cyan
    try {
        $facts = Get-DcFacts -ComputerName $target -Credential $Credential
        foreach ($row in (Test-DcComputerControls -ComputerName $target -Facts $facts -PatchAgeDays $PatchAgeDays)) {
            $allRows.Add($row)
        }
    }
    catch {
        $allRows.Add((New-ControlRow -ControlId 'DC-000' -Control 'Fact collection' -Target $target -Status 'Error' -Evidence $_.Exception.Message))
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
