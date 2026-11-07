#Requires -Version 5.1
<#
.SYNOPSIS
GPO startup deployment for a one-time, two-phase migration broker repair.
.DESCRIPTION
Startup (SYSTEM): installs a protected local copy, clears selected unloaded
profiles, records per-SID completion, and creates per-user logon tasks.
UserRepair (interactive user): re-registers AAD BrokerPlugin unconditionally
once for that repair version. No elevation, credentials, or UAC prompts.
With -CloseApplications, first force-closes Outlook, Teams, OneDrive and
Office applications in that user's current session. Unsaved work can be lost.
-ForceUserRepair allows an intentional manual rerun, including registration-only
when startup cleanup is unconfirmed. It never clears live caches or controls the
service and is not forwarded to GPO-created tasks.

Apply AFTER device migration and profile/SID reassignment are complete.
ProfilePathPattern is REQUIRED for Startup to make profile scope explicit.
Use a new RepairId only when deliberately requesting another cache reset.
See the accompanying deployment guide before configuring the GPO.
.EXAMPLE
.\Repair-MigrationTokenBroker-GPO.ps1 -Mode Startup -RepairId Alumtran-2026-11-v1 -ProfilePathPattern 'C:\Users\*' -FullBrokerReset -CloseApplications
#>
[CmdletBinding()]
param(
    [ValidateSet('Startup', 'UserRepair')]
    [string]$Mode = 'Startup',
    [ValidatePattern('^[A-Za-z0-9][A-Za-z0-9._-]{0,63}$')]
    [string]$RepairId = 'Alumtran-2026-11-v1',
    [string[]]$ProfilePathPattern,
    [switch]$FullBrokerReset,
    [switch]$CloseApplications,
    [switch]$ForceUserRepair
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if ($env:OS -ne 'Windows_NT' -or $PSVersionTable.PSEdition -ne 'Desktop' -or
    -not [Environment]::Is64BitProcess) {
    throw 'Use 64-bit Windows PowerShell 5.1 on Windows.'
}
if ($ForceUserRepair -and $Mode -ne 'UserRepair') {
    throw '-ForceUserRepair is only supported with -Mode UserRepair; do not add it to the startup GPO.'
}
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$sid = $identity.User.Value
$installRoot = Join-Path $env:ProgramFiles 'Alumtran\MigrationTokenBrokerRepair'
$localScript = Join-Path $installRoot 'Repair-MigrationTokenBroker-GPO.ps1'
$stateRoot = Join-Path $installRoot "State\$RepairId"
$manifest = Join-Path $env:windir 'SystemApps\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy\AppxManifest.xml'

function Assert-NoReparsePoint {
    param([Parameter(Mandatory)][string]$Path)
    # Walk up as well, so an AppData/Local/profile junction cannot redirect
    # a SYSTEM deletion outside the selected profile.
    $cursor = [IO.Path]::GetFullPath($Path)
    while ($cursor) {
        if (Test-Path -LiteralPath $cursor) {
            $item = Get-Item -LiteralPath $cursor -Force
            if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) {
                throw "Refusing reparse point: $cursor"
            }
        }
        $cursor = [IO.Path]::GetDirectoryName($cursor)
    }
}

function Remove-BrokerCachePath {
    param([Parameter(Mandatory)][string]$Path)
    Assert-NoReparsePoint -Path $Path
    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Host "Already absent: $Path"
        return
    }
    # Inspect descendants without traversing junctions/symlinks.
    $pending = New-Object 'System.Collections.Generic.Stack[string]'
    $pending.Push($Path)
    while ($pending.Count -gt 0) {
        $item = Get-Item -LiteralPath ($pending.Pop()) -Force
        if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) {
            throw "Refusing reparse point within cache: $($item.FullName)"
        }
        if ($item.PSIsContainer) {
            foreach ($child in [IO.Directory]::GetFileSystemEntries($item.FullName)) {
                $pending.Push($child)
            }
        }
    }
    Write-Host "Removing: $Path"
    Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
}

function Test-ProfileSelected {
    param([string]$Path, [string[]]$Patterns)
    foreach ($pattern in $Patterns) {
        if ($Path -like $pattern) { return $true }
    }
    return $false
}

function Get-OfflineProfileStatus {
    param([string]$ProfileSid, [string]$ProfilePath)
    $current = Get-CimInstance Win32_UserProfile -Filter "SID='$ProfileSid'" -ErrorAction Stop
    if (-not $current) {
        return [pscustomobject]@{ State='Deferred'; Reason='Profile record disappeared during this run' }
    }
    if ([IO.Path]::GetFullPath($current.LocalPath).TrimEnd('\') -ine $ProfilePath) {
        throw "Profile path changed during this run for $ProfileSid. Retry after profile migration completes."
    }
    $userHive = Test-Path -LiteralPath "Registry::HKEY_USERS\$ProfileSid"
    $classesHive = Test-Path -LiteralPath "Registry::HKEY_USERS\${ProfileSid}_Classes"
    if ($current.Loaded -or $userHive -or $classesHive) {
        return [pscustomobject]@{
            State='Deferred'
            Reason="CIM Loaded=$($current.Loaded); user hive loaded=$userHive; classes hive loaded=$classesHive"
        }
    }
    return [pscustomobject]@{ State='Ready'; Reason='Profile and registry hives are unloaded' }
}

function Get-MigrationApplications {
    param([int]$SessionId)
    $appNames = @('OUTLOOK', 'olk', 'Teams', 'ms-teams', 'msteams', 'OneDrive',
        'WINWORD', 'EXCEL', 'POWERPNT', 'MSACCESS', 'MSPUB', 'ONENOTE',
        'onenotem', 'VISIO', 'WINPROJ')
    Get-Process | Where-Object {
        $_.SessionId -eq $SessionId -and $_.ProcessName -in $appNames
    }
}

function Close-MigrationApplications {
    param([int]$SessionId)
    # Bound retries for apps that relaunch automatically. Never terminate
    # arbitrary processes, Windows shell processes, or another user session.
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        $apps = @(Get-MigrationApplications -SessionId $SessionId)
        if ($apps.Count -eq 0) {
            Write-Host "No matching Microsoft 365 apps remain in session $SessionId."
            return
        }
        foreach ($app in $apps) {
            $label = '{0} [{1}]' -f $app.ProcessName, $app.Id
            Write-Warning "Force-closing $label in session $SessionId (attempt $attempt)."
            try {
                $app | Stop-Process -Force -ErrorAction Stop
            } catch {
                # An app may exit naturally between enumeration and Stop-Process.
                if (-not $app.HasExited) { throw }
            }
            if (-not $app.WaitForExit(10000)) { throw "Application did not exit: $label" }
        }
    }
    $remaining = @(Get-MigrationApplications -SessionId $SessionId)
    if ($remaining.Count -gt 0) {
        throw "Applications remain/restarted: $(($remaining.ProcessName | Sort-Object -Unique) -join ', '). Broker registration was not attempted."
    }
}

function Write-UserDeviceStatus {
    try {
        & "$env:windir\System32\dsregcmd.exe" /status | Out-Host
        if ($LASTEXITCODE -ne 0) { Write-Warning "dsregcmd exited with code $LASTEXITCODE." }
    } catch {
        Write-Warning "Could not capture device status: $($_.Exception.Message)"
    }
}

function Invoke-UserBrokerRepair {
    param([switch]$CloseApplications, [switch]$ForceUserRepair)
    $sessionId = (Get-Process -Id $PID).SessionId
    Write-Host "UserRepair: user=$($identity.Name); SID=$sid; profile=$env:USERPROFILE; session=$sessionId"
    Write-Host "RepairId=$RepairId; CloseApplications=$CloseApplications; ForceUserRepair=$ForceUserRepair"
    if ($identity.IsSystem -or $sessionId -eq 0) {
        throw 'UserRepair must run as the affected user in their interactive Windows session, not SYSTEM.'
    }
    $machineMarker = Join-Path $stateRoot "$sid.json"
    $startupCompleted = Test-Path -LiteralPath $machineMarker
    if (-not $startupCompleted) {
        if (-not $ForceUserRepair) {
            Write-Warning "UserRepair blocked: no startup cleanup marker for this SID and RepairId at $machineMarker. No apps were closed. Startup may have deferred this profile. Exit code 2."
            Write-Host 'For a manual app-close and broker registration attempt, use -Mode UserRepair -CloseApplications -ForceUserRepair. This does not clear caches or stop TokenBroker.'
            return 2
        }
        Write-Warning 'Manual registration-only run: startup cache cleanup is NOT confirmed. No normal completion marker will be written.'
    } else {
        $state = Get-Content -LiteralPath $machineMarker -Raw | ConvertFrom-Json
        if ($state.RepairId -ne $RepairId -or $state.Sid -ne $sid -or
            $state.ProfilePath.TrimEnd('\') -ine $env:USERPROFILE.TrimEnd('\')) {
            throw 'Startup marker does not match this user/profile. Rerun startup after profile reassignment. ForceUserRepair does not bypass a mismatched marker.'
        }
    }
    $userKey = "HKCU:\Software\Alumtran\MigrationTokenBrokerRepair\$RepairId\$env:COMPUTERNAME"
    if ($startupCompleted -and (Test-Path -LiteralPath $userKey)) {
        $done = Get-ItemProperty -LiteralPath $userKey
        if ($done.PSObject.Properties['CompletedComputer'] -and
            $done.CompletedComputer -eq $env:COMPUTERNAME) {
            $appsAlreadyClosed = $done.PSObject.Properties['ApplicationsClosedComputer'] -and
                $done.ApplicationsClosedComputer -eq $env:COMPUTERNAME
            if ((-not $CloseApplications -or $appsAlreadyClosed) -and -not $ForceUserRepair) {
                Write-Host 'UserRepair skipped: this repair is already complete for this user and computer. No apps were closed. Use -ForceUserRepair for an intentional manual rerun.'
                return 0
            }
        }
    }
    if (-not (Test-Path -LiteralPath $manifest -PathType Leaf)) { throw 'AAD BrokerPlugin system manifest is missing.' }
    if ($CloseApplications) {
        Close-MigrationApplications -SessionId $sessionId
    } else {
        Write-Warning 'App closure was NOT requested. Add -CloseApplications to close the named Microsoft 365 apps in this session.'
    }
    Write-Host "Registering broker for $($identity.Name) [$sid] on $env:COMPUTERNAME"
    Add-AppxPackage -Register $manifest -DisableDevelopmentMode -ForceApplicationShutdown -ErrorAction Stop | Out-Null
    $package = Get-AppxPackage -Name Microsoft.AAD.BrokerPlugin -ErrorAction Stop
    if (-not $package -or $package.Status -ne 'Ok') { throw 'Broker package is missing or its status is not Ok.' }
    $package | Select-Object Name, PackageFullName, Status | Format-List | Out-Host
    Write-UserDeviceStatus
    if ($startupCompleted) {
        New-Item -Path $userKey -Force | Out-Null
        New-ItemProperty -LiteralPath $userKey -Name CompletedUtc -Value ([DateTime]::UtcNow.ToString('o')) -PropertyType String -Force | Out-Null
        New-ItemProperty -LiteralPath $userKey -Name CompletedComputer -Value $env:COMPUTERNAME -PropertyType String -Force | Out-Null
        if ($CloseApplications) {
            New-ItemProperty -LiteralPath $userKey -Name ApplicationsClosedComputer -Value $env:COMPUTERNAME -PropertyType String -Force | Out-Null
        }
    } else {
        Write-Warning 'Manual broker registration completed, but startup cache cleanup remains unconfirmed. No normal completion marker was written.'
    }
    Write-Host 'Registration completed. Fresh sign-in/MFA may be needed. Test Alumtran app sign-in.'
    return 0
}

if ($Mode -eq 'UserRepair') {
    # Start logging BEFORE checking markers so skipped/blocked manual attempts
    # produce evidence rather than silently exiting with no transcript.
    $logRoot = Join-Path $env:LOCALAPPDATA 'Microsoft\MigrationTokenBrokerRepair\Logs'
    New-Item -ItemType Directory -Path $logRoot -Force | Out-Null
    $log = Join-Path $logRoot ("{0}-{1}-{2}-UserRepair.log" -f $RepairId, (Get-Date -Format 'yyyyMMdd-HHmmss'), $PID)
    $exitCode = 1
    Start-Transcript -LiteralPath $log | Out-Null
    try {
        $exitCode = Invoke-UserBrokerRepair -CloseApplications:$CloseApplications -ForceUserRepair:$ForceUserRepair
    } catch {
        $exitCode = 1
        Write-Warning ($_ | Out-String)
    } finally {
        Write-Host "UserRepair exit code: $exitCode. Log: $log"
        Stop-Transcript | Out-Null
    }
    exit $exitCode
}

if (-not $identity.IsSystem) { throw 'Startup mode must run as SYSTEM through computer GPO startup.' }
if (-not $ProfilePathPattern -or @($ProfilePathPattern | Where-Object { -not $_ }).Count -gt 0) {
    throw 'Supply -ProfilePathPattern explicitly, for example C:\Users\* or C:\Users\jsmith*.'
}
if (-not (Test-Path -LiteralPath $manifest -PathType Leaf)) { throw 'AAD BrokerPlugin system manifest is missing; no caches were removed.' }

Assert-NoReparsePoint -Path $installRoot
New-Item -ItemType Directory -Path $installRoot -Force | Out-Null
# Fresh DACL: SYSTEM/admins can write, users can only read/execute.
$acl = New-Object Security.AccessControl.DirectorySecurity
$acl.SetAccessRuleProtection($true, $false)
$acl.SetOwner((New-Object Security.Principal.SecurityIdentifier('S-1-5-32-544')))
foreach ($entry in @(@('S-1-5-18', 'FullControl'), @('S-1-5-32-544', 'FullControl'), @('S-1-5-32-545', 'ReadAndExecute'))) {
    $account = New-Object Security.Principal.SecurityIdentifier($entry[0])
    $rule = New-Object Security.AccessControl.FileSystemAccessRule($account, $entry[1], 'ContainerInherit,ObjectInherit', 'None', 'Allow')
    $acl.AddAccessRule($rule)
}
Set-Acl -LiteralPath $installRoot -AclObject $acl
Assert-NoReparsePoint -Path $localScript
if ([IO.Path]::GetFullPath($PSCommandPath) -ine $localScript) {
    Copy-Item -LiteralPath $PSCommandPath -Destination $localScript -Force
}
Assert-NoReparsePoint -Path $stateRoot
New-Item -ItemType Directory -Path $stateRoot -Force | Out-Null
$logRoot = Join-Path $installRoot 'Logs'
Assert-NoReparsePoint -Path $logRoot
New-Item -ItemType Directory -Path $logRoot -Force | Out-Null
$log = Join-Path $logRoot ("{0}-{1}-Startup.log" -f $RepairId, (Get-Date -Format 'yyyyMMdd-HHmmss'))
$restoreService = $false
$servicePrepared = $false
$failed = $false
$summary = @{ Missing=0; Deferred=0; Cleaned=0; AlreadyClean=0; Tasks=0; Errors=0 }
Start-Transcript -LiteralPath $log | Out-Null
try {
    Write-Host "Startup repair $RepairId; full reset: $FullBrokerReset; profile scope: $($ProfilePathPattern -join ', ')"
    $profiles = @(Get-CimInstance Win32_UserProfile | Where-Object {
        -not $_.Special -and $_.LocalPath -and
        $_.SID -match '^S-1-(5-21|12-1)-' -and
        (Test-ProfileSelected -Path $_.LocalPath -Patterns $ProfilePathPattern)
    })
    foreach ($profile in $profiles) {
        try {
            $profileSid = $profile.SID
            $profilePath = [IO.Path]::GetFullPath($profile.LocalPath).TrimEnd('\')
            Assert-NoReparsePoint -Path $profilePath
            if (-not (Test-Path -LiteralPath $profilePath)) {
                $summary.Missing++
                Write-Warning "Skipping profile record with missing directory: $profileSid at $profilePath. No cache, marker, task or profile registration was changed for this record."
                continue
            }
            if (-not (Test-Path -LiteralPath $profilePath -PathType Container)) {
                throw "Profile path is not a directory: $profilePath"
            }
            $marker = Join-Path $stateRoot "$profileSid.json"
            $cleaned = $false
            if (Test-Path -LiteralPath $marker) {
                $previous = Get-Content -LiteralPath $marker -Raw | ConvertFrom-Json
                if ($previous.ProfilePath -ine $profilePath -or $previous.Sid -ne $profileSid -or
                    $previous.RepairId -ne $RepairId -or $previous.FullBrokerReset -ne [bool]$FullBrokerReset) {
                    throw 'Existing marker scope/reset mode differs. Choose a new RepairId for an intentional reset.'
                }
                $cleaned = $true
            }
            if (-not $cleaned) {
                $availability = Get-OfflineProfileStatus -ProfileSid $profileSid -ProfilePath $profilePath
                if ($availability.State -ne 'Ready') {
                    $summary.Deferred++
                    Write-Warning "Deferring profile $profileSid at $profilePath. $($availability.Reason). Cache cleanup and task creation are pending; retry at startup with this profile fully signed out."
                    $failed = $true
                    continue
                }
                if (-not $servicePrepared) {
                    $service = Get-Service -Name TokenBroker
                    Write-Host "TokenBroker status before cleanup: $($service.Status)"
                    if ($service.Status -ne 'Stopped') { $restoreService = $true }
                    Stop-Service -Name TokenBroker -Force
                    (Get-Service -Name TokenBroker).WaitForStatus('Stopped', [TimeSpan]::FromSeconds(30))
                    Write-Host 'TokenBroker confirmed Stopped.'
                    $servicePrepared = $true
                }
                Write-Host "Cleaning $profileSid at $profilePath"
                $localData = Join-Path $profilePath 'AppData\Local'
                Remove-BrokerCachePath -Path (Join-Path $localData 'Microsoft\TokenBroker\Cache')
                $brokerData = Join-Path $localData 'Packages\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy'
                if ($FullBrokerReset) {
                    Remove-BrokerCachePath -Path $brokerData
                } else {
                    Remove-BrokerCachePath -Path (Join-Path $brokerData 'AC\TokenBroker')
                }
                [ordered]@{
                    RepairId = $RepairId; Sid = $profileSid; ProfilePath = $profilePath
                    FullBrokerReset = [bool]$FullBrokerReset; CleanedUtc = [DateTime]::UtcNow.ToString('o')
                } | ConvertTo-Json | Set-Content -LiteralPath "$marker.tmp" -Encoding UTF8
                Move-Item -LiteralPath "$marker.tmp" -Destination $marker -Force
                $summary.Cleaned++
            } else {
                $summary.AlreadyClean++
                Write-Host "Already cleaned: $profileSid. Leaving caches intact."
            }

            # Exact SID for both trigger and interactive principal. No password
            # and no SYSTEM AppX registration. Success is tracked inside HKCU.
            $closeArgument = if ($CloseApplications) { ' -CloseApplications' } else { '' }
            $action = New-ScheduledTaskAction -Execute "$PSHOME\powershell.exe" -Argument (
                '-NoProfile -NonInteractive -WindowStyle Hidden -File "{0}" -Mode UserRepair -RepairId "{1}"{2}' -f $localScript, $RepairId, $closeArgument)
            $trigger = New-ScheduledTaskTrigger -AtLogOn -User $profileSid
            $taskPrincipal = New-ScheduledTaskPrincipal -UserId $profileSid -LogonType Interactive -RunLevel Limited
            $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
                -ExecutionTimeLimit (New-TimeSpan -Minutes 5) -MultipleInstances IgnoreNew
            Register-ScheduledTask -TaskName "Alumtran-TokenBroker-$profileSid" -Action $action `
                -Trigger $trigger -Principal $taskPrincipal -Settings $settings -Force | Out-Null
            $summary.Tasks++
            Write-Host "Logon registration task ready for $profileSid"
        } catch {
            $failed = $true
            $summary.Errors++
            Write-Warning "Profile repair failed for $($profile.SID): $($_ | Out-String)"
        }
    }
    Write-Host "Selected $($profiles.Count) profile(s). Close applications at user logon: $CloseApplications. No restart is forced."
} catch {
    $failed = $true
    $summary.Errors++
    Write-Warning ($_ | Out-String)
} finally {
    if ($restoreService) {
        try {
            Start-Service -Name TokenBroker
            (Get-Service -Name TokenBroker).WaitForStatus('Running', [TimeSpan]::FromSeconds(30))
            Write-Host 'TokenBroker restored and confirmed Running.'
        } catch {
            $failed = $true
            $summary.Errors++
            Write-Warning "TokenBroker restoration failed: $($_ | Out-String)"
        }
    }
    Write-Host ('Summary: newly processed={0}; previously cleaned={1}; tasks ready={2}; missing directories skipped={3}; profiles deferred={4}; errors={5}.' -f
        $summary.Cleaned, $summary.AlreadyClean, $summary.Tasks, $summary.Missing, $summary.Deferred, $summary.Errors)
    Write-Host 'Startup completion does not confirm user-phase registration or Alumtran authentication.'
    Stop-Transcript | Out-Null
}
if ($failed) { exit 1 }
exit 0
