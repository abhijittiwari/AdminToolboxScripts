#requires -Version 5.1
<#
.SYNOPSIS
Exports one GPO or the GPOs linked to an OU, and imports them into another domain.
.DESCRIPTION
Run in Windows PowerShell 5.1 with the GroupPolicy module (GPMC/RSAT).
Uses the current Windows account: run Export in the source domain and Import
in the destination domain, using an authorized account for each domain.
No password is stored and no trust or PowerShell remoting is required when
each phase runs locally on a management computer in its respective domain.

The two modes run independently, including on different hosts. Export only
contacts the source domain. Import only contacts the destination domain;
SourceDomain and SourceServer in the manifest are informational only.

For portable transfer, supply PackagePath (a ZIP filename) in both modes.
Export: Path is a NEW working folder; PackagePath is a NEW ZIP outside it.
Import: PackagePath is the copied ZIP; Path is a NEW extraction folder.
Copy the ZIP and this script to the destination host; no source-host access
is needed during import. Paths may be completely different on the two hosts.
Without PackagePath, copy the ENTIRE export folder and import from that folder.
Packages contain a native GPO backup, manifest and HTML/XML reports. Protect
them as policy data. Credentials are never included in the package.

By default Import creates a new, unlinked GPO. Existing names are refused.
AllowOverwrite explicitly permits replacement of settings after a backup.
Existing links/filtering remain in effect, so overwrite can affect live clients.
This is not a transactional operation: a failed import may leave partial settings.
The script retains backups and reports and never automatically deletes a GPO.

GPO links, security filtering/delegation, WMI filters and Central Store ADMX/ADML
files are not migrated by this script. Review these separately. A .migtable
can map supported security-principal/UNC references; it is not a general
search/replace for domain strings, DNS names, certificates, or script contents.
External files referenced by policy must be migrated separately.

Use SourceOU with the OU distinguished name to export its direct links, including
disabled links. IncludeChildOUs adds all descendant OUs and requires the
ActiveDirectory PowerShell module on the source host. IncludeInherited adds
applicable parent OU and domain links for each selected OU.
Link metadata is recorded for review. It is never replayed on the destination.
Use TargetNamePrefix to give the imported GPOs a consistent staging prefix.

.EXAMPLE
.\Move-GpoAcrossDomains.ps1 -Mode Export -Domain old.example.com -Server dc01.old.example.com -SourceOU 'OU=HQ,DC=old,DC=example,DC=com' -IncludeChildOUs -Path C:\GPOTransfer\HQExport -PackagePath C:\GPOTransfer\HQ.zip -WhatIf
.EXAMPLE
.\Move-GpoAcrossDomains.ps1 -Mode Export -Domain old.example.com -Server dc01.old.example.com -SourceOU 'OU=Workstations,DC=old,DC=example,DC=com' -Path C:\GPOTransfer\OUExport -PackagePath C:\GPOTransfer\Workstations.zip
.EXAMPLE
.\Move-GpoAcrossDomains.ps1 -Mode Import -Domain new.example.com -Server dc01.new.example.com -Path D:\GPOTransfer\OUImport -PackagePath D:\GPOTransfer\Workstations.zip -TargetNamePrefix 'Migrated - ' -WhatIf
.EXAMPLE
.\Move-GpoAcrossDomains.ps1 -Mode Export -Domain old.example.com -Server dc01.old.example.com -GpoName 'Workstation Policy' -Path C:\GPOTransfer\Export1 -PackagePath C:\GPOTransfer\Workstation.zip
.EXAMPLE
.\Move-GpoAcrossDomains.ps1 -Mode Import -Domain new.example.com -Server dc01.new.example.com -Path D:\GPOTransfer\Import1 -PackagePath D:\GPOTransfer\Workstation.zip -TargetName 'Migrated - Workstation Policy'
.EXAMPLE
.\Move-GpoAcrossDomains.ps1 -Mode Export -Domain old.example.com -Server dc01.old.example.com -GpoName 'Workstation Policy' -Path C:\GPOTransfer\Workstation
.EXAMPLE
.\Move-GpoAcrossDomains.ps1 -Mode Import -Domain new.example.com -Server dc01.new.example.com -Path C:\GPOTransfer\Workstation -TargetName 'Migrated - Workstation Policy'
.EXAMPLE
.\Move-GpoAcrossDomains.ps1 -Mode Import -Domain new.example.com -Server dc01.new.example.com -Path C:\GPOTransfer\Workstation -TargetName 'Workstation Policy' -MigrationTable C:\GPOTransfer\DomainMap.migtable -AllowOverwrite
.NOTES
WhatIf performs read-only checks but does not prove write permissions or simulate
the full import. Import success and report generation are not an RSoP test.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory)][ValidateSet('Export','Import')][string]$Mode,
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Domain,
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Server,
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Path,
    [string]$PackagePath,
    [string]$GpoName,
    [guid]$GpoId = [guid]::Empty,
    [string]$TargetName,
    [string]$MigrationTable,
    [switch]$AllowOverwrite,
    [string]$SourceOU,
    [switch]$IncludeInherited,
    [string]$TargetNamePrefix,
    [switch]$IncludeChildOUs
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module GroupPolicy -ErrorAction Stop
$context = @{ Domain = $Domain; Server = $Server; ErrorAction = 'Stop' }
$Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
if ($PackagePath) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $PackagePath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($PackagePath)
    if ([IO.Path]::GetExtension($PackagePath) -ine '.zip') { throw 'PackagePath must have a .zip extension.' }
}

if ($Mode -eq 'Export') {
    if ([string]::IsNullOrWhiteSpace($SourceOU)) { $SourceOU = '' }
    $selectorCount = [int](-not [string]::IsNullOrWhiteSpace($GpoName)) +
                     [int]($GpoId -ne [guid]::Empty) +
                     [int][bool]$SourceOU
    if ($selectorCount -ne 1) {
        throw 'Specify exactly one of -GpoName, -GpoId or -SourceOU for Export.'
    }
    if (($IncludeInherited -or $IncludeChildOUs) -and -not $SourceOU) {
        throw '-IncludeInherited and -IncludeChildOUs require -SourceOU.'
    }
    if ($TargetName -or $TargetNamePrefix -or $MigrationTable -or $AllowOverwrite) {
        throw 'TargetName, TargetNamePrefix, MigrationTable and AllowOverwrite are Import-only parameters.'
    }
    if (Test-Path -LiteralPath $Path) {
        throw "Export folder already exists. Choose a new folder: $Path"
    }
    if ($PackagePath) {
        if (Test-Path -LiteralPath $PackagePath) { throw "Package already exists: $PackagePath" }
        $exportPrefix = $Path.TrimEnd('\','/') + [IO.Path]::DirectorySeparatorChar
        if ($PackagePath.Equals($Path, [StringComparison]::OrdinalIgnoreCase) -or
            $PackagePath.StartsWith($exportPrefix, [StringComparison]::OrdinalIgnoreCase)) {
            throw 'PackagePath must be outside the export folder.'
        }
    }
    $sourceScope = $null
    $linkRows = [System.Collections.Generic.List[object]]::new()
    if ($SourceOU) {
        $SourceOU = $SourceOU.Trim()
        if ($SourceOU -notmatch '^OU=') { throw '-SourceOU must be an OU distinguished name.' }
        $sourceScope = Get-GPInheritance -Target $SourceOU @context
        $scopes = @(
            $sourceScope
            if ($IncludeChildOUs) {
                Import-Module ActiveDirectory -ErrorAction Stop
                $childOUs = Get-ADOrganizationalUnit -Filter * -SearchBase $sourceScope.Path -SearchScope Subtree -Server $Server -ErrorAction Stop |
                    Sort-Object -Property DistinguishedName
                foreach ($childOU in $childOUs) {
                    if ($childOU.DistinguishedName -ine $sourceScope.Path) {
                        Get-GPInheritance -Target $childOU.DistinguishedName @context
                    }
                }
            }
        )
        $seenGpos = @{}
        $gpos = @(foreach ($scope in $scopes) {
            $links = @($scope.GpoLinks | Where-Object { $null -ne $_ })
            $directCount = $links.Count
            if ($IncludeInherited) {
                $links += @($scope.InheritedGpoLinks | Where-Object { $null -ne $_ })
            }

            # An inherited-link result can repeat a direct link for this OU.
            $seenLinks = @{}
            for ($i = 0; $i -lt $links.Count; $i++) {
                $link = $links[$i]
                if ($link.GpoDomainName -ine $Domain) {
                    throw "GPO '$($link.DisplayName)' is stored in '$($link.GpoDomainName)'. Export it separately from its owning domain."
                }
                $linkId = [guid]$link.GpoId
                $linkKey = "$linkId|$($link.Target)"
                if ($seenLinks.ContainsKey($linkKey)) { continue }
                $seenLinks[$linkKey] = $true
                $linkRows.Add([pscustomobject]@{
                    SourceGpoId = $linkId.ToString()
                    DisplayName = $link.DisplayName
                    LinkedAt = $link.Target
                    Inherited = $i -ge $directCount
                    Order = $link.Order
                    Enabled = $link.Enabled
                    Enforced = $link.Enforced
                    ScopeOU = $scope.Path
                })
                if (-not $seenGpos.ContainsKey($linkId)) {
                    $seenGpos[$linkId] = $true
                    Get-GPO -Guid $linkId @context
                }
            }
        })
        if ($gpos.Count -eq 0) { throw "No GPO links matched the requested scope: $SourceOU" }
        foreach ($linkRow in $linkRows) {
            Write-Host ("{0} | OU {1} | Linked at {2} | Order {3} | Enabled {4} | Enforced {5}" -f
                $linkRow.DisplayName, $linkRow.ScopeOU, $linkRow.LinkedAt, $linkRow.Order, $linkRow.Enabled, $linkRow.Enforced)
        }
        $exportTarget = "$Domain / $SourceOU ($($gpos.Count) GPOs across $($scopes.Count) OU(s))"
    } else {
        if ($GpoId -ne [guid]::Empty) { $gpos = @(Get-GPO -Guid $GpoId @context) }
        else { $gpos = @(Get-GPO -Name $GpoName @context) }
        $exportTarget = "$Domain / $($gpos[0].DisplayName)"
    }

    if (-not $PSCmdlet.ShouldProcess($exportTarget, "Export to $Path")) { return }
    $backupPath = Join-Path $Path 'Backup'
    $null = New-Item -ItemType Directory -Path $backupPath -Force
    $manifestData = [ordered]@{
        SchemaVersion = 1
        SourceDomain = $Domain
        SourceServer = $Server
        SourceHost = $env:COMPUTERNAME
        ExportedUtc = [DateTime]::UtcNow.ToString('o')
    }
    $exported = @(foreach ($gpo in $gpos) {
        $backup = Backup-GPO -Guid $gpo.Id -Path $backupPath @context
        $reportPath = $Path
        if ($SourceOU) {
            $reportPath = Join-Path (Join-Path $Path 'SourceReports') $gpo.Id.ToString()
            $null = New-Item -ItemType Directory -Path $reportPath -Force
        }
        $null = Get-GPOReport -Guid $gpo.Id -ReportType Html -Path (Join-Path $reportPath 'Source.html') @context
        $null = Get-GPOReport -Guid $gpo.Id -ReportType Xml -Path (Join-Path $reportPath 'Source.xml') @context
        [pscustomobject]@{
            SourceName = $gpo.DisplayName
            SourceGpoId = $gpo.Id.ToString()
            BackupId = $backup.Id.ToString()
        }
    })
    if ($SourceOU) {
        $manifestData.SchemaVersion = 2
        $manifestData['SourceOU'] = $sourceScope.Path
        $manifestData['IncludeInherited'] = [bool]$IncludeInherited
        $manifestData['GpoInheritanceBlocked'] = [string]$sourceScope.GpoInheritanceBlocked
        if ($IncludeChildOUs) {
            $manifestData['IncludeChildOUs'] = $true
            $manifestData['SourceOUs'] = @(foreach ($scope in $scopes) {
                [pscustomobject]@{
                    DistinguishedName = $scope.Path
                    GpoInheritanceBlocked = [string]$scope.GpoInheritanceBlocked
                }
            })
        }
        $manifestData['Gpos'] = $exported
        $linkRows | Export-Csv -LiteralPath (Join-Path $Path 'SourceLinks.csv') -NoTypeInformation -Encoding UTF8
    } else {
        $manifestData['SourceName'] = $exported[0].SourceName
        $manifestData['SourceGpoId'] = $exported[0].SourceGpoId
        $manifestData['BackupId'] = $exported[0].BackupId
    }
    $manifestData | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $Path 'Manifest.json') -Encoding UTF8
    if ($PackagePath) {
        $null = New-Item -ItemType Directory -Path (Split-Path -Parent $PackagePath) -Force
        try {
            [IO.Compression.ZipFile]::CreateFromDirectory($Path, $PackagePath)
        } catch {
            Write-Warning "Export folder is retained at $Path. ZIP creation failed; any partial ZIP must not be used."
            throw
        }
        Write-Host "Export complete. Copy the ZIP and this script to the destination host: $PackagePath"
    } else {
        Write-Host "Export complete. Copy this entire folder and this script to the destination host: $Path"
    }
    foreach ($item in $exported) {
        [pscustomobject]@{ Mode = 'Export'; GpoName = $item.SourceName; BackupId = [guid]$item.BackupId; Path = $Path; PackagePath = $PackagePath }
    }
    return
}

if ($GpoName -or $GpoId -ne [guid]::Empty -or $SourceOU -or $IncludeInherited -or $IncludeChildOUs) {
    throw 'GpoName, GpoId, SourceOU, IncludeInherited and IncludeChildOUs are Export-only parameters.'
}
if ($TargetName -and $TargetNamePrefix) {
    throw 'Specify either -TargetName or -TargetNamePrefix.'
}
$manifestPath = Join-Path $Path 'Manifest.json'
$zipEntryPaths = @{}
if ($PackagePath) {
    if (-not (Test-Path -LiteralPath $PackagePath -PathType Leaf)) { throw "ZIP not found: $PackagePath" }
    if (Test-Path -LiteralPath $Path) { throw "Extraction folder already exists. Choose a new -Path, or omit -PackagePath to reuse an extracted export: $Path" }
    # Read without extraction so -WhatIf leaves the filesystem unchanged.
    $archive = [IO.Compression.ZipFile]::OpenRead($PackagePath)
    try {
        $entry = $archive.GetEntry('Manifest.json')
        if (-not $entry) { throw 'ZIP must contain Manifest.json at its root.' }
        $reader = [IO.StreamReader]::new($entry.Open())
        try { $manifest = $reader.ReadToEnd() | ConvertFrom-Json }
        finally { $reader.Dispose() }
        $extractPrefix = $Path.TrimEnd('\','/') + [IO.Path]::DirectorySeparatorChar
        foreach ($zipEntry in $archive.Entries) {
            $entryPath = [IO.Path]::GetFullPath([IO.Path]::Combine($Path, $zipEntry.FullName))
            if (-not $entryPath.StartsWith($extractPrefix, [StringComparison]::OrdinalIgnoreCase)) {
                throw "ZIP entry points outside the extraction folder: $($zipEntry.FullName)"
            }
            $normalizedPath = $zipEntry.FullName.Replace('\','/')
            if ($zipEntryPaths.ContainsKey($normalizedPath)) { throw "Duplicate ZIP entry: $normalizedPath" }
            $zipEntryPaths[$normalizedPath] = $true
        }
    } finally { $archive.Dispose() }
} else {
    $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
}
$isOuPackage = $manifest.SchemaVersion -eq 2
switch ($manifest.SchemaVersion) {
    1 { $exported = @($manifest) }
    2 { $exported = @($manifest.Gpos | Where-Object { $null -ne $_ }) }
    default { throw 'Unsupported export manifest version.' }
}
if ($exported.Count -eq 0) { throw 'The manifest does not contain any GPOs.' }
if ($TargetName -and $exported.Count -gt 1) { throw '-TargetName applies to one GPO. Use -TargetNamePrefix for a batch.' }

$backupPath = Join-Path $Path 'Backup'
$seenBackupIds = @{}
$seenSourceIds = @{}
$seenTargetNames = @{}
$plan = @(foreach ($item in $exported) {
    $backupId = [guid]$item.BackupId
    $sourceId = [guid]$item.SourceGpoId
    if ($backupId -eq [guid]::Empty -or $sourceId -eq [guid]::Empty -or
        $seenBackupIds.ContainsKey($backupId) -or $seenSourceIds.ContainsKey($sourceId)) {
        throw 'The manifest contains an empty or duplicate GPO or backup ID.'
    }
    $seenBackupIds[$backupId] = $true
    $seenSourceIds[$sourceId] = $true
    $backupFolderName = $backupId.ToString('B')
    if ($PackagePath) {
        if (-not $zipEntryPaths.ContainsKey("Backup/$backupFolderName/bkupInfo.xml")) {
            throw "The package is missing backup $backupId."
        }
    } else {
        $backupInfo = Join-Path (Join-Path $backupPath $backupFolderName) 'bkupInfo.xml'
        if (-not (Test-Path -LiteralPath $backupInfo -PathType Leaf)) { throw "The export is missing backup $backupId." }
    }
    $name = $TargetNamePrefix + [string]$item.SourceName
    if (-not [string]::IsNullOrWhiteSpace($TargetName)) { $name = $TargetName }
    if ([string]::IsNullOrWhiteSpace($name)) { throw 'Target GPO name is empty.' }
    if ($seenTargetNames.ContainsKey($name)) { throw "The package maps multiple GPOs to '$name'." }
    $seenTargetNames[$name] = $true
    [pscustomobject]@{ BackupId = $backupId; SourceId = $sourceId; TargetName = $name; Existing = $null; ReportsPath = $null }
})
if ($MigrationTable) {
    $MigrationTable = (Resolve-Path -LiteralPath $MigrationTable).ProviderPath
    if (-not (Test-Path -LiteralPath $MigrationTable -PathType Leaf)) { throw 'MigrationTable must be a file.' }
} else {
    Write-Warning 'No migration table supplied. Review source-domain accounts, UNC paths and other domain-specific settings before linking the imported GPO.'
}

# Enumerate successfully before deciding that the target does not exist.
# Do not treat an arbitrary access/connectivity failure as a missing GPO.
$destinationByName = @{}
foreach ($destinationGpo in (Get-GPO -All @context)) {
    $displayName = $destinationGpo.DisplayName
    if (-not $destinationByName.ContainsKey($displayName)) {
        $destinationByName[$displayName] = @()
    }
    $destinationByName[$displayName] += $destinationGpo
}
$conflicts = @(foreach ($entry in $plan) {
    if (-not $destinationByName.ContainsKey($entry.TargetName)) { continue }
    $matchedGpos = $destinationByName[$entry.TargetName]
    if ($matchedGpos.Count -gt 1) { throw "Multiple target GPOs named '$($entry.TargetName)'; resolve the duplicate names first." }
    if ($matchedGpos.Count -eq 1) {
        $entry.Existing = $matchedGpos[0]
        if (-not $AllowOverwrite) { $entry.TargetName }
    }
})
if ($conflicts.Count -gt 0) {
    throw "Target GPOs already exist: $($conflicts -join ', '). Choose different names or explicitly use -AllowOverwrite."
}
foreach ($entry in $plan) {
    if ($entry.Existing) { $operation = 'Back up and overwrite' }
    else { $operation = 'Create unlinked and import' }
    Write-Host ("{0} | {1} | Source GPO {2}" -f $operation, $entry.TargetName, $entry.SourceId)
}
$action = 'Import GPO settings'
if ($AllowOverwrite) { $action += ' with overwrite enabled (existing links remain active)' }
if ($PackagePath) { $action = "Extract $PackagePath into $Path; $action" }
if (-not $PSCmdlet.ShouldProcess("$Domain ($($plan.Count) GPOs)", $action)) { return }

if ($PackagePath) {
    # All input comes from the copied ZIP. No connection to the source is made.
    [IO.Compression.ZipFile]::ExtractToDirectory($PackagePath, $Path)
}
$runName = 'Import-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '-' + [guid]::NewGuid().ToString('N')
$runPath = Join-Path $Path $runName
$null = New-Item -ItemType Directory -Path $runPath
Write-Host "Import run folder: $runPath"

# Back up every overwrite target before changing any destination GPO.
foreach ($entry in $plan) {
    $entry.ReportsPath = $runPath
    if ($isOuPackage) {
        $entry.ReportsPath = Join-Path $runPath $entry.SourceId.ToString()
        $null = New-Item -ItemType Directory -Path $entry.ReportsPath
    }
    if ($entry.Existing) {
        $rollbackPath = Join-Path $entry.ReportsPath 'BeforeImport'
        $null = New-Item -ItemType Directory -Path $rollbackPath
        $rollback = Backup-GPO -Guid $entry.Existing.Id -Path $rollbackPath @context
        [pscustomobject]@{ Domain = $Domain; Server = $Server; TargetId = $entry.Existing.Id.ToString(); BackupId = $rollback.Id.ToString() } |
            ConvertTo-Json | Set-Content -LiteralPath (Join-Path $entry.ReportsPath 'Rollback.json') -Encoding UTF8
        Write-Host "Pre-import backup for '$($entry.TargetName)' retained at: $rollbackPath"
    }
}

$completed = 0
foreach ($entry in $plan) {
    $target = $entry.Existing
    try {
        if (-not $target) {
            # New-GPO refuses an existing name. Import by GUID to pin the selected object.
            $target = New-GPO -Name $entry.TargetName @context
        }
        $importParams = @{ BackupId = $entry.BackupId; Path = $backupPath; TargetGuid = $target.Id }
        if ($MigrationTable) { $importParams.MigrationTable = $MigrationTable }
        $null = Import-GPO @importParams @context
    } catch {
        Write-Warning "Stopped at '$($entry.TargetName)' after $completed completed import(s). Backups and reports are retained at $runPath."
        if ($target) { Write-Warning "Inspect target $($target.Id) for partial settings before retrying." }
        throw
    }
    $completed++
    try {
        $null = Get-GPOReport -Guid $target.Id -ReportType Html -Path (Join-Path $entry.ReportsPath 'Destination.html') @context
        $null = Get-GPOReport -Guid $target.Id -ReportType Xml -Path (Join-Path $entry.ReportsPath 'Destination.xml') @context
    } catch {
        Write-Warning "Settings imported for '$($entry.TargetName)', but report generation failed: $($_.Exception.Message)"
    }
    [pscustomobject]@{
        Mode = 'Import'; Domain = $Domain; GpoName = $entry.TargetName
        GpoId = $target.Id; OverwroteExisting = [bool]$entry.Existing; ReportsPath = $entry.ReportsPath
        SourceGpoId = $entry.SourceId
    }
}
Write-Host "Imported $completed GPO(s). Review settings, filtering, delegation and WMI requirements before linking. Reports: $runPath"
