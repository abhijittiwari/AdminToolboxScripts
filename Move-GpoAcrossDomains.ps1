#requires -Version 5.1
<#
.SYNOPSIS
Exports one GPO, or imports a previously exported GPO into another domain.
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
    [switch]$AllowOverwrite
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
    if (([bool]$GpoName) -eq ($GpoId -ne [guid]::Empty)) {
        throw 'Specify exactly one of -GpoName or -GpoId for Export.'
    }
    if ($TargetName -or $MigrationTable -or $AllowOverwrite) {
        throw 'TargetName, MigrationTable and AllowOverwrite are Import-only parameters.'
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
    if ($GpoId -ne [guid]::Empty) { $gpo = Get-GPO -Guid $GpoId @context }
    else { $gpo = Get-GPO -Name $GpoName @context }

    if (-not $PSCmdlet.ShouldProcess("$Domain / $($gpo.DisplayName)", "Export to $Path")) { return }
    $backupPath = Join-Path $Path 'Backup'
    $null = New-Item -ItemType Directory -Path $backupPath -Force
    $backup = Backup-GPO -Guid $gpo.Id -Path $backupPath @context
    Get-GPOReport -Guid $gpo.Id -ReportType Html -Path (Join-Path $Path 'Source.html') @context
    Get-GPOReport -Guid $gpo.Id -ReportType Xml -Path (Join-Path $Path 'Source.xml') @context
    [ordered]@{
        SchemaVersion = 1
        SourceDomain = $Domain
        SourceServer = $Server
        SourceHost = $env:COMPUTERNAME
        SourceName = $gpo.DisplayName
        SourceGpoId = $gpo.Id.ToString()
        BackupId = $backup.Id.ToString()
        ExportedUtc = [DateTime]::UtcNow.ToString('o')
    } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $Path 'Manifest.json') -Encoding UTF8
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
    [pscustomobject]@{ Mode = 'Export'; GpoName = $gpo.DisplayName; BackupId = $backup.Id; Path = $Path; PackagePath = $PackagePath }
    return
}

if ($GpoName -or $GpoId -ne [guid]::Empty) {
    throw 'GpoName and GpoId are Export-only parameters; use TargetName for Import.'
}
$manifestPath = Join-Path $Path 'Manifest.json'
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
        $expectedEntry = 'Backup/{' + ([guid]$manifest.BackupId).ToString() + '}/bkupInfo.xml'
        $foundBackup = @($archive.Entries | Where-Object { $_.FullName.Replace('\','/') -ieq $expectedEntry })
        if ($foundBackup.Count -ne 1) { throw 'The backup referenced by Manifest.json is missing or duplicated in the ZIP.' }
        $extractPrefix = $Path.TrimEnd('\','/') + [IO.Path]::DirectorySeparatorChar
        foreach ($zipEntry in $archive.Entries) {
            $entryPath = [IO.Path]::GetFullPath([IO.Path]::Combine($Path, $zipEntry.FullName))
            if (-not $entryPath.StartsWith($extractPrefix, [StringComparison]::OrdinalIgnoreCase)) {
                throw "ZIP entry points outside the extraction folder: $($zipEntry.FullName)"
            }
        }
    } finally { $archive.Dispose() }
} else {
    $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
}
if ($manifest.SchemaVersion -ne 1) { throw 'Unsupported export manifest version.' }
$backupId = [guid]$manifest.BackupId
$backupPath = Join-Path $Path 'Backup'
if (-not $PackagePath) {
    $backupInfo = Join-Path $backupPath ('{' + $backupId.ToString() + '}\bkupInfo.xml')
    if (-not (Test-Path -LiteralPath $backupInfo -PathType Leaf)) { throw 'The backup referenced by Manifest.json is missing.' }
}
if ([string]::IsNullOrWhiteSpace($TargetName)) { $TargetName = [string]$manifest.SourceName }
if ([string]::IsNullOrWhiteSpace($TargetName)) { throw 'Target GPO name is empty.' }
if ($MigrationTable) {
    $MigrationTable = (Resolve-Path -LiteralPath $MigrationTable).ProviderPath
    if (-not (Test-Path -LiteralPath $MigrationTable -PathType Leaf)) { throw 'MigrationTable must be a file.' }
} else {
    Write-Warning 'No migration table supplied. Review source-domain accounts, UNC paths and other domain-specific settings before linking the imported GPO.'
}

# Enumerate successfully before deciding that the target does not exist.
# Do not treat an arbitrary access/connectivity failure as a missing GPO.
$matches = @(Get-GPO -All @context | Where-Object { $_.DisplayName -eq $TargetName })
if ($matches.Count -gt 1) { throw "Multiple target GPOs named '$TargetName'; resolve the duplicate names first." }
$existing = $null
if ($matches.Count -eq 1) { $existing = $matches[0] }
if ($existing -and -not $AllowOverwrite) {
    throw "Target '$TargetName' already exists. Choose another name or explicitly use -AllowOverwrite."
}
$action = 'Create an unlinked GPO and import settings'
if ($existing) { $action = 'Back up target and overwrite its settings (existing links remain active)' }
if ($PackagePath) { $action = "Extract $PackagePath into $Path; $action" }
if (-not $PSCmdlet.ShouldProcess("$Domain / $TargetName", $action)) { return }

if ($PackagePath) {
    # All input comes from the copied ZIP. No connection to the source is made.
    [IO.Compression.ZipFile]::ExtractToDirectory($PackagePath, $Path)
}
$runName = 'Import-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '-' + [guid]::NewGuid().ToString('N')
$runPath = Join-Path $Path $runName
$null = New-Item -ItemType Directory -Path $runPath
if ($existing) {
    $rollbackPath = Join-Path $runPath 'BeforeImport'
    $null = New-Item -ItemType Directory -Path $rollbackPath
    $rollback = Backup-GPO -Guid $existing.Id -Path $rollbackPath @context
    [pscustomobject]@{ Domain = $Domain; Server = $Server; TargetId = $existing.Id.ToString(); BackupId = $rollback.Id.ToString() } |
        ConvertTo-Json | Set-Content -LiteralPath (Join-Path $runPath 'Rollback.json') -Encoding UTF8
    Write-Host "Pre-import backup retained at: $rollbackPath"
    $target = $existing
} else {
    # New-GPO refuses an existing name; import by GUID to pin the selected object.
    $target = New-GPO -Name $TargetName @context
}

$importParams = @{ BackupId = $backupId; Path = $backupPath; TargetGuid = $target.Id }
if ($MigrationTable) { $importParams.MigrationTable = $MigrationTable }
try {
    $null = Import-GPO @importParams @context
} catch {
    Write-Warning "Import failed. Target GUID: $($target.Id). It may contain partial settings; inspect it before retrying. Backups/reports retained at $runPath."
    throw
}
# Reports are a review aid, not proof of settings equivalence or client application.
try {
    Get-GPOReport -Guid $target.Id -ReportType Html -Path (Join-Path $runPath 'Destination.html') @context
    Get-GPOReport -Guid $target.Id -ReportType Xml -Path (Join-Path $runPath 'Destination.xml') @context
} catch {
    Write-Warning "Settings imported, but report generation failed: $($_.Exception.Message)"
}
Write-Host "Import complete. Review settings, filtering, delegation and WMI requirements before linking. Reports: $runPath"
[pscustomobject]@{
    Mode = 'Import'; Domain = $Domain; GpoName = $TargetName
    GpoId = $target.Id; OverwroteExisting = [bool]$existing; ReportsPath = $runPath
}
