# Move a GPO between domains

[Script](../Move-GpoAcrossDomains.ps1) · [Script reference](script-reference.md#move-gpoacrossdomainsps1)

`Move-GpoAcrossDomains.ps1` backs up one GPO and imports its settings into another domain. Despite its name, it does **not** delete the source GPO. Export and import are independent operations and can run on different Windows hosts, at different times, with different accounts.

## Requirements and credentials

- Use Windows PowerShell 5.1 with GPMC/RSAT Group Policy Management Tools installed on both hosts. Verify availability with `Get-Module -ListAvailable GroupPolicy`.
- Each host needs DNS resolution and AD/SYSVOL connectivity to its own domain and the DC supplied in `-Server`. Use a writable destination DC.
- The source account needs access to read and back up the selected GPO.
- The destination account needs permission to enumerate GPOs and create a GPO, or read/back up and edit the existing target when overwriting.
- The script uses the current Windows security context. It does not accept `-Credential` or store passwords. Sign in with the appropriate account on each host, or open a separate shell as that account.

On the source-domain host:

```cmd
runas /user:OLDDOMAIN\AdminUser "powershell.exe -NoProfile"
```

On the destination-domain host:

```cmd
runas /user:NEWDOMAIN\AdminUser "powershell.exe -NoProfile"
```

Use your actual NetBIOS domain and account names. Run the following commands in the resulting shells from the folder containing the script. Alternate-account shells do not automatically grant elevated privileges or GPO permissions.

When each stage runs on a management host in its respective domain, no trust between the source and destination domains and no WinRM connection between the hosts are required. The destination stage reads its copied package locally; it does not contact the source host or source domain. The source-domain names in the manifest are informational.

## Two-host ZIP workflow

### 1. Export on the source host

Replace all example domains, DCs and GPO names with your values.

```powershell
.\Move-GpoAcrossDomains.ps1 `
    -Mode Export `
    -Domain 'old.example.com' `
    -Server 'dc01.old.example.com' `
    -GpoName 'Workstation Policy' `
    -Path 'C:\GPOTransfer\Export1' `
    -PackagePath 'C:\GPOTransfer\Workstation.zip'
```

The working folder and ZIP must both be new. The ZIP must be outside the working folder. To select a GPO by GUID, replace `-GpoName` with `-GpoId '<source-GPO-GUID>'`; do not supply both.

### 2. Transfer the package

Copy **Workstation.zip and Move-GpoAcrossDomains.ps1** to the destination host using your approved transfer method. The script is not embedded in the ZIP. Keep the complete package; do not copy only its HTML/XML reports. Paths may differ between hosts. If using a migration table, transfer that reviewed file separately.

The ZIP is not encrypted. It contains policy data and potentially embedded policy scripts; restrict access and do not commit generated packages to this repository.

### 3. Preview on the destination host

```powershell
.\Move-GpoAcrossDomains.ps1 `
    -Mode Import `
    -Domain 'new.example.com' `
    -Server 'dc01.new.example.com' `
    -Path 'D:\GPOTransfer\Import1' `
    -PackagePath 'D:\GPOTransfer\Workstation.zip' `
    -TargetName 'Migrated - Workstation Policy' `
    -WhatIf
```

This reads the package metadata and checks the destination GPO name without extraction or GPO changes. It does not prove write permissions, fully validate all backup contents, or simulate policy application.

### 4. Import

Run the same command without `-WhatIf`. The extraction folder must not exist. A new GPO is created unlinked; omitting `-TargetName` uses the source display name.

Review the imported settings, security filtering, delegation, WMI requirements and user/computer configuration status before linking to a test OU. Configure the intended scope and verify resultant policy on test clients before wider deployment.

## Folder transfer alternative

Omit `-PackagePath` on export. Copy the entire export folder and the script to the destination host, then import from the copied folder:

```powershell
.\Move-GpoAcrossDomains.ps1 -Mode Import -Domain 'new.example.com' -Server 'dc01.new.example.com' -Path 'D:\GPOTransfer\CopiedExport' -TargetName 'Migrated - Workstation Policy'
```

With no `-PackagePath`, import expects an existing folder containing Manifest.json and Backup. This also lets you reuse a successfully extracted package without attempting to extract it again.

## Parameters

| Parameter | Mode | Meaning |
| --- | --- | --- |
| `-Mode` | Both, required | Export or Import. |
| `-Domain` | Both, required | DNS name of the domain used by the current stage. |
| `-Server` | Both, required | DC used by the current stage; use a writable destination DC. |
| `-Path` | Both, required | New export folder; new extraction folder when importing a ZIP; existing copied folder when importing without a ZIP. |
| `-PackagePath` | Both, optional | New ZIP on export, existing copied ZIP on import. |
| `-GpoName` | Export | Source display name. Specify this or GpoId, not both. |
| `-GpoId` | Export | Source GPO GUID, not a backup GUID. |
| `-TargetName` | Import | Destination display name; defaults to the source name in the manifest. |
| `-MigrationTable` | Import | Local path to a reviewed .migtable file. |
| `-AllowOverwrite` | Import | Permit replacing settings of an existing target after backing it up. |
| `-WhatIf` | Both | Read-only preflight and action description. |
| `-Confirm` | Both | Request confirmation before the stage performs writes. |

## Migration tables and scope

Add `-MigrationTable 'D:\GPOTransfer\DomainMap.migtable'` to the import command when you need to map supported source security principals and UNC paths to destination values. Prepare and review the table using the GPMC Migration Table Editor; destination accounts/groups and resources must be provisioned separately.

Without a table, the script warns and proceeds. A migration table is not a general replacement of source-domain strings. Review scripts, preferences, DNS names, application paths and other embedded references independently.

The script does not migrate:

- Domain, site or OU links, link order or enforcement.
- Source security filtering or delegation.
- WMI filter objects or their destination associations.
- Central Store ADMX/ADML files.
- External scripts, installation packages, shares, certificates or other referenced resources.

New targets receive destination defaults. On an existing target, its deployment scope can remain active; importing settings does not make it safe to treat that GPO as unlinked.

## Existing GPOs, failures and recovery

By default, a matching destination display name stops the import. Choose a new target name for a review copy. Use `-AllowOverwrite` only when you intend to replace the existing GPO's settings. The script backs it up first and saves a Rollback.json record.

**Overwriting a linked GPO can affect clients through normal policy refresh.** The operation is not transactional. An import failure can leave partial settings; the script does not delete a new GPO or automatically roll back an existing one.

For an overwritten target, inspect the retained backup and restore it deliberately in the destination domain if needed:

```powershell
# Use the actual import run folder printed by the script.
$runPath = 'D:\GPOTransfer\Import1\Import-<timestamp>-<unique-id>'
$rollback = Get-Content -LiteralPath (Join-Path $runPath 'Rollback.json') -Raw |
    ConvertFrom-Json

Restore-GPO -BackupId ([guid]$rollback.BackupId) -Path (Join-Path $runPath 'BeforeImport') -Domain $rollback.Domain -Server $rollback.Server
```

This is recovery of the destination GPO backed up before overwrite, not a cross-domain restore of the source backup. Review the scope and validate clients after recovery. A newly created target has no pre-import backup; inspect its GUID and settings in GPMC before deciding whether to retry or remove it.

## Output

| Location | Contents |
| --- | --- |
| Export folder / ZIP: Manifest.json | Source metadata, source GPO ID, exact backup ID, export timestamp. No credentials. |
| Export folder / ZIP: Backup | Native GPO backup used for import. |
| Export folder / ZIP: Source.html, Source.xml | Source policy reports for review. |
| Import folder: Import-&lt;timestamp&gt;-&lt;unique-id&gt; | Destination run folder. |
| Run folder: Destination.html, Destination.xml | Imported policy reports when report generation succeeds. |
| Run folder: BeforeImport and Rollback.json | Existing destination backup and recovery metadata when overwriting. |

An import can succeed even if report generation fails; check warnings. Compare source and destination reports in light of intended mappings. Reports alone do not establish settings equivalence or successful client application.

## Troubleshooting

| Symptom | Action |
| --- | --- |
| GroupPolicy module is missing | Install GPMC/RSAT Group Policy Management Tools and run Windows PowerShell 5.1. |
| Access denied or DC unavailable | Verify the current account, delegated GPO rights, domain/DC FQDN, DNS and AD/SYSVOL connectivity for this stage. |
| Export folder / ZIP already exists | Choose new output names. Existing packages are not overwritten. |
| Extraction folder already exists | Choose a new folder, or omit PackagePath to reuse a complete extracted export. |
| Target GPO already exists | Choose a new TargetName, or intentionally use AllowOverwrite. |
| Manifest or referenced backup missing | Recopy the complete original package. The HTML/XML reports alone cannot be imported. |
| ZIP creation or extraction failed | Preserve the original export, resolve the storage/path error and use a fresh folder/package. Do not use a partial ZIP or incomplete extraction. |
| Import failed after target creation | Inspect the printed target GUID; it can contain partial settings. Review before retrying. |
| Source-domain references remain | Review migration-table coverage and manually adjust unsupported embedded references before linking. |

## Validation status and references

The published script has been statically reviewed. It has not been executed against live AD in this authoring environment. Test export/import and policy application in a lab before production use.

- [Microsoft: Backup-GPO](https://learn.microsoft.com/en-us/powershell/module/grouppolicy/backup-gpo)
- [Microsoft: Import-GPO](https://learn.microsoft.com/en-us/powershell/module/grouppolicy/import-gpo)
- [Microsoft: Restore-GPO](https://learn.microsoft.com/en-us/powershell/module/grouppolicy/restore-gpo)
- [Microsoft: Copying and importing GPOs across domains](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/gpmc/copying-and-importing-gpos-across-domains)
