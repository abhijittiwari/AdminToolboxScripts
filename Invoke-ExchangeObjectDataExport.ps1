<#
.SYNOPSIS
    Runs the full modular Exchange object export: inventory, memberships, permissions, workbook.

.DESCRIPTION
    Orchestrates the five modular export scripts in dependency order:
      1. Export-ExchangeObjectInventory.ps1        (Exchange Online -> Objects.csv, ProxyAddresses.csv)
      2. Export-ExchangeGroupMemberships.ps1       (Microsoft Graph -> membership CSVs)
      3. Export-ExchangeSendAsPermissions.ps1      (Exchange Online -> SendAs.csv)
      4. Export-ExchangeFullAccessPermissions.ps1  (Exchange Online -> FullAccess.csv)
      5. Export-ExchangeObjectDataWorkbook.ps1     (offline -> ExchangeObjectData.xlsx)

    The job scripts must sit in the same folder as this script. If the inventory job fails,
    the run aborts; if a later job fails, the remaining jobs still run, the workbook is built
    from whatever CSVs exist, and the script exits with code 1. Sessions opened by the jobs
    stay open, so the whole run authenticates to Exchange Online and Microsoft Graph once.

.PARAMETER All
    Export all recipients of the selected -RecipientType.

.PARAMETER InputCsv
    Path to a CSV listing the recipients to export, passed through to the inventory
    job. The CSV needs one identity column (named by -IdentityColumn) whose values
    Exchange can resolve: UPN, primary SMTP address, alias, or GUID. Other columns
    are ignored.

.PARAMETER IdentityColumn
    Name of the identity column in -InputCsv. Defaults to Identity.

.PARAMETER Identity
    A single recipient identity to export.

.PARAMETER RecipientType
    Recipient scope: Mailbox, Group, MailUser, MailContact, or All. Defaults to Mailbox.

.PARAMETER OutputFolder
    Destination folder shared by all jobs. Defaults to ./ExchangeObjectExport_<timestamp>.

.PARAMETER Force
    Overwrite an existing workbook when building ExchangeObjectData.xlsx.

.NOTES
    Requires: ExchangeOnlineManagement, Microsoft.Graph.Authentication, ImportExcel
    This script is read-only. It does not modify Exchange or Entra objects.

.EXAMPLE
    ./Invoke-ExchangeObjectDataExport.ps1 -All

.EXAMPLE
    ./Invoke-ExchangeObjectDataExport.ps1 -InputCsv ./objects.csv -RecipientType Group -OutputFolder ./export

.EXAMPLE
    ./Invoke-ExchangeObjectDataExport.ps1 -Identity user@contoso.com -Force
#>

[CmdletBinding(DefaultParameterSetName = 'All')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'All')]
    [switch]$All,

    [Parameter(Mandatory = $true, ParameterSetName = 'Csv')]
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
    [string]$InputCsv,

    [Parameter(Mandatory = $true, ParameterSetName = 'Identity')]
    [string]$Identity,

    [Parameter(Mandatory = $false, ParameterSetName = 'Csv')]
    [string]$IdentityColumn = 'Identity',

    [Parameter(Mandatory = $false)]
    [ValidateSet('Mailbox', 'Group', 'MailUser', 'MailContact', 'All')]
    [string]$RecipientType = 'Mailbox',

    [Parameter(Mandatory = $false)]
    [string]$OutputFolder = ".\ExchangeObjectExport_$(Get-Date -Format 'yyyyMMdd_HHmmss')",

    [Parameter(Mandatory = $false)]
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

function Invoke-ExportJob {
    param(
        [Parameter(Mandatory = $true)] [string]$Name,
        [Parameter(Mandatory = $true)] [scriptblock]$Action
    )

    Write-Host ''
    Write-Host "=== $Name ===" -ForegroundColor Cyan
    try {
        & $Action
        return $true
    }
    catch {
        Write-Warning "Job '$Name' failed: $($_.Exception.Message)"
        return $false
    }
}

$jobScripts = [ordered]@{
    Inventory        = Join-Path $PSScriptRoot 'Export-ExchangeObjectInventory.ps1'
    GroupMemberships = Join-Path $PSScriptRoot 'Export-ExchangeGroupMemberships.ps1'
    SendAs           = Join-Path $PSScriptRoot 'Export-ExchangeSendAsPermissions.ps1'
    FullAccess       = Join-Path $PSScriptRoot 'Export-ExchangeFullAccessPermissions.ps1'
    Workbook         = Join-Path $PSScriptRoot 'Export-ExchangeObjectDataWorkbook.ps1'
}

foreach ($jobScript in $jobScripts.Values) {
    if (-not (Test-Path -Path $jobScript -PathType Leaf)) {
        throw "Job script '$jobScript' not found. All five export scripts must sit in the same folder as this script."
    }
}

$resolvedOutputFolder = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputFolder)
$objectsCsvPath = Join-Path $resolvedOutputFolder 'Objects.csv'
$failedJobs = [System.Collections.Generic.List[string]]::new()

$inventoryArgs = @{
    RecipientType = $RecipientType
    OutputFolder  = $resolvedOutputFolder
}
switch ($PSCmdlet.ParameterSetName) {
    'All' { $inventoryArgs.All = $true }
    'Csv' { $inventoryArgs.InputCsv = $InputCsv; $inventoryArgs.IdentityColumn = $IdentityColumn }
    'Identity' { $inventoryArgs.Identity = $Identity }
}

# 1. Inventory - everything downstream depends on Objects.csv, so a failure here aborts.
if (-not (Invoke-ExportJob -Name 'Recipient inventory (Export-ExchangeObjectInventory.ps1)' -Action { & $jobScripts.Inventory @inventoryArgs })) {
    throw 'Recipient inventory failed; aborting the export run.'
}

$objectRowCount = @(Import-Csv -Path $objectsCsvPath).Count

if ($objectRowCount -eq 0) {
    Write-Warning 'The inventory selected no objects. Skipping the membership and permission jobs.'
}
else {
    # 2. Group memberships (Microsoft Graph).
    if (-not (Invoke-ExportJob -Name 'Group memberships (Export-ExchangeGroupMemberships.ps1)' -Action { & $jobScripts.GroupMemberships -ObjectsCsv $objectsCsvPath })) {
        $failedJobs.Add('GroupMemberships') | Out-Null
    }

    # 3. SendAs permissions (Exchange Online).
    if (-not (Invoke-ExportJob -Name 'SendAs permissions (Export-ExchangeSendAsPermissions.ps1)' -Action { & $jobScripts.SendAs -ObjectsCsv $objectsCsvPath })) {
        $failedJobs.Add('SendAs') | Out-Null
    }

    # 4. FullAccess permissions (Exchange Online).
    if (-not (Invoke-ExportJob -Name 'FullAccess permissions (Export-ExchangeFullAccessPermissions.ps1)' -Action { & $jobScripts.FullAccess -ObjectsCsv $objectsCsvPath })) {
        $failedJobs.Add('FullAccess') | Out-Null
    }
}

# 5. Workbook - built from whatever CSVs the jobs produced.
$workbookArgs = @{ InputFolder = $resolvedOutputFolder }
if ($Force) { $workbookArgs.Force = $true }
if (-not (Invoke-ExportJob -Name 'Excel workbook (Export-ExchangeObjectDataWorkbook.ps1)' -Action { & $jobScripts.Workbook @workbookArgs })) {
    $failedJobs.Add('Workbook') | Out-Null
}

Write-Host ''
Write-Host "=== Export run complete ===" -ForegroundColor Cyan
Write-Host "Output folder: $resolvedOutputFolder" -ForegroundColor Green
Write-Host "Objects exported: $objectRowCount" -ForegroundColor Green
if ($failedJobs.Count -gt 0) {
    Write-Warning "Jobs with failures: $($failedJobs -join ', '). Check the Errors-*.csv files in the output folder."
    exit 1
}
Write-Host 'All jobs completed successfully.' -ForegroundColor Green
