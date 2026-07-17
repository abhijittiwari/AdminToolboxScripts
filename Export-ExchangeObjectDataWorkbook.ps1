<#
.SYNOPSIS
    Collates Exchange object export CSVs into a single Excel workbook.

.DESCRIPTION
    Final job of the modular Exchange object export. Reads the CSVs written by the export job
    scripts (Export-ExchangeObjectInventory.ps1, Export-ExchangeGroupMemberships.ps1,
    Export-ExchangeSendAsPermissions.ps1, Export-ExchangeFullAccessPermissions.ps1) - or by
    the all-in-one Export-ExchangeObjectData.ps1, which uses the same file names - and writes
    one .xlsx with a Summary sheet, a Consolidated per-object sheet, and one sheet per
    non-empty dataset. All Errors*.csv files in the folder are merged into a single Errors
    sheet; the Stage column identifies the source job.

    Fully offline: needs no Exchange Online or Microsoft Graph connection.

.PARAMETER InputFolder
    Folder containing the export CSVs. Objects.csv is mandatory; ProxyAddresses.csv,
    FullAccess.csv, SendAs.csv, ExchangeGroupMemberships.csv, EntraGroupMemberships.csv,
    and Errors*.csv are included when present. The CSVs are consumed with the column
    layout the export job scripts write; no extra columns are required.

.PARAMETER OutputPath
    Path of the .xlsx to write. Defaults to ExchangeObjectData.xlsx inside -InputFolder.

.PARAMETER Force
    Overwrite an existing workbook at -OutputPath.

.NOTES
    Requires: ImportExcel (Install-Module ImportExcel -Scope CurrentUser). Excel itself is
    not required; the workbook is generated cross-platform.

.EXAMPLE
    ./Export-ExchangeObjectDataWorkbook.ps1 -InputFolder ./ExchangeObjectInventory_20260714_120000

.EXAMPLE
    ./Export-ExchangeObjectDataWorkbook.ps1 -InputFolder ./ExchangeObjectData_20260714_120000 -OutputPath ./report.xlsx -Force
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -Path $_ -PathType Container })]
    [string]$InputFolder,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

if (-not (Get-Module -ListAvailable ImportExcel)) {
    throw 'The ImportExcel module is required to build the workbook. Install it with: Install-Module ImportExcel -Scope CurrentUser'
}
Import-Module ImportExcel -ErrorAction Stop

function Import-DataSet {
    param(
        [Parameter(Mandatory = $true)] [string]$Folder,
        [Parameter(Mandatory = $true)] [string]$FileName
    )

    $path = Join-Path $Folder $FileName
    if (-not (Test-Path -Path $path -PathType Leaf)) {
        Write-Warning "Optional dataset '$FileName' not found in '$Folder'; its sheet will be skipped."
        return @()
    }
    return @(Import-Csv -Path $path)
}

function Join-Values {
    param([Parameter(Mandatory = $false)] $Values)

    return (@($Values) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Sort-Object -Unique) -join '; '
}

function Group-RowsByIdentity {
    param([Parameter(Mandatory = $false)] [AllowEmptyCollection()] [object[]]$Rows)

    $index = @{}
    foreach ($row in @($Rows)) {
        if ($null -eq $row) { continue }
        $key = [string]$row.Identity
        if (-not $index.ContainsKey($key)) { $index[$key] = [System.Collections.Generic.List[object]]::new() }
        $index[$key].Add($row) | Out-Null
    }
    return $index
}

function New-ConsolidatedRows {
    param(
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] [object[]]$Objects,
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] [object[]]$ProxyRows,
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] [object[]]$FullAccessRows,
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] [object[]]$SendAsRows,
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] [object[]]$ExchangeGroupRows,
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] [object[]]$EntraGroupRows,
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] [object[]]$ErrorRows
    )

    $proxyIndex = Group-RowsByIdentity -Rows $ProxyRows
    $fullAccessIndex = Group-RowsByIdentity -Rows $FullAccessRows
    $sendAsIndex = Group-RowsByIdentity -Rows $SendAsRows
    $exchangeGroupIndex = Group-RowsByIdentity -Rows $ExchangeGroupRows
    $entraGroupIndex = Group-RowsByIdentity -Rows $EntraGroupRows
    $errorIndex = Group-RowsByIdentity -Rows $ErrorRows

    foreach ($object in $Objects) {
        $identity = [string]$object.Identity
        [pscustomobject]@{
            Identity                  = $object.Identity
            DisplayName               = $object.DisplayName
            RecipientType             = $object.RecipientType
            RecipientTypeDetails      = $object.RecipientTypeDetails
            MailboxType               = $object.MailboxType
            PrimarySmtpAddress        = $object.PrimarySmtpAddress
            Alias                     = $object.Alias
            ExternalDirectoryObjectId = $object.ExternalDirectoryObjectId
            ExchangeObjectId          = $object.ExchangeObjectId
            ProxyAddresses            = Join-Values (@($proxyIndex[$identity]) | ForEach-Object { $_.RawProxyAddress })
            FullAccessTrustees        = Join-Values (@($fullAccessIndex[$identity]) | ForEach-Object { $_.Trustee })
            SendAsTrustees            = Join-Values (@($sendAsIndex[$identity]) | ForEach-Object { $_.Trustee })
            ExchangeGroupMemberships  = Join-Values (@($exchangeGroupIndex[$identity]) | ForEach-Object { $_.GroupIdentity })
            EntraGroupMemberships     = Join-Values (@($entraGroupIndex[$identity]) | ForEach-Object { $_.GroupDisplayName })
            HasErrors                 = [bool]($errorIndex.ContainsKey($identity) -and $errorIndex[$identity].Count -gt 0)
        }
    }
}

$resolvedInputFolder = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($InputFolder)
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path $resolvedInputFolder 'ExchangeObjectData.xlsx'
}
$resolvedOutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)

if (Test-Path -Path $resolvedOutputPath) {
    if (-not $Force) {
        throw "Workbook '$resolvedOutputPath' already exists. Use -Force to overwrite it."
    }
    # Export-Excel appends sheets to an existing workbook; remove it so stale sheets never survive.
    Remove-Item -Path $resolvedOutputPath -Force
}

Write-Host "Input folder: $resolvedInputFolder" -ForegroundColor Cyan
Write-Host "Workbook: $resolvedOutputPath" -ForegroundColor Cyan

$objectsPath = Join-Path $resolvedInputFolder 'Objects.csv'
if (-not (Test-Path -Path $objectsPath -PathType Leaf)) {
    throw "Objects.csv not found in '$resolvedInputFolder'. Run Export-ExchangeObjectInventory.ps1 (or Export-ExchangeObjectData.ps1) against this folder first."
}
$objects = @(Import-Csv -Path $objectsPath)

$proxyRows = @(Import-DataSet -Folder $resolvedInputFolder -FileName 'ProxyAddresses.csv')
$fullAccessRows = @(Import-DataSet -Folder $resolvedInputFolder -FileName 'FullAccess.csv')
$sendAsRows = @(Import-DataSet -Folder $resolvedInputFolder -FileName 'SendAs.csv')
$exchangeGroupRows = @(Import-DataSet -Folder $resolvedInputFolder -FileName 'ExchangeGroupMemberships.csv')
$entraGroupRows = @(Import-DataSet -Folder $resolvedInputFolder -FileName 'EntraGroupMemberships.csv')

$errorRows = @(
    Get-ChildItem -Path $resolvedInputFolder -Filter 'Errors*.csv' -File |
        Sort-Object -Property Name |
        ForEach-Object { Import-Csv -Path $_.FullName }
)

Write-Host 'Building consolidated rows...' -ForegroundColor Cyan
$consolidatedRows = @(
    New-ConsolidatedRows `
        -Objects $objects `
        -ProxyRows $proxyRows `
        -FullAccessRows $fullAccessRows `
        -SendAsRows $sendAsRows `
        -ExchangeGroupRows $exchangeGroupRows `
        -EntraGroupRows $entraGroupRows `
        -ErrorRows $errorRows
)

$dataSets = @(
    @{ Name = 'Consolidated';             Rows = $consolidatedRows;   Source = 'derived from Objects.csv and related CSVs' }
    @{ Name = 'Objects';                  Rows = $objects;            Source = 'Objects.csv' }
    @{ Name = 'ProxyAddresses';           Rows = $proxyRows;          Source = 'ProxyAddresses.csv' }
    @{ Name = 'FullAccess';               Rows = $fullAccessRows;     Source = 'FullAccess.csv' }
    @{ Name = 'SendAs';                   Rows = $sendAsRows;         Source = 'SendAs.csv' }
    @{ Name = 'ExchangeGroupMemberships'; Rows = $exchangeGroupRows;  Source = 'ExchangeGroupMemberships.csv' }
    @{ Name = 'EntraGroupMemberships';    Rows = $entraGroupRows;     Source = 'EntraGroupMemberships.csv' }
    @{ Name = 'Errors';                   Rows = $errorRows;          Source = 'Errors*.csv (merged)' }
)

$summaryRows = @(
    foreach ($dataSet in $dataSets) {
        [pscustomobject]@{
            Dataset    = [string]$dataSet.Name
            Rows       = @($dataSet.Rows).Count
            SourceFile = [string]$dataSet.Source
        }
    }
)

$summaryRows | Export-Excel -Path $resolvedOutputPath -WorksheetName 'Summary' -AutoFilter -FreezeTopRow -BoldTopRow

$sheetIndex = 0
foreach ($dataSet in $dataSets) {
    $sheetIndex++
    $rows = @($dataSet.Rows)
    if ($rows.Count -eq 0) {
        Write-Host "Skipping empty dataset: $($dataSet.Name)" -ForegroundColor Yellow
        continue
    }
    Write-Progress -Id 1 -Activity 'Writing workbook sheets' -Status "$sheetIndex of $($dataSets.Count): $($dataSet.Name)" -PercentComplete (($sheetIndex / $dataSets.Count) * 100)
    $rows | Export-Excel -Path $resolvedOutputPath -WorksheetName ([string]$dataSet.Name) -AutoFilter -FreezeTopRow -BoldTopRow
}
Write-Progress -Id 1 -Completed

Write-Host "Workbook written: $resolvedOutputPath" -ForegroundColor Green
Write-Host "Objects: $($objects.Count); consolidated rows: $($consolidatedRows.Count); error rows: $($errorRows.Count)" -ForegroundColor Green
