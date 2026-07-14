<#
.SYNOPSIS
    Exports FullAccess mailbox permissions for recipients in an Objects.csv inventory.

.DESCRIPTION
    FullAccess job of the modular Exchange object export. Reads the Objects.csv produced by
    Export-ExchangeObjectInventory.ps1 (or Export-ExchangeObjectData.ps1) and collects
    non-inherited FullAccess trustees (SELF excluded) for each mailbox-type recipient.
    Non-mailbox recipients are skipped without a query.
    Reuses an existing Exchange Online session when one exists and leaves the session open
    so chained job runs authenticate once.

.NOTES
    Requires: ExchangeOnlineManagement
    This script is read-only. It does not modify Exchange or Entra objects.

.EXAMPLE
    ./Export-ExchangeFullAccessPermissions.ps1 -ObjectsCsv ./ExchangeObjectInventory_20260714_120000/Objects.csv
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
    [string]$ObjectsCsv,

    [Parameter(Mandatory = $false)]
    [string]$OutputFolder
)

$ErrorActionPreference = 'Stop'

$script:Errors = [System.Collections.Generic.List[object]]::new()

function Add-ExportError {
    param(
        [Parameter(Mandatory = $false)] [string]$Identity,
        [Parameter(Mandatory = $true)] [string]$Stage,
        [Parameter(Mandatory = $true)] [string]$Operation,
        [Parameter(Mandatory = $true)] [string]$Message
    )

    $script:Errors.Add([pscustomobject]@{
        Identity  = $Identity
        Stage     = $Stage
        Operation = $Operation
        Message   = $Message
    }) | Out-Null

    Write-Warning "[$Stage/$Operation] $Identity $Message"
}

function Initialize-OutputFolder {
    param([Parameter(Mandatory = $true)] [string]$OutputFolder)

    $resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputFolder)
    if (-not (Test-Path -Path $resolvedPath)) {
        New-Item -Path $resolvedPath -ItemType Directory -Force | Out-Null
    }
    return $resolvedPath
}

function Import-ObjectsCsv {
    param(
        [Parameter(Mandatory = $true)] [string]$Path,
        [Parameter(Mandatory = $true)] [string[]]$RequiredColumns
    )

    $rows = @(Import-Csv -Path $Path)
    if ($rows.Count -eq 0) {
        throw "Objects CSV '$Path' contains no rows. Run Export-ExchangeObjectInventory.ps1 first."
    }

    foreach ($column in $RequiredColumns) {
        if ($null -eq $rows[0].PSObject.Properties[$column]) {
            throw "Objects CSV '$Path' does not contain required column '$column'. Expected an Objects.csv produced by Export-ExchangeObjectInventory.ps1."
        }
    }

    return $rows
}

function Export-CsvWithHeaders {
    param(
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] [object[]]$Rows,
        [Parameter(Mandatory = $true)] [string]$Path,
        [Parameter(Mandatory = $true)] [string[]]$Headers
    )

    if ($Rows.Count -gt 0) {
        $Rows | Select-Object -Property $Headers | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
        return
    }

    $headerLine = ($Headers | ForEach-Object { '"' + ($_ -replace '"', '""') + '"' }) -join ','
    Set-Content -Path $Path -Value $headerLine -Encoding UTF8
}

function Connect-ExchangeOnlineIfNeeded {
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    if (@(Get-ConnectionInformation).Count -gt 0) {
        Write-Host 'Reusing existing Exchange Online connection.' -ForegroundColor Cyan
        return
    }
    Connect-ExchangeOnline -ShowBanner:$false
}

function Get-MailboxType {
    param([Parameter(Mandatory = $true)] $Recipient)

    $typeDetails = [string]$Recipient.RecipientTypeDetails
    if ($typeDetails -match 'Mailbox$') {
        return $typeDetails
    }

    if ([string]$Recipient.RecipientType -eq 'UserMailbox') {
        return 'UserMailbox'
    }

    return ''
}

function Get-FullAccessRows {
    param([Parameter(Mandatory = $true)] $Recipient)

    $mailboxType = Get-MailboxType -Recipient $Recipient
    if ([string]::IsNullOrWhiteSpace($mailboxType)) {
        return @()
    }

    $lookupIdentity = [string]$Recipient.PrimarySmtpAddress
    if ([string]::IsNullOrWhiteSpace($lookupIdentity)) { $lookupIdentity = [string]$Recipient.ExternalDirectoryObjectId }
    if ([string]::IsNullOrWhiteSpace($lookupIdentity)) { $lookupIdentity = [string]$Recipient.Guid }

    try {
        return @(
            Get-EXOMailboxPermission -Identity $lookupIdentity -ErrorAction Stop |
                Where-Object { -not $_.IsInherited -and $_.User -notmatch '^NT AUTHORITY\\SELF$' -and ($_.AccessRights -contains 'FullAccess') } |
                ForEach-Object {
                    [pscustomobject]@{
                        Identity           = [string]$Recipient.Identity
                        PrimarySmtpAddress = [string]$Recipient.PrimarySmtpAddress
                        Trustee            = [string]$_.User
                        AccessRights       = ($_.AccessRights -join '; ')
                        Deny               = [bool]$_.Deny
                        IsInherited        = [bool]$_.IsInherited
                    }
                }
        )
    }
    catch {
        Add-ExportError -Identity ([string]$Recipient.Identity) -Stage 'FullAccess' -Operation 'Get-EXOMailboxPermission' -Message "$($_.Exception.Message) (lookup: $lookupIdentity)"
        return @()
    }
}

$resolvedObjectsCsv = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ObjectsCsv)
if ([string]::IsNullOrWhiteSpace($OutputFolder)) {
    $OutputFolder = Split-Path -Parent $resolvedObjectsCsv
}
$resolvedOutputFolder = Initialize-OutputFolder -OutputFolder $OutputFolder

Write-Host "Objects CSV: $resolvedObjectsCsv" -ForegroundColor Cyan
Write-Host "Output folder: $resolvedOutputFolder" -ForegroundColor Cyan

$recipients = @(Import-ObjectsCsv -Path $resolvedObjectsCsv -RequiredColumns @('Identity', 'PrimarySmtpAddress', 'RecipientTypeDetails', 'RecipientType'))
Write-Host "Objects loaded: $($recipients.Count)" -ForegroundColor Green

Connect-ExchangeOnlineIfNeeded

Write-Host 'Collecting FullAccess permissions...' -ForegroundColor Cyan
$fullAccessRows = [System.Collections.Generic.List[object]]::new()
$mailboxRecipients = @($recipients | Where-Object { -not [string]::IsNullOrWhiteSpace((Get-MailboxType -Recipient $_)) })
$index = 0
foreach ($recipient in $mailboxRecipients) {
    $index++
    Write-Progress -Id 1 -Activity 'Collecting FullAccess permissions' -Status "$index of $($mailboxRecipients.Count): $($recipient.Identity)" -PercentComplete (($index / [Math]::Max($mailboxRecipients.Count, 1)) * 100)
    foreach ($row in @(Get-FullAccessRows -Recipient $recipient)) { $fullAccessRows.Add($row) | Out-Null }
}
Write-Progress -Id 1 -Completed

Export-CsvWithHeaders -Rows @($fullAccessRows) -Path (Join-Path $resolvedOutputFolder 'FullAccess.csv') -Headers @('Identity', 'PrimarySmtpAddress', 'Trustee', 'AccessRights', 'Deny', 'IsInherited')
Export-CsvWithHeaders -Rows @($script:Errors) -Path (Join-Path $resolvedOutputFolder 'Errors-FullAccess.csv') -Headers @('Identity', 'Stage', 'Operation', 'Message')

Write-Host "Mailboxes queried: $($mailboxRecipients.Count) of $($recipients.Count) objects" -ForegroundColor Green
Write-Host "FullAccess rows: $($fullAccessRows.Count)" -ForegroundColor Green
Write-Host "Errors logged: $($script:Errors.Count)" -ForegroundColor Yellow
