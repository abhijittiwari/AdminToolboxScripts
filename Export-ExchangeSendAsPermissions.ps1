<#
.SYNOPSIS
    Exports SendAs permissions for recipients in an Objects.csv inventory.

.DESCRIPTION
    SendAs job of the modular Exchange object export. Reads the Objects.csv produced by
    Export-ExchangeObjectInventory.ps1 (or Export-ExchangeObjectData.ps1) and collects
    non-inherited SendAs trustees (SELF excluded) for each recipient.
    Reuses an existing Exchange Online session when one exists and leaves the session open
    so chained job runs authenticate once.

.PARAMETER ObjectsCsv
    Path to the Objects.csv inventory produced by Export-ExchangeObjectInventory.ps1
    (or Export-ExchangeObjectData.ps1). Mandatory. Required columns: Identity,
    PrimarySmtpAddress, RecipientTypeDetails, RecipientType.

.PARAMETER OutputFolder
    Destination folder for SendAs.csv and Errors-SendAs.csv. Defaults to the folder
    containing Objects.csv.

.NOTES
    Requires: ExchangeOnlineManagement
    This script is read-only. It does not modify Exchange or Entra objects.

.EXAMPLE
    ./Export-ExchangeSendAsPermissions.ps1 -ObjectsCsv ./ExchangeObjectInventory_20260714_120000/Objects.csv
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

function Test-IsSelfTrustee {
    param([Parameter(Mandatory = $false)] [string]$Trustee)

    $normalized = $Trustee.Trim()
    return $normalized -in @('NT AUTHORITY\SELF', 'SELF', 'S-1-5-10')
}

function Get-SendAsPermissionCommandName {
    if (Get-Command Get-EXORecipientPermission -ErrorAction SilentlyContinue) {
        return 'Get-EXORecipientPermission'
    }

    if (Get-Command Get-RecipientPermission -ErrorAction SilentlyContinue) {
        return 'Get-RecipientPermission'
    }

    return ''
}

function Get-SendAsRows {
    param(
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] [object[]]$Recipients,
        [Parameter(Mandatory = $true)] [bool]$UseOrgWideQuery
    )

    $rows = [System.Collections.Generic.List[object]]::new()
    $permissionCommand = Get-SendAsPermissionCommandName

    if ([string]::IsNullOrWhiteSpace($permissionCommand)) {
        Add-ExportError -Identity '' -Stage 'SendAs' -Operation 'Get-EXORecipientPermission' -Message 'Neither Get-EXORecipientPermission nor Get-RecipientPermission is available in this Exchange Online session.'
        return @($rows)
    }

    if ($UseOrgWideQuery -and $permissionCommand -eq 'Get-RecipientPermission') {
        $recipientsByIdentity = @{}
        foreach ($recipient in @($Recipients)) { $recipientsByIdentity[[string]$recipient.Identity] = $recipient }
        try {
            Write-Progress -Id 1 -Activity 'Collecting SendAs permissions' -Status 'Running org-wide Get-RecipientPermission query (this can take a while)...'
            $allPermissions = @(& $permissionCommand -ResultSize Unlimited -ErrorAction Stop)
            $permissionIndex = 0
            foreach ($permission in $allPermissions) {
                $permissionIndex++
                if ($permissionIndex % 100 -eq 0 -or $permissionIndex -eq $allPermissions.Count) {
                    Write-Progress -Id 1 -Activity 'Collecting SendAs permissions' -Status "Processing $permissionIndex of $($allPermissions.Count) permission entries" -PercentComplete (($permissionIndex / [Math]::Max($allPermissions.Count, 1)) * 100)
                }
                $identity = [string]$permission.Identity
                if (-not $recipientsByIdentity.ContainsKey($identity)) { continue }
                if ($permission.IsInherited) { continue }
                if (Test-IsSelfTrustee -Trustee ([string]$permission.Trustee)) { continue }
                if ($permission.AccessRights -notcontains 'SendAs') { continue }
                $rows.Add([pscustomobject]@{
                    Identity           = $identity
                    PrimarySmtpAddress = [string]$recipientsByIdentity[$identity].PrimarySmtpAddress
                    Trustee            = [string]$permission.Trustee
                    AccessRights       = ($permission.AccessRights -join '; ')
                    IsInherited        = [bool]$permission.IsInherited
                }) | Out-Null
            }
        }
        catch {
            Add-ExportError -Identity '' -Stage 'SendAs' -Operation "$permissionCommand (org-wide)" -Message $_.Exception.Message
        }
        Write-Progress -Id 1 -Completed
        return @($rows)
    }

    $recipientIndex = 0
    foreach ($recipient in @($Recipients)) {
        $recipientIndex++
        Write-Progress -Id 1 -Activity 'Collecting SendAs permissions' -Status "$recipientIndex of $($Recipients.Count): $($recipient.Identity)" -PercentComplete (($recipientIndex / [Math]::Max($Recipients.Count, 1)) * 100)
        $lookupIdentity = [string]$recipient.PrimarySmtpAddress
        if ([string]::IsNullOrWhiteSpace($lookupIdentity)) { $lookupIdentity = [string]$recipient.ExternalDirectoryObjectId }
        if ([string]::IsNullOrWhiteSpace($lookupIdentity)) { $lookupIdentity = [string]$recipient.Guid }
        try {
            foreach ($permission in @(& $permissionCommand -Identity $lookupIdentity -ErrorAction Stop)) {
                if ($permission.IsInherited) { continue }
                if (Test-IsSelfTrustee -Trustee ([string]$permission.Trustee)) { continue }
                if ($permission.AccessRights -notcontains 'SendAs') { continue }
                $rows.Add([pscustomobject]@{
                    Identity           = [string]$recipient.Identity
                    PrimarySmtpAddress = [string]$recipient.PrimarySmtpAddress
                    Trustee            = [string]$permission.Trustee
                    AccessRights       = ($permission.AccessRights -join '; ')
                    IsInherited        = [bool]$permission.IsInherited
                }) | Out-Null
            }
        }
        catch {
            Add-ExportError -Identity ([string]$recipient.Identity) -Stage 'SendAs' -Operation $permissionCommand -Message "$($_.Exception.Message) (lookup: $lookupIdentity)"
        }
    }
    Write-Progress -Id 1 -Completed
    return @($rows)
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

Write-Host 'Collecting SendAs permissions...' -ForegroundColor Cyan
$sendAsRows = @(Get-SendAsRows -Recipients $recipients -UseOrgWideQuery ($recipients.Count -gt 1))

Export-CsvWithHeaders -Rows $sendAsRows -Path (Join-Path $resolvedOutputFolder 'SendAs.csv') -Headers @('Identity', 'PrimarySmtpAddress', 'Trustee', 'AccessRights', 'IsInherited')
Export-CsvWithHeaders -Rows @($script:Errors) -Path (Join-Path $resolvedOutputFolder 'Errors-SendAs.csv') -Headers @('Identity', 'Stage', 'Operation', 'Message')

Write-Host "SendAs rows: $($sendAsRows.Count)" -ForegroundColor Green
Write-Host "Errors logged: $($script:Errors.Count)" -ForegroundColor Yellow
