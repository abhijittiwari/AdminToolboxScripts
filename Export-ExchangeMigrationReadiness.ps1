<#
.SYNOPSIS
    Identifies Exchange mail objects blocked, risky, or unsuitable for migration.

.DESCRIPTION
    Analyzes legal hold, retention, mailbox type, licensing, and compliance
    constraints for each in-scope object. Produces a migration readiness report
    flagging objects as Ready, Risky, Blocked, or Review. Can operate on all
    recipients, recipients from a CSV, or those in an Objects.csv inventory
    from Export-ExchangeObjectInventory.ps1.

.NOTES
    Requires: ExchangeOnlineManagement, Microsoft.Graph.Authentication
    This script is read-only. It does not modify Exchange or Entra objects.

.EXAMPLE
    ./Export-ExchangeMigrationReadiness.ps1 -All

.EXAMPLE
    ./Export-ExchangeMigrationReadiness.ps1 -ObjectsCsv ./export/Objects.csv

.EXAMPLE
    ./Export-ExchangeMigrationReadiness.ps1 -InputCsv ./objects.csv -IncludeLicensing
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

    [Parameter(Mandatory = $true, ParameterSetName = 'ObjectsCsv')]
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
    [string]$ObjectsCsv,

    [Parameter(Mandatory = $false)]
    [string]$OutputFolder,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeLicensing
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

function Connect-GraphIfNeeded {
    if (-not $IncludeLicensing) { return }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    if (Get-MgContext) {
        Write-Host 'Reusing existing Microsoft Graph context.' -ForegroundColor Cyan
        return
    }
    Connect-MgGraph -Scopes @('User.Read.All') -NoWelcome
}

function Invoke-MigrationReadinessCheck {
    param(
        [string]$Identity,
        [string]$DisplayName,
        [string]$RecipientType,
        [string]$MailboxType,
        [bool]$LitigationHold,
        [string]$RetentionPolicy,
        [object[]]$ComplianceHolds,
        [bool]$Licensing,
        [string[]]$Licenses = @()
    )

    $blockers = @()
    $status = 'Ready'

    if ($LitigationHold) {
        $blockers += 'litigation hold'
        $status = 'Blocked'
    }

    if (@($ComplianceHolds).Count -gt 0) {
        $blockers += "compliance hold ($(@($ComplianceHolds) -join ', '))"
        $status = 'Blocked'
    }

    if ($status -ne 'Blocked' -and -not [string]::IsNullOrWhiteSpace($RetentionPolicy)) {
        $blockers += "retention policy ($RetentionPolicy)"
        $status = 'Risky'
    }

    if (-not $Licensing) {
        $blockers += 'no license'
        $status = 'Blocked'
    }

    if ($status -ne 'Blocked' -and -not [string]::IsNullOrWhiteSpace($MailboxType)) {
        if ($MailboxType -in @('DiscoveryMailbox', 'LegacyMailbox', 'LinkedMailbox', 'LinkedRoomMailbox')) {
            $blockers += $MailboxType
            $status = 'Review'
        }
    }

    [pscustomobject]@{
        Identity           = $Identity
        DisplayName        = $DisplayName
        RecipientType      = $RecipientType
        MailboxType        = $MailboxType
        LitigationHold     = $LitigationHold
        RetentionPolicy    = $RetentionPolicy
        ComplianceHolds    = @($ComplianceHolds) -join '; '
        Licensing          = $Licensing
        Licenses           = @($Licenses) -join '; '
        MigrationStatus    = $status
        BlockingReasons    = $blockers -join '; '
        HasErrors          = $false
    }
}

function Get-TargetRecipients {
    if ($PSCmdlet.ParameterSetName -eq 'ObjectsCsv') {
        return @(Import-Csv -Path $ObjectsCsv)
    }

    if ($PSCmdlet.ParameterSetName -eq 'All') {
        Write-Host "Loading Exchange recipients..." -ForegroundColor Cyan
        return @(Get-EXORecipient -ResultSize Unlimited -Properties RecipientTypeDetails, ExternalDirectoryObjectId)
    }

    if ($PSCmdlet.ParameterSetName -eq 'Csv') {
        $rows = @(Import-Csv -Path $InputCsv)
        if ($rows.Count -eq 0) { throw "Input CSV '$InputCsv' contains no rows." }
        if ($null -eq $rows[0].PSObject.Properties[$IdentityColumn]) {
            throw "Input CSV '$InputCsv' does not contain column '$IdentityColumn'."
        }
        return @($rows | ForEach-Object { $_.PSObject.Properties[$IdentityColumn].Value } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { Get-EXORecipient -Identity $_ })
    }

    if ($PSCmdlet.ParameterSetName -eq 'Identity') {
        return @(Get-EXORecipient -Identity $Identity)
    }

    return @()
}

if ([string]::IsNullOrWhiteSpace($OutputFolder)) {
    if ($PSCmdlet.ParameterSetName -eq 'ObjectsCsv') {
        $OutputFolder = Split-Path -Parent $ObjectsCsv
    }
    else {
        $OutputFolder = ".\ExchangeMigrationReadiness_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    }
}
$resolvedOutputFolder = Initialize-OutputFolder -OutputFolder $OutputFolder

Write-Host "Output folder: $resolvedOutputFolder" -ForegroundColor Cyan
Write-Host "Include licensing: $IncludeLicensing" -ForegroundColor Cyan

Connect-ExchangeOnlineIfNeeded
Connect-GraphIfNeeded

$recipients = @(Get-TargetRecipients)
Write-Host "Objects to assess: $($recipients.Count)" -ForegroundColor Green

$readinessRows = [System.Collections.Generic.List[object]]::new()
$licenseCache = @{}

$index = 0
foreach ($recipient in $recipients) {
    $index++
    Write-Progress -Id 1 -Activity 'Assessing migration readiness' -Status "$index of $($recipients.Count): $($recipient.DisplayName)" -PercentComplete (($index / [Math]::Max($recipients.Count, 1)) * 100)

    try {
        # Try multiple identity types to maximize hit rate: SMTP first, then GUID, then raw identity.
        $lookupIdentity = $null
        $lookupMethod = $null

        if (-not [string]::IsNullOrWhiteSpace([string]$recipient.PrimarySmtpAddress)) {
            $lookupIdentity = [string]$recipient.PrimarySmtpAddress
            $lookupMethod = 'SMTP'
        }
        elseif (-not [string]::IsNullOrWhiteSpace([string]$recipient.ExternalDirectoryObjectId)) {
            $lookupIdentity = [string]$recipient.ExternalDirectoryObjectId
            $lookupMethod = 'GUID'
        }
        else {
            $lookupIdentity = [string]$recipient.Identity
            $lookupMethod = 'Identity'
        }

        $mailbox = Get-EXOMailbox -Identity $lookupIdentity -ErrorAction Stop -Properties LitigationHoldEnabled, RetentionPolicy, InPlaceHolds
        $litigationHold = [bool]$mailbox.LitigationHoldEnabled
        $retentionPolicy = [string]$mailbox.RetentionPolicy
        $complianceHolds = @($mailbox.InPlaceHolds) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

        $hasLicense = $true
        $licenseNames = @()
        if ($IncludeLicensing) {
            $externalId = [string]$recipient.ExternalDirectoryObjectId
            if (-not [string]::IsNullOrWhiteSpace($externalId)) {
                if (-not $licenseCache.ContainsKey($externalId)) {
                    try {
                        $licenseDetails = Invoke-MgGraphRequest -Method GET -Uri "v1.0/users/$externalId/licenseDetails" -ErrorAction Stop
                        $licenseCache[$externalId] = @(@($licenseDetails.value) | ForEach-Object { [string]$_.skuPartNumber } | Where-Object { $_ } | Sort-Object)
                    }
                    catch {
                        Add-ExportError -Identity ([string]$recipient.Identity) -Stage 'Readiness' -Operation 'Get-LicenseDetails' -Message $_.Exception.Message
                        $licenseCache[$externalId] = @()
                    }
                }
                $licenseNames = @($licenseCache[$externalId])
                $hasLicense = $licenseNames.Count -gt 0
            }
        }

        $mailboxType = [string]$mailbox.RecipientTypeDetails
        $row = Invoke-MigrationReadinessCheck -Identity ([string]$recipient.Identity) -DisplayName ([string]$recipient.DisplayName) -RecipientType ([string]$recipient.RecipientType) -MailboxType $mailboxType -LitigationHold $litigationHold -RetentionPolicy $retentionPolicy -ComplianceHolds $complianceHolds -Licensing $hasLicense -Licenses $licenseNames
        $readinessRows.Add($row) | Out-Null
    }
    catch {
        Add-ExportError -Identity ([string]$recipient.Identity) -Stage 'Readiness' -Operation 'Get-EXOMailbox' -Message $_.Exception.Message
    }
}
Write-Progress -Id 1 -Completed

$blockedRows = @($readinessRows | Where-Object { $_.MigrationStatus -eq 'Blocked' })

Export-CsvWithHeaders -Rows @($readinessRows) -Path (Join-Path $resolvedOutputFolder 'MigrationReadiness.csv') -Headers @('Identity', 'DisplayName', 'RecipientType', 'MailboxType', 'LitigationHold', 'RetentionPolicy', 'ComplianceHolds', 'Licensing', 'Licenses', 'MigrationStatus', 'BlockingReasons', 'HasErrors')
Export-CsvWithHeaders -Rows $blockedRows -Path (Join-Path $resolvedOutputFolder 'BlockedObjects.csv') -Headers @('Identity', 'DisplayName', 'RecipientType', 'MailboxType', 'Licenses', 'MigrationStatus', 'BlockingReasons')
Export-CsvWithHeaders -Rows @($script:Errors) -Path (Join-Path $resolvedOutputFolder 'Errors-MigrationReadiness.csv') -Headers @('Identity', 'Stage', 'Operation', 'Message')

Write-Host "Migration readiness report: $(@($readinessRows).Count) objects assessed" -ForegroundColor Green
Write-Host "  Ready: $(@($readinessRows | Where-Object { $_.MigrationStatus -eq 'Ready' }).Count)" -ForegroundColor Green
Write-Host "  Risky: $(@($readinessRows | Where-Object { $_.MigrationStatus -eq 'Risky' }).Count)" -ForegroundColor Yellow
Write-Host "  Review: $(@($readinessRows | Where-Object { $_.MigrationStatus -eq 'Review' }).Count)" -ForegroundColor Yellow
Write-Host "  Blocked: $($blockedRows.Count)" -ForegroundColor Red
Write-Host "Errors logged: $($script:Errors.Count)" -ForegroundColor Yellow
