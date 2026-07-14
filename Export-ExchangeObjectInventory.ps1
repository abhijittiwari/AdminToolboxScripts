<#
.SYNOPSIS
    Exports the Exchange Online recipient inventory (Objects.csv) and proxy addresses.

.DESCRIPTION
    First job of the modular Exchange object export. Selects recipients (all of a type,
    identities from CSV, or a single identity) and writes Objects.csv plus ProxyAddresses.csv.
    The Objects.csv output feeds the other job scripts:
      Export-ExchangeGroupMemberships.ps1, Export-ExchangeSendAsPermissions.ps1,
      Export-ExchangeFullAccessPermissions.ps1
    and Export-ExchangeObjectDataWorkbook.ps1 collates the folder into an Excel workbook.
    Reuses an existing Exchange Online session when one exists and leaves the session open
    so chained job runs authenticate once.

.NOTES
    Requires: ExchangeOnlineManagement
    This script is read-only. It does not modify Exchange or Entra objects.

.EXAMPLE
    ./Export-ExchangeObjectInventory.ps1 -All

.EXAMPLE
    ./Export-ExchangeObjectInventory.ps1 -InputCsv ./objects.csv -IdentityColumn Identity -RecipientType Group

.EXAMPLE
    ./Export-ExchangeObjectInventory.ps1 -Identity user@contoso.com
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
    [string]$OutputFolder = ".\ExchangeObjectInventory_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
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

function Get-InputIdentityList {
    if ($PSCmdlet.ParameterSetName -eq 'Identity') {
        return @($Identity)
    }

    if ($PSCmdlet.ParameterSetName -eq 'Csv') {
        $rows = @(Import-Csv -Path $InputCsv)
        if ($rows.Count -eq 0) {
            throw "Input CSV '$InputCsv' contains no rows."
        }

        if ($null -eq $rows[0].PSObject.Properties[$IdentityColumn]) {
            throw "Input CSV '$InputCsv' does not contain identity column '$IdentityColumn'."
        }

        return @(
            $rows |
                ForEach-Object { [string]$_.PSObject.Properties[$IdentityColumn].Value } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        )
    }

    return @()
}

function Test-RecipientMatchesType {
    param(
        [Parameter(Mandatory = $true)] $Recipient,
        [Parameter(Mandatory = $true)] [string]$SelectedRecipientType
    )

    if ($SelectedRecipientType -eq 'All') {
        return $true
    }

    $typeDetails = [string]$Recipient.RecipientTypeDetails
    $type = [string]$Recipient.RecipientType

    switch ($SelectedRecipientType) {
        'Mailbox' {
            return ($typeDetails -ne 'GroupMailbox' -and ($typeDetails -match 'Mailbox$' -or $type -eq 'UserMailbox'))
        }
        'Group' {
            # $type catches group recipients whose details do not end in Group, e.g. RoomList.
            return ($typeDetails -eq 'GroupMailbox' -or $typeDetails -match 'Group$' -or $typeDetails -match 'DistributionGroup' -or $type -match 'Group')
        }
        'MailUser' {
            return ($typeDetails -eq 'MailUser' -or $type -eq 'MailUser')
        }
        'MailContact' {
            return ($typeDetails -eq 'MailContact' -or $type -eq 'MailContact')
        }
    }

    return $false
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

function ConvertTo-ExportObject {
    param([Parameter(Mandatory = $true)] $Recipient)

    [pscustomobject]@{
        Identity                  = [string]$Recipient.Identity
        DisplayName               = [string]$Recipient.DisplayName
        RecipientType             = [string]$Recipient.RecipientType
        RecipientTypeDetails      = [string]$Recipient.RecipientTypeDetails
        MailboxType               = Get-MailboxType -Recipient $Recipient
        PrimarySmtpAddress        = [string]$Recipient.PrimarySmtpAddress
        Alias                     = [string]$Recipient.Alias
        ExternalDirectoryObjectId = [string]$Recipient.ExternalDirectoryObjectId
        ExchangeObjectId          = [string]$Recipient.Guid
        DistinguishedName         = [string]$Recipient.DistinguishedName
    }
}

function ConvertTo-ProxyRows {
    param([Parameter(Mandatory = $true)] $Recipient)

    $rows = @()
    foreach ($address in @($Recipient.EmailAddresses)) {
        if ($null -eq $address) { continue }
        $value = $address.ToString()
        $parts = $value -split ':', 2
        $prefix = if ($parts.Count -eq 2) { $parts[0] } else { '' }
        $addressValue = if ($parts.Count -eq 2) { $parts[1] } else { $value }

        $rows += [pscustomobject]@{
            Identity           = [string]$Recipient.Identity
            PrimarySmtpAddress = [string]$Recipient.PrimarySmtpAddress
            AddressType        = $prefix
            ProxyAddress       = $addressValue
            RawProxyAddress    = $value
            IsPrimarySmtp      = $value.StartsWith('SMTP:')
        }
    }
    return @($rows)
}

function Get-TargetRecipients {
    if ($PSCmdlet.ParameterSetName -eq 'All') {
        Write-Host "Loading Exchange recipients..." -ForegroundColor Cyan
        return @(
            Get-EXORecipient -ResultSize Unlimited -Properties RecipientTypeDetails,ExternalDirectoryObjectId,PrimarySmtpAddress,Alias,Guid,DistinguishedName,EmailAddresses |
                Where-Object { Test-RecipientMatchesType -Recipient $_ -SelectedRecipientType $RecipientType }
        )
    }

    $identities = @(Get-InputIdentityList)
    $recipients = [System.Collections.Generic.List[object]]::new()

    $lookupIndex = 0
    foreach ($inputIdentity in $identities) {
        $lookupIndex++
        Write-Progress -Id 1 -Activity 'Resolving input identities' -Status "$lookupIndex of $($identities.Count): $inputIdentity" -PercentComplete (($lookupIndex / [Math]::Max($identities.Count, 1)) * 100)
        try {
            $recipient = Get-EXORecipient -Identity $inputIdentity -Properties RecipientTypeDetails,ExternalDirectoryObjectId,PrimarySmtpAddress,Alias,Guid,DistinguishedName,EmailAddresses -ErrorAction Stop
            if (Test-RecipientMatchesType -Recipient $recipient -SelectedRecipientType $RecipientType) {
                $recipients.Add($recipient) | Out-Null
            }
            else {
                Add-ExportError -Identity $inputIdentity -Stage 'RecipientTypeFilter' -Operation 'FilterRecipient' -Message "Recipient '$($recipient.DisplayName)' does not match selected recipient type '$RecipientType'."
            }
        }
        catch {
            Add-ExportError -Identity $inputIdentity -Stage 'RecipientLookup' -Operation 'Get-EXORecipient' -Message $_.Exception.Message
        }
    }

    Write-Progress -Id 1 -Completed
    return @($recipients)
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

$resolvedOutputFolder = Initialize-OutputFolder -OutputFolder $OutputFolder
Write-Host "Output folder: $resolvedOutputFolder" -ForegroundColor Cyan
Write-Host "Recipient type: $RecipientType" -ForegroundColor Cyan

Connect-ExchangeOnlineIfNeeded

$targetRecipients = @(Get-TargetRecipients)

$objectList = [System.Collections.Generic.List[object]]::new()
$proxyRows = [System.Collections.Generic.List[object]]::new()
$objectIndex = 0
foreach ($recipient in $targetRecipients) {
    $objectIndex++
    if ($objectIndex % 50 -eq 0 -or $objectIndex -eq $targetRecipients.Count) {
        Write-Progress -Id 1 -Activity 'Building object inventory' -Status "$objectIndex of $($targetRecipients.Count)" -PercentComplete (($objectIndex / [Math]::Max($targetRecipients.Count, 1)) * 100)
    }
    $objectList.Add((ConvertTo-ExportObject -Recipient $recipient)) | Out-Null
    foreach ($row in @(ConvertTo-ProxyRows -Recipient $recipient)) { $proxyRows.Add($row) | Out-Null }
}
Write-Progress -Id 1 -Completed
$objects = @($objectList)

$objectsCsvPath = Join-Path $resolvedOutputFolder 'Objects.csv'
Export-CsvWithHeaders -Rows $objects -Path $objectsCsvPath -Headers @('Identity', 'DisplayName', 'RecipientType', 'RecipientTypeDetails', 'MailboxType', 'PrimarySmtpAddress', 'Alias', 'ExternalDirectoryObjectId', 'ExchangeObjectId', 'DistinguishedName')
Export-CsvWithHeaders -Rows @($proxyRows) -Path (Join-Path $resolvedOutputFolder 'ProxyAddresses.csv') -Headers @('Identity', 'PrimarySmtpAddress', 'AddressType', 'ProxyAddress', 'RawProxyAddress', 'IsPrimarySmtp')
Export-CsvWithHeaders -Rows @($script:Errors) -Path (Join-Path $resolvedOutputFolder 'Errors-Inventory.csv') -Headers @('Identity', 'Stage', 'Operation', 'Message')

Write-Host "Objects exported: $($objects.Count)" -ForegroundColor Green
Write-Host "Proxy address rows: $($proxyRows.Count)" -ForegroundColor Green
Write-Host "Errors logged: $($script:Errors.Count)" -ForegroundColor Yellow
Write-Host "Next: run the membership/permission jobs with -ObjectsCsv '$objectsCsvPath'" -ForegroundColor Cyan
