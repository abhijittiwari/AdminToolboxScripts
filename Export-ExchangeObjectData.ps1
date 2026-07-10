<#
.SYNOPSIS
    Exports Exchange Online recipient object data, permissions, memberships, and dashboard output.

.DESCRIPTION
    Exports proxy addresses, FullAccess permissions, SendAs permissions, Exchange group memberships,
    and Entra ID group memberships for Exchange Online recipients. Supports all recipients of a
    selected type, identities from CSV, or a single identity.

.NOTES
    Requires: ExchangeOnlineManagement
    Optional Graph enrichment requires: Microsoft.Graph
    This script is read-only. It does not modify Exchange or Entra objects.

.EXAMPLE
    ./Export-ExchangeObjectData.ps1 -All

.EXAMPLE
    ./Export-ExchangeObjectData.ps1 -All -RecipientType All

.EXAMPLE
    ./Export-ExchangeObjectData.ps1 -InputCsv ./objects.csv -IdentityColumn Identity -RecipientType Group

.EXAMPLE
    ./Export-ExchangeObjectData.ps1 -Identity user@contoso.com
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
    [string]$OutputFolder = ".\ExchangeObjectData_$(Get-Date -Format 'yyyyMMdd_HHmmss')",

    [Parameter(Mandatory = $false)]
    [switch]$SkipGraph,

    [Parameter(Mandatory = $false)]
    [switch]$NoDashboard
)

$ErrorActionPreference = 'Stop'

$script:Errors = New-Object System.Collections.Generic.List[object]

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
            return ($typeDetails -match 'Mailbox$' -or $type -eq 'UserMailbox')
        }
        'Group' {
            return ($typeDetails -match 'Group$' -or $typeDetails -match 'DistributionGroup' -or $typeDetails -eq 'MailUniversalSecurityGroup')
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

function Get-TargetRecipients {
    if ($PSCmdlet.ParameterSetName -eq 'All') {
        Write-Host "Loading Exchange recipients..." -ForegroundColor Cyan
        return @(
            Get-EXORecipient -ResultSize Unlimited -Properties RecipientTypeDetails,ExternalDirectoryObjectId,PrimarySmtpAddress,Alias,Guid,DistinguishedName |
                Where-Object { Test-RecipientMatchesType -Recipient $_ -SelectedRecipientType $RecipientType }
        )
    }

    $identities = @(Get-InputIdentityList)
    $recipients = New-Object System.Collections.Generic.List[object]

    foreach ($inputIdentity in $identities) {
        try {
            $recipient = Get-EXORecipient -Identity $inputIdentity -Properties RecipientTypeDetails,ExternalDirectoryObjectId,PrimarySmtpAddress,Alias,Guid,DistinguishedName -ErrorAction Stop
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

    return @($recipients)
}

$resolvedOutputFolder = Initialize-OutputFolder -OutputFolder $OutputFolder
$isMultiObjectRun = $PSCmdlet.ParameterSetName -in @('All', 'Csv')

Write-Host "Output folder: $resolvedOutputFolder" -ForegroundColor Cyan
Write-Host "Recipient type: $RecipientType" -ForegroundColor Cyan

Import-Module ExchangeOnlineManagement -ErrorAction Stop
Connect-ExchangeOnline -ShowBanner:$false

$targetRecipients = @(Get-TargetRecipients)
$objects = @($targetRecipients | ForEach-Object { ConvertTo-ExportObject -Recipient $_ })

Write-Host "Objects selected: $($objects.Count)" -ForegroundColor Green
