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

function Get-ProxyAddressRows {
    param([Parameter(Mandatory = $true)] $Recipient)

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($address in @($Recipient.EmailAddresses)) {
        if ($null -eq $address) { continue }
        $value = [string]$address
        $parts = $value -split ':', 2
        $prefix = if ($parts.Count -eq 2) { $parts[0] } else { '' }
        $addressValue = if ($parts.Count -eq 2) { $parts[1] } else { $value }

        $rows.Add([pscustomobject]@{
            Identity           = [string]$Recipient.Identity
            PrimarySmtpAddress = [string]$Recipient.PrimarySmtpAddress
            AddressType        = $prefix
            ProxyAddress       = $addressValue
            RawProxyAddress    = $value
            IsPrimarySmtp      = $value.StartsWith('SMTP:')
        }) | Out-Null
    }
    return @($rows)
}

function Get-FullAccessRows {
    param([Parameter(Mandatory = $true)] $Recipient)

    $mailboxType = Get-MailboxType -Recipient $Recipient
    if ([string]::IsNullOrWhiteSpace($mailboxType)) {
        return @()
    }

    try {
        return @(
            Get-EXOMailboxPermission -Identity $Recipient.Identity -ErrorAction Stop |
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
        Add-ExportError -Identity ([string]$Recipient.Identity) -Stage 'FullAccess' -Operation 'Get-EXOMailboxPermission' -Message $_.Exception.Message
        return @()
    }
}

function Get-SendAsRows {
    param([Parameter(Mandatory = $true)] $Recipient)

    try {
        return @(
            Get-RecipientPermission -Identity $Recipient.Identity -ErrorAction Stop |
                Where-Object { -not $_.IsInherited -and ($_.AccessRights -contains 'SendAs') } |
                ForEach-Object {
                    [pscustomobject]@{
                        Identity           = [string]$Recipient.Identity
                        PrimarySmtpAddress = [string]$Recipient.PrimarySmtpAddress
                        Trustee            = [string]$_.Trustee
                        AccessRights       = ($_.AccessRights -join '; ')
                        IsInherited        = [bool]$_.IsInherited
                    }
                }
        )
    }
    catch {
        Add-ExportError -Identity ([string]$Recipient.Identity) -Stage 'SendAs' -Operation 'Get-RecipientPermission' -Message $_.Exception.Message
        return @()
    }
}

function Get-ExchangeGroupMembershipRows {
    param([Parameter(Mandatory = $true)] $Recipient)

    try {
        return @(
            Get-EXORecipient -Identity $Recipient.Identity -Properties MemberOfGroup -ErrorAction Stop |
                Select-Object -ExpandProperty MemberOfGroup -ErrorAction SilentlyContinue |
                ForEach-Object {
                    [pscustomobject]@{
                        Identity           = [string]$Recipient.Identity
                        PrimarySmtpAddress = [string]$Recipient.PrimarySmtpAddress
                        GroupIdentity      = [string]$_
                    }
                }
        )
    }
    catch {
        Add-ExportError -Identity ([string]$Recipient.Identity) -Stage 'ExchangeGroupMemberships' -Operation 'Get-EXORecipient MemberOfGroup' -Message $_.Exception.Message
        return @()
    }
}

function Connect-GraphIfNeeded {
    if ($SkipGraph) {
        Write-Host 'Skipping Microsoft Graph connection because -SkipGraph was specified.' -ForegroundColor Yellow
        return $false
    }

    try {
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
        Import-Module Microsoft.Graph.DirectoryObjects -ErrorAction Stop
        Connect-MgGraph -Scopes @('User.Read.All', 'Group.Read.All', 'Directory.Read.All') -NoWelcome
        return $true
    }
    catch {
        Add-ExportError -Identity '' -Stage 'GraphConnect' -Operation 'Connect-MgGraph' -Message $_.Exception.Message
        return $false
    }
}

function Get-EntraGroupMembershipRows {
    param([Parameter(Mandatory = $true)] $Recipient)

    $externalId = [string]$Recipient.ExternalDirectoryObjectId
    if ([string]::IsNullOrWhiteSpace($externalId)) {
        return @()
    }

    try {
        return @(
            Get-MgDirectoryObjectMemberOf -DirectoryObjectId $externalId -All -ErrorAction Stop |
                Where-Object { $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group' } |
                ForEach-Object {
                    [pscustomobject]@{
                        Identity             = [string]$Recipient.Identity
                        PrimarySmtpAddress   = [string]$Recipient.PrimarySmtpAddress
                        GroupId              = [string]$_.Id
                        GroupDisplayName     = [string]$_.AdditionalProperties.displayName
                        GroupMail            = [string]$_.AdditionalProperties.mail
                        GroupSecurityEnabled = [string]$_.AdditionalProperties.securityEnabled
                    }
                }
        )
    }
    catch {
        Add-ExportError -Identity ([string]$Recipient.Identity) -Stage 'EntraGroupMemberships' -Operation 'Get-MgDirectoryObjectMemberOf' -Message $_.Exception.Message
        return @()
    }
}

function Get-TargetRecipients {
    if ($PSCmdlet.ParameterSetName -eq 'All') {
        Write-Host "Loading Exchange recipients..." -ForegroundColor Cyan
        return @(
            Get-EXORecipient -ResultSize Unlimited -Properties RecipientTypeDetails,ExternalDirectoryObjectId,PrimarySmtpAddress,Alias,Guid,DistinguishedName,EmailAddresses,MemberOfGroup |
                Where-Object { Test-RecipientMatchesType -Recipient $_ -SelectedRecipientType $RecipientType }
        )
    }

    $identities = @(Get-InputIdentityList)
    $recipients = New-Object System.Collections.Generic.List[object]

    foreach ($inputIdentity in $identities) {
        try {
            $recipient = Get-EXORecipient -Identity $inputIdentity -Properties RecipientTypeDetails,ExternalDirectoryObjectId,PrimarySmtpAddress,Alias,Guid,DistinguishedName,EmailAddresses,MemberOfGroup -ErrorAction Stop
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

function Join-Values {
    param([Parameter(Mandatory = $false)] $Values)

    return (@($Values) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Sort-Object -Unique) -join '; '
}

function New-ConsolidatedRows {
    param(
        [Parameter(Mandatory = $true)] [object[]]$Objects,
        [Parameter(Mandatory = $true)] [object[]]$ProxyRows,
        [Parameter(Mandatory = $true)] [object[]]$FullAccessRows,
        [Parameter(Mandatory = $true)] [object[]]$SendAsRows,
        [Parameter(Mandatory = $true)] [object[]]$ExchangeGroupRows,
        [Parameter(Mandatory = $true)] [object[]]$EntraGroupRows,
        [Parameter(Mandatory = $true)] [object[]]$ErrorRows
    )

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
            ProxyAddresses            = Join-Values (($ProxyRows | Where-Object { $_.Identity -eq $identity }).RawProxyAddress)
            FullAccessTrustees        = Join-Values (($FullAccessRows | Where-Object { $_.Identity -eq $identity }).Trustee)
            SendAsTrustees            = Join-Values (($SendAsRows | Where-Object { $_.Identity -eq $identity }).Trustee)
            ExchangeGroupMemberships  = Join-Values (($ExchangeGroupRows | Where-Object { $_.Identity -eq $identity }).GroupIdentity)
            EntraGroupMemberships     = Join-Values (($EntraGroupRows | Where-Object { $_.Identity -eq $identity }).GroupDisplayName)
            HasErrors                 = [bool](@($ErrorRows | Where-Object { $_.Identity -eq $identity }).Count -gt 0)
        }
    }
}

function Export-ReportCsvs {
    param(
        [Parameter(Mandatory = $true)] [string]$Folder,
        [Parameter(Mandatory = $true)] [object[]]$Objects,
        [Parameter(Mandatory = $true)] [object[]]$ConsolidatedRows,
        [Parameter(Mandatory = $true)] [object[]]$ProxyRows,
        [Parameter(Mandatory = $true)] [object[]]$FullAccessRows,
        [Parameter(Mandatory = $true)] [object[]]$SendAsRows,
        [Parameter(Mandatory = $true)] [object[]]$ExchangeGroupRows,
        [Parameter(Mandatory = $true)] [object[]]$EntraGroupRows,
        [Parameter(Mandatory = $true)] [object[]]$ErrorRows
    )

    $ConsolidatedRows | Export-Csv -Path (Join-Path $Folder 'Objects-Consolidated.csv') -NoTypeInformation -Encoding UTF8
    $Objects | Export-Csv -Path (Join-Path $Folder 'Objects.csv') -NoTypeInformation -Encoding UTF8
    $ProxyRows | Export-Csv -Path (Join-Path $Folder 'ProxyAddresses.csv') -NoTypeInformation -Encoding UTF8
    $FullAccessRows | Export-Csv -Path (Join-Path $Folder 'FullAccess.csv') -NoTypeInformation -Encoding UTF8
    $SendAsRows | Export-Csv -Path (Join-Path $Folder 'SendAs.csv') -NoTypeInformation -Encoding UTF8
    $ExchangeGroupRows | Export-Csv -Path (Join-Path $Folder 'ExchangeGroupMemberships.csv') -NoTypeInformation -Encoding UTF8
    $EntraGroupRows | Export-Csv -Path (Join-Path $Folder 'EntraGroupMemberships.csv') -NoTypeInformation -Encoding UTF8
    $ErrorRows | Export-Csv -Path (Join-Path $Folder 'Errors.csv') -NoTypeInformation -Encoding UTF8
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

$graphConnected = Connect-GraphIfNeeded
$entraGroupMembershipRows = New-Object System.Collections.Generic.List[object]

$proxyRows = New-Object System.Collections.Generic.List[object]
$fullAccessRows = New-Object System.Collections.Generic.List[object]
$sendAsRows = New-Object System.Collections.Generic.List[object]
$exchangeGroupMembershipRows = New-Object System.Collections.Generic.List[object]

$index = 0
foreach ($recipient in $targetRecipients) {
    $index++
    Write-Progress -Activity 'Collecting Exchange object data' -Status "$index of $($targetRecipients.Count): $($recipient.Identity)" -PercentComplete (($index / [Math]::Max($targetRecipients.Count, 1)) * 100)

    foreach ($row in @(Get-ProxyAddressRows -Recipient $recipient)) { $proxyRows.Add($row) | Out-Null }
    foreach ($row in @(Get-FullAccessRows -Recipient $recipient)) { $fullAccessRows.Add($row) | Out-Null }
    foreach ($row in @(Get-SendAsRows -Recipient $recipient)) { $sendAsRows.Add($row) | Out-Null }
    foreach ($row in @(Get-ExchangeGroupMembershipRows -Recipient $recipient)) { $exchangeGroupMembershipRows.Add($row) | Out-Null }
    if ($graphConnected) {
        foreach ($row in @(Get-EntraGroupMembershipRows -Recipient $recipient)) { $entraGroupMembershipRows.Add($row) | Out-Null }
    }
}
Write-Progress -Activity 'Collecting Exchange object data' -Completed

$consolidatedRows = @(
    New-ConsolidatedRows `
        -Objects $objects `
        -ProxyRows @($proxyRows) `
        -FullAccessRows @($fullAccessRows) `
        -SendAsRows @($sendAsRows) `
        -ExchangeGroupRows @($exchangeGroupMembershipRows) `
        -EntraGroupRows @($entraGroupMembershipRows) `
        -ErrorRows @($script:Errors)
)

Export-ReportCsvs `
    -Folder $resolvedOutputFolder `
    -Objects $objects `
    -ConsolidatedRows $consolidatedRows `
    -ProxyRows @($proxyRows) `
    -FullAccessRows @($fullAccessRows) `
    -SendAsRows @($sendAsRows) `
    -ExchangeGroupRows @($exchangeGroupMembershipRows) `
    -EntraGroupRows @($entraGroupMembershipRows) `
    -ErrorRows @($script:Errors)

Write-Host "CSV exports written to: $resolvedOutputFolder" -ForegroundColor Green
