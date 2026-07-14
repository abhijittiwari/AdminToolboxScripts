<#
.SYNOPSIS
    Exports Exchange Online recipient object data, permissions, memberships, and dashboard output.

.DESCRIPTION
    Exports proxy addresses, FullAccess permissions, SendAs permissions, Exchange group memberships,
    and Entra ID group memberships for Exchange Online recipients. Supports all recipients of a
    selected type, identities from CSV, or a single identity.

.NOTES
    Requires: ExchangeOnlineManagement
    Group membership collection requires: Microsoft.Graph.Authentication
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

# FIX: was `New-Object System.Collections.Generic.List[object]`. New-Object returns the list
# wrapped in a PSObject; in PowerShell 7, @() over an EMPTY PSObject-wrapped generic list throws
# "Argument types do not match", which killed clean (zero-error) runs at New-ConsolidatedRows.
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

$script:TotalPhases = 7

function Write-PhaseProgress {
    param(
        [Parameter(Mandatory = $true)] [int]$Phase,
        [Parameter(Mandatory = $true)] [string]$Name
    )

    Write-Progress -Id 0 -Activity 'Export-ExchangeObjectData' -Status "Step $Phase of $($script:TotalPhases): $Name" -PercentComplete ((($Phase - 1) / $script:TotalPhases) * 100)
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
            Write-Progress -Id 1 -ParentId 0 -Activity 'Collecting SendAs permissions' -Status 'Running org-wide Get-RecipientPermission query (this can take a while)...'
            $allPermissions = @(& $permissionCommand -ResultSize Unlimited -ErrorAction Stop)
            $permissionIndex = 0
            foreach ($permission in $allPermissions) {
                $permissionIndex++
                if ($permissionIndex % 100 -eq 0 -or $permissionIndex -eq $allPermissions.Count) {
                    Write-Progress -Id 1 -ParentId 0 -Activity 'Collecting SendAs permissions' -Status "Processing $permissionIndex of $($allPermissions.Count) permission entries" -PercentComplete (($permissionIndex / [Math]::Max($allPermissions.Count, 1)) * 100)
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
        Write-Progress -Id 1 -ParentId 0 -Activity 'Collecting SendAs permissions' -Status "$recipientIndex of $($Recipients.Count): $($recipient.Identity)" -PercentComplete (($recipientIndex / [Math]::Max($Recipients.Count, 1)) * 100)
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

function Connect-GraphIfNeeded {
    if ($SkipGraph) {
        Write-Host 'Skipping Microsoft Graph connection because -SkipGraph was specified.' -ForegroundColor Yellow
        return $false
    }

    try {
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
        Connect-MgGraph -Scopes @('User.Read.All', 'Group.Read.All', 'Directory.Read.All') -NoWelcome
        return $true
    }
    catch {
        Add-ExportError -Identity '' -Stage 'GraphConnect' -Operation 'Connect-MgGraph' -Message $_.Exception.Message
        return $false
    }
}

function New-GraphRequestExecutor {
    return {
        param($Method, $Uri, $Body)
        if ($null -ne $Body) {
            Invoke-MgGraphRequest -Method $Method -Uri $Uri -Body ($Body | ConvertTo-Json -Depth 10) -ContentType 'application/json'
        }
        else {
            Invoke-MgGraphRequest -Method $Method -Uri $Uri
        }
    }
}

function Invoke-GraphBatch {
    param(
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] [object[]]$Requests,
        [Parameter(Mandatory = $false)] [scriptblock]$GraphRequest,
        [Parameter(Mandatory = $false)] [int]$MaxAttempts = 5,
        [Parameter(Mandatory = $false)] [scriptblock]$OnItemError,
        [Parameter(Mandatory = $false)] [string]$Activity = 'Calling Microsoft Graph'
    )

    if (-not $GraphRequest) { $GraphRequest = New-GraphRequestExecutor }

    $results = @{}
    $pending = [System.Collections.Generic.List[object]]::new()
    foreach ($request in @($Requests)) { $pending.Add($request) | Out-Null }
    $totalCount = [Math]::Max($pending.Count, 1)
    $attempt = 1

    while ($pending.Count -gt 0 -and $attempt -le $MaxAttempts) {
        $retryQueue = [System.Collections.Generic.List[object]]::new()
        $delaySeconds = [Math]::Pow(2, $attempt)

        for ($offset = 0; $offset -lt $pending.Count; $offset += 20) {
            $chunk = @($pending[$offset..([Math]::Min($offset + 19, $pending.Count - 1))])
            Write-Progress -Id 1 -ParentId 0 -Activity $Activity -Status "attempt $attempt, $([Math]::Min($offset + 20, $pending.Count)) of $($pending.Count)" -PercentComplete ((($totalCount - $pending.Count + $offset) / $totalCount) * 100)
            $body = @{ requests = @($chunk | ForEach-Object { @{ id = [string]$_.Id; method = 'GET'; url = [string]$_.Url } }) }

            try {
                $response = & $GraphRequest 'POST' 'https://graph.microsoft.com/v1.0/$batch' $body
            }
            catch {
                foreach ($item in $chunk) { $retryQueue.Add($item) | Out-Null }
                continue
            }

            if (@($response.responses | Where-Object { $null -ne $_ }).Count -eq 0 -and $chunk.Count -gt 0) {
                foreach ($item in $chunk) { $retryQueue.Add($item) | Out-Null }
                continue
            }

            foreach ($item in @($response.responses)) {
                $itemId = [string]$item.id
                $status = [int]$item.status
                if ($status -ge 200 -and $status -lt 300) {
                    $results[$itemId] = $item.body
                }
                elseif ($status -in @(429, 502, 503, 504)) {
                    $retryQueue.Add(($chunk | Where-Object { [string]$_.Id -eq $itemId } | Select-Object -First 1)) | Out-Null
                    if ($item.headers -and $item.headers.'Retry-After') {
                        $headerDelay = 0.0
                        if ([double]::TryParse([string]$item.headers.'Retry-After', [ref]$headerDelay) -and $headerDelay -gt $delaySeconds) { $delaySeconds = $headerDelay }
                    }
                }
                else {
                    $message = if ($item.body -and $item.body.error -and $item.body.error.message) { [string]$item.body.error.message } else { "HTTP $status" }
                    if ($OnItemError) { & $OnItemError $itemId $message }
                }
            }
        }

        $pending = $retryQueue
        if ($pending.Count -gt 0 -and $attempt -lt $MaxAttempts) { Start-Sleep -Seconds $delaySeconds }
        $attempt++
    }

    Write-Progress -Id 1 -Activity $Activity -Completed
    foreach ($leftover in $pending) {
        if ($OnItemError) { & $OnItemError ([string]$leftover.Id) "Gave up after $MaxAttempts attempts (throttled or failed)" }
    }

    return $results
}

# FIX: memberOf is only defined on concrete Graph types (user, group, orgContact) - the
# directoryObject base type has no memberOf navigation in v1.0 $metadata, so
# /directoryObjects/{id}/memberOf fails with 400 for every recipient and the membership
# CSVs came back empty.
function Get-GraphMemberOfResourceSegment {
    param([Parameter(Mandatory = $true)] $Recipient)

    $typeDetails = [string]$Recipient.RecipientTypeDetails
    $type = [string]$Recipient.RecipientType

    if ($typeDetails -eq 'MailContact' -or $type -eq 'MailContact') {
        return 'contacts'
    }

    if ($typeDetails -eq 'GroupMailbox' -or $typeDetails -match 'Group$' -or $type -match 'Group') {
        return 'groups'
    }

    return 'users'
}

function Get-MembershipRows {
    param(
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] [object[]]$Recipients,
        [Parameter(Mandatory = $false)] [scriptblock]$GraphRequest
    )

    if (-not $GraphRequest) { $GraphRequest = New-GraphRequestExecutor }

    $entraRows = [System.Collections.Generic.List[object]]::new()
    $exchangeRows = [System.Collections.Generic.List[object]]::new()

    $recipientsById = @{}
    $requests = [System.Collections.Generic.List[object]]::new()
    foreach ($recipient in @($Recipients)) {
        $externalId = [string]$recipient.ExternalDirectoryObjectId
        if ([string]::IsNullOrWhiteSpace($externalId)) { continue }
        if ($recipientsById.ContainsKey($externalId)) { continue }
        $recipientsById[$externalId] = $recipient
        $segment = Get-GraphMemberOfResourceSegment -Recipient $recipient
        $requests.Add(@{
            Id  = $externalId
            Url = "/$segment/$externalId/memberOf?`$select=id,displayName,mail,mailEnabled,securityEnabled&`$top=999"
        }) | Out-Null
    }

    $onItemError = {
        param($id, $message)
        $identity = if ($recipientsById.ContainsKey($id)) { [string]$recipientsById[$id].Identity } else { [string]$id }
        Add-ExportError -Identity $identity -Stage 'GroupMemberships' -Operation 'Graph memberOf' -Message $message
    }

    $responses = Invoke-GraphBatch -Requests @($requests) -GraphRequest $GraphRequest -OnItemError $onItemError -Activity 'Collecting group memberships via Microsoft Graph'

    $responseKeys = @($responses.Keys)
    $membershipIndex = 0
    foreach ($externalId in $responseKeys) {
        $membershipIndex++
        if ($membershipIndex % 25 -eq 0 -or $membershipIndex -eq $responseKeys.Count) {
            Write-Progress -Id 1 -ParentId 0 -Activity 'Processing group membership results' -Status "$membershipIndex of $($responseKeys.Count)" -PercentComplete (($membershipIndex / [Math]::Max($responseKeys.Count, 1)) * 100)
        }
        $recipient = $recipientsById[$externalId]
        $body = $responses[$externalId]

        $groups = [System.Collections.Generic.List[object]]::new()
        while ($true) {
            foreach ($group in @($body.value)) {
                if ($null -ne $group) { $groups.Add($group) | Out-Null }
            }
            $nextLink = [string]$body.'@odata.nextLink'
            if ([string]::IsNullOrWhiteSpace($nextLink)) { break }
            try {
                $body = & $GraphRequest 'GET' $nextLink $null
            }
            catch {
                Add-ExportError -Identity ([string]$recipient.Identity) -Stage 'GroupMemberships' -Operation 'Graph memberOf paging' -Message $_.Exception.Message
                break
            }
        }

        foreach ($group in $groups) {
            $odataType = [string]$group.'@odata.type'
            if (-not [string]::IsNullOrWhiteSpace($odataType) -and $odataType -ne '#microsoft.graph.group') { continue }

            $groupMail = [string]$group.mail
            $entraRows.Add([pscustomobject]@{
                Identity             = [string]$recipient.Identity
                PrimarySmtpAddress   = [string]$recipient.PrimarySmtpAddress
                GroupId              = [string]$group.id
                GroupDisplayName     = [string]$group.displayName
                GroupMail            = $groupMail
                GroupSecurityEnabled = [string]$group.securityEnabled
            }) | Out-Null

            if ([bool]$group.mailEnabled) {
                $groupIdentity = [string]$group.displayName
                if (-not [string]::IsNullOrWhiteSpace($groupMail)) { $groupIdentity = "$groupIdentity ($groupMail)" }
                $exchangeRows.Add([pscustomobject]@{
                    Identity           = [string]$recipient.Identity
                    PrimarySmtpAddress = [string]$recipient.PrimarySmtpAddress
                    GroupIdentity      = $groupIdentity
                }) | Out-Null
            }
        }
    }

    Write-Progress -Id 1 -Completed
    return @{ Entra = @($entraRows); Exchange = @($exchangeRows) }
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
    # FIX: was `New-Object System.Collections.Generic.List[object]` - same PSObject-wrapping issue
    # as $script:Errors; `return @($recipients)` would throw if zero identities resolved.
    $recipients = [System.Collections.Generic.List[object]]::new()

    $lookupIndex = 0
    foreach ($inputIdentity in $identities) {
        $lookupIndex++
        Write-Progress -Id 1 -ParentId 0 -Activity 'Resolving input identities' -Status "$lookupIndex of $($identities.Count): $inputIdentity" -PercentComplete (($lookupIndex / [Math]::Max($identities.Count, 1)) * 100)
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

    $consolidateIndex = 0
    foreach ($object in $Objects) {
        $consolidateIndex++
        if ($consolidateIndex % 25 -eq 0 -or $consolidateIndex -eq $Objects.Count) {
            Write-Progress -Id 1 -ParentId 0 -Activity 'Consolidating object data' -Status "$consolidateIndex of $($Objects.Count)" -PercentComplete (($consolidateIndex / [Math]::Max($Objects.Count, 1)) * 100)
        }
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
    Write-Progress -Id 1 -Completed
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

function Export-ReportCsvs {
    param(
        [Parameter(Mandatory = $true)] [string]$Folder,
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] [object[]]$Objects,
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] [object[]]$ConsolidatedRows,
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] [object[]]$ProxyRows,
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] [object[]]$FullAccessRows,
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] [object[]]$SendAsRows,
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] [object[]]$ExchangeGroupRows,
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] [object[]]$EntraGroupRows,
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] [object[]]$ErrorRows
    )

    $exports = @(
        @{ Name = 'Objects-Consolidated.csv';     Rows = $ConsolidatedRows;  Headers = @('Identity', 'DisplayName', 'RecipientType', 'RecipientTypeDetails', 'MailboxType', 'PrimarySmtpAddress', 'Alias', 'ExternalDirectoryObjectId', 'ExchangeObjectId', 'ProxyAddresses', 'FullAccessTrustees', 'SendAsTrustees', 'ExchangeGroupMemberships', 'EntraGroupMemberships', 'HasErrors') }
        @{ Name = 'Objects.csv';                  Rows = $Objects;           Headers = @('Identity', 'DisplayName', 'RecipientType', 'RecipientTypeDetails', 'MailboxType', 'PrimarySmtpAddress', 'Alias', 'ExternalDirectoryObjectId', 'ExchangeObjectId', 'DistinguishedName') }
        @{ Name = 'ProxyAddresses.csv';           Rows = $ProxyRows;         Headers = @('Identity', 'PrimarySmtpAddress', 'AddressType', 'ProxyAddress', 'RawProxyAddress', 'IsPrimarySmtp') }
        @{ Name = 'FullAccess.csv';               Rows = $FullAccessRows;    Headers = @('Identity', 'PrimarySmtpAddress', 'Trustee', 'AccessRights', 'Deny', 'IsInherited') }
        @{ Name = 'SendAs.csv';                   Rows = $SendAsRows;        Headers = @('Identity', 'PrimarySmtpAddress', 'Trustee', 'AccessRights', 'IsInherited') }
        @{ Name = 'ExchangeGroupMemberships.csv'; Rows = $ExchangeGroupRows; Headers = @('Identity', 'PrimarySmtpAddress', 'GroupIdentity') }
        @{ Name = 'EntraGroupMemberships.csv';    Rows = $EntraGroupRows;    Headers = @('Identity', 'PrimarySmtpAddress', 'GroupId', 'GroupDisplayName', 'GroupMail', 'GroupSecurityEnabled') }
        @{ Name = 'Errors.csv';                   Rows = $ErrorRows;         Headers = @('Identity', 'Stage', 'Operation', 'Message') }
    )

    $exportIndex = 0
    foreach ($export in $exports) {
        $exportIndex++
        Write-Progress -Id 1 -ParentId 0 -Activity 'Writing CSV exports' -Status "$exportIndex of $($exports.Count): $($export.Name)" -PercentComplete (($exportIndex / $exports.Count) * 100)
        Export-CsvWithHeaders -Rows @($export.Rows) -Path (Join-Path $Folder $export.Name) -Headers $export.Headers
    }
    Write-Progress -Id 1 -Completed
}

function ConvertTo-JsonLiteral {
    param([Parameter(Mandatory = $false)] $Value)

    return ((ConvertTo-Json -InputObject @($Value) -Depth 10 -Compress) `
        -replace '&', '\u0026' `
        -replace '<', '\u003c' `
        -replace '>', '\u003e' `
        -replace [char]0x2028, '\u2028' `
        -replace [char]0x2029, '\u2029')
}

function Export-DashboardHtml {
    param(
        [Parameter(Mandatory = $true)] [string]$Folder,
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] [object[]]$ConsolidatedRows,
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] [object[]]$ProxyRows,
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] [object[]]$FullAccessRows,
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] [object[]]$SendAsRows,
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] [object[]]$ExchangeGroupRows,
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] [object[]]$EntraGroupRows
    )

    $objectsJson = ConvertTo-JsonLiteral -Value $ConsolidatedRows
    $proxyJson = ConvertTo-JsonLiteral -Value $ProxyRows
    $fullAccessJson = ConvertTo-JsonLiteral -Value $FullAccessRows
    $sendAsJson = ConvertTo-JsonLiteral -Value $SendAsRows
    $exchangeGroupsJson = ConvertTo-JsonLiteral -Value $ExchangeGroupRows
    $entraGroupsJson = ConvertTo-JsonLiteral -Value $EntraGroupRows

    $html = @"
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Exchange Object Data Dashboard</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 24px; color: #1f2933; background: #f7f9fb; }
    h1 { margin-bottom: 8px; }
    .warning { background: #fff3cd; border: 1px solid #ffe69c; padding: 12px; border-radius: 6px; margin: 12px 0; }
    .toolbar { display: flex; gap: 8px; flex-wrap: wrap; margin: 16px 0; }
    input[type="search"] { padding: 8px; min-width: 320px; }
    button { padding: 8px 10px; cursor: pointer; }
    table { border-collapse: collapse; width: 100%; background: white; }
    th, td { border: 1px solid #d9e2ec; padding: 8px; text-align: left; vertical-align: top; }
    th { background: #e6f0ff; position: sticky; top: 0; }
    details { margin-top: 6px; }
    .muted { color: #627d98; }
  </style>
</head>
<body>
  <h1>Exchange Object Data Dashboard</h1>
  <p class="muted">Search, select objects, and export selected data from this local static report.</p>
  <div class="warning">This dashboard may contain sensitive proxy addresses, mailbox permissions, group memberships, and object identifiers. Handle it as sensitive administrative data.</div>
  <div class="toolbar">
    <input id="search" type="search" placeholder="Search objects, proxies, permissions, groups">
    <button onclick="selectVisible(true)">Select visible</button>
    <button onclick="selectVisible(false)">Clear visible</button>
    <button onclick="exportSelected('objects')">Export selected objects</button>
    <button onclick="exportSelected('proxies')">Export selected proxies</button>
    <button onclick="exportSelected('fullaccess')">Export selected FullAccess</button>
    <button onclick="exportSelected('sendas')">Export selected SendAs</button>
    <button onclick="exportSelected('exchangegroups')">Export selected Exchange groups</button>
    <button onclick="exportSelected('entragroups')">Export selected Entra groups</button>
  </div>
  <div id="summary"></div>
  <table>
    <thead><tr><th>Select</th><th>Object</th><th>Type</th><th>Primary SMTP</th><th>Details</th></tr></thead>
    <tbody id="rows"></tbody>
  </table>
<script>
const objects = $objectsJson;
const proxies = $proxyJson;
const fullaccess = $fullAccessJson;
const sendas = $sendAsJson;
const exchangegroups = $exchangeGroupsJson;
const entragroups = $entraGroupsJson;
const selected = new Set();

function textForObject(o) {
  return Object.values(o).join(' ').toLowerCase();
}

function relatedRows(data, identity) {
  return data.filter(r => String(r.Identity) === String(identity));
}

function render() {
  const q = document.getElementById('search').value.toLowerCase();
  const body = document.getElementById('rows');
  body.innerHTML = '';
  const visible = objects.filter(o => textForObject(o).includes(q));
  document.getElementById('summary').textContent = visible.length + ' visible of ' + objects.length + ' objects';
  for (const o of visible) {
    const id = String(o.Identity);
    const tr = document.createElement('tr');
    tr.dataset.identity = id;
    tr.innerHTML = '<td><input type="checkbox" ' + (selected.has(id) ? 'checked' : '') + '></td>' +
      '<td><strong>' + escapeHtml(o.DisplayName || '') + '</strong><br><span class="muted">' + escapeHtml(id) + '</span></td>' +
      '<td>' + escapeHtml(o.RecipientTypeDetails || o.RecipientType || '') + '<br>' + escapeHtml(o.MailboxType || '') + '</td>' +
      '<td>' + escapeHtml(o.PrimarySmtpAddress || '') + '</td>' +
      '<td>' + detailsHtml(o) + '</td>';
    tr.querySelector('input').addEventListener('change', event => toggleSelected(event.target, id));
    body.appendChild(tr);
  }
}

function detailsHtml(o) {
  const id = String(o.Identity);
  return section('Proxies', relatedRows(proxies, id), 'RawProxyAddress') +
    section('FullAccess', relatedRows(fullaccess, id), 'Trustee') +
    section('SendAs', relatedRows(sendas, id), 'Trustee') +
    section('Exchange Groups', relatedRows(exchangegroups, id), 'GroupIdentity') +
    section('Entra Groups', relatedRows(entragroups, id), 'GroupDisplayName');
}

function section(title, rows, field) {
  const values = rows.map(r => r[field]).filter(Boolean).map(escapeHtml).join('<br>');
  return '<details><summary>' + title + ' (' + rows.length + ')</summary>' + (values || '<span class="muted">None</span>') + '</details>';
}

function toggleSelected(cb, id) {
  if (cb.checked) selected.add(id); else selected.delete(id);
}

function selectVisible(value) {
  document.querySelectorAll('#rows tr').forEach(tr => { if (value) selected.add(tr.dataset.identity); else selected.delete(tr.dataset.identity); });
  render();
}

function exportSelected(kind) {
  const ids = selected;
  let data = objects;
  if (kind === 'proxies') data = proxies;
  if (kind === 'fullaccess') data = fullaccess;
  if (kind === 'sendas') data = sendas;
  if (kind === 'exchangegroups') data = exchangegroups;
  if (kind === 'entragroups') data = entragroups;
  const rows = data.filter(r => ids.has(String(r.Identity)));
  downloadCsv(kind + '-selected.csv', rows);
}

function downloadCsv(filename, rows) {
  if (!rows.length) { alert('No selected rows to export.'); return; }
  const headers = Object.keys(rows[0]);
  const csv = [headers.join(',')].concat(rows.map(row => headers.map(h => csvCell(row[h])).join(','))).join('\n');
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = filename;
  link.click();
  URL.revokeObjectURL(link.href);
}

function csvCell(value) {
  const s = value == null ? '' : String(value);
  return '"' + s.replace(/"/g, '""') + '"';
}

function escapeHtml(value) {
  return String(value == null ? '' : value).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

document.getElementById('search').addEventListener('input', render);
render();
</script>
</body>
</html>
"@

    Set-Content -Path (Join-Path $Folder 'Dashboard.html') -Value $html -Encoding UTF8
}

$resolvedOutputFolder = Initialize-OutputFolder -OutputFolder $OutputFolder
$isMultiObjectRun = $PSCmdlet.ParameterSetName -in @('All', 'Csv')

Write-Host "Output folder: $resolvedOutputFolder" -ForegroundColor Cyan
Write-Host "Recipient type: $RecipientType" -ForegroundColor Cyan

$graphConnected = Connect-GraphIfNeeded

Import-Module ExchangeOnlineManagement -ErrorAction Stop
Connect-ExchangeOnline -ShowBanner:$false

Write-PhaseProgress -Phase 1 -Name 'Loading recipient inventory'
$targetRecipients = @(Get-TargetRecipients)

$objectList = [System.Collections.Generic.List[object]]::new()
$objectIndex = 0
foreach ($recipient in $targetRecipients) {
    $objectIndex++
    if ($objectIndex % 50 -eq 0 -or $objectIndex -eq $targetRecipients.Count) {
        Write-Progress -Id 1 -ParentId 0 -Activity 'Building object inventory' -Status "$objectIndex of $($targetRecipients.Count)" -PercentComplete (($objectIndex / [Math]::Max($targetRecipients.Count, 1)) * 100)
    }
    $objectList.Add((ConvertTo-ExportObject -Recipient $recipient)) | Out-Null
}
$objects = @($objectList)
Write-Progress -Id 1 -Completed

Write-Host "Objects selected: $($objects.Count)" -ForegroundColor Green

$proxyRows = [System.Collections.Generic.List[object]]::new()
$fullAccessRows = [System.Collections.Generic.List[object]]::new()
$sendAsRows = [System.Collections.Generic.List[object]]::new()
$exchangeGroupMembershipRows = [System.Collections.Generic.List[object]]::new()
$entraGroupMembershipRows = [System.Collections.Generic.List[object]]::new()

# Phase 2: proxy addresses (in-memory, from the Phase 1 inventory)
Write-Host 'Deriving proxy address rows from inventory...' -ForegroundColor Cyan
Write-PhaseProgress -Phase 2 -Name 'Deriving proxy addresses'
$proxyRecipientIndex = 0
foreach ($recipient in $targetRecipients) {
    $proxyRecipientIndex++
    if ($proxyRecipientIndex % 50 -eq 0 -or $proxyRecipientIndex -eq $targetRecipients.Count) {
        Write-Progress -Id 1 -ParentId 0 -Activity 'Deriving proxy addresses' -Status "$proxyRecipientIndex of $($targetRecipients.Count)" -PercentComplete (($proxyRecipientIndex / [Math]::Max($targetRecipients.Count, 1)) * 100)
    }
    foreach ($row in @(ConvertTo-ProxyRows -Recipient $recipient)) { $proxyRows.Add($row) | Out-Null }
}
Write-Progress -Id 1 -Completed

# Phase 3: group memberships (Microsoft Graph $batch)
Write-PhaseProgress -Phase 3 -Name 'Collecting group memberships (Graph)'
if ($graphConnected) {
    Write-Host 'Collecting group memberships via Microsoft Graph...' -ForegroundColor Cyan
    $membershipRows = Get-MembershipRows -Recipients $targetRecipients
    foreach ($row in @($membershipRows.Exchange)) { $exchangeGroupMembershipRows.Add($row) | Out-Null }
    foreach ($row in @($membershipRows.Entra)) { $entraGroupMembershipRows.Add($row) | Out-Null }
}
else {
    Write-Host 'Microsoft Graph not connected: membership CSVs will contain headers only.' -ForegroundColor Yellow
}

# Phase 4a: SendAs permissions (org-wide sweep on -All)
Write-Host 'Collecting SendAs permissions...' -ForegroundColor Cyan
Write-PhaseProgress -Phase 4 -Name 'Collecting SendAs permissions'
foreach ($row in @(Get-SendAsRows -Recipients $targetRecipients -UseOrgWideQuery ($PSCmdlet.ParameterSetName -eq 'All'))) { $sendAsRows.Add($row) | Out-Null }

# Phase 4b: FullAccess permissions (per mailbox; no bulk API exists)
Write-Host 'Collecting FullAccess permissions...' -ForegroundColor Cyan
Write-PhaseProgress -Phase 5 -Name 'Collecting FullAccess permissions'
$mailboxRecipients = @($targetRecipients | Where-Object { -not [string]::IsNullOrWhiteSpace((Get-MailboxType -Recipient $_)) })
$index = 0
foreach ($recipient in $mailboxRecipients) {
    $index++
    Write-Progress -Id 1 -ParentId 0 -Activity 'Collecting FullAccess permissions' -Status "$index of $($mailboxRecipients.Count): $($recipient.Identity)" -PercentComplete (($index / [Math]::Max($mailboxRecipients.Count, 1)) * 100)
    foreach ($row in @(Get-FullAccessRows -Recipient $recipient)) { $fullAccessRows.Add($row) | Out-Null }
}
Write-Progress -Id 1 -Activity 'Collecting FullAccess permissions' -Completed

Write-Host 'Consolidating and exporting...' -ForegroundColor Cyan
Write-PhaseProgress -Phase 6 -Name 'Consolidating object data'
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

Write-PhaseProgress -Phase 7 -Name 'Writing CSV exports and dashboard'
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

if ($isMultiObjectRun -and -not $NoDashboard) {
    try {
        Write-Progress -Id 1 -ParentId 0 -Activity 'Writing dashboard' -Status 'Dashboard.html'
        Export-DashboardHtml `
            -Folder $resolvedOutputFolder `
            -ConsolidatedRows $consolidatedRows `
            -ProxyRows @($proxyRows) `
            -FullAccessRows @($fullAccessRows) `
            -SendAsRows @($sendAsRows) `
            -ExchangeGroupRows @($exchangeGroupMembershipRows) `
            -EntraGroupRows @($entraGroupMembershipRows)
        Write-Host "Dashboard written to: $(Join-Path $resolvedOutputFolder 'Dashboard.html')" -ForegroundColor Green
    }
    catch {
        Add-ExportError -Identity '' -Stage 'Dashboard' -Operation 'Export-DashboardHtml' -Message $_.Exception.Message
        @($script:Errors) | Export-Csv -Path (Join-Path $resolvedOutputFolder 'Errors.csv') -NoTypeInformation -Encoding UTF8
    }
}

Write-Progress -Id 1 -Completed
Write-Progress -Id 0 -Activity 'Export-ExchangeObjectData' -Completed

Write-Host "Objects exported: $($objects.Count)" -ForegroundColor Green
Write-Host "Errors logged: $($script:Errors.Count)" -ForegroundColor Yellow

try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch { }
if ($graphConnected) {
    try { Disconnect-MgGraph | Out-Null } catch { }
}