<#
.SYNOPSIS
    Inventories tenant resources that depend on a specific SMTP domain.

.DESCRIPTION
    Finds distribution lists, mail-enabled security groups, Microsoft 365 groups,
    dynamic distribution groups, shared mailboxes, room/equipment mailboxes, and
    address rewrite entries that reference the target domain through primary SMTP
    addresses, aliases/proxies, membership, sender restrictions, moderation,
    forwarding, routing, or rewrite behavior.

    The script is read-only. It writes CSV evidence for migration planning,
    reconfiguration, owner decisions, exclusions, downstream testing, and comms.

.PARAMETER Domain
    SMTP domain to inventory. Mandatory.

.PARAMETER OutputFolder
    Destination folder. Defaults to .\TenantDomainDependencyInventory_<timestamp>.

.PARAMETER IncludeAll
    Include all scoped objects, not just objects with a target-domain dependency.

.PARAMETER SkipMemberExpansion
    Do not expand group membership. Useful for very large tenants.

.PARAMETER SkipAddressRewriteRules
    Do not attempt Get-AddressRewriteEntry.

.EXAMPLE
    .\Export-TenantDomainDependencyInventory.ps1 -Domain contoso.com

.EXAMPLE
    .\Export-TenantDomainDependencyInventory.ps1 -Domain contoso.com -OutputFolder .\ContosoDependencyInventory
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Domain,
    [string]$OutputFolder = ".\TenantDomainDependencyInventory_$(Get-Date -Format 'yyyyMMdd_HHmmss')",
    [switch]$IncludeAll,
    [switch]$SkipMemberExpansion,
    [switch]$SkipAddressRewriteRules
)

$ErrorActionPreference = 'Stop'
$script:Errors = [System.Collections.Generic.List[object]]::new()

function Add-ExportError {
    param(
        [Parameter(Mandatory = $false)] [string]$Identity,
        [Parameter(Mandatory = $true)] [string]$Stage,
        [Parameter(Mandatory = $true)] [string]$Operation,
        [Parameter(Mandatory = $true)] [string]$Message,
        [switch]$NoConsoleWarning
    )

    $script:Errors.Add([pscustomobject]@{
        Identity  = $Identity
        Stage     = $Stage
        Operation = $Operation
        Message   = $Message
    }) | Out-Null

    if (-not $NoConsoleWarning) {
        Write-Warning "[$Stage/$Operation] $Identity $Message"
    }
}

function Add-ExportWarnings {
    param(
        [Parameter(Mandatory = $false)] [string]$Identity,
        [Parameter(Mandatory = $true)] [string]$Stage,
        [Parameter(Mandatory = $true)] [string]$Operation,
        [Parameter(Mandatory = $false)] [AllowEmptyCollection()] [object[]]$Warnings
    )

    foreach ($warning in @($Warnings)) {
        $message = ([string]$warning).Trim()
        if ([string]::IsNullOrWhiteSpace($message)) { continue }
        Add-ExportError -Identity $Identity -Stage $Stage -Operation $Operation -Message $message -NoConsoleWarning
    }
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

function Connect-GraphForGroupFallbackIfNeeded {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    $context = Get-MgContext -ErrorAction SilentlyContinue
    if ($context) { return }

    Connect-MgGraph -Scopes @('Group.Read.All', 'GroupMember.Read.All', 'User.Read.All') -NoWelcome
}

function New-GraphRequestExecutor {
    return {
        param(
            [Parameter(Mandatory = $true)] [string]$Method,
            [Parameter(Mandatory = $true)] [string]$Uri,
            [Parameter(Mandatory = $false)] $Body
        )

        Connect-GraphForGroupFallbackIfNeeded
        if ($null -ne $Body) {
            Invoke-MgGraphRequest -Method $Method -Uri $Uri -Body ($Body | ConvertTo-Json -Depth 10) -ContentType 'application/json'
        }
        else {
            Invoke-MgGraphRequest -Method $Method -Uri $Uri
        }
    }
}

function Normalize-Domain {
    param([Parameter(Mandatory = $true)] [string]$Domain)
    return $Domain.Trim().TrimStart('@').ToLowerInvariant()
}

function Join-ReportValues {
    param([AllowEmptyCollection()] [object[]]$Values)

    $cleanValues = @($Values | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | ForEach-Object { ([string]$_).Trim() } | Sort-Object -Unique)
    return ($cleanValues -join '; ')
}

function Get-ObjectPropertyValue {
    param(
        [Parameter(Mandatory = $true)] $InputObject,
        [Parameter(Mandatory = $true)] [string]$PropertyName
    )

    if ($InputObject -is [hashtable] -and $InputObject.ContainsKey($PropertyName)) { return $InputObject[$PropertyName] }
    $property = $InputObject.PSObject.Properties[$PropertyName]
    if ($property) { return $property.Value }
    return $null
}

function Get-GraphCollectionValues {
    param(
        [Parameter(Mandatory = $true)] [string]$Uri,
        [Parameter(Mandatory = $false)] [scriptblock]$GraphRequest
    )

    if (-not $GraphRequest) { $GraphRequest = New-GraphRequestExecutor }

    $values = [System.Collections.Generic.List[object]]::new()
    $nextUri = $Uri
    while (-not [string]::IsNullOrWhiteSpace($nextUri)) {
        $response = & $GraphRequest 'GET' $nextUri $null
        foreach ($item in @(Get-ObjectPropertyValue -InputObject $response -PropertyName 'value')) {
            $values.Add($item) | Out-Null
        }
        $nextUri = [string](Get-ObjectPropertyValue -InputObject $response -PropertyName '@odata.nextLink')
    }

    return @($values)
}

function ConvertFrom-GraphDirectoryObjectLink {
    param([Parameter(Mandatory = $true)] $DirectoryObject)

    $mail = [string](Get-ObjectPropertyValue -InputObject $DirectoryObject -PropertyName 'mail')
    $upn = [string](Get-ObjectPropertyValue -InputObject $DirectoryObject -PropertyName 'userPrincipalName')
    $displayName = [string](Get-ObjectPropertyValue -InputObject $DirectoryObject -PropertyName 'displayName')
    $address = if (-not [string]::IsNullOrWhiteSpace($mail)) { $mail } else { $upn }

    return [pscustomobject]@{
        PrimarySmtpAddress  = $address
        WindowsEmailAddress = $upn
        ExternalEmailAddress = ''
        UserPrincipalName   = $upn
        Mail                = $mail
        DisplayName         = $displayName
        Name                = $displayName
    }
}

function ConvertTo-StringValues {
    param([AllowNull()] $Value)

    if ($null -eq $Value) { return @() }
    if ($Value -is [string]) { return @($Value) }
    if ($Value -is [System.Collections.IEnumerable]) {
        return @($Value | ForEach-Object { if ($null -ne $_) { [string]$_ } })
    }
    return @([string]$Value)
}

function Get-AddressValue {
    param([AllowNull()] $Address)

    if ($null -eq $Address) { return '' }
    $value = ([string]$Address).Trim()
    $parts = $value -split ':', 2
    if ($parts.Count -eq 2) { return $parts[1] }
    return $value
}

function Test-ValueUsesDomain {
    param(
        [AllowNull()] $Value,
        [Parameter(Mandatory = $true)] [string]$Domain
    )

    $normalizedDomain = Normalize-Domain -Domain $Domain
    foreach ($item in @(ConvertTo-StringValues -Value $Value)) {
        $addressValue = (Get-AddressValue -Address $item).ToLowerInvariant()
        if ($addressValue -match "@([a-z0-9._-]+)$") {
            $addressDomain = $Matches[1]
            if ($addressDomain -eq $normalizedDomain -or $addressDomain.EndsWith(".$normalizedDomain")) { return $true }
        }
        elseif ($addressValue -match "(^|[^a-z0-9._-])$([regex]::Escape($normalizedDomain))($|[^a-z0-9._-])") {
            return $true
        }
    }
    return $false
}

function Get-TargetDomainValues {
    param(
        [AllowNull()] $Values,
        [Parameter(Mandatory = $true)] [string]$Domain
    )

    return @(ConvertTo-StringValues -Value $Values | Where-Object { Test-ValueUsesDomain -Value $_ -Domain $Domain })
}

function Add-Reason {
    param(
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] [System.Collections.Generic.List[string]]$Reasons,
        [Parameter(Mandatory = $true)] [string]$Reason,
        [Parameter(Mandatory = $true)] [bool]$Condition
    )

    if ($Condition -and -not $Reasons.Contains($Reason)) { $Reasons.Add($Reason) | Out-Null }
}

function Get-ReviewDisposition {
    param(
        [Parameter(Mandatory = $true)] [string]$ObjectType,
        [Parameter(Mandatory = $false)] [AllowEmptyCollection()] [string[]]$Reasons = @()
    )

    $actions = [System.Collections.Generic.List[string]]::new()
    if ($Reasons -contains 'TargetDomainPrimarySmtp' -or $Reasons -contains 'TargetDomainProxy') { $actions.Add('Migration') | Out-Null }
    if ($Reasons -contains 'TargetDomainForwarding' -or $Reasons -contains 'TargetDomainRouting' -or $Reasons -contains 'TargetDomainAddressRewriteRule') { $actions.Add('Reconfiguration') | Out-Null }
    if ($Reasons -contains 'TargetDomainMember' -or $Reasons -contains 'TargetDomainOwner' -or $Reasons -contains 'TargetDomainSenderRestriction' -or $Reasons -contains 'ModerationEnabled') { $actions.Add('OwnerDecision') | Out-Null }
    if ($ObjectType -match 'RoomMailbox|EquipmentMailbox') { $actions.Add('OwnerDecision') | Out-Null }
    if ($actions.Count -eq 0) { $actions.Add('ExclusionCandidate') | Out-Null }
    return Join-ReportValues -Values $actions
}

function Get-TestingAndCommsNeeds {
    param(
        [Parameter(Mandatory = $true)] [string]$ObjectType,
        [Parameter(Mandatory = $false)] [AllowEmptyCollection()] [string[]]$Reasons = @(),
        [Parameter(Mandatory = $false)] [bool]$AllowsExternalSenders
    )

    $needs = [System.Collections.Generic.List[string]]::new()
    if ($Reasons -contains 'TargetDomainPrimarySmtp' -or $Reasons -contains 'TargetDomainProxy') { $needs.Add('Test inbound delivery to target-domain address and aliases') | Out-Null }
    if ($Reasons -contains 'TargetDomainMember') { $needs.Add('Validate group expansion after member migration') | Out-Null }
    if ($Reasons -contains 'TargetDomainForwarding' -or $Reasons -contains 'TargetDomainRouting') { $needs.Add('Test forwarding and routing path') | Out-Null }
    if ($Reasons -contains 'TargetDomainAddressRewriteRule') { $needs.Add('Test rewritten sender/recipient address flow') | Out-Null }
    if ($Reasons -contains 'ModerationEnabled') { $needs.Add('Validate moderator approval workflow') | Out-Null }
    if ($AllowsExternalSenders) { $needs.Add('Test external sender delivery') | Out-Null }
    if ($ObjectType -match 'RoomMailbox|EquipmentMailbox') { $needs.Add('Test booking, delegates, and resource policy') | Out-Null }
    if ($Reasons.Count -gt 0) { $needs.Add('Communicate impact and decision request to owner') | Out-Null }
    return Join-ReportValues -Values $needs
}

function Get-RecipientObjectType {
    param([Parameter(Mandatory = $true)] $Recipient)

    $details = [string](Get-ObjectPropertyValue -InputObject $Recipient -PropertyName 'RecipientTypeDetails')
    $type = [string](Get-ObjectPropertyValue -InputObject $Recipient -PropertyName 'RecipientType')
    if ($details -eq 'MailUniversalSecurityGroup') { return 'MailEnabledSecurityGroup' }
    if ($details -match 'DistributionGroup|RoomList' -or $type -match 'DistributionGroup') { return 'DistributionList' }
    if ($details -eq 'DynamicDistributionGroup') { return 'DynamicDistributionGroup' }
    if ($details -eq 'GroupMailbox') { return 'Microsoft365Group' }
    if ($details -in @('SharedMailbox','RoomMailbox','EquipmentMailbox')) { return $details }
    return $details
}

function ConvertTo-DependencyRow {
    param(
        [Parameter(Mandatory = $true)] $Object,
        [Parameter(Mandatory = $true)] [string]$ObjectType,
        [Parameter(Mandatory = $true)] [string]$Domain,
        [Parameter(Mandatory = $false)] [object[]]$Members = @(),
        [Parameter(Mandatory = $false)] [object[]]$Owners = @(),
        [Parameter(Mandatory = $false)] [string]$MembershipSource = '',
        [Parameter(Mandatory = $false)] [object[]]$FullAccessTrustees = @(),
        [Parameter(Mandatory = $false)] [object[]]$SendAsTrustees = @(),
        [Parameter(Mandatory = $false)] $CalendarProcessing
    )

    $identity = [string](Get-ObjectPropertyValue -InputObject $Object -PropertyName 'Identity')
    $displayName = [string](Get-ObjectPropertyValue -InputObject $Object -PropertyName 'DisplayName')
    $primarySmtp = [string](Get-ObjectPropertyValue -InputObject $Object -PropertyName 'PrimarySmtpAddress')
    $emailAddresses = @(ConvertTo-StringValues -Value (Get-ObjectPropertyValue -InputObject $Object -PropertyName 'EmailAddresses'))
    $targetProxies = @(Get-TargetDomainValues -Values $emailAddresses -Domain $Domain)
    $ownerValues = @(
        $Owners | ForEach-Object {
            if ($_ -is [string]) { $_ }
            else { Join-ReportValues -Values @($_.PrimarySmtpAddress, $_.WindowsEmailAddress, $_.ExternalEmailAddress, $_.UserPrincipalName, $_.Mail, $_.DisplayName, $_.Name) }
        }
        Get-ObjectPropertyValue -InputObject $Object -PropertyName 'ManagedBy'
    )
    $memberValues = @($Members | ForEach-Object { Join-ReportValues -Values @($_.PrimarySmtpAddress, $_.WindowsEmailAddress, $_.ExternalEmailAddress, $_.UserPrincipalName, $_.Mail, $_.DisplayName, $_.Name) })

    $forwardingValues = @(
        Get-ObjectPropertyValue -InputObject $Object -PropertyName 'ForwardingAddress'
        Get-ObjectPropertyValue -InputObject $Object -PropertyName 'ForwardingSmtpAddress'
        Get-ObjectPropertyValue -InputObject $Object -PropertyName 'DeliverToMailboxAndForward'
        Get-ObjectPropertyValue -InputObject $Object -PropertyName 'TargetAddress'
    )

    $senderRestrictionValues = @(
        Get-ObjectPropertyValue -InputObject $Object -PropertyName 'AcceptMessagesOnlyFrom'
        Get-ObjectPropertyValue -InputObject $Object -PropertyName 'AcceptMessagesOnlyFromDLMembers'
        Get-ObjectPropertyValue -InputObject $Object -PropertyName 'AcceptMessagesOnlyFromSendersOrMembers'
        Get-ObjectPropertyValue -InputObject $Object -PropertyName 'RejectMessagesFrom'
        Get-ObjectPropertyValue -InputObject $Object -PropertyName 'RejectMessagesFromDLMembers'
        Get-ObjectPropertyValue -InputObject $Object -PropertyName 'RejectMessagesFromSendersOrMembers'
    )

    $moderationValues = @(
        Get-ObjectPropertyValue -InputObject $Object -PropertyName 'ModeratedBy'
        Get-ObjectPropertyValue -InputObject $Object -PropertyName 'BypassModerationFromSendersOrMembers'
    )

    $resourceDelegateValues = @()
    $processExternalMeetings = ''
    if ($null -ne $CalendarProcessing) {
        $resourceDelegateValues = @(Get-ObjectPropertyValue -InputObject $CalendarProcessing -PropertyName 'ResourceDelegates')
        $processExternalMeetings = [string](Get-ObjectPropertyValue -InputObject $CalendarProcessing -PropertyName 'ProcessExternalMeetingMessages')
    }

    $reasons = [System.Collections.Generic.List[string]]::new()
    Add-Reason -Reasons $reasons -Reason 'TargetDomainPrimarySmtp' -Condition (Test-ValueUsesDomain -Value $primarySmtp -Domain $Domain)
    Add-Reason -Reasons $reasons -Reason 'TargetDomainProxy' -Condition ($targetProxies.Count -gt 0)
    Add-Reason -Reasons $reasons -Reason 'TargetDomainOwner' -Condition (Test-ValueUsesDomain -Value $ownerValues -Domain $Domain)
    Add-Reason -Reasons $reasons -Reason 'TargetDomainMember' -Condition (Test-ValueUsesDomain -Value $memberValues -Domain $Domain)
    Add-Reason -Reasons $reasons -Reason 'TargetDomainForwarding' -Condition (Test-ValueUsesDomain -Value $forwardingValues -Domain $Domain)
    Add-Reason -Reasons $reasons -Reason 'TargetDomainSenderRestriction' -Condition (Test-ValueUsesDomain -Value $senderRestrictionValues -Domain $Domain)
    Add-Reason -Reasons $reasons -Reason 'TargetDomainModeration' -Condition (Test-ValueUsesDomain -Value $moderationValues -Domain $Domain)
    Add-Reason -Reasons $reasons -Reason 'TargetDomainResourceDelegate' -Condition (Test-ValueUsesDomain -Value $resourceDelegateValues -Domain $Domain)

    $moderationEnabled = [bool](Get-ObjectPropertyValue -InputObject $Object -PropertyName 'ModerationEnabled')
    Add-Reason -Reasons $reasons -Reason 'ModerationEnabled' -Condition $moderationEnabled

    $requireAuth = Get-ObjectPropertyValue -InputObject $Object -PropertyName 'RequireSenderAuthenticationEnabled'
    $allowsExternalSenders = ($null -ne $requireAuth -and -not [bool]$requireAuth)
    $reasonArray = @($reasons.ToArray())
    $recommendedDisposition = if ($reasonArray.Count -gt 0) { Get-ReviewDisposition -ObjectType $ObjectType -Reasons $reasonArray } else { Get-ReviewDisposition -ObjectType $ObjectType }
    $testingAndCommsNeeds = if ($reasonArray.Count -gt 0) { Get-TestingAndCommsNeeds -ObjectType $ObjectType -Reasons $reasonArray -AllowsExternalSenders $allowsExternalSenders } else { Get-TestingAndCommsNeeds -ObjectType $ObjectType -AllowsExternalSenders $allowsExternalSenders }

    return [pscustomobject]@{
        ObjectType                       = $ObjectType
        Identity                         = $identity
        DisplayName                      = $displayName
        RecipientTypeDetails             = [string](Get-ObjectPropertyValue -InputObject $Object -PropertyName 'RecipientTypeDetails')
        PrimarySmtpAddress               = $primarySmtp
        Alias                            = [string](Get-ObjectPropertyValue -InputObject $Object -PropertyName 'Alias')
        ExternalDirectoryObjectId        = [string](Get-ObjectPropertyValue -InputObject $Object -PropertyName 'ExternalDirectoryObjectId')
        TargetDomainAddresses            = Join-ReportValues -Values @($targetProxies + $primarySmtp)
        Owners                           = Join-ReportValues -Values $ownerValues
        MembershipSource                 = $MembershipSource
        TargetDomainMemberCount          = @(Get-TargetDomainValues -Values $memberValues -Domain $Domain).Count
        TargetDomainMemberSample         = Join-ReportValues -Values (@(Get-TargetDomainValues -Values $memberValues -Domain $Domain) | Select-Object -First 20)
        AllowsExternalSenders            = $allowsExternalSenders
        AcceptMessagesOnlyFrom           = Join-ReportValues -Values (Get-ObjectPropertyValue -InputObject $Object -PropertyName 'AcceptMessagesOnlyFromSendersOrMembers')
        RejectMessagesFrom               = Join-ReportValues -Values (Get-ObjectPropertyValue -InputObject $Object -PropertyName 'RejectMessagesFromSendersOrMembers')
        ModerationEnabled                = $moderationEnabled
        ModeratedBy                      = Join-ReportValues -Values (Get-ObjectPropertyValue -InputObject $Object -PropertyName 'ModeratedBy')
        BypassModeration                 = Join-ReportValues -Values (Get-ObjectPropertyValue -InputObject $Object -PropertyName 'BypassModerationFromSendersOrMembers')
        ForwardingOrRoutingTargets       = Join-ReportValues -Values $forwardingValues
        FullAccessTrustees               = Join-ReportValues -Values $FullAccessTrustees
        SendAsTrustees                   = Join-ReportValues -Values $SendAsTrustees
        SendOnBehalfTo                   = Join-ReportValues -Values (Get-ObjectPropertyValue -InputObject $Object -PropertyName 'GrantSendOnBehalfTo')
        ResourceDelegates                = Join-ReportValues -Values $resourceDelegateValues
        ProcessExternalMeetingMessages   = $processExternalMeetings
        DependencyReasons                = Join-ReportValues -Values $reasonArray
        RecommendedDisposition           = $recommendedDisposition
        TestingAndCommsNeeds             = $testingAndCommsNeeds
    }
}

function Get-DistributionGroupMembersSafe {
    param([Parameter(Mandatory = $true)] $Group)

    if ($SkipMemberExpansion) { return @() }
    $warnings = @()
    try {
        $members = @(Get-DistributionGroupMember -Identity $Group.Identity -ResultSize Unlimited -ErrorAction Stop -WarningAction SilentlyContinue -WarningVariable warnings)
        Add-ExportWarnings -Identity ([string]$Group.Identity) -Stage 'Membership' -Operation 'Get-DistributionGroupMember' -Warnings $warnings
        return $members
    }
    catch {
        Add-ExportWarnings -Identity ([string]$Group.Identity) -Stage 'Membership' -Operation 'Get-DistributionGroupMember' -Warnings $warnings
        Add-ExportError -Identity ([string]$Group.Identity) -Stage 'Membership' -Operation 'Get-DistributionGroupMember' -Message $_.Exception.Message
        return @()
    }
}

function Get-UnifiedGroupLinksSafe {
    param(
        [Parameter(Mandatory = $true)] $Group,
        [Parameter(Mandatory = $true)] [ValidateSet('Members','Owners')] [string]$LinkType
    )

    if ($SkipMemberExpansion -and $LinkType -eq 'Members') { return @() }
    $warnings = @()
    try {
        $links = @(Get-UnifiedGroupLinks -Identity $Group.Identity -LinkType $LinkType -ResultSize Unlimited -ErrorAction Stop -WarningAction SilentlyContinue -WarningVariable warnings)
        Add-ExportWarnings -Identity ([string]$Group.Identity) -Stage 'Membership' -Operation "Get-UnifiedGroupLinks:$LinkType" -Warnings $warnings
        return $links
    }
    catch {
        Add-ExportWarnings -Identity ([string]$Group.Identity) -Stage 'Membership' -Operation "Get-UnifiedGroupLinks:$LinkType" -Warnings $warnings
        $exchangeErrorMessage = $_.Exception.Message
        try {
            return @(Get-GraphUnifiedGroupLinksSafe -Group $Group -LinkType $LinkType)
        }
        catch {
            Add-ExportError -Identity ([string]$Group.Identity) -Stage 'Membership' -Operation "Get-UnifiedGroupLinks:$LinkType" -Message $exchangeErrorMessage
            Add-ExportError -Identity ([string]$Group.Identity) -Stage 'Membership' -Operation "GraphGroupLinks:$LinkType" -Message $_.Exception.Message
        }
        return @()
    }
}

function Get-GraphUnifiedGroupLinksSafe {
    param(
        [Parameter(Mandatory = $true)] $Group,
        [Parameter(Mandatory = $true)] [ValidateSet('Members','Owners')] [string]$LinkType,
        [Parameter(Mandatory = $false)] [scriptblock]$GraphRequest
    )

    $groupId = [string](Get-ObjectPropertyValue -InputObject $Group -PropertyName 'ExternalDirectoryObjectId')
    if ([string]::IsNullOrWhiteSpace($groupId)) {
        throw 'Group ExternalDirectoryObjectId is blank; cannot use Microsoft Graph fallback.'
    }

    $segment = if ($LinkType -eq 'Owners') { 'owners' } else { 'members' }
    $uri = "https://graph.microsoft.com/v1.0/groups/$groupId/${segment}?`$select=id,displayName,mail,userPrincipalName&`$top=999"
    return @(Get-GraphCollectionValues -Uri $uri -GraphRequest $GraphRequest | ForEach-Object { ConvertFrom-GraphDirectoryObjectLink -DirectoryObject $_ })
}

function Get-MailboxPermissionValuesSafe {
    param(
        [Parameter(Mandatory = $true)] $Mailbox,
        [Parameter(Mandatory = $true)] [ValidateSet('FullAccess','SendAs')] [string]$PermissionType
    )

    $warnings = @()
    try {
        if ($PermissionType -eq 'FullAccess') {
            $permissions = @(Get-MailboxPermission -Identity $Mailbox.Identity -ErrorAction Stop -WarningAction SilentlyContinue -WarningVariable warnings)
            Add-ExportWarnings -Identity ([string]$Mailbox.Identity) -Stage 'Permissions' -Operation "Get-$PermissionType" -Warnings $warnings
            return @($permissions | Where-Object { -not $_.IsInherited -and -not $_.Deny -and [string]$_.User -ne 'NT AUTHORITY\SELF' } | ForEach-Object { [string]$_.User })
        }
        $permissions = @(Get-RecipientPermission -Identity $Mailbox.Identity -ErrorAction Stop -WarningAction SilentlyContinue -WarningVariable warnings)
        Add-ExportWarnings -Identity ([string]$Mailbox.Identity) -Stage 'Permissions' -Operation "Get-$PermissionType" -Warnings $warnings
        return @($permissions | Where-Object { -not $_.IsInherited -and -not $_.Deny -and [string]$_.Trustee -ne 'NT AUTHORITY\SELF' } | ForEach-Object { [string]$_.Trustee })
    }
    catch {
        Add-ExportWarnings -Identity ([string]$Mailbox.Identity) -Stage 'Permissions' -Operation "Get-$PermissionType" -Warnings $warnings
        Add-ExportError -Identity ([string]$Mailbox.Identity) -Stage 'Permissions' -Operation "Get-$PermissionType" -Message $_.Exception.Message
        return @()
    }
}

function Get-CalendarProcessingSafe {
    param([Parameter(Mandatory = $true)] $Mailbox)

    if ([string](Get-ObjectPropertyValue -InputObject $Mailbox -PropertyName 'RecipientTypeDetails') -notin @('RoomMailbox','EquipmentMailbox')) { return $null }
    $warnings = @()
    try {
        $calendarProcessing = Get-CalendarProcessing -Identity $Mailbox.Identity -ErrorAction Stop -WarningAction SilentlyContinue -WarningVariable warnings
        Add-ExportWarnings -Identity ([string]$Mailbox.Identity) -Stage 'ResourceMailbox' -Operation 'Get-CalendarProcessing' -Warnings $warnings
        return $calendarProcessing
    }
    catch {
        Add-ExportWarnings -Identity ([string]$Mailbox.Identity) -Stage 'ResourceMailbox' -Operation 'Get-CalendarProcessing' -Warnings $warnings
        Add-ExportError -Identity ([string]$Mailbox.Identity) -Stage 'ResourceMailbox' -Operation 'Get-CalendarProcessing' -Message $_.Exception.Message
        return $null
    }
}

function Get-AddressRewriteRows {
    param([Parameter(Mandatory = $true)] [string]$Domain)

    if ($SkipAddressRewriteRules) { return @() }
    if (-not (Get-Command Get-AddressRewriteEntry -ErrorAction SilentlyContinue)) {
        Add-ExportError -Identity $Domain -Stage 'AddressRewrite' -Operation 'Get-AddressRewriteEntry' -Message 'Get-AddressRewriteEntry is not available in this Exchange session.'
        return @()
    }

    $entries = @()
    $warnings = @()
    try {
        $entries = @(Get-AddressRewriteEntry -ErrorAction Stop -WarningAction SilentlyContinue -WarningVariable warnings)
        Add-ExportWarnings -Identity $Domain -Stage 'AddressRewrite' -Operation 'Get-AddressRewriteEntry' -Warnings $warnings
    }
    catch {
        Add-ExportWarnings -Identity $Domain -Stage 'AddressRewrite' -Operation 'Get-AddressRewriteEntry' -Warnings $warnings
        Add-ExportError -Identity $Domain -Stage 'AddressRewrite' -Operation 'Get-AddressRewriteEntry' -Message $_.Exception.Message
        return @()
    }

    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($entry in $entries) {
        $values = @(
            Get-ObjectPropertyValue -InputObject $entry -PropertyName 'Name'
            Get-ObjectPropertyValue -InputObject $entry -PropertyName 'InternalAddress'
            Get-ObjectPropertyValue -InputObject $entry -PropertyName 'ExternalAddress'
            Get-ObjectPropertyValue -InputObject $entry -PropertyName 'Pattern'
            Get-ObjectPropertyValue -InputObject $entry -PropertyName 'ExceptionList'
        )
        $usesTargetDomain = Test-ValueUsesDomain -Value $values -Domain $Domain
        if (-not $usesTargetDomain -and -not $IncludeAll) { continue }

        $rows.Add([pscustomobject]@{
            ObjectType             = 'AddressRewriteRule'
            Identity               = [string](Get-ObjectPropertyValue -InputObject $entry -PropertyName 'Identity')
            Name                   = [string](Get-ObjectPropertyValue -InputObject $entry -PropertyName 'Name')
            InternalAddress        = [string](Get-ObjectPropertyValue -InputObject $entry -PropertyName 'InternalAddress')
            ExternalAddress        = [string](Get-ObjectPropertyValue -InputObject $entry -PropertyName 'ExternalAddress')
            Pattern                = [string](Get-ObjectPropertyValue -InputObject $entry -PropertyName 'Pattern')
            ExceptionList          = Join-ReportValues -Values (Get-ObjectPropertyValue -InputObject $entry -PropertyName 'ExceptionList')
            OutboundOnly           = [string](Get-ObjectPropertyValue -InputObject $entry -PropertyName 'OutboundOnly')
            DependencyReasons      = if ($usesTargetDomain) { 'TargetDomainAddressRewriteRule' } else { '' }
            RecommendedDisposition = if ($usesTargetDomain) { 'Reconfiguration' } else { 'ExclusionCandidate' }
            TestingAndCommsNeeds   = if ($usesTargetDomain) { 'Test rewritten sender/recipient address flow; Communicate rewrite change to application and mail-flow owners' } else { '' }
        }) | Out-Null
    }

    return @($rows)
}

function Get-DistributionGroupsSafe {
    $warnings = @()
    try {
        $groups = @(Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop -WarningAction SilentlyContinue -WarningVariable warnings)
        Add-ExportWarnings -Identity '' -Stage 'Groups' -Operation 'Get-DistributionGroup' -Warnings $warnings
        return $groups
    }
    catch {
        Add-ExportWarnings -Identity '' -Stage 'Groups' -Operation 'Get-DistributionGroup' -Warnings $warnings
        Add-ExportError -Identity '' -Stage 'Groups' -Operation 'Get-DistributionGroup' -Message $_.Exception.Message
        return @()
    }
}

function Get-DynamicDistributionGroupsFromRecipientFallback {
    param([Parameter(Mandatory = $true)] [string[]]$RecipientProperties)

    if (-not (Get-Command Get-EXORecipient -ErrorAction SilentlyContinue)) {
        Add-ExportError -Identity '' -Stage 'Groups' -Operation 'Get-EXORecipient:DynamicDistributionGroup' -Message 'Get-EXORecipient is not available for dynamic distribution group fallback.'
        return @()
    }

    $recipientWarnings = @()
    try {
        $recipients = @(Get-EXORecipient -ResultSize Unlimited -RecipientTypeDetails DynamicDistributionGroup -Properties $RecipientProperties -ErrorAction Stop -WarningAction SilentlyContinue -WarningVariable recipientWarnings)
        Add-ExportWarnings -Identity '' -Stage 'Groups' -Operation 'Get-EXORecipient:DynamicDistributionGroup' -Warnings $recipientWarnings
    }
    catch {
        Add-ExportWarnings -Identity '' -Stage 'Groups' -Operation 'Get-EXORecipient:DynamicDistributionGroup' -Warnings $recipientWarnings
        Add-ExportError -Identity '' -Stage 'Groups' -Operation 'Get-EXORecipient:DynamicDistributionGroup' -Message $_.Exception.Message
        return @()
    }

    $groups = [System.Collections.Generic.List[object]]::new()
    foreach ($recipient in $recipients) {
        $identity = [string](Get-ObjectPropertyValue -InputObject $recipient -PropertyName 'Identity')
        if ([string]::IsNullOrWhiteSpace($identity)) {
            $groups.Add($recipient) | Out-Null
            continue
        }

        $singleWarnings = @()
        try {
            $group = Get-DynamicDistributionGroup -Identity $identity -ErrorAction Stop -WarningAction SilentlyContinue -WarningVariable singleWarnings
            Add-ExportWarnings -Identity $identity -Stage 'Groups' -Operation 'Get-DynamicDistributionGroup:Identity' -Warnings $singleWarnings
            $groups.Add($group) | Out-Null
        }
        catch {
            Add-ExportWarnings -Identity $identity -Stage 'Groups' -Operation 'Get-DynamicDistributionGroup:Identity' -Warnings $singleWarnings
            Add-ExportError -Identity $identity -Stage 'Groups' -Operation 'Get-DynamicDistributionGroup:Identity' -Message "Using Get-EXORecipient fallback row because object-specific lookup failed: $($_.Exception.Message)"
            $groups.Add($recipient) | Out-Null
        }
    }

    return @($groups)
}

function Get-DynamicDistributionGroupsSafe {
    param([Parameter(Mandatory = $true)] [string[]]$RecipientProperties)

    $warnings = @()
    try {
        $groups = @(Get-DynamicDistributionGroup -ResultSize Unlimited -ErrorAction Stop -WarningAction SilentlyContinue -WarningVariable warnings)
        Add-ExportWarnings -Identity '' -Stage 'Groups' -Operation 'Get-DynamicDistributionGroup' -Warnings $warnings
        return $groups
    }
    catch {
        Add-ExportWarnings -Identity '' -Stage 'Groups' -Operation 'Get-DynamicDistributionGroup' -Warnings $warnings
        Add-ExportError -Identity '' -Stage 'Groups' -Operation 'Get-DynamicDistributionGroup' -Message $_.Exception.Message
        return @(Get-DynamicDistributionGroupsFromRecipientFallback -RecipientProperties $RecipientProperties)
    }
}

function Get-MailboxesSafe {
    param([Parameter(Mandatory = $true)] [string[]]$RecipientProperties)

    $warnings = @()
    try {
        $mailboxes = @(Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox,RoomMailbox,EquipmentMailbox -Properties $RecipientProperties -ErrorAction Stop -WarningAction SilentlyContinue -WarningVariable warnings)
        Add-ExportWarnings -Identity '' -Stage 'Mailboxes' -Operation 'Get-EXOMailbox' -Warnings $warnings
        return $mailboxes
    }
    catch {
        Add-ExportWarnings -Identity '' -Stage 'Mailboxes' -Operation 'Get-EXOMailbox' -Warnings $warnings
        Add-ExportError -Identity '' -Stage 'Mailboxes' -Operation 'Get-EXOMailbox' -Message $_.Exception.Message
        return @()
    }
}

function Get-DependencyInventory {
    param([Parameter(Mandatory = $true)] [string]$Domain)

    $rows = [System.Collections.Generic.List[object]]::new()
    $recipientProperties = @(
        'EmailAddresses','Alias','PrimarySmtpAddress','ExternalDirectoryObjectId','RecipientTypeDetails',
        'RequireSenderAuthenticationEnabled','AcceptMessagesOnlyFrom','AcceptMessagesOnlyFromDLMembers',
        'AcceptMessagesOnlyFromSendersOrMembers','RejectMessagesFrom','RejectMessagesFromDLMembers','RejectMessagesFromSendersOrMembers',
        'ModerationEnabled','ModeratedBy','BypassModerationFromSendersOrMembers','GrantSendOnBehalfTo',
        'ForwardingAddress','ForwardingSmtpAddress','DeliverToMailboxAndForward'
    )

    Write-Host 'Collecting distribution lists and mail-enabled security groups...' -ForegroundColor Cyan
    $distributionGroups = @(Get-DistributionGroupsSafe)
    foreach ($group in $distributionGroups) {
        $members = @(Get-DistributionGroupMembersSafe -Group $group)
        $row = ConvertTo-DependencyRow -Object $group -ObjectType (Get-RecipientObjectType -Recipient $group) -Domain $Domain -Members $members -Owners (Get-ObjectPropertyValue -InputObject $group -PropertyName 'ManagedBy') -MembershipSource 'StaticDistributionGroup'
        if ($IncludeAll -or -not [string]::IsNullOrWhiteSpace($row.DependencyReasons)) { $rows.Add($row) | Out-Null }
    }

    Write-Host 'Collecting dynamic distribution groups...' -ForegroundColor Cyan
    $dynamicGroups = @(Get-DynamicDistributionGroupsSafe -RecipientProperties $recipientProperties)
    foreach ($group in $dynamicGroups) {
        $row = ConvertTo-DependencyRow -Object $group -ObjectType 'DynamicDistributionGroup' -Domain $Domain -Owners (Get-ObjectPropertyValue -InputObject $group -PropertyName 'ManagedBy') -MembershipSource "DynamicRecipientFilter: $([string](Get-ObjectPropertyValue -InputObject $group -PropertyName 'RecipientFilter'))"
        if (Test-ValueUsesDomain -Value (Get-ObjectPropertyValue -InputObject $group -PropertyName 'RecipientFilter') -Domain $Domain) {
            $row.DependencyReasons = Join-ReportValues -Values @($row.DependencyReasons, 'TargetDomainMembershipFilter')
            $row.RecommendedDisposition = Get-ReviewDisposition -ObjectType 'DynamicDistributionGroup' -Reasons @($row.DependencyReasons -split '; ')
            $row.TestingAndCommsNeeds = Join-ReportValues -Values @($row.TestingAndCommsNeeds, 'Validate dynamic membership filter after migration')
        }
        if ($IncludeAll -or -not [string]::IsNullOrWhiteSpace($row.DependencyReasons)) { $rows.Add($row) | Out-Null }
    }

    Write-Host 'Collecting Microsoft 365 groups...' -ForegroundColor Cyan
    try {
        $unifiedGroupWarnings = @()
        $unifiedGroups = @(Get-UnifiedGroup -ResultSize Unlimited -ErrorAction Stop -WarningAction SilentlyContinue -WarningVariable unifiedGroupWarnings)
        Add-ExportWarnings -Identity '' -Stage 'Groups' -Operation 'Get-UnifiedGroup' -Warnings $unifiedGroupWarnings
        foreach ($group in $unifiedGroups) {
            $members = @(Get-UnifiedGroupLinksSafe -Group $group -LinkType Members)
            $owners = @(Get-UnifiedGroupLinksSafe -Group $group -LinkType Owners)
            $row = ConvertTo-DependencyRow -Object $group -ObjectType 'Microsoft365Group' -Domain $Domain -Members $members -Owners $owners -MembershipSource 'Microsoft365GroupLinks'
            if ($IncludeAll -or -not [string]::IsNullOrWhiteSpace($row.DependencyReasons)) { $rows.Add($row) | Out-Null }
        }
    }
    catch {
        Add-ExportWarnings -Identity '' -Stage 'Groups' -Operation 'Get-UnifiedGroup' -Warnings $unifiedGroupWarnings
        Add-ExportError -Identity '' -Stage 'Groups' -Operation 'Get-UnifiedGroup' -Message $_.Exception.Message
    }

    Write-Host 'Collecting shared and resource mailboxes...' -ForegroundColor Cyan
    $mailboxes = @(Get-MailboxesSafe -RecipientProperties $recipientProperties)
    foreach ($mailbox in $mailboxes) {
        $fullAccess = @(Get-MailboxPermissionValuesSafe -Mailbox $mailbox -PermissionType FullAccess)
        $sendAs = @(Get-MailboxPermissionValuesSafe -Mailbox $mailbox -PermissionType SendAs)
        $calendarProcessing = Get-CalendarProcessingSafe -Mailbox $mailbox
        $row = ConvertTo-DependencyRow -Object $mailbox -ObjectType (Get-RecipientObjectType -Recipient $mailbox) -Domain $Domain -FullAccessTrustees $fullAccess -SendAsTrustees $sendAs -CalendarProcessing $calendarProcessing -MembershipSource 'MailboxDelegation'
        if ($IncludeAll -or -not [string]::IsNullOrWhiteSpace($row.DependencyReasons)) { $rows.Add($row) | Out-Null }
    }

    return @($rows)
}

$targetDomain = Normalize-Domain -Domain $Domain
$resolvedOutputFolder = Initialize-OutputFolder -OutputFolder $OutputFolder

Connect-ExchangeOnlineIfNeeded

$inventoryRows = @(Get-DependencyInventory -Domain $targetDomain)
$rewriteRows = @(Get-AddressRewriteRows -Domain $targetDomain)

$inventoryHeaders = @(
    'ObjectType','Identity','DisplayName','RecipientTypeDetails','PrimarySmtpAddress','Alias','ExternalDirectoryObjectId',
    'TargetDomainAddresses','Owners','MembershipSource','TargetDomainMemberCount','TargetDomainMemberSample','AllowsExternalSenders',
    'AcceptMessagesOnlyFrom','RejectMessagesFrom','ModerationEnabled','ModeratedBy','BypassModeration','ForwardingOrRoutingTargets',
    'FullAccessTrustees','SendAsTrustees','SendOnBehalfTo','ResourceDelegates','ProcessExternalMeetingMessages',
    'DependencyReasons','RecommendedDisposition','TestingAndCommsNeeds'
)
$rewriteHeaders = @('ObjectType','Identity','Name','InternalAddress','ExternalAddress','Pattern','ExceptionList','OutboundOnly','DependencyReasons','RecommendedDisposition','TestingAndCommsNeeds')
$errorHeaders = @('Identity','Stage','Operation','Message')

Export-CsvWithHeaders -Rows $inventoryRows -Path (Join-Path $resolvedOutputFolder 'DomainDependencyInventory.csv') -Headers $inventoryHeaders
Export-CsvWithHeaders -Rows $rewriteRows -Path (Join-Path $resolvedOutputFolder 'AddressRewriteRules.csv') -Headers $rewriteHeaders
Export-CsvWithHeaders -Rows @($script:Errors) -Path (Join-Path $resolvedOutputFolder 'Errors-DomainDependencyInventory.csv') -Headers $errorHeaders

$requiresMigration = @($inventoryRows | Where-Object { $_.RecommendedDisposition -match 'Migration' }).Count
$requiresReconfiguration = @($inventoryRows | Where-Object { $_.RecommendedDisposition -match 'Reconfiguration' }).Count + @($rewriteRows | Where-Object { $_.RecommendedDisposition -match 'Reconfiguration' }).Count
$requiresOwnerDecision = @($inventoryRows | Where-Object { $_.RecommendedDisposition -match 'OwnerDecision' }).Count

$summaryRows = @(
    [pscustomobject]@{ Metric = 'TargetDomain'; Value = $targetDomain }
    [pscustomobject]@{ Metric = 'InventoryRows'; Value = $inventoryRows.Count }
    [pscustomobject]@{ Metric = 'AddressRewriteRows'; Value = $rewriteRows.Count }
    [pscustomobject]@{ Metric = 'Errors'; Value = $script:Errors.Count }
    [pscustomobject]@{ Metric = 'RequiresMigration'; Value = $requiresMigration }
    [pscustomobject]@{ Metric = 'RequiresReconfiguration'; Value = $requiresReconfiguration }
    [pscustomobject]@{ Metric = 'RequiresOwnerDecision'; Value = $requiresOwnerDecision }
)
Export-CsvWithHeaders -Rows $summaryRows -Path (Join-Path $resolvedOutputFolder 'Summary.csv') -Headers @('Metric','Value')

Write-Host "Domain dependency inventory exported to: $resolvedOutputFolder" -ForegroundColor Green
Write-Host "Inventory rows: $($inventoryRows.Count); address rewrite rows: $($rewriteRows.Count); errors: $($script:Errors.Count)" -ForegroundColor Green
