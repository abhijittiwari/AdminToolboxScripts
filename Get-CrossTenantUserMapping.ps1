<#
.SYNOPSIS
    Maps users from a source tenant to their corresponding target tenant accounts
    using synced on-premises extension attributes 13 and 14.

.DESCRIPTION
    Connects to Microsoft Graph in the source tenant and exports all users, then
    disconnects and connects to the target tenant. Each source user is matched to
    a target account by comparing the source user principal name against the
    on-premises synced extensionAttribute13 and extensionAttribute14 values on
    target users.

    Matching is case-insensitive. An exact attribute value match is tried first;
    if none is found, a fallback scan looks for target attribute values that
    contain the source UPN (for values that embed the UPN in surrounding text).
    Multiple target candidates are reported as MatchStatus=Ambiguous rather than
    guessed.

    The script is read-only. It writes one CSV row per source user with the
    match status, matched attribute, and target account details.

.PARAMETER SourceTenantId
    Optional tenant ID (GUID or domain) of the source tenant, passed to
    Connect-MgGraph to force sign-in against the correct tenant.

.PARAMETER TargetTenantId
    Optional tenant ID (GUID or domain) of the target tenant, passed to
    Connect-MgGraph to force sign-in against the correct tenant.

.PARAMETER IncludeGuests
    Include guest users when enumerating both tenants. By default only
    userType=Member accounts are processed.

.PARAMETER OutputPath
    CSV output path. Defaults to .\CrossTenantUserMapping_yyyyMMdd_HHmmss.csv.

.EXAMPLE
    .\Get-CrossTenantUserMapping.ps1

.EXAMPLE
    .\Get-CrossTenantUserMapping.ps1 -SourceTenantId contoso.onmicrosoft.com -TargetTenantId fabrikam.onmicrosoft.com
#>

[CmdletBinding()]
param(
    [string]$SourceTenantId,
    [string]$TargetTenantId,
    [switch]$IncludeGuests,
    [string]$OutputPath = ".\CrossTenantUserMapping_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

$ErrorActionPreference = 'Stop'

$script:MatchAttributeNames = @('ExtensionAttribute13', 'ExtensionAttribute14')

function ConvertTo-UpnKey {
    param([AllowEmptyString()] [string]$Value)

    return ([string]$Value).Trim().ToLowerInvariant()
}

function Join-ReportValues {
    param([AllowEmptyCollection()] [object[]]$Values)

    $cleanValues = @($Values | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | ForEach-Object { ([string]$_).Trim() } | Sort-Object -Unique)
    return ($cleanValues -join '; ')
}

function Get-UserExtensionAttributeValue {
    param(
        [Parameter(Mandatory = $true)] $User,
        [Parameter(Mandatory = $true)] [string]$AttributeName
    )

    $attributes = $User.OnPremisesExtensionAttributes
    if ($null -eq $attributes) { return '' }

    if ($attributes -is [hashtable]) {
        return [string]$attributes[$AttributeName]
    }

    $property = $attributes.PSObject.Properties[$AttributeName]
    if ($null -eq $property) { return '' }
    return [string]$property.Value
}

function New-TargetAttributeIndex {
    param([AllowEmptyCollection()] [object[]]$TargetUsers)

    $index = @{}
    foreach ($user in $TargetUsers) {
        foreach ($attributeName in $script:MatchAttributeNames) {
            $key = ConvertTo-UpnKey -Value (Get-UserExtensionAttributeValue -User $user -AttributeName $attributeName)
            if ([string]::IsNullOrWhiteSpace($key)) { continue }

            if (-not $index.ContainsKey($key)) {
                $index[$key] = [System.Collections.Generic.List[object]]::new()
            }
            $index[$key].Add([pscustomobject]@{
                User      = $user
                Attribute = $attributeName
                MatchType = 'Exact'
            }) | Out-Null
        }
    }
    return $index
}

function Find-TargetUserMatches {
    param(
        [Parameter(Mandatory = $true)] [string]$SourceUserPrincipalName,
        [Parameter(Mandatory = $true)] [hashtable]$AttributeIndex,
        [AllowEmptyCollection()] [object[]]$TargetUsers
    )

    $key = ConvertTo-UpnKey -Value $SourceUserPrincipalName
    if ([string]::IsNullOrWhiteSpace($key)) { return @() }

    $candidates = [System.Collections.Generic.List[object]]::new()
    if ($AttributeIndex.ContainsKey($key)) {
        foreach ($entry in $AttributeIndex[$key]) {
            $candidates.Add($entry) | Out-Null
        }
    }
    else {
        $pattern = "*$([System.Management.Automation.WildcardPattern]::Escape($key))*"
        foreach ($user in $TargetUsers) {
            foreach ($attributeName in $script:MatchAttributeNames) {
                $value = ConvertTo-UpnKey -Value (Get-UserExtensionAttributeValue -User $user -AttributeName $attributeName)
                if (-not [string]::IsNullOrWhiteSpace($value) -and $value -like $pattern) {
                    $candidates.Add([pscustomobject]@{
                        User      = $user
                        Attribute = $attributeName
                        MatchType = 'Contains'
                    }) | Out-Null
                }
            }
        }
    }

    # Collapse hits on the same target user (e.g. UPN present in both attributes)
    $results = [System.Collections.Generic.List[object]]::new()
    foreach ($group in @($candidates | Group-Object { [string]$_.User.Id })) {
        $attributeNames = @($group.Group | ForEach-Object Attribute | Sort-Object -Unique)
        $results.Add([pscustomobject]@{
            User              = $group.Group[0].User
            MatchedAttributes = ($attributeNames -join '; ')
            MatchType         = $group.Group[0].MatchType
        }) | Out-Null
    }
    return $results.ToArray()
}

function New-UserMappingRow {
    param(
        [Parameter(Mandatory = $true)] $SourceUser,
        [AllowEmptyCollection()] [object[]]$MatchResults
    )

    $matchStatus = switch (@($MatchResults).Count) {
        0 { 'NotFound' }
        1 { 'Matched' }
        default { 'Ambiguous' }
    }

    $matchedUsers = @($MatchResults | ForEach-Object User)

    return [pscustomobject]@{
        SourceUserPrincipalName    = [string]$SourceUser.UserPrincipalName
        SourceDisplayName          = [string]$SourceUser.DisplayName
        SourceUserId               = [string]$SourceUser.Id
        SourceAccountEnabled       = $SourceUser.AccountEnabled
        MatchStatus                = $matchStatus
        MatchType                  = Join-ReportValues -Values @($MatchResults | ForEach-Object MatchType)
        MatchedAttribute           = Join-ReportValues -Values @($MatchResults | ForEach-Object MatchedAttributes)
        TargetUserPrincipalName    = Join-ReportValues -Values @($matchedUsers | ForEach-Object UserPrincipalName)
        TargetDisplayName          = Join-ReportValues -Values @($matchedUsers | ForEach-Object DisplayName)
        TargetUserId               = Join-ReportValues -Values @($matchedUsers | ForEach-Object Id)
        TargetAccountEnabled       = Join-ReportValues -Values @($matchedUsers | ForEach-Object { $_.AccountEnabled })
        TargetOnPremisesSyncEnabled = Join-ReportValues -Values @($matchedUsers | ForEach-Object { $_.OnPremisesSyncEnabled })
        TargetExtensionAttribute13 = Join-ReportValues -Values @($matchedUsers | ForEach-Object { Get-UserExtensionAttributeValue -User $_ -AttributeName 'ExtensionAttribute13' })
        TargetExtensionAttribute14 = Join-ReportValues -Values @($matchedUsers | ForEach-Object { Get-UserExtensionAttributeValue -User $_ -AttributeName 'ExtensionAttribute14' })
    }
}

function Connect-TenantGraph {
    param(
        [Parameter(Mandatory = $true)] [string]$TenantLabel,
        [AllowEmptyString()] [string]$TenantId
    )

    Write-Host "Sign in to the $TenantLabel tenant when prompted." -ForegroundColor Yellow
    $connectParams = @{
        Scopes      = @('User.Read.All')
        NoWelcome   = $true
        ErrorAction = 'Stop'
    }
    if (-not [string]::IsNullOrWhiteSpace($TenantId)) {
        $connectParams['TenantId'] = $TenantId
    }
    Connect-MgGraph @connectParams

    $context = Get-MgContext
    Write-Host "Connected to $TenantLabel tenant $($context.TenantId) as $($context.Account)." -ForegroundColor Cyan
    return $context
}

function Get-TenantUsers {
    param(
        [Parameter(Mandatory = $true)] [string[]]$Properties,
        [switch]$IncludeGuests
    )

    $userParams = @{
        All         = $true
        Property    = $Properties
        ErrorAction = 'Stop'
    }
    if (-not $IncludeGuests) {
        $userParams['Filter'] = "userType eq 'Member'"
        $userParams['ConsistencyLevel'] = 'eventual'
        $userParams['CountVariable'] = 'tenantUserCount'
    }
    return @(Get-MgUser @userParams)
}

$sourceProperties = @('id', 'displayName', 'userPrincipalName', 'accountEnabled', 'userType')
$targetProperties = @('id', 'displayName', 'userPrincipalName', 'accountEnabled', 'userType', 'onPremisesSyncEnabled', 'onPremisesExtensionAttributes')

Connect-TenantGraph -TenantLabel 'source' -TenantId $SourceTenantId | Out-Null
try {
    Write-Host 'Collecting source tenant users...' -ForegroundColor Cyan
    $sourceUsers = Get-TenantUsers -Properties $sourceProperties -IncludeGuests:$IncludeGuests
}
finally {
    Disconnect-MgGraph | Out-Null
}
Write-Host "Source users collected: $($sourceUsers.Count)" -ForegroundColor Cyan

if ($sourceUsers.Count -eq 0) {
    throw 'No users were returned from the source tenant.'
}

Write-Host ''
Connect-TenantGraph -TenantLabel 'target' -TenantId $TargetTenantId | Out-Null
try {
    Write-Host 'Collecting target tenant users with extension attributes...' -ForegroundColor Cyan
    $targetUsers = Get-TenantUsers -Properties $targetProperties -IncludeGuests:$IncludeGuests
}
finally {
    Disconnect-MgGraph | Out-Null
}
Write-Host "Target users collected: $($targetUsers.Count)" -ForegroundColor Cyan

$attributeIndex = New-TargetAttributeIndex -TargetUsers $targetUsers

$reportRows = [System.Collections.Generic.List[object]]::new()
$count = 0
foreach ($sourceUser in $sourceUsers) {
    $count++
    Write-Progress -Activity 'Matching source users against target extension attributes' `
        -Status ("{0} of {1} - {2}" -f $count, $sourceUsers.Count, $sourceUser.UserPrincipalName) `
        -PercentComplete (($count / $sourceUsers.Count) * 100)

    $matchResults = @(Find-TargetUserMatches -SourceUserPrincipalName ([string]$sourceUser.UserPrincipalName) -AttributeIndex $attributeIndex -TargetUsers $targetUsers)
    $reportRows.Add((New-UserMappingRow -SourceUser $sourceUser -MatchResults $matchResults)) | Out-Null
}
Write-Progress -Activity 'Matching source users against target extension attributes' -Completed

$reportRows |
    Sort-Object MatchStatus, SourceUserPrincipalName |
    Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

$matchedCount = @($reportRows | Where-Object MatchStatus -eq 'Matched').Count
$ambiguousCount = @($reportRows | Where-Object MatchStatus -eq 'Ambiguous').Count
$notFoundCount = @($reportRows | Where-Object MatchStatus -eq 'NotFound').Count

Write-Host ''
Write-Host "Report exported: $OutputPath ($($reportRows.Count) rows)" -ForegroundColor Green
Write-Host "Matched: $matchedCount | Ambiguous: $ambiguousCount | NotFound: $notFoundCount" -ForegroundColor Green
