<#
.SYNOPSIS
    Finds duplicate proxy/email address values across on-prem Active Directory objects.

.DESCRIPTION
    Scans AD objects that contain one or more selected mail-related attributes.
    By default, checks:
        - proxyAddresses
        - mail

    Output is pair-based: each row shows the duplicate value and both AD objects
    that share it, including attribute type, attribute value, object type,
    objectGUID, SID if available, name, sAMAccountName, and DN.

.NOTES
    Requires: RSAT Active Directory PowerShell module
    This script is read-only. It does not modify AD.

.EXAMPLE
    .\Find-ADDuplicateEmailProxyAddresses.ps1

.EXAMPLE
    .\Find-ADDuplicateEmailProxyAddresses.ps1 `
        -SearchBase "DC=contoso,DC=com" `
        -Server "dc01.contoso.com"

.EXAMPLE
    .\Find-ADDuplicateEmailProxyAddresses.ps1 `
        -IncludeTargetAddress `
        -IncludeUserPrincipalName `
        -SmtpOnly

.EXAMPLE
    .\Find-ADDuplicateEmailProxyAddresses.ps1 `
        -OutputCsv "C:\Temp\AD-Duplicate-Address-Pairs.csv" `
        -OccurrenceCsv "C:\Temp\AD-Duplicate-Address-Occurrences.csv"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$SearchBase,

    [Parameter(Mandatory = $false)]
    [string]$Server,

    [Parameter(Mandatory = $false)]
    [string]$OutputCsv = ".\AD-Duplicate-Address-Pairs.csv",

    [Parameter(Mandatory = $false)]
    [string]$OccurrenceCsv = ".\AD-Duplicate-Address-Occurrences.csv",

    [Parameter(Mandatory = $false)]
    [string[]]$AdditionalAttributes = @(),

    [Parameter(Mandatory = $false)]
    [switch]$IncludeTargetAddress,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeUserPrincipalName,

    # When set, only SMTP proxyAddresses are checked.
    # Without this switch, non-SMTP proxyAddresses such as X500, X400, SIP are also checked.
    [Parameter(Mandatory = $false)]
    [switch]$SmtpOnly,

    [Parameter(Mandatory = $false)]
    [switch]$NoOccurrenceCsv
)

begin {
    Import-Module ActiveDirectory -ErrorAction Stop

    function ConvertTo-LdapEscapedValue {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Value
        )

        $Value.Replace('\', '\5c').
               Replace('*', '\2a').
               Replace('(', '\28').
               Replace(')', '\29').
               Replace(([string][char]0), '\00')
    }

    function ConvertTo-SidString {
        param(
            [Parameter(Mandatory = $false)]
            $Sid
        )

        if ($null -eq $Sid) {
            return $null
        }

        try {
            if ($Sid -is [System.Security.Principal.SecurityIdentifier]) {
                return $Sid.Value
            }

            if ($Sid -is [byte[]]) {
                return (New-Object System.Security.Principal.SecurityIdentifier($Sid, 0)).Value
            }

            return ([System.Security.Principal.SecurityIdentifier]$Sid).Value
        }
        catch {
            try {
                return [string]$Sid
            }
            catch {
                return $null
            }
        }
    }

    function Get-MostSpecificObjectClass {
        param(
            [Parameter(Mandatory = $false)]
            $ObjectClass
        )

        $classes = @($ObjectClass) | Where-Object { $_ }

        if ($classes.Count -eq 0) {
            return $null
        }

        return [string]$classes[-1]
    }

    function Test-ADAttributeExists {
        param(
            [Parameter(Mandatory = $true)]
            [string]$AttributeName,

            [Parameter(Mandatory = $true)]
            [string]$SchemaNamingContext,

            [Parameter(Mandatory = $false)]
            [string]$Server
        )

        $escapedAttributeName = ConvertTo-LdapEscapedValue -Value $AttributeName

        $params = @{
            SearchBase    = $SchemaNamingContext
            LDAPFilter    = "(&(objectClass=attributeSchema)(lDAPDisplayName=$escapedAttributeName))"
            Properties    = @("lDAPDisplayName")
            ResultSetSize = 1
            ErrorAction   = "Stop"
        }

        if (-not [string]::IsNullOrWhiteSpace($Server)) {
            $params["Server"] = $Server
        }

        try {
            $result = Get-ADObject @params
            return [bool]$result
        }
        catch {
            Write-Warning "Could not validate schema attribute '$AttributeName'. Error: $($_.Exception.Message)"
            return $false
        }
    }

    function ConvertTo-CheckedAddressValue {
        param(
            [Parameter(Mandatory = $true)]
            [string]$AttributeName,

            [Parameter(Mandatory = $true)]
            [string]$RawValue,

            [Parameter(Mandatory = $false)]
            [switch]$SmtpOnly
        )

        $value = ([string]$RawValue).Trim()

        if ([string]::IsNullOrWhiteSpace($value)) {
            return $null
        }

        $prefix = $null
        $body = $value

        if ($value -match '^([^:]+):(.+)$') {
            $prefix = $Matches[1].Trim()
            $body = $Matches[2].Trim()
        }

        if ([string]::IsNullOrWhiteSpace($body)) {
            return $null
        }

        $attributeType = $AttributeName
        $matchValue = $null

        if ($AttributeName -ieq "proxyAddresses") {
            if ($prefix) {
                if ($SmtpOnly -and ($prefix -ine "smtp")) {
                    return $null
                }

                if ($prefix -ieq "smtp") {
                    # SMTP proxyAddresses are normalized to the bare email address
                    # so they can match the mail attribute.
                    $attributeType = "proxyAddresses:smtp"
                    $matchValue = $body.ToLowerInvariant()
                }
                else {
                    # Non-SMTP proxies are compared with their prefix to avoid
                    # matching unrelated values across different proxy address types.
                    $attributeType = "proxyAddresses:$($prefix.ToLowerInvariant())"
                    $matchValue = "$($prefix.ToLowerInvariant()):$($body.ToLowerInvariant())"
                }
            }
            else {
                if ($SmtpOnly) {
                    return $null
                }

                $attributeType = "proxyAddresses"
                $matchValue = $value.ToLowerInvariant()
            }
        }
        else {
            # For mail, targetAddress, userPrincipalName, or custom attributes:
            # If the value is SMTP-prefixed, normalize to the bare address.
            # Otherwise compare the trimmed lower-case value.
            if ($prefix -and ($prefix -ieq "smtp")) {
                $matchValue = $body.ToLowerInvariant()
            }
            elseif ($prefix -and ($AttributeName -ieq "targetAddress")) {
                if ($SmtpOnly -and ($prefix -ine "smtp")) {
                    return $null
                }

                $matchValue = "$($prefix.ToLowerInvariant()):$($body.ToLowerInvariant())"
            }
            else {
                $matchValue = $value.ToLowerInvariant()
            }
        }

        if ([string]::IsNullOrWhiteSpace($matchValue)) {
            return $null
        }

        [pscustomobject]@{
            AttributeType  = $attributeType
            AttributeValue = $value
            MatchValue     = $matchValue
        }
    }

    function Ensure-ParentFolder {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Path
        )

        $resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
        $parent = Split-Path -Path $resolvedPath -Parent

        if (-not [string]::IsNullOrWhiteSpace($parent) -and -not (Test-Path -Path $parent)) {
            New-Item -Path $parent -ItemType Directory -Force | Out-Null
        }

        return $resolvedPath
    }
}

process {
    $rootDseParams = @{
        ErrorAction = "Stop"
    }

    if (-not [string]::IsNullOrWhiteSpace($Server)) {
        $rootDseParams["Server"] = $Server
    }

    $rootDse = Get-ADRootDSE @rootDseParams

    if ([string]::IsNullOrWhiteSpace($SearchBase)) {
        $SearchBase = $rootDse.DefaultNamingContext
    }

    $requestedAttributes = @(
        "proxyAddresses"
        "mail"
    )

    if ($IncludeTargetAddress) {
        $requestedAttributes += "targetAddress"
    }

    if ($IncludeUserPrincipalName) {
        $requestedAttributes += "userPrincipalName"
    }

    if ($AdditionalAttributes.Count -gt 0) {
        $requestedAttributes += $AdditionalAttributes
    }

    $requestedAttributes = $requestedAttributes |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        ForEach-Object { $_.Trim() } |
        Sort-Object -Unique

    $availableAttributes = foreach ($attribute in $requestedAttributes) {
        if (Test-ADAttributeExists `
                -AttributeName $attribute `
                -SchemaNamingContext $rootDse.SchemaNamingContext `
                -Server $Server) {
            $attribute
        }
        else {
            Write-Warning "Skipping '$attribute' because it was not found in the AD schema."
        }
    }

    $availableAttributes = @($availableAttributes)

    if ($availableAttributes.Count -eq 0) {
        throw "None of the requested attributes exist in this AD schema. Nothing to scan."
    }

    $ldapAttributeFilters = foreach ($attribute in $availableAttributes) {
        "($attribute=*)"
    }

    if ($ldapAttributeFilters.Count -eq 1) {
        $ldapFilter = $ldapAttributeFilters[0]
    }
    else {
        $ldapFilter = "(|$($ldapAttributeFilters -join ''))"
    }

    $propertiesToLoad = @(
        $availableAttributes
        "name"
        "objectClass"
        "objectGUID"
        "objectSid"
        "sAMAccountName"
        "distinguishedName"
    ) |
        ForEach-Object { $_ } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Sort-Object -Unique

    $getAdObjectParams = @{
        LDAPFilter    = $ldapFilter
        SearchBase    = $SearchBase
        SearchScope   = "Subtree"
        Properties    = $propertiesToLoad
        ResultPageSize = 1000
        ResultSetSize = $null
        ErrorAction   = "Stop"
    }

    if (-not [string]::IsNullOrWhiteSpace($Server)) {
        $getAdObjectParams["Server"] = $Server
    }

    Write-Host "Search base: $SearchBase"
    Write-Host "Attributes checked: $($availableAttributes -join ', ')"
    Write-Host "LDAP filter: $ldapFilter"

    $occurrences = New-Object System.Collections.Generic.List[object]
    $objectCount = 0

    foreach ($adObject in Get-ADObject @getAdObjectParams) {
        $objectCount++

        $objectType = Get-MostSpecificObjectClass -ObjectClass $adObject.objectClass
        $objectId = [string]$adObject.objectGUID
        $sid = ConvertTo-SidString -Sid $adObject.objectSid

        foreach ($attribute in $availableAttributes) {
            $rawValues = @($adObject.$attribute) | Where-Object { $null -ne $_ }

            foreach ($rawValue in $rawValues) {
                $checkedValue = ConvertTo-CheckedAddressValue `
                    -AttributeName $attribute `
                    -RawValue ([string]$rawValue) `
                    -SmtpOnly:$SmtpOnly

                if ($null -eq $checkedValue) {
                    continue
                }

                $occurrences.Add([pscustomobject]@{
                    MatchValue        = $checkedValue.MatchValue
                    AttributeType     = $checkedValue.AttributeType
                    AttributeValue    = $checkedValue.AttributeValue
                    ObjectType        = $objectType
                    ObjectID          = $objectId
                    SID               = $sid
                    Name              = [string]$adObject.Name
                    SamAccountName    = [string]$adObject.sAMAccountName
                    DistinguishedName = [string]$adObject.DistinguishedName
                }) | Out-Null
            }
        }
    }

    $duplicatePairs = New-Object System.Collections.Generic.List[object]
    $duplicateOccurrences = New-Object System.Collections.Generic.List[object]

    $duplicateGroups = $occurrences |
        Group-Object -Property MatchValue |
        Where-Object {
            @($_.Group | Select-Object -ExpandProperty ObjectID -Unique).Count -gt 1
        }

    foreach ($duplicateGroup in $duplicateGroups) {
        $objectsForValue = @($duplicateGroup.Group | Group-Object -Property ObjectID)
        $duplicateObjectCount = $objectsForValue.Count

        foreach ($occurrence in $duplicateGroup.Group) {
            $duplicateOccurrences.Add($occurrence) | Out-Null
        }

        for ($i = 0; $i -lt $objectsForValue.Count; $i++) {
            for ($j = $i + 1; $j -lt $objectsForValue.Count; $j++) {
                $leftGroup = @($objectsForValue[$i].Group)
                $rightGroup = @($objectsForValue[$j].Group)

                $left = $leftGroup[0]
                $right = $rightGroup[0]

                $leftAttributeTypes = (
                    $leftGroup |
                        Select-Object -ExpandProperty AttributeType -Unique |
                        Sort-Object
                ) -join "; "

                $leftAttributeValues = (
                    $leftGroup |
                        Select-Object -ExpandProperty AttributeValue -Unique |
                        Sort-Object
                ) -join "; "

                $rightAttributeTypes = (
                    $rightGroup |
                        Select-Object -ExpandProperty AttributeType -Unique |
                        Sort-Object
                ) -join "; "

                $rightAttributeValues = (
                    $rightGroup |
                        Select-Object -ExpandProperty AttributeValue -Unique |
                        Sort-Object
                ) -join "; "

                $duplicatePairs.Add([pscustomobject]@{
                    DuplicateMatchValue      = $duplicateGroup.Name
                    DuplicateObjectCount     = $duplicateObjectCount

                    Object1AttributeType     = $leftAttributeTypes
                    Object1AttributeValue    = $leftAttributeValues
                    Object1ObjectType        = $left.ObjectType
                    Object1ObjectID          = $left.ObjectID
                    Object1SID               = $left.SID
                    Object1Name              = $left.Name
                    Object1SamAccountName    = $left.SamAccountName
                    Object1DistinguishedName = $left.DistinguishedName

                    Object2AttributeType     = $rightAttributeTypes
                    Object2AttributeValue    = $rightAttributeValues
                    Object2ObjectType        = $right.ObjectType
                    Object2ObjectID          = $right.ObjectID
                    Object2SID               = $right.SID
                    Object2Name              = $right.Name
                    Object2SamAccountName    = $right.SamAccountName
                    Object2DistinguishedName = $right.DistinguishedName
                }) | Out-Null
            }
        }
    }

    $duplicatePairsSorted = @(
        $duplicatePairs |
            Sort-Object `
                DuplicateMatchValue,
                Object1ObjectType,
                Object1Name,
                Object2ObjectType,
                Object2Name
    )

    $duplicateOccurrencesSorted = @(
        $duplicateOccurrences |
            Sort-Object `
                MatchValue,
                ObjectType,
                Name,
                AttributeType
    )

    $resolvedOutputCsv = Ensure-ParentFolder -Path $OutputCsv

    $duplicatePairsSorted |
        Export-Csv -Path $resolvedOutputCsv -NoTypeInformation -Encoding UTF8

    if (-not $NoOccurrenceCsv) {
        $resolvedOccurrenceCsv = Ensure-ParentFolder -Path $OccurrenceCsv

        $duplicateOccurrencesSorted |
            Export-Csv -Path $resolvedOccurrenceCsv -NoTypeInformation -Encoding UTF8
    }

    Write-Host ""
    Write-Host "Objects scanned: $objectCount"
    Write-Host "Address/proxy occurrences scanned: $($occurrences.Count)"
    Write-Host "Duplicate values found: $(@($duplicateGroups).Count)"
    Write-Host "Duplicate object-pair rows exported: $($duplicatePairsSorted.Count)"
    Write-Host "Pair CSV: $resolvedOutputCsv"

    if (-not $NoOccurrenceCsv) {
        Write-Host "Occurrence CSV: $resolvedOccurrenceCsv"
    }

    # Also write pair results to the pipeline for console review or further processing.
    $duplicatePairsSorted
}