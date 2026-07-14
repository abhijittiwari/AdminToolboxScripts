<#
.SYNOPSIS
    Exports WAE and Fortescue admin account Entra ID directory role and license details.

.DESCRIPTION
    Reads a CSV containing WAEUPN, Prefix, and FortescueUPN columns. The script connects
    to Microsoft Graph for WAE, exports the WAE account details, disconnects, prompts for
    Fortescue sign-in, exports the Fortescue account details, and writes one combined CSV.
    If a Fortescue UPN is blank or stale, the script strips (Admin) from the matching WAE
    display name and tries exact Fortescue displayName lookups for Admin <name> first,
    then <name>, when a candidate finds exactly one user.

    The script is read-only. It reports Entra ID directory roles only, including active
    and PIM-eligible assignments when available.

.PARAMETER InputCsv
    Path to the admin mapping CSV. Defaults to .\Admin.csv.

.PARAMETER OutputPath
    Combined CSV output path. Defaults to .\CrossTenantAdminRoles_yyyyMMdd_HHmmss.csv.

.EXAMPLE
    .\Get-CrossTenantAdminRoleAssignments.ps1 -InputCsv .\Admin.csv
#>

[CmdletBinding()]
param(
    [string]$InputCsv = '.\Admin.csv',
    [string]$OutputPath = ".\CrossTenantAdminRoles_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

$ErrorActionPreference = 'Stop'

function Test-RequiredCsvColumns {
    param(
        [Parameter(Mandatory = $true)] [object[]]$Rows,
        [Parameter(Mandatory = $true)] [string[]]$RequiredColumns
    )

    if ($Rows.Count -eq 0) {
        throw 'Input CSV contains no data rows.'
    }

    $actualColumns = @($Rows[0].PSObject.Properties.Name)
    $missingColumns = @($RequiredColumns | Where-Object { $actualColumns -notcontains $_ })
    if ($missingColumns.Count -gt 0) {
        throw "Input CSV is missing required column(s): $($missingColumns -join ', ')"
    }
}

function Join-ReportValues {
    param([object[]]$Values)

    $cleanValues = @($Values | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Sort-Object -Unique)
    return ($cleanValues -join '; ')
}

function Add-UserRole {
    param(
        [Parameter(Mandatory = $true)] [hashtable]$RoleMap,
        [Parameter(Mandatory = $true)] [string]$UserId,
        [Parameter(Mandatory = $true)] [string]$RoleName
    )

    if (-not $RoleMap.ContainsKey($UserId)) {
        $RoleMap[$UserId] = [System.Collections.Generic.List[string]]::new()
    }

    if (-not $RoleMap[$UserId].Contains($RoleName)) {
        $RoleMap[$UserId].Add($RoleName) | Out-Null
    }
}

function Test-GraphNotFoundError {
    param([Parameter(Mandatory = $true)] [System.Management.Automation.ErrorRecord]$ErrorRecord)

    $message = $ErrorRecord.Exception.Message
    $errorId = $ErrorRecord.FullyQualifiedErrorId
    return ($errorId -match 'Request_ResourceNotFound|NotFound|404' -or $message -match 'Request_ResourceNotFound|NotFound|404|does not exist')
}

function ConvertTo-AdminlessDisplayName {
    param([AllowEmptyString()] [string]$DisplayName)

    $cleanName = ([string]$DisplayName) -replace '\s*\(Admin\)\s*', ' '
    $cleanName = $cleanName -replace '\s+', ' '
    return $cleanName.Trim()
}

function Find-MgUserByDisplayName {
    param([Parameter(Mandatory = $true)] [string]$DisplayName)

    $escapedDisplayName = $DisplayName -replace "'", "''"
    $filter = "displayName eq '$escapedDisplayName'"
    return @(Get-MgUser -Filter $filter -Property 'id,displayName,userPrincipalName,onPremisesImmutableId' -All -ErrorAction Stop)
}

function Get-DisplayNameFallbackCandidates {
    param([AllowEmptyString()] [string]$DisplayName)

    $baseDisplayName = ConvertTo-AdminlessDisplayName -DisplayName $DisplayName
    if ([string]::IsNullOrWhiteSpace($baseDisplayName)) { return @() }

    $candidates = [System.Collections.Generic.List[string]]::new()
    if ($baseDisplayName -notmatch '^Admin\s+') {
        $candidates.Add("Admin $baseDisplayName") | Out-Null
    }
    $candidates.Add($baseDisplayName) | Out-Null

    return @($candidates | Select-Object -Unique)
}

function Resolve-RolePrincipal {
    param(
        [Parameter(Mandatory = $true)] [string]$PrincipalId,
        [Parameter(Mandatory = $true)] [string]$RoleName,
        [Parameter(Mandatory = $true)] [hashtable]$RoleMap,
        [System.Collections.Generic.List[string]]$EnvironmentWarnings
    )

    $principal = Get-MgDirectoryObjectById -Ids $PrincipalId -ErrorAction SilentlyContinue
    if (-not $principal) { return }

    switch ($principal.AdditionalProperties.'@odata.type') {
        '#microsoft.graph.user' {
            Add-UserRole -RoleMap $RoleMap -UserId $PrincipalId -RoleName $RoleName
        }
        '#microsoft.graph.group' {
            $groupName = $principal.AdditionalProperties.displayName
            try {
                Get-MgGroupTransitiveMember -GroupId $PrincipalId -All -ErrorAction Stop |
                    Where-Object { $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.user' } |
                    ForEach-Object {
                        Add-UserRole -RoleMap $RoleMap -UserId $_.Id -RoleName "$RoleName via group: $groupName"
                    }
            }
            catch {
                $groupLabel = if ([string]::IsNullOrWhiteSpace($groupName)) { $PrincipalId } else { $groupName }
                $warning = "Group role expansion failed for ${groupLabel}: $($_.Exception.Message)"
                if ($null -ne $EnvironmentWarnings) {
                    $EnvironmentWarnings.Add($warning) | Out-Null
                }
                Write-Warning $warning
            }
        }
    }
}

function Get-TenantRoleIndex {
    $roleDefinitions = @{}
    Get-MgRoleManagementDirectoryRoleDefinition -All | ForEach-Object {
        $roleDefinitions[$_.Id] = $_.DisplayName
    }

    $activeRoles = @{}
    $eligibleRoles = @{}
    $pimWarning = ''
    $environmentWarnings = [System.Collections.Generic.List[string]]::new()

    Write-Host 'Collecting active Entra ID role assignments...' -ForegroundColor Cyan
    $activeAssignments = @(Get-MgRoleManagementDirectoryRoleAssignment -All)
    foreach ($assignment in $activeAssignments) {
        $roleName = $roleDefinitions[$assignment.RoleDefinitionId]
        if ($roleName) {
            Resolve-RolePrincipal -PrincipalId $assignment.PrincipalId -RoleName "$roleName (Active)" -RoleMap $activeRoles -EnvironmentWarnings $environmentWarnings
        }
    }

    Write-Host 'Collecting PIM-eligible Entra ID role assignments...' -ForegroundColor Cyan
    try {
        $eligibleAssignments = @(Get-MgRoleManagementDirectoryRoleEligibilityScheduleInstance -All)
        foreach ($assignment in $eligibleAssignments) {
            $roleName = $roleDefinitions[$assignment.RoleDefinitionId]
            if ($roleName) {
                Resolve-RolePrincipal -PrincipalId $assignment.PrincipalId -RoleName "$roleName (PIM Eligible)" -RoleMap $eligibleRoles -EnvironmentWarnings $environmentWarnings
            }
        }
    }
    catch {
        $pimWarning = "Could not read PIM eligibility schedules: $($_.Exception.Message)"
        Write-Warning $pimWarning
    }

    return [pscustomobject]@{
        ActiveRoles  = $activeRoles
        EligibleRoles = $eligibleRoles
        PimWarning   = $pimWarning
        EnvironmentWarnings = Join-ReportValues -Values $environmentWarnings
    }
}

function Get-UserReportRow {
    param(
        [Parameter(Mandatory = $true)] [string]$Environment,
        [Parameter(Mandatory = $true)] [string]$Prefix,
        [Parameter(Mandatory = $true)] [AllowEmptyString()] [string]$InputUPN,
        [Parameter(Mandatory = $true)] [object]$RoleIndex,
        [AllowEmptyString()] [string]$DisplayNameHint = ''
    )

    $errorMessages = [System.Collections.Generic.List[string]]::new()
    if (-not [string]::IsNullOrWhiteSpace($RoleIndex.PimWarning)) {
        $errorMessages.Add($RoleIndex.PimWarning) | Out-Null
    }
    if (-not [string]::IsNullOrWhiteSpace($RoleIndex.EnvironmentWarnings)) {
        $errorMessages.Add($RoleIndex.EnvironmentWarnings) | Out-Null
    }

    $user = $null
    $lookupStatus = 'Found'
    $upnLookupNotFound = $false
    $displayNameFallbackError = $false

    if (-not [string]::IsNullOrWhiteSpace($InputUPN)) {
        try {
            $user = Get-MgUser -UserId $InputUPN -Property 'id,displayName,userPrincipalName,onPremisesImmutableId' -ErrorAction Stop
        }
        catch {
            $upnLookupNotFound = Test-GraphNotFoundError -ErrorRecord $_
            $errorMessages.Add("User lookup failed for input UPN '$InputUPN': $($_.Exception.Message)") | Out-Null
            if (-not $upnLookupNotFound) {
                return [pscustomobject]@{
                    Environment               = $Environment
                    Prefix                    = $Prefix
                    InputUPN                  = $InputUPN
                    UserPrincipalName         = ''
                    DisplayName               = ''
                    ImmutableId               = ''
                    ActiveDirectoryRoles      = ''
                    PimEligibleDirectoryRoles = ''
                    AllDirectoryRoles         = ''
                    LicensesAssigned          = ''
                    LookupStatus              = 'Error'
                    Error                     = Join-ReportValues -Values $errorMessages
                }
            }
        }
    }

    $fallbackDisplayName = ConvertTo-AdminlessDisplayName -DisplayName $DisplayNameHint
    $fallbackDisplayNames = @(Get-DisplayNameFallbackCandidates -DisplayName $DisplayNameHint)
    if (-not $user -and $fallbackDisplayNames.Count -gt 0) {
        try {
            foreach ($candidateDisplayName in $fallbackDisplayNames) {
                $matches = @(Find-MgUserByDisplayName -DisplayName $candidateDisplayName)
                if ($matches.Count -eq 1) {
                    $user = $matches[0]
                    $lookupStatus = 'FoundByDisplayName'
                    $fallbackReason = if ([string]::IsNullOrWhiteSpace($InputUPN)) { 'input UPN was blank' } else { 'input UPN was not found' }
                    $errorMessages.Add("Matched by display name '$candidateDisplayName' because $fallbackReason.") | Out-Null
                    break
                }
                elseif ($matches.Count -eq 0) {
                    $errorMessages.Add("Display name fallback found no user for '$candidateDisplayName'.") | Out-Null
                }
                else {
                    $matchedUpns = Join-ReportValues -Values @($matches | ForEach-Object { $_.UserPrincipalName })
                    $errorMessages.Add("Display name fallback for '$candidateDisplayName' matched multiple users: $matchedUpns") | Out-Null
                    $displayNameFallbackError = $true
                    break
                }
            }
        }
        catch {
            $errorMessages.Add("Display name fallback failed for '$fallbackDisplayName': $($_.Exception.Message)") | Out-Null
            $displayNameFallbackError = $true
        }
    }

    if (-not $user) {
        if ([string]::IsNullOrWhiteSpace($InputUPN) -and [string]::IsNullOrWhiteSpace($fallbackDisplayName)) {
            $errorMessages.Add('Input UPN is blank and no display name hint is available.') | Out-Null
        }

        return [pscustomobject]@{
            Environment               = $Environment
            Prefix                    = $Prefix
            InputUPN                  = $InputUPN
            UserPrincipalName         = ''
            DisplayName               = ''
            ImmutableId               = ''
            ActiveDirectoryRoles      = ''
            PimEligibleDirectoryRoles = ''
            AllDirectoryRoles         = ''
            LicensesAssigned          = ''
            LookupStatus              = if ($displayNameFallbackError) { 'Error' } elseif ($upnLookupNotFound -or -not [string]::IsNullOrWhiteSpace($fallbackDisplayName)) { 'NotFound' } else { 'Error' }
            Error                     = Join-ReportValues -Values $errorMessages
        }
    }

    $licenses = ''
    try {
        $licenses = Join-ReportValues -Values @(Get-MgUserLicenseDetail -UserId $user.Id -ErrorAction Stop | Select-Object -ExpandProperty SkuPartNumber)
    }
    catch {
        $errorMessages.Add("License lookup failed: $($_.Exception.Message)") | Out-Null
        $lookupStatus = 'Error'
    }

    $activeRoles = if ($RoleIndex.ActiveRoles.ContainsKey($user.Id)) { Join-ReportValues -Values $RoleIndex.ActiveRoles[$user.Id] } else { '' }
    $eligibleRoles = if ($RoleIndex.EligibleRoles.ContainsKey($user.Id)) { Join-ReportValues -Values $RoleIndex.EligibleRoles[$user.Id] } else { '' }
    $allRoleValues = [System.Collections.Generic.List[string]]::new()
    if (-not [string]::IsNullOrWhiteSpace($activeRoles)) {
        ($activeRoles -split '; ') | ForEach-Object { $allRoleValues.Add($_) | Out-Null }
    }
    if (-not [string]::IsNullOrWhiteSpace($eligibleRoles)) {
        ($eligibleRoles -split '; ') | ForEach-Object { $allRoleValues.Add($_) | Out-Null }
    }
    $allRoles = Join-ReportValues -Values $allRoleValues

    return [pscustomobject]@{
        Environment               = $Environment
        Prefix                    = $Prefix
        InputUPN                  = $InputUPN
        UserPrincipalName         = $user.UserPrincipalName
        DisplayName               = $user.DisplayName
        ImmutableId               = $user.OnPremisesImmutableId
        ActiveDirectoryRoles      = $activeRoles
        PimEligibleDirectoryRoles = $eligibleRoles
        AllDirectoryRoles         = $allRoles
        LicensesAssigned          = $licenses
        LookupStatus              = $lookupStatus
        Error                     = Join-ReportValues -Values $errorMessages
    }
}

function Invoke-TenantExport {
    param(
        [Parameter(Mandatory = $true)] [string]$Environment,
        [Parameter(Mandatory = $true)] [object[]]$Mappings,
        [Parameter(Mandatory = $true)] [string]$UpnColumn,
        [hashtable]$DisplayNameHints = @{}
    )

    Write-Host "Connecting to Microsoft Graph for $Environment..." -ForegroundColor Cyan
    Connect-MgGraph -Scopes @('Directory.Read.All', 'RoleManagement.Read.Directory') -NoWelcome

    try {
        $roleIndex = Get-TenantRoleIndex
        $rows = [System.Collections.Generic.List[object]]::new()
        $count = 0

        foreach ($mapping in $Mappings) {
            $count++
            $upn = [string]$mapping.$UpnColumn
            $displayNameHint = ''
            if ($DisplayNameHints.ContainsKey([string]$mapping.Prefix)) {
                $displayNameHint = [string]$DisplayNameHints[[string]$mapping.Prefix]
            }
            Write-Progress -Activity "Processing $Environment admin accounts" `
                -Status ("{0} of {1} - {2}" -f $count, $Mappings.Count, $upn) `
                -PercentComplete (($count / $Mappings.Count) * 100)

            $rows.Add((Get-UserReportRow -Environment $Environment -Prefix $mapping.Prefix -InputUPN $upn -RoleIndex $roleIndex -DisplayNameHint $displayNameHint)) | Out-Null
        }

        Write-Progress -Activity "Processing $Environment admin accounts" -Completed
        return $rows
    }
    finally {
        Disconnect-MgGraph | Out-Null
    }
}

if (-not (Test-Path -Path $InputCsv -PathType Leaf)) {
    throw "Input CSV not found: $InputCsv"
}

$mappings = @(Import-Csv -Path $InputCsv)
Test-RequiredCsvColumns -Rows $mappings -RequiredColumns @('WAEUPN', 'Prefix', 'FortescueUPN')

$reportRows = [System.Collections.Generic.List[object]]::new()
$waeRows = [object[]](Invoke-TenantExport -Environment 'WAE' -Mappings $mappings -UpnColumn 'WAEUPN')
$reportRows.AddRange($waeRows)

$fortescueDisplayNameHints = @{}
foreach ($row in $waeRows) {
    $displayNameHint = ConvertTo-AdminlessDisplayName -DisplayName $row.DisplayName
    if (-not [string]::IsNullOrWhiteSpace($displayNameHint)) {
        $fortescueDisplayNameHints[[string]$row.Prefix] = $displayNameHint
    }
}

Write-Host ''
Write-Host 'Sign in to the Fortescue tenant when prompted.' -ForegroundColor Yellow
$reportRows.AddRange([object[]](Invoke-TenantExport -Environment 'Fortescue' -Mappings $mappings -UpnColumn 'FortescueUPN' -DisplayNameHints $fortescueDisplayNameHints))

$columns = 'Environment','Prefix','InputUPN','UserPrincipalName','DisplayName','ImmutableId',
           'ActiveDirectoryRoles','PimEligibleDirectoryRoles','AllDirectoryRoles',
           'LicensesAssigned','LookupStatus','Error'

$reportRows |
    Sort-Object Prefix, Environment |
    Select-Object $columns |
    Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

Write-Host "Report exported: $OutputPath ($($reportRows.Count) rows)" -ForegroundColor Green
