<#
.SYNOPSIS
    Checks Entra ID directory synchronization hard-match and soft-match protection flags using Microsoft Graph.

.DESCRIPTION
    Uses Microsoft Graph v1.0 to read /directory/onPremisesSynchronization and reports:
      - blockCloudObjectTakeoverThroughHardMatchEnabled
      - blockSoftMatchEnabled

    The script is read-only and does not use the MSOnline module.

    Microsoft Learn reference:
      https://learn.microsoft.com/graph/api/resources/onpremisesdirectorysynchronizationfeature
      https://learn.microsoft.com/graph/api/onpremisesdirectorysynchronization-get

.PARAMETER OutputPath
    Optional CSV output path. If omitted, results are only written to the pipeline/table.

.EXAMPLE
    .\Get-EntraDirSyncProtectionSettings.ps1

.EXAMPLE
    .\Get-EntraDirSyncProtectionSettings.ps1 -OutputPath .\DirSyncProtectionSettings.csv
#>

[CmdletBinding()]
param(
    [string]$OutputPath = ''
)

$ErrorActionPreference = 'Stop'

function Connect-GraphIfNeeded {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    $context = Get-MgContext -ErrorAction SilentlyContinue
    if ($context) {
        Write-Host "Reusing existing Microsoft Graph context for tenant $($context.TenantId)." -ForegroundColor Cyan
        return
    }

    Connect-MgGraph -Scopes @('OnPremDirectorySynchronization.Read.All') -NoWelcome
}

function New-GraphRequestExecutor {
    return {
        param(
            [Parameter(Mandatory = $true)] [string]$Method,
            [Parameter(Mandatory = $true)] [string]$Uri,
            [Parameter(Mandatory = $false)] $Body
        )

        if ($null -ne $Body) {
            Invoke-MgGraphRequest -Method $Method -Uri $Uri -Body ($Body | ConvertTo-Json -Depth 10) -ContentType 'application/json'
        }
        else {
            Invoke-MgGraphRequest -Method $Method -Uri $Uri
        }
    }
}

function Get-ObjectPropertyValue {
    param(
        [Parameter(Mandatory = $true)] $InputObject,
        [Parameter(Mandatory = $true)] [string]$PropertyName
    )

    if ($InputObject -is [hashtable] -and $InputObject.ContainsKey($PropertyName)) {
        return $InputObject[$PropertyName]
    }

    $property = $InputObject.PSObject.Properties[$PropertyName]
    if ($property) { return $property.Value }

    $additionalProperties = $InputObject.PSObject.Properties['AdditionalProperties']
    if ($additionalProperties -and $additionalProperties.Value -is [hashtable] -and $additionalProperties.Value.ContainsKey($PropertyName)) {
        return $additionalProperties.Value[$PropertyName]
    }

    return $null
}

function Get-OnPremisesSynchronizationObject {
    param([Parameter(Mandatory = $true)] [scriptblock]$GraphRequest)

    $uri = 'https://graph.microsoft.com/v1.0/directory/onPremisesSynchronization?$select=id,features'
    $response = & $GraphRequest 'GET' $uri $null
    $value = Get-ObjectPropertyValue -InputObject $response -PropertyName 'value'

    if ($null -ne $value) {
        $objects = @($value)
        if ($objects.Count -eq 0) {
            throw 'Microsoft Graph returned no onPremisesSynchronization objects for this tenant.'
        }
        return $objects[0]
    }

    return $response
}

function ConvertTo-FeatureStatus {
    param([AllowNull()] $Value)

    if ($null -eq $Value) { return 'Unknown' }
    if ([bool]$Value) { return 'Enabled' }
    return 'Disabled'
}

function New-DirSyncProtectionRow {
    param(
        [Parameter(Mandatory = $true)] [string]$TenantId,
        [Parameter(Mandatory = $true)] $Features,
        [Parameter(Mandatory = $true)] [string]$Setting,
        [Parameter(Mandatory = $true)] [string]$GraphProperty,
        [Parameter(Mandatory = $true)] [string]$Notes
    )

    $enabled = Get-ObjectPropertyValue -InputObject $Features -PropertyName $GraphProperty
    $status = ConvertTo-FeatureStatus -Value $enabled

    return [pscustomobject]@{
        TenantId         = $TenantId
        Setting          = $Setting
        GraphProperty    = $GraphProperty
        Enabled          = $enabled
        Status           = $status
        RecommendedState = 'Enabled'
        Notes            = $Notes
    }
}

function Get-DirSyncProtectionSettings {
    param([Parameter(Mandatory = $false)] [scriptblock]$GraphRequest)

    if (-not $GraphRequest) { $GraphRequest = New-GraphRequestExecutor }

    $sync = Get-OnPremisesSynchronizationObject -GraphRequest $GraphRequest
    $tenantId = [string](Get-ObjectPropertyValue -InputObject $sync -PropertyName 'id')
    $features = Get-ObjectPropertyValue -InputObject $sync -PropertyName 'features'

    if (-not $features) {
        throw 'Microsoft Graph response did not include onPremisesSynchronization.features.'
    }

    return @(
        New-DirSyncProtectionRow `
            -TenantId $tenantId `
            -Features $features `
            -Setting 'BlockCloudObjectTakeover' `
            -GraphProperty 'blockCloudObjectTakeoverThroughHardMatchEnabled' `
            -Notes 'Blocks cloud object takeover through source-anchor hard match when enabled.'
        New-DirSyncProtectionRow `
            -TenantId $tenantId `
            -Features $features `
            -Setting 'BlockSoftMatch' `
            -GraphProperty 'blockSoftMatchEnabled' `
            -Notes 'Blocks soft match for all objects when enabled; Microsoft recommends keeping this enabled except during planned soft-match activity.'
    )
}

Connect-GraphIfNeeded
$rows = @(Get-DirSyncProtectionSettings)

$rows | Format-Table -AutoSize

if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
    $rows | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "Report exported: $OutputPath" -ForegroundColor Green
}
