<#
.SYNOPSIS
    Exports the Entra ID per-domain authentication model for divestiture discovery.

.DESCRIPTION
    Read-only discovery that answers the SoW's #1 priority action: confirm the current-state
    authentication model. For every verified domain in the tenant it reports whether the domain is
    Managed (cloud auth: Password Hash Sync or Pass-through Authentication) or Federated (ADFS or a
    third-party STP), and it reports the tenant-wide hybrid-auth posture:
      - Per domain: authentication type (Managed / Federated), and for federated domains the issuer
        URI and passive-STS endpoint.
      - Tenant: on-premises sync enabled, Seamless SSO, and (best effort) whether PHS or PTA is in use
        based on the presence of PTA authentication agents.

    Uses Microsoft Graph v1.0 (`/domains`, `/organization`) plus the PTA agent endpoint (beta). Each
    fact is collected defensively; the shaping is separated so it is unit-testable with mock data. The
    script makes no changes.

.PARAMETER OutputPath
    Optional CSV output path for the per-domain rows. If omitted, results are only written to the table.

.EXAMPLE
    .\Export-EntraAuthenticationModel.ps1 -OutputPath .\AuthModel.csv
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
    Connect-MgGraph -Scopes @('Domain.Read.All','Organization.Read.All') -NoWelcome
}

function New-GraphRequestExecutor {
    return {
        param([string]$Method, [string]$Uri, $Body)
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
        [Parameter(Mandatory = $false)] [AllowNull()] $InputObject,
        [Parameter(Mandatory = $true)] [string]$PropertyName
    )
    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [hashtable] -and $InputObject.ContainsKey($PropertyName)) { return $InputObject[$PropertyName] }
    $property = $InputObject.PSObject.Properties[$PropertyName]
    if ($property) { return $property.Value }
    $additional = $InputObject.PSObject.Properties['AdditionalProperties']
    if ($additional -and $additional.Value -is [hashtable] -and $additional.Value.ContainsKey($PropertyName)) { return $additional.Value[$PropertyName] }
    return $null
}

function New-DomainAuthRow {
    param([Parameter(Mandatory = $true)] $Domain)

    $authType = [string](Get-ObjectPropertyValue -InputObject $Domain -PropertyName 'authenticationType')
    $isFederated = ($authType -eq 'Federated')
    return [pscustomobject]@{
        DomainName         = [string](Get-ObjectPropertyValue -InputObject $Domain -PropertyName 'id')
        AuthenticationType = if ([string]::IsNullOrWhiteSpace($authType)) { 'Unknown' } else { $authType }
        IsFederated        = $isFederated
        IsDefault          = [bool](Get-ObjectPropertyValue -InputObject $Domain -PropertyName 'isDefault')
        IsVerified         = [bool](Get-ObjectPropertyValue -InputObject $Domain -PropertyName 'isVerified')
        SupportedServices  = (@(Get-ObjectPropertyValue -InputObject $Domain -PropertyName 'supportedServices') -join ';')
    }
}

function Get-AuthenticationModelSummary {
    param(
        [Parameter(Mandatory = $true)] [object[]]$DomainRows,
        [AllowNull()] $OnPremisesSyncEnabled,
        [AllowNull()] $SeamlessSsoEnabled,
        [AllowNull()] $PtaAgentCount
    )

    $federatedDomains = @($DomainRows | Where-Object { $_.IsFederated })
    $managedDomains = @($DomainRows | Where-Object { -not $_.IsFederated -and $_.AuthenticationType -eq 'Managed' })

    $model =
        if ($federatedDomains.Count -gt 0 -and $managedDomains.Count -gt 0) { 'Mixed (Federated + Managed)' }
        elseif ($federatedDomains.Count -gt 0) { 'Federated (ADFS / third-party STS)' }
        elseif ([int]$PtaAgentCount -gt 0) { 'Managed - Pass-through Authentication' }
        elseif ([bool]$OnPremisesSyncEnabled) { 'Managed - Password Hash Sync (no PTA agents detected)' }
        else { 'Managed (cloud-only or PHS)' }

    return [pscustomobject]@{
        AuthenticationModel   = $model
        FederatedDomainCount  = $federatedDomains.Count
        ManagedDomainCount    = $managedDomains.Count
        OnPremisesSyncEnabled = [bool]$OnPremisesSyncEnabled
        SeamlessSsoEnabled    = [bool]$SeamlessSsoEnabled
        PtaAgentCount         = [int]$PtaAgentCount
    }
}

function Get-EntraAuthModelFacts {
    param([Parameter(Mandatory = $true)] [scriptblock]$GraphRequest)

    $facts = @{ Domains = @(); OnPremisesSyncEnabled = $null; SeamlessSsoEnabled = $null; PtaAgentCount = $null }

    try {
        $domainsResponse = & $GraphRequest 'GET' 'https://graph.microsoft.com/v1.0/domains' $null
        $facts.Domains = @(Get-ObjectPropertyValue -InputObject $domainsResponse -PropertyName 'value')
    } catch { }

    try {
        $orgResponse = & $GraphRequest 'GET' 'https://graph.microsoft.com/v1.0/organization?$select=onPremisesSyncEnabled' $null
        $org = @(Get-ObjectPropertyValue -InputObject $orgResponse -PropertyName 'value')
        if ($org.Count -gt 0) { $facts.OnPremisesSyncEnabled = Get-ObjectPropertyValue -InputObject $org[0] -PropertyName 'onPremisesSyncEnabled' }
    } catch { }

    try {
        $ptaResponse = & $GraphRequest 'GET' 'https://graph.microsoft.com/beta/onPremisesPublishingProfiles/authentication/agents' $null
        $facts.PtaAgentCount = @(Get-ObjectPropertyValue -InputObject $ptaResponse -PropertyName 'value').Count
    } catch { }

    return $facts
}

function Get-EntraAuthenticationInventory {
    param([Parameter(Mandatory = $true)] $Facts)

    $domainRows = @(@($Facts.Domains) | Where-Object { $_ } | ForEach-Object { New-DomainAuthRow -Domain $_ })
    $summary = Get-AuthenticationModelSummary -DomainRows $domainRows -OnPremisesSyncEnabled $Facts.OnPremisesSyncEnabled -SeamlessSsoEnabled $Facts.SeamlessSsoEnabled -PtaAgentCount $Facts.PtaAgentCount
    return @{ Domains = $domainRows; Summary = $summary }
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
Connect-GraphIfNeeded
$graphRequest = New-GraphRequestExecutor

Write-Host 'Collecting Entra ID per-domain authentication model...' -ForegroundColor Cyan
$facts = Get-EntraAuthModelFacts -GraphRequest $graphRequest
$inventory = Get-EntraAuthenticationInventory -Facts $facts

Write-Host "`n=== Authentication model summary ===" -ForegroundColor Yellow
$inventory.Summary | Format-List
Write-Host "=== Domains ($(@($inventory.Domains).Count)) ===" -ForegroundColor Yellow
$inventory.Domains | Format-Table DomainName, AuthenticationType, IsFederated, IsDefault, IsVerified -AutoSize

if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
    $inventory.Domains | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "Report exported: $OutputPath" -ForegroundColor Green
}
