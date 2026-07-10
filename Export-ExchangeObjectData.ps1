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

$resolvedOutputFolder = Initialize-OutputFolder -OutputFolder $OutputFolder
$isMultiObjectRun = $PSCmdlet.ParameterSetName -in @('All', 'Csv')

Write-Host "Output folder: $resolvedOutputFolder" -ForegroundColor Cyan
Write-Host "Recipient type: $RecipientType" -ForegroundColor Cyan
