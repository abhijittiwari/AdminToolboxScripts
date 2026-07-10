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
            return ($typeDetails -ne 'GroupMailbox' -and ($typeDetails -match 'Mailbox$' -or $type -eq 'UserMailbox'))
        }
        'Group' {
            return ($typeDetails -eq 'GroupMailbox' -or $typeDetails -match 'Group$' -or $typeDetails -match 'DistributionGroup' -or $typeDetails -eq 'MailUniversalSecurityGroup')
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

function Get-ExchangeLookupIdentity {
    param([Parameter(Mandatory = $true)] $Recipient)

    foreach ($candidate in @(
            $Recipient.PrimarySmtpAddress,
            $Recipient.Alias,
            $Recipient.DistinguishedName,
            $Recipient.Guid,
            $Recipient.Identity
        )) {
        $value = [string]$candidate
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            return $value
        }
    }

    return [string]$Recipient.Identity
}

function Get-ProxyAddressRows {
    param([Parameter(Mandatory = $true)] $Recipient)

    $lookupIdentity = Get-ExchangeLookupIdentity -Recipient $Recipient

    try {
        $recipientDetails = Get-Recipient -Identity $lookupIdentity -ErrorAction Stop
        $rows = New-Object System.Collections.Generic.List[object]
        foreach ($address in @($recipientDetails.EmailAddresses)) {
            if ($null -eq $address) { continue }
            $value = $address.ToString()
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
    catch {
        Add-ExportError -Identity ([string]$Recipient.Identity) -Stage 'ProxyAddresses' -Operation 'Get-Recipient EmailAddresses' -Message "$($_.Exception.Message) (lookup: $lookupIdentity)"
        return @()
    }
}

function Get-FullAccessRows {
    param([Parameter(Mandatory = $true)] $Recipient)

    $mailboxType = Get-MailboxType -Recipient $Recipient
    if ([string]::IsNullOrWhiteSpace($mailboxType)) {
        return @()
    }

    $lookupIdentity = Get-ExchangeLookupIdentity -Recipient $Recipient

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

function Get-SendAsRows {
    param([Parameter(Mandatory = $true)] $Recipient)

    $lookupIdentity = Get-ExchangeLookupIdentity -Recipient $Recipient

    try {
        return @(
            Get-RecipientPermission -Identity $lookupIdentity -ErrorAction Stop |
                Where-Object { -not $_.IsInherited -and -not (Test-IsSelfTrustee -Trustee ([string]$_.Trustee)) -and ($_.AccessRights -contains 'SendAs') } |
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
        Add-ExportError -Identity ([string]$Recipient.Identity) -Stage 'SendAs' -Operation 'Get-RecipientPermission' -Message "$($_.Exception.Message) (lookup: $lookupIdentity)"
        return @()
    }
}

function Get-ExchangeGroupMembershipRows {
    param([Parameter(Mandatory = $true)] $Recipient)

    $lookupIdentity = Get-ExchangeLookupIdentity -Recipient $Recipient

    try {
        return @(
            Get-Recipient -Identity $lookupIdentity -ErrorAction Stop |
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
        Add-ExportError -Identity ([string]$Recipient.Identity) -Stage 'ExchangeGroupMemberships' -Operation 'Get-Recipient MemberOfGroup' -Message "$($_.Exception.Message) (lookup: $lookupIdentity)"
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

function Join-Values {
    param([Parameter(Mandatory = $false)] $Values)

    return (@($Values) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Sort-Object -Unique) -join '; '
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

    Export-CsvWithHeaders -Rows $ConsolidatedRows -Path (Join-Path $Folder 'Objects-Consolidated.csv') -Headers @('Identity', 'DisplayName', 'RecipientType', 'RecipientTypeDetails', 'MailboxType', 'PrimarySmtpAddress', 'Alias', 'ExternalDirectoryObjectId', 'ExchangeObjectId', 'ProxyAddresses', 'FullAccessTrustees', 'SendAsTrustees', 'ExchangeGroupMemberships', 'EntraGroupMemberships', 'HasErrors')
    Export-CsvWithHeaders -Rows $Objects -Path (Join-Path $Folder 'Objects.csv') -Headers @('Identity', 'DisplayName', 'RecipientType', 'RecipientTypeDetails', 'MailboxType', 'PrimarySmtpAddress', 'Alias', 'ExternalDirectoryObjectId', 'ExchangeObjectId', 'DistinguishedName')
    Export-CsvWithHeaders -Rows $ProxyRows -Path (Join-Path $Folder 'ProxyAddresses.csv') -Headers @('Identity', 'PrimarySmtpAddress', 'AddressType', 'ProxyAddress', 'RawProxyAddress', 'IsPrimarySmtp')
    Export-CsvWithHeaders -Rows $FullAccessRows -Path (Join-Path $Folder 'FullAccess.csv') -Headers @('Identity', 'PrimarySmtpAddress', 'Trustee', 'AccessRights', 'Deny', 'IsInherited')
    Export-CsvWithHeaders -Rows $SendAsRows -Path (Join-Path $Folder 'SendAs.csv') -Headers @('Identity', 'PrimarySmtpAddress', 'Trustee', 'AccessRights', 'IsInherited')
    Export-CsvWithHeaders -Rows $ExchangeGroupRows -Path (Join-Path $Folder 'ExchangeGroupMemberships.csv') -Headers @('Identity', 'PrimarySmtpAddress', 'GroupIdentity')
    Export-CsvWithHeaders -Rows $EntraGroupRows -Path (Join-Path $Folder 'EntraGroupMemberships.csv') -Headers @('Identity', 'PrimarySmtpAddress', 'GroupId', 'GroupDisplayName', 'GroupMail', 'GroupSecurityEnabled')
    Export-CsvWithHeaders -Rows $ErrorRows -Path (Join-Path $Folder 'Errors.csv') -Headers @('Identity', 'Stage', 'Operation', 'Message')
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

if ($isMultiObjectRun -and -not $NoDashboard) {
    try {
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

Write-Host "Objects exported: $($objects.Count)" -ForegroundColor Green
Write-Host "Errors logged: $($script:Errors.Count)" -ForegroundColor Yellow

try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch { }
if ($graphConnected) {
    try { Disconnect-MgGraph | Out-Null } catch { }
}
