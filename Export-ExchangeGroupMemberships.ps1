<#
.SYNOPSIS
    Exports Entra ID and Exchange group memberships for recipients in an Objects.csv inventory.

.DESCRIPTION
    Group membership job of the modular Exchange object export. Reads the Objects.csv produced
    by Export-ExchangeObjectInventory.ps1 (or Export-ExchangeObjectData.ps1) and collects group
    memberships via Microsoft Graph $batch memberOf queries. No Exchange session is needed.
    Reuses an existing Microsoft Graph context when one exists and leaves it connected so
    chained job runs authenticate once.

.NOTES
    Requires: Microsoft.Graph.Authentication
    This script is read-only. It does not modify Exchange or Entra objects.

.EXAMPLE
    ./Export-ExchangeGroupMemberships.ps1 -ObjectsCsv ./ExchangeObjectInventory_20260714_120000/Objects.csv
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
    [string]$ObjectsCsv,

    [Parameter(Mandatory = $false)]
    [string]$OutputFolder
)

$ErrorActionPreference = 'Stop'

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

function Initialize-OutputFolder {
    param([Parameter(Mandatory = $true)] [string]$OutputFolder)

    $resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputFolder)
    if (-not (Test-Path -Path $resolvedPath)) {
        New-Item -Path $resolvedPath -ItemType Directory -Force | Out-Null
    }
    return $resolvedPath
}

function Import-ObjectsCsv {
    param(
        [Parameter(Mandatory = $true)] [string]$Path,
        [Parameter(Mandatory = $true)] [string[]]$RequiredColumns
    )

    $rows = @(Import-Csv -Path $Path)
    if ($rows.Count -eq 0) {
        throw "Objects CSV '$Path' contains no rows. Run Export-ExchangeObjectInventory.ps1 first."
    }

    foreach ($column in $RequiredColumns) {
        if ($null -eq $rows[0].PSObject.Properties[$column]) {
            throw "Objects CSV '$Path' does not contain required column '$column'. Expected an Objects.csv produced by Export-ExchangeObjectInventory.ps1."
        }
    }

    return $rows
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

function Connect-GraphIfNeeded {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    if (Get-MgContext) {
        Write-Host 'Reusing existing Microsoft Graph context.' -ForegroundColor Cyan
        return
    }
    Connect-MgGraph -Scopes @('User.Read.All', 'Group.Read.All', 'Directory.Read.All') -NoWelcome
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
            Write-Progress -Id 1 -Activity $Activity -Status "attempt $attempt, $([Math]::Min($offset + 20, $pending.Count)) of $($pending.Count)" -PercentComplete ((($totalCount - $pending.Count + $offset) / $totalCount) * 100)
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

# memberOf is only defined on concrete Graph types (user, group, orgContact) - the
# directoryObject base type has no memberOf navigation in v1.0 $metadata.
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
            Write-Progress -Id 1 -Activity 'Processing group membership results' -Status "$membershipIndex of $($responseKeys.Count)" -PercentComplete (($membershipIndex / [Math]::Max($responseKeys.Count, 1)) * 100)
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

$resolvedObjectsCsv = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ObjectsCsv)
if ([string]::IsNullOrWhiteSpace($OutputFolder)) {
    $OutputFolder = Split-Path -Parent $resolvedObjectsCsv
}
$resolvedOutputFolder = Initialize-OutputFolder -OutputFolder $OutputFolder

Write-Host "Objects CSV: $resolvedObjectsCsv" -ForegroundColor Cyan
Write-Host "Output folder: $resolvedOutputFolder" -ForegroundColor Cyan

$recipients = @(Import-ObjectsCsv -Path $resolvedObjectsCsv -RequiredColumns @('Identity', 'PrimarySmtpAddress', 'ExternalDirectoryObjectId', 'RecipientTypeDetails', 'RecipientType'))
Write-Host "Objects loaded: $($recipients.Count)" -ForegroundColor Green

Connect-GraphIfNeeded

Write-Host 'Collecting group memberships via Microsoft Graph...' -ForegroundColor Cyan
$membershipRows = Get-MembershipRows -Recipients $recipients

Export-CsvWithHeaders -Rows @($membershipRows.Entra) -Path (Join-Path $resolvedOutputFolder 'EntraGroupMemberships.csv') -Headers @('Identity', 'PrimarySmtpAddress', 'GroupId', 'GroupDisplayName', 'GroupMail', 'GroupSecurityEnabled')
Export-CsvWithHeaders -Rows @($membershipRows.Exchange) -Path (Join-Path $resolvedOutputFolder 'ExchangeGroupMemberships.csv') -Headers @('Identity', 'PrimarySmtpAddress', 'GroupIdentity')
Export-CsvWithHeaders -Rows @($script:Errors) -Path (Join-Path $resolvedOutputFolder 'Errors-GroupMemberships.csv') -Headers @('Identity', 'Stage', 'Operation', 'Message')

Write-Host "Entra membership rows: $(@($membershipRows.Entra).Count)" -ForegroundColor Green
Write-Host "Exchange membership rows: $(@($membershipRows.Exchange).Count)" -ForegroundColor Green
Write-Host "Errors logged: $($script:Errors.Count)" -ForegroundColor Yellow
