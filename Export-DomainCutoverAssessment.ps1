<#
.SYNOPSIS
    Assesses an accepted domain ahead of a tenant-to-tenant domain cutover.

.DESCRIPTION
    Confirms the accepted-domain configuration for a target domain (live via
    Get-AcceptedDomain or from a previously exported CSV), resolves the domain's
    MX, SPF, Autodiscover, and DKIM DNS records, and analyzes shared mailboxes
    from an Objects.csv inventory: which use the target domain as primary SMTP
    or alias, and (when a MigrationReadiness.csv is supplied) which carry
    litigation or compliance holds that constrain a MailUser conversion.

    Produces the evidence needed to design a shared mailbox to MailUser
    migration and a later domain move to another tenant.

.NOTES
    Requires: ExchangeOnlineManagement (unless -SkipExchange or -AcceptedDomainsCsv is used)
    DNS lookups use Resolve-DnsName on Windows or dig on macOS/Linux.
    This script is read-only. It does not modify Exchange, Entra, or DNS.

.EXAMPLE
    ./Export-DomainCutoverAssessment.ps1 -Domain wae.com -ObjectsCsv ./export/Objects.csv

.EXAMPLE
    ./Export-DomainCutoverAssessment.ps1 -Domain wae.com -ObjectsCsv ./export/Objects.csv -MigrationReadinessCsv ./export/MigrationReadiness.csv

.EXAMPLE
    ./Export-DomainCutoverAssessment.ps1 -Domain wae.com -ObjectsCsv ./export/Objects.csv -AcceptedDomainsCsv ./WaeDomainExchangeStatus.csv -SkipExchange
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Domain,

    [Parameter(Mandatory = $false)]
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
    [string]$ObjectsCsv,

    [Parameter(Mandatory = $false)]
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
    [string]$ProxyAddressesCsv,

    [Parameter(Mandatory = $false)]
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
    [string]$MigrationReadinessCsv,

    [Parameter(Mandatory = $false)]
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
    [string]$AcceptedDomainsCsv,

    [Parameter(Mandatory = $false)]
    [string]$OutputFolder,

    [Parameter(Mandatory = $false)]
    [switch]$SkipExchange
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

function Get-DomainFromSmtpAddress {
    param([Parameter(Mandatory = $false)] [string]$Address)

    if ([string]::IsNullOrWhiteSpace($Address)) { return '' }
    $parts = $Address -split '@', 2
    if ($parts.Count -eq 2) { return $parts[1].ToLowerInvariant() }
    return ''
}

function Test-MxIsExchangeOnline {
    param([Parameter(Mandatory = $false)] [string]$MxHost)

    return ([string]$MxHost).TrimEnd('.') -match '\.mail\.protection\.outlook\.com$'
}

function Get-SpfIncludes {
    param([Parameter(Mandatory = $false)] [string]$SpfRecord)

    if ([string]::IsNullOrWhiteSpace($SpfRecord)) { return @() }
    return @([regex]::Matches($SpfRecord, 'include:(\S+)') | ForEach-Object { $_.Groups[1].Value })
}

function Get-DnsShortAnswer {
    param(
        [Parameter(Mandatory = $true)] [string]$Name,
        [Parameter(Mandatory = $true)] [ValidateSet('MX', 'TXT', 'CNAME')] [string]$Type
    )

    if (Get-Command Resolve-DnsName -ErrorAction SilentlyContinue) {
        try {
            $records = @(Resolve-DnsName -Name $Name -Type $Type -ErrorAction Stop)
            switch ($Type) {
                'MX' { return @($records | Where-Object { $_.Type -eq 'MX' } | ForEach-Object { "$($_.Preference) $($_.NameExchange)" }) }
                'TXT' { return @($records | Where-Object { $_.Type -eq 'TXT' } | ForEach-Object { ($_.Strings -join '') }) }
                'CNAME' { return @($records | Where-Object { $_.Type -eq 'CNAME' } | ForEach-Object { $_.NameHost }) }
            }
        }
        catch { return @() }
    }

    if (Get-Command dig -ErrorAction SilentlyContinue) {
        try {
            $answers = @((& dig +short $Type $Name) 2>$null | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            return @($answers | ForEach-Object { ([string]$_).Trim().Trim('"').TrimEnd('.') })
        }
        catch { return @() }
    }

    throw 'Neither Resolve-DnsName nor dig is available for DNS lookups.'
}

function Get-DnsRecordRows {
    param([Parameter(Mandatory = $true)] [string]$Domain)

    $rows = [System.Collections.Generic.List[object]]::new()

    try {
        foreach ($mx in @(Get-DnsShortAnswer -Name $Domain -Type MX)) {
            $mxHost = (($mx -split '\s+')[-1]).TrimEnd('.')
            $assessment = if (Test-MxIsExchangeOnline -MxHost $mxHost) { 'Exchange Online Protection (Microsoft 365)' } else { 'Not Exchange Online Protection - verify mail routing' }
            $rows.Add([pscustomobject]@{ Domain = $Domain; RecordType = 'MX'; Query = $Domain; Value = $mx.TrimEnd('.'); Assessment = $assessment }) | Out-Null
        }

        foreach ($txt in @(Get-DnsShortAnswer -Name $Domain -Type TXT)) {
            if ($txt -notmatch '^v=spf1') { continue }
            $includes = @(Get-SpfIncludes -SpfRecord $txt)
            $thirdParty = @($includes | Where-Object { $_ -ne 'spf.protection.outlook.com' })
            $assessment = if ($thirdParty.Count -gt 0) { "Third-party senders: $($thirdParty -join ', ')" } else { 'Exchange Online only' }
            $rows.Add([pscustomobject]@{ Domain = $Domain; RecordType = 'TXT (SPF)'; Query = $Domain; Value = $txt; Assessment = $assessment }) | Out-Null
        }

        foreach ($cname in @(Get-DnsShortAnswer -Name "autodiscover.$Domain" -Type CNAME)) {
            $value = $cname.TrimEnd('.')
            $assessment = if ($value -eq 'autodiscover.outlook.com') { 'Exchange Online Autodiscover' } else { 'Custom Autodiscover target - verify client connectivity' }
            $rows.Add([pscustomobject]@{ Domain = $Domain; RecordType = 'CNAME'; Query = "autodiscover.$Domain"; Value = $value; Assessment = $assessment }) | Out-Null
        }

        foreach ($selector in @('selector1', 'selector2')) {
            foreach ($cname in @(Get-DnsShortAnswer -Name "$selector._domainkey.$Domain" -Type CNAME)) {
                $rows.Add([pscustomobject]@{ Domain = $Domain; RecordType = 'CNAME (DKIM)'; Query = "$selector._domainkey.$Domain"; Value = $cname.TrimEnd('.'); Assessment = 'DKIM selector - must be re-created in the target tenant after cutover' }) | Out-Null
            }
        }
    }
    catch {
        Add-ExportError -Identity $Domain -Stage 'Dns' -Operation 'Get-DnsShortAnswer' -Message $_.Exception.Message
    }

    return @($rows)
}

function Get-AcceptedDomainRows {
    param([Parameter(Mandatory = $true)] [string]$TargetDomain)

    $domains = @()
    if (-not [string]::IsNullOrWhiteSpace($AcceptedDomainsCsv)) {
        $domains = @(Import-Csv -Path $AcceptedDomainsCsv)
    }
    elseif (-not $SkipExchange) {
        Connect-ExchangeOnlineIfNeeded
        $domains = @(Get-AcceptedDomain)
    }
    else {
        Write-Warning 'SkipExchange set and no -AcceptedDomainsCsv supplied; AcceptedDomains.csv will be headers-only.'
        return @()
    }

    return @($domains | ForEach-Object {
        [pscustomobject]@{
            DomainName      = [string]$_.DomainName
            DomainType      = [string]$_.DomainType
            Default         = [string]$_.Default
            InitialDomain   = [string]$_.InitialDomain
            MatchSubDomains = [string]$_.MatchSubDomains
            PendingRemoval  = [string]$_.PendingRemoval
            WhenCreated     = [string]$_.WhenCreated
            IsTargetDomain  = ([string]$_.DomainName).Equals($TargetDomain, [StringComparison]::OrdinalIgnoreCase)
        }
    })
}

function Get-SharedMailboxUsageRows {
    param(
        [Parameter(Mandatory = $true)] [string]$TargetDomain,
        [Parameter(Mandatory = $true)] [object[]]$Objects,
        [Parameter(Mandatory = $false)] [object[]]$ProxyRows,
        [Parameter(Mandatory = $false)] [object[]]$ReadinessRows
    )

    $proxyByIdentity = @{}
    foreach ($proxy in @($ProxyRows)) {
        if ($null -eq $proxy) { continue }
        $key = [string]$proxy.Identity
        if (-not $proxyByIdentity.ContainsKey($key)) { $proxyByIdentity[$key] = [System.Collections.Generic.List[object]]::new() }
        $proxyByIdentity[$key].Add($proxy) | Out-Null
    }

    $readinessByIdentity = @{}
    foreach ($readiness in @($ReadinessRows)) {
        if ($null -eq $readiness) { continue }
        $readinessByIdentity[[string]$readiness.Identity] = $readiness
    }

    $rows = [System.Collections.Generic.List[object]]::new()
    $sharedMailboxes = @($Objects | Where-Object { [string]$_.RecipientTypeDetails -eq 'SharedMailbox' })

    foreach ($mailbox in $sharedMailboxes) {
        $identity = [string]$mailbox.Identity
        $primaryDomain = Get-DomainFromSmtpAddress -Address ([string]$mailbox.PrimarySmtpAddress)

        $targetAliases = @()
        $otherAliases = @()
        foreach ($proxy in @($proxyByIdentity[$identity])) {
            if ($null -eq $proxy) { continue }
            if ([string]$proxy.AddressType -notmatch '^smtp$') { continue }
            $aliasDomain = Get-DomainFromSmtpAddress -Address ([string]$proxy.ProxyAddress)
            if ($aliasDomain -eq $TargetDomain.ToLowerInvariant()) { $targetAliases += [string]$proxy.ProxyAddress }
            else { $otherAliases += [string]$proxy.ProxyAddress }
        }

        $readiness = $readinessByIdentity[$identity]
        # HasArchive/ArchiveState only exist in readiness CSVs generated after the
        # archive check was added; tolerate older exports.
        $hasArchive = if ($readiness -and $null -ne $readiness.PSObject.Properties['HasArchive']) { [string]$readiness.HasArchive } else { '' }
        $archiveState = if ($readiness -and $null -ne $readiness.PSObject.Properties['ArchiveState']) { [string]$readiness.ArchiveState } else { '' }

        $rows.Add([pscustomobject]@{
            Identity                = $identity
            DisplayName             = [string]$mailbox.DisplayName
            PrimarySmtpAddress      = [string]$mailbox.PrimarySmtpAddress
            PrimarySmtpDomain       = $primaryDomain
            PrimaryUsesTargetDomain = ($primaryDomain -eq $TargetDomain.ToLowerInvariant())
            TargetDomainAliases     = ($targetAliases | Sort-Object -Unique) -join '; '
            OtherDomainAliases      = ($otherAliases | Sort-Object -Unique) -join '; '
            LitigationHold          = if ($readiness) { [string]$readiness.LitigationHold } else { '' }
            ComplianceHolds         = if ($readiness) { [string]$readiness.ComplianceHolds } else { '' }
            RetentionPolicy         = if ($readiness) { [string]$readiness.RetentionPolicy } else { '' }
            HasArchive              = $hasArchive
            ArchiveState            = $archiveState
            MigrationStatus         = if ($readiness) { [string]$readiness.MigrationStatus } else { '' }
            BlockingReasons         = if ($readiness) { [string]$readiness.BlockingReasons } else { '' }
        }) | Out-Null
    }

    return @($rows)
}

if ([string]::IsNullOrWhiteSpace($OutputFolder)) {
    if (-not [string]::IsNullOrWhiteSpace($ObjectsCsv)) {
        $OutputFolder = Split-Path -Parent $ObjectsCsv
    }
    else {
        $OutputFolder = ".\DomainCutoverAssessment_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    }
}
$resolvedOutputFolder = Initialize-OutputFolder -OutputFolder $OutputFolder

Write-Host "Target domain: $Domain" -ForegroundColor Cyan
Write-Host "Output folder: $resolvedOutputFolder" -ForegroundColor Cyan

# Accepted domains ---------------------------------------------------------
$acceptedDomainRows = @()
try {
    $acceptedDomainRows = @(Get-AcceptedDomainRows -TargetDomain $Domain)
}
catch {
    Add-ExportError -Identity $Domain -Stage 'AcceptedDomains' -Operation 'Get-AcceptedDomain' -Message $_.Exception.Message
}
$targetDomainRow = $acceptedDomainRows | Where-Object { $_.IsTargetDomain } | Select-Object -First 1
if ($targetDomainRow) {
    Write-Host "Accepted domain '$Domain': Type=$($targetDomainRow.DomainType) Default=$($targetDomainRow.Default)" -ForegroundColor Green
}
elseif ($acceptedDomainRows.Count -gt 0) {
    Add-ExportError -Identity $Domain -Stage 'AcceptedDomains' -Operation 'MatchTargetDomain' -Message "Domain '$Domain' was not found among the accepted domains."
}

# DNS -----------------------------------------------------------------------
$dnsRows = @(Get-DnsRecordRows -Domain $Domain)
Write-Host "DNS records resolved: $($dnsRows.Count)" -ForegroundColor Green

# Shared mailbox usage -------------------------------------------------------
$usageRows = @()
if (-not [string]::IsNullOrWhiteSpace($ObjectsCsv)) {
    $objects = @(Import-Csv -Path $ObjectsCsv)

    if ([string]::IsNullOrWhiteSpace($ProxyAddressesCsv)) {
        $candidate = Join-Path (Split-Path -Parent $ObjectsCsv) 'ProxyAddresses.csv'
        if (Test-Path -Path $candidate -PathType Leaf) { $ProxyAddressesCsv = $candidate }
    }
    if ([string]::IsNullOrWhiteSpace($MigrationReadinessCsv)) {
        $candidate = Join-Path (Split-Path -Parent $ObjectsCsv) 'MigrationReadiness.csv'
        if (Test-Path -Path $candidate -PathType Leaf) { $MigrationReadinessCsv = $candidate }
    }

    $proxyRows = if ($ProxyAddressesCsv) { @(Import-Csv -Path $ProxyAddressesCsv) } else { @() }
    $readinessRows = if ($MigrationReadinessCsv) { @(Import-Csv -Path $MigrationReadinessCsv) } else { @() }

    $usageRows = @(Get-SharedMailboxUsageRows -TargetDomain $Domain -Objects $objects -ProxyRows $proxyRows -ReadinessRows $readinessRows)
    Write-Host "Shared mailboxes analyzed: $($usageRows.Count)" -ForegroundColor Green
}
else {
    Write-Warning 'No -ObjectsCsv supplied; SharedMailboxDomainUsage.csv will be headers-only.'
}

# Summary --------------------------------------------------------------------
$mxValues = @($dnsRows | Where-Object { $_.RecordType -eq 'MX' } | ForEach-Object { $_.Value })
$spfRow = $dnsRows | Where-Object { $_.RecordType -eq 'TXT (SPF)' } | Select-Object -First 1
$autodiscoverRow = $dnsRows | Where-Object { $_.RecordType -eq 'CNAME' } | Select-Object -First 1
$spfThirdParty = @(Get-SpfIncludes -SpfRecord ([string]$spfRow.Value) | Where-Object { $_ -ne 'spf.protection.outlook.com' })
$mxIsEop = @($mxValues | Where-Object { Test-MxIsExchangeOnline -MxHost (($_ -split '\s+')[-1]) }).Count -eq $mxValues.Count -and $mxValues.Count -gt 0

$holdCount = @($usageRows | Where-Object { $_.LitigationHold -eq 'True' -or -not [string]::IsNullOrWhiteSpace($_.ComplianceHolds) }).Count
$summaryRows = @(
    [pscustomobject]@{ Metric = 'TargetDomain'; Value = $Domain }
    [pscustomobject]@{ Metric = 'AcceptedDomainType'; Value = [string]$targetDomainRow.DomainType }
    [pscustomobject]@{ Metric = 'IsDefaultAcceptedDomain'; Value = [string]$targetDomainRow.Default }
    [pscustomobject]@{ Metric = 'MxRecords'; Value = $mxValues -join '; ' }
    [pscustomobject]@{ Metric = 'MxIsExchangeOnlineProtection'; Value = [string]$mxIsEop }
    [pscustomobject]@{ Metric = 'SpfRecord'; Value = [string]$spfRow.Value }
    [pscustomobject]@{ Metric = 'SpfThirdPartyIncludes'; Value = $spfThirdParty -join '; ' }
    [pscustomobject]@{ Metric = 'AutodiscoverTarget'; Value = [string]$autodiscoverRow.Value }
    [pscustomobject]@{ Metric = 'SharedMailboxCount'; Value = [string]$usageRows.Count }
    [pscustomobject]@{ Metric = 'SharedMailboxesWithTargetDomainPrimary'; Value = [string]@($usageRows | Where-Object { $_.PrimaryUsesTargetDomain }).Count }
    [pscustomobject]@{ Metric = 'SharedMailboxesWithTargetDomainAlias'; Value = [string]@($usageRows | Where-Object { -not [string]::IsNullOrWhiteSpace($_.TargetDomainAliases) }).Count }
    [pscustomobject]@{ Metric = 'SharedMailboxesWithHolds'; Value = [string]$holdCount }
    [pscustomobject]@{ Metric = 'SharedMailboxesWithArchives'; Value = [string]@($usageRows | Where-Object { $_.HasArchive -eq 'True' }).Count }
    [pscustomobject]@{ Metric = 'SharedMailboxesBlocked'; Value = [string]@($usageRows | Where-Object { $_.MigrationStatus -eq 'Blocked' }).Count }
)

Export-CsvWithHeaders -Rows $acceptedDomainRows -Path (Join-Path $resolvedOutputFolder 'AcceptedDomains.csv') -Headers @('DomainName', 'DomainType', 'Default', 'InitialDomain', 'MatchSubDomains', 'PendingRemoval', 'WhenCreated', 'IsTargetDomain')
Export-CsvWithHeaders -Rows $dnsRows -Path (Join-Path $resolvedOutputFolder 'DnsRecords.csv') -Headers @('Domain', 'RecordType', 'Query', 'Value', 'Assessment')
Export-CsvWithHeaders -Rows $usageRows -Path (Join-Path $resolvedOutputFolder 'SharedMailboxDomainUsage.csv') -Headers @('Identity', 'DisplayName', 'PrimarySmtpAddress', 'PrimarySmtpDomain', 'PrimaryUsesTargetDomain', 'TargetDomainAliases', 'OtherDomainAliases', 'LitigationHold', 'ComplianceHolds', 'RetentionPolicy', 'HasArchive', 'ArchiveState', 'MigrationStatus', 'BlockingReasons')
Export-CsvWithHeaders -Rows $summaryRows -Path (Join-Path $resolvedOutputFolder 'CutoverImpactSummary.csv') -Headers @('Metric', 'Value')
Export-CsvWithHeaders -Rows @($script:Errors) -Path (Join-Path $resolvedOutputFolder 'Errors-DomainCutover.csv') -Headers @('Identity', 'Stage', 'Operation', 'Message')

Write-Host ''
Write-Host 'Cutover impact summary:' -ForegroundColor Cyan
$summaryRows | ForEach-Object { Write-Host ("  {0}: {1}" -f $_.Metric, $_.Value) }
Write-Host "Errors logged: $($script:Errors.Count)" -ForegroundColor Yellow
