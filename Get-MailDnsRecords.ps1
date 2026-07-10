<#
.SYNOPSIS
    Fetches MX, SPF, DKIM, and DMARC DNS records for a domain.

.NOTES
    DKIM records require one or more selectors. They are not reliably discoverable
    from DNS alone, so this script checks a default list and lets you pass your own.

.EXAMPLES
    .\Get-MailDnsRecords.ps1 -Domain example.com

    .\Get-MailDnsRecords.ps1 -Domain example.com -DkimSelectors selector1,selector2,google

    .\Get-MailDnsRecords.ps1 -Domain example.com -Server 8.8.8.8 -Json
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Domain,

    [string[]]$DkimSelectors = @(
        "selector1",
        "selector2",
        "google",
        "default",
        "dkim",
        "mail",
        "smtp",
        "s1",
        "s2",
        "k1",
        "k2"
    ),

    [string]$Server,

    [switch]$Json
)

function Normalize-Domain {
    param([string]$InputDomain)

    $clean = $InputDomain.Trim().ToLowerInvariant()
    $clean = $clean -replace '^https?://', ''
    $clean = $clean -replace '/.*$', ''
    $clean = $clean.TrimEnd('.')

    return $clean
}

function Resolve-Record {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [string]$Type
    )

    $params = @{
        Name        = $Name
        Type        = $Type
        ErrorAction = "Stop"
    }

    if (-not [string]::IsNullOrWhiteSpace($Server)) {
        $params["Server"] = $Server
    }

    try {
        return @(Resolve-DnsName @params)
    }
    catch {
        return @()
    }
}

function Get-RecordProperty {
    param(
        [object]$Record,
        [string]$PropertyName
    )

    $property = $Record.PSObject.Properties[$PropertyName]

    if ($null -ne $property) {
        return $property.Value
    }

    return $null
}

function Get-TxtValue {
    param([object]$Record)

    $strings = Get-RecordProperty -Record $Record -PropertyName "Strings"

    if ($strings) {
        return ($strings -join "")
    }

    $text = Get-RecordProperty -Record $Record -PropertyName "Text"

    if ($text) {
        return [string]$text
    }

    $descriptiveText = Get-RecordProperty -Record $Record -PropertyName "DescriptiveText"

    if ($descriptiveText) {
        return ([string]$descriptiveText).Trim('"')
    }

    return $null
}

function Get-TxtRecords {
    param([string]$Name)

    return @(
        Resolve-Record -Name $Name -Type TXT |
            Where-Object { $_.Type -eq "TXT" } |
            ForEach-Object {
                $txt = Get-TxtValue -Record $_

                if (-not [string]::IsNullOrWhiteSpace($txt)) {
                    [pscustomobject]@{
                        Name = $_.Name
                        TTL  = $_.TTL
                        Text = $txt
                    }
                }
            }
    )
}

function Get-CnameRecord {
    param([string]$Name)

    $record = Resolve-Record -Name $Name -Type CNAME |
        Where-Object { $_.Type -eq "CNAME" } |
        Select-Object -First 1

    if ($null -eq $record) {
        return $null
    }

    return [pscustomobject]@{
        Name   = $record.Name
        Target = $record.NameHost.TrimEnd(".")
        TTL    = $record.TTL
    }
}

function Get-DkimQueryName {
    param(
        [string]$Selector,
        [string]$BaseDomain
    )

    $selector = $Selector.Trim().TrimEnd(".")

    if ($selector -match "\._domainkey\.") {
        return $selector
    }

    if ($selector -match "\._domainkey$") {
        return "$selector.$BaseDomain"
    }

    return "$selector._domainkey.$BaseDomain"
}

function Get-DkimRecords {
    param(
        [string]$BaseDomain,
        [string[]]$Selectors
    )

    $uniqueSelectors = $Selectors |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Select-Object -Unique

    foreach ($selector in $uniqueSelectors) {
        $queryName = Get-DkimQueryName -Selector $selector -BaseDomain $BaseDomain

        $txtRecords = @(Get-TxtRecords -Name $queryName)
        $cname = Get-CnameRecord -Name $queryName

        if ($txtRecords.Count -eq 0 -and $null -ne $cname) {
            $txtRecords = @(Get-TxtRecords -Name $cname.Target)
        }

        $txtValues = @($txtRecords | ForEach-Object { $_.Text })

        $looksLikeDkim = $false

        foreach ($txt in $txtValues) {
            if ($txt -match "^\s*v=DKIM1\b" -or $txt -match ";\s*p=" -or $txt -match "^\s*k=rsa\s*;") {
                $looksLikeDkim = $true
                break
            }
        }

        $status = if ($txtValues.Count -eq 0) {
            "Not found"
        }
        elseif ($looksLikeDkim) {
            "Found"
        }
        else {
            "TXT found; verify manually"
        }

        [pscustomobject]@{
            Selector  = $selector
            QueryName = $queryName
            CNAME     = if ($null -ne $cname) { $cname.Target } else { $null }
            Status    = $status
            TXT       = $txtValues
        }
    }
}

if (-not (Get-Command Resolve-DnsName -ErrorAction SilentlyContinue)) {
    throw "Resolve-DnsName was not found. This script requires the DnsClient module, which is available on Windows."
}

$Domain = Normalize-Domain -InputDomain $Domain

$mxRecords = @(
    Resolve-Record -Name $Domain -Type MX |
        Where-Object { $_.Type -eq "MX" } |
        Sort-Object Preference, NameExchange |
        ForEach-Object {
            [pscustomobject]@{
                Preference = $_.Preference
                Exchange   = $_.NameExchange.TrimEnd(".")
                TTL        = $_.TTL
            }
        }
)

$domainTxtRecords = @(Get-TxtRecords -Name $Domain)

$spfRecords = @(
    $domainTxtRecords |
        Where-Object { $_.Text -match "^\s*v=spf1(\s|$)" }
)

$dmarcName = "_dmarc.$Domain"

$dmarcRecords = @(
    Get-TxtRecords -Name $dmarcName |
        Where-Object { $_.Text -match "^\s*v=DMARC1(\s*;|$)" }
)

$dkimRecords = @(Get-DkimRecords -BaseDomain $Domain -Selectors $DkimSelectors)

$warnings = New-Object System.Collections.Generic.List[string]

if ($spfRecords.Count -gt 1) {
    $warnings.Add("Multiple SPF records found. A domain should normally publish exactly one SPF TXT record.")
}

if ($dmarcRecords.Count -gt 1) {
    $warnings.Add("Multiple DMARC records found. A domain should normally publish exactly one DMARC TXT record at _dmarc.$Domain.")
}

if ($dkimRecords.Count -eq 0) {
    $warnings.Add("No DKIM selectors were checked.")
}

$result = [pscustomobject][ordered]@{
    Domain    = $Domain
    QueriedAt = Get-Date
    DnsServer = if ([string]::IsNullOrWhiteSpace($Server)) { "System default" } else { $Server }
    MX        = $mxRecords
    SPF       = $spfRecords
    DMARC     = $dmarcRecords
    DKIM      = $dkimRecords
    Warnings  = @($warnings)
}

if ($Json) {
    $result | ConvertTo-Json -Depth 10
    exit
}

Write-Host ""
Write-Host "Domain:    $($result.Domain)"
Write-Host "Queried:   $($result.QueriedAt)"
Write-Host "DNS:       $($result.DnsServer)"
Write-Host ""

Write-Host "MX Records"
Write-Host "----------"
if ($result.MX.Count -gt 0) {
    $result.MX | Format-Table -AutoSize
}
else {
    Write-Host "No MX records found."
}
Write-Host ""

Write-Host "SPF Records"
Write-Host "-----------"
if ($result.SPF.Count -gt 0) {
    $result.SPF | Format-Table Name, TTL, Text -Wrap -AutoSize
}
else {
    Write-Host "No SPF record found at $Domain."
}
Write-Host ""

Write-Host "DMARC Records"
Write-Host "-------------"
if ($result.DMARC.Count -gt 0) {
    $result.DMARC | Format-Table Name, TTL, Text -Wrap -AutoSize
}
else {
    Write-Host "No DMARC record found at $dmarcName."
}
Write-Host ""

Write-Host "DKIM Records"
Write-Host "------------"
if ($result.DKIM.Count -gt 0) {
    foreach ($record in $result.DKIM) {
        Write-Host "Selector:  $($record.Selector)"
        Write-Host "Query:     $($record.QueryName)"
        Write-Host "Status:    $($record.Status)"

        if ($record.CNAME) {
            Write-Host "CNAME:     $($record.CNAME)"
        }

        if ($record.TXT.Count -gt 0) {
            Write-Host "TXT:"
            foreach ($txt in $record.TXT) {
                Write-Host "  $txt"
            }
        }

        Write-Host ""
    }
}
else {
    Write-Host "No DKIM records checked."
}
Write-Host ""

if ($result.Warnings.Count -gt 0) {
    Write-Host "Warnings"
    Write-Host "--------"
    foreach ($warning in $result.Warnings) {
        Write-Host "- $warning"
    }
    Write-Host ""
}