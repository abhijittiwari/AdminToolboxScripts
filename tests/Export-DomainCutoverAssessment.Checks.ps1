#Requires -Version 7
$ErrorActionPreference = 'Stop'

$script:ScriptPath = Join-Path $PSScriptRoot '..' 'Export-DomainCutoverAssessment.ps1'
$script:Failures = [System.Collections.Generic.List[string]]::new()
$script:Passes = 0

function Get-ScriptAst {
    $tokens = $null
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:ScriptPath, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count -gt 0) {
        throw "Parse errors in $($script:ScriptPath): $($parseErrors[0].Message)"
    }
    return $ast
}

function Import-ScriptFunctions {
    $ast = Get-ScriptAst
    $functions = $ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)
    foreach ($functionAst in $functions) {
        . ([scriptblock]::Create($functionAst.Extent.Text))
    }
}

function Assert-True {
    param(
        [Parameter(Mandatory = $true)] [bool]$Condition,
        [Parameter(Mandatory = $true)] [string]$Name
    )
    if ($Condition) {
        $script:Passes++
        Write-Host "PASS $Name" -ForegroundColor Green
    }
    else {
        $script:Failures.Add($Name) | Out-Null
        Write-Host "FAIL $Name" -ForegroundColor Red
    }
}

$script:Errors = [System.Collections.Generic.List[object]]::new()

. Import-ScriptFunctions

# ---------------------------------------------------------------------------
# Checks
# ---------------------------------------------------------------------------

function Test-DomainParsing {
    Assert-True -Condition ((Get-DomainFromSmtpAddress -Address 'Info@WAE.com') -eq 'wae.com') -Name 'Get-DomainFromSmtpAddress lowercases the domain'
    Assert-True -Condition ((Get-DomainFromSmtpAddress -Address '') -eq '') -Name 'Get-DomainFromSmtpAddress returns empty for blank input'
    Assert-True -Condition ((Get-DomainFromSmtpAddress -Address 'not-an-address') -eq '') -Name 'Get-DomainFromSmtpAddress returns empty without @'
}
Test-DomainParsing

function Test-MxClassification {
    Assert-True -Condition (Test-MxIsExchangeOnline -MxHost 'wae-com.mail.protection.outlook.com.') -Name 'Test-MxIsExchangeOnline matches EOP host with trailing dot'
    Assert-True -Condition (-not (Test-MxIsExchangeOnline -MxHost 'mx1.mimecast.com')) -Name 'Test-MxIsExchangeOnline rejects non-EOP host'
    Assert-True -Condition (-not (Test-MxIsExchangeOnline -MxHost 'evil-mail.protection.outlook.com.example.net')) -Name 'Test-MxIsExchangeOnline anchors the suffix match'
}
Test-MxClassification

function Test-SpfParsing {
    $includes = @(Get-SpfIncludes -SpfRecord 'v=spf1 include:spf.protection.outlook.com include:mailgun.org include:docebosaas.com -all')
    Assert-True -Condition ($includes.Count -eq 3 -and 'mailgun.org' -in $includes) -Name 'Get-SpfIncludes extracts all include mechanisms'
    Assert-True -Condition (@(Get-SpfIncludes -SpfRecord '').Count -eq 0) -Name 'Get-SpfIncludes returns empty for blank record'
}
Test-SpfParsing

function Test-SharedMailboxUsage {
    $objects = @(
        [pscustomobject]@{ Identity = 'sm1'; DisplayName = 'Shared One'; RecipientTypeDetails = 'SharedMailbox'; PrimarySmtpAddress = 'one@wae.com' }
        [pscustomobject]@{ Identity = 'sm2'; DisplayName = 'Shared Two'; RecipientTypeDetails = 'SharedMailbox'; PrimarySmtpAddress = 'two@waetechnology.com' }
        [pscustomobject]@{ Identity = 'u1'; DisplayName = 'User One'; RecipientTypeDetails = 'UserMailbox'; PrimarySmtpAddress = 'user@wae.com' }
    )
    $proxies = @(
        [pscustomobject]@{ Identity = 'sm2'; AddressType = 'smtp'; ProxyAddress = 'two@wae.com' }
        [pscustomobject]@{ Identity = 'sm2'; AddressType = 'smtp'; ProxyAddress = 'two@tenant.onmicrosoft.com' }
        [pscustomobject]@{ Identity = 'sm2'; AddressType = 'X500'; ProxyAddress = '/o=ExchangeLabs/ou=x' }
    )
    $readiness = @(
        [pscustomobject]@{ Identity = 'sm1'; LitigationHold = 'True'; ComplianceHolds = ''; RetentionPolicy = ''; MigrationStatus = 'Blocked'; BlockingReasons = 'litigation hold' }
    )

    $rows = @(Get-SharedMailboxUsageRows -TargetDomain 'WAE.com' -Objects $objects -ProxyRows $proxies -ReadinessRows $readiness)

    Assert-True -Condition ($rows.Count -eq 2) -Name 'Get-SharedMailboxUsageRows includes only shared mailboxes'
    $sm1 = $rows | Where-Object { $_.Identity -eq 'sm1' }
    Assert-True -Condition ($sm1.PrimaryUsesTargetDomain -eq $true) -Name 'Get-SharedMailboxUsageRows flags target-domain primary SMTP'
    Assert-True -Condition ($sm1.LitigationHold -eq 'True' -and $sm1.MigrationStatus -eq 'Blocked') -Name 'Get-SharedMailboxUsageRows joins readiness hold data'
    $sm2 = $rows | Where-Object { $_.Identity -eq 'sm2' }
    Assert-True -Condition ($sm2.PrimaryUsesTargetDomain -eq $false -and $sm2.TargetDomainAliases -eq 'two@wae.com') -Name 'Get-SharedMailboxUsageRows captures target-domain aliases'
    Assert-True -Condition ($sm2.OtherDomainAliases -eq 'two@tenant.onmicrosoft.com') -Name 'Get-SharedMailboxUsageRows keeps non-target smtp aliases and drops X500'

    # Readiness rows without archive columns (older exports) must not fail.
    Assert-True -Condition ($sm1.HasArchive -eq '' -and $sm1.ArchiveState -eq '') -Name 'Get-SharedMailboxUsageRows tolerates readiness CSVs without archive columns'

    # Archive columns pass through when present.
    $readinessWithArchive = @(
        [pscustomobject]@{ Identity = 'sm1'; LitigationHold = 'False'; ComplianceHolds = ''; RetentionPolicy = ''; HasArchive = 'True'; ArchiveState = 'HostedProvisioned'; MigrationStatus = 'Review'; BlockingReasons = 'online archive (resolve before MailUser conversion)' }
    )
    $rows = @(Get-SharedMailboxUsageRows -TargetDomain 'wae.com' -Objects $objects -ProxyRows $proxies -ReadinessRows $readinessWithArchive)
    $sm1a = $rows | Where-Object { $_.Identity -eq 'sm1' }
    Assert-True -Condition ($sm1a.HasArchive -eq 'True' -and $sm1a.ArchiveState -eq 'HostedProvisioned') -Name 'Get-SharedMailboxUsageRows passes archive columns through'
}
Test-SharedMailboxUsage

function Test-ScriptStaticContract {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw
    Assert-True -Condition ($scriptText -match 'AcceptedDomains\.csv') -Name 'Script writes AcceptedDomains.csv'
    Assert-True -Condition ($scriptText -match 'DnsRecords\.csv') -Name 'Script writes DnsRecords.csv'
    Assert-True -Condition ($scriptText -match 'SharedMailboxDomainUsage\.csv') -Name 'Script writes SharedMailboxDomainUsage.csv'
    Assert-True -Condition ($scriptText -match 'CutoverImpactSummary\.csv') -Name 'Script writes CutoverImpactSummary.csv'
    Assert-True -Condition ($scriptText -match 'Errors-DomainCutover\.csv') -Name 'Script writes job-specific errors CSV'
    Assert-True -Condition ($scriptText -match 'Get-AcceptedDomain') -Name 'Script queries accepted domains'
    Assert-True -Condition ($scriptText -match 'Resolve-DnsName' -and $scriptText -match '\bdig\b') -Name 'Script supports Resolve-DnsName with dig fallback'
    Assert-True -Condition ($scriptText -match '\$SkipExchange' -and $scriptText -match '\$AcceptedDomainsCsv') -Name 'Script supports offline runs (-SkipExchange / -AcceptedDomainsCsv)'
}
Test-ScriptStaticContract

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
if ($script:Failures.Count -gt 0) {
    Write-Host "$($script:Failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host "All $($script:Passes) checks passed." -ForegroundColor Green
exit 0
