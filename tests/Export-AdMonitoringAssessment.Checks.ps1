#Requires -Version 7
$ErrorActionPreference = 'Stop'

$script:ScriptPath = Join-Path $PSScriptRoot '..' 'Export-AdMonitoringAssessment.ps1'
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

. Import-ScriptFunctions

function Test-RequiredFunctionsExist {
    $ast = Get-ScriptAst
    $functionNames = @($ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true) | ForEach-Object Name)
    foreach ($name in @('New-ControlRow','Get-MonDcFacts','Test-MonDcControls','Test-MonDomainControls','Get-MonDomainFacts')) {
        Assert-True -Condition ($functionNames -contains $name) -Name "Function exists: $name"
    }
}
Test-RequiredFunctionsExist

function Test-StaticScriptContent {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw
    Assert-True -Condition ($scriptText -match 'auditpol') -Name 'Audit policy is read'
    Assert-True -Condition ($scriptText -match 'AuditReceivingNTLMTraffic') -Name 'NTLM auditing is checked'
    Assert-True -Condition ($scriptText -match 'Kerberoasting') -Name 'Kerberoasting detection is referenced'
    foreach ($controlId in 1..10 | ForEach-Object { 'MON-{0:d3}' -f $_ }) {
        Assert-True -Condition ($scriptText -match [regex]::Escape($controlId)) -Name "Control $controlId is present"
    }
    Assert-True -Condition ($scriptText -notmatch 'Set-ItemProperty' -and $scriptText -notmatch 'auditpol /set') -Name 'Script performs no audit configuration changes'
}
Test-StaticScriptContent

function New-FullAuditPolicy {
    $subcategories = @(
        'Credential Validation','Kerberos Authentication Service','Kerberos Service Ticket Operations',
        'Computer Account Management','Security Group Management','User Account Management',
        'Directory Service Access','Directory Service Changes','Audit Policy Change',
        'Special Logon','Logon','Logoff','Account Lockout','Sensitive Privilege Use','Security System Extension'
    )
    $policy = @{}
    foreach ($subcategory in $subcategories) { $policy[$subcategory] = 'Success and Failure' }
    return $policy
}

function New-HardenedMonDcFacts {
    return @{
        AuditPolicy = New-FullAuditPolicy
        WefSubscriptionManager = 'Server=http://wef.contoso.com'
        SiemAgentsRunning = @('AzureMonitorAgent')
        AuditReceivingNtlm = 2
        AuditNtlmInDomain = 7
        TimeSyncType = 'NT5DS'
        TimeNtpServer = $null
    }
}

function Test-HardenedMonDc {
    $rows = @(Test-MonDcControls -ServerName 'DC02' -Facts (New-HardenedMonDcFacts) -IsPdcEmulator $false)
    $byId = @{}
    foreach ($row in $rows) { $byId[$row.ControlId] = $row }

    Assert-True -Condition ($rows.Count -eq 5) -Name 'Five per-DC monitoring rows are produced'
    Assert-True -Condition ($byId['MON-001'].Status -eq 'Pass') -Name 'Log forwarding passes MON-001'
    Assert-True -Condition ($byId['MON-002'].Status -eq 'Pass') -Name 'Full audit policy passes MON-002'
    Assert-True -Condition ($byId['MON-003'].Status -eq 'Manual') -Name 'Group auditing enabled makes MON-003 Manual (confirm alert)'
    Assert-True -Condition ($byId['MON-006'].Status -eq 'Pass') -Name 'NT5DS on non-PDC passes MON-006'
    Assert-True -Condition ($byId['MON-009'].Status -eq 'Pass') -Name 'NTLM auditing passes MON-009'
}
Test-HardenedMonDc

function Test-PdcTimeSource {
    $facts = New-HardenedMonDcFacts
    $facts.TimeSyncType = 'NTP'
    $facts.TimeNtpServer = 'time.windows.com,0x9'
    $rows = @(Test-MonDcControls -ServerName 'PDC01' -Facts $facts -IsPdcEmulator $true)
    $timeRow = $rows | Where-Object ControlId -eq 'MON-006'
    Assert-True -Condition ($timeRow.Status -eq 'Pass' -and $timeRow.Evidence -match 'external NTP') -Name 'PDC with external NTP passes MON-006'

    $facts.TimeSyncType = 'NT5DS'
    $facts.TimeNtpServer = $null
    $badPdc = @(Test-MonDcControls -ServerName 'PDC01' -Facts $facts -IsPdcEmulator $true)
    $badTimeRow = $badPdc | Where-Object ControlId -eq 'MON-006'
    Assert-True -Condition ($badTimeRow.Status -eq 'Warning') -Name 'PDC syncing from hierarchy warns MON-006'
}
Test-PdcTimeSource

function Test-WeakMonDc {
    $facts = New-HardenedMonDcFacts
    $facts.WefSubscriptionManager = $null
    $facts.SiemAgentsRunning = @()
    $facts.AuditPolicy.Remove('Security Group Management')
    $facts.AuditPolicy['Directory Service Changes'] = 'No Auditing'
    $facts.AuditReceivingNtlm = $null
    $facts.AuditNtlmInDomain = $null

    $rows = @(Test-MonDcControls -ServerName 'DC03' -Facts $facts -IsPdcEmulator $false)
    $byId = @{}
    foreach ($row in $rows) { $byId[$row.ControlId] = $row }

    Assert-True -Condition ($byId['MON-001'].Status -eq 'Warning') -Name 'No log forwarding warns MON-001'
    Assert-True -Condition ($byId['MON-002'].Status -eq 'Fail' -and $byId['MON-002'].Evidence -match 'Directory Service Changes') -Name 'Missing audit subcategory fails MON-002 with name'
    Assert-True -Condition ($byId['MON-003'].Status -eq 'Fail') -Name 'No group auditing fails MON-003'
    Assert-True -Condition ($byId['MON-009'].Status -eq 'Warning') -Name 'No NTLM auditing warns MON-009'
}
Test-WeakMonDc

function Test-DomainControls {
    $facts = @{ DomainName = 'contoso.com'; HoneytokenAccounts = @('HNY-svc-sql','HNY-admin') }
    $rows = @(Test-MonDomainControls -Facts $facts -HoneytokenNamePattern 'HNY-*')
    $byId = @{}
    foreach ($row in $rows) { $byId[$row.ControlId] = $row }

    Assert-True -Condition ($rows.Count -eq 5) -Name 'Five domain-level monitoring rows are produced'
    Assert-True -Condition ($byId['MON-007'].Status -eq 'Pass' -and $byId['MON-007'].Evidence -match 'HNY-svc-sql') -Name 'Honeytoken accounts pass MON-007 with names'
    foreach ($controlId in @('MON-004','MON-005','MON-008','MON-010')) {
        Assert-True -Condition ($byId[$controlId].Status -eq 'Manual') -Name "Detection-content control $controlId is Manual"
    }

    $noPatternRows = @(Test-MonDomainControls -Facts $facts -HoneytokenNamePattern '')
    Assert-True -Condition (($noPatternRows | Where-Object ControlId -eq 'MON-007').Status -eq 'Manual') -Name 'MON-007 is Manual without a honeytoken pattern'

    $missingFacts = @{ DomainName = 'contoso.com'; HoneytokenAccounts = @() }
    $missingRows = @(Test-MonDomainControls -Facts $missingFacts -HoneytokenNamePattern 'HNY-*')
    Assert-True -Condition (($missingRows | Where-Object ControlId -eq 'MON-007').Status -eq 'Fail') -Name 'No matching honeytoken accounts fail MON-007'
}
Test-DomainControls

if ($script:Failures.Count -gt 0) {
    Write-Host "$($script:Failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host "All $($script:Passes) checks passed." -ForegroundColor Green
exit 0
