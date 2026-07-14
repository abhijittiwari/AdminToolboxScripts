#Requires -Version 7
$ErrorActionPreference = 'Stop'

$script:ScriptPath = Join-Path $PSScriptRoot '..' 'Export-ExchangeMigrationReadiness.ps1'
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

function Test-MigrationStatusLogic {
    # Blocked when litigation hold is present.
    $row = Invoke-MigrationReadinessCheck -Identity 'user1' -DisplayName 'User One' -RecipientType 'UserMailbox' -LitigationHold $true -RetentionPolicy $null -ComplianceHolds @() -Licensing $true
    Assert-True -Condition ($row.MigrationStatus -eq 'Blocked' -and $row.BlockingReasons -match 'litigation hold') -Name 'Invoke-MigrationReadinessCheck blocks on litigation hold'

    # Blocked when compliance holds present.
    $row = Invoke-MigrationReadinessCheck -Identity 'user2' -DisplayName 'User Two' -RecipientType 'UserMailbox' -LitigationHold $false -RetentionPolicy $null -ComplianceHolds @('eDiscoveryHold') -Licensing $true
    Assert-True -Condition ($row.MigrationStatus -eq 'Blocked' -and $row.BlockingReasons -match 'compliance hold') -Name 'Invoke-MigrationReadinessCheck blocks on compliance holds'

    # Risky when retention policy present (but not blocked).
    $row = Invoke-MigrationReadinessCheck -Identity 'user3' -DisplayName 'User Three' -RecipientType 'UserMailbox' -LitigationHold $false -RetentionPolicy 'Default' -ComplianceHolds @() -Licensing $true
    Assert-True -Condition ($row.MigrationStatus -eq 'Risky' -and $row.BlockingReasons -match 'retention policy') -Name 'Invoke-MigrationReadinessCheck marks as risky with retention policy'

    # Blocked when licensing missing.
    $row = Invoke-MigrationReadinessCheck -Identity 'user4' -DisplayName 'User Four' -RecipientType 'UserMailbox' -LitigationHold $false -RetentionPolicy $null -ComplianceHolds @() -Licensing $false
    Assert-True -Condition ($row.MigrationStatus -eq 'Blocked' -and $row.BlockingReasons -match 'no license') -Name 'Invoke-MigrationReadinessCheck blocks on missing license'

    # Review when mailbox type is uncommon.
    $row = Invoke-MigrationReadinessCheck -Identity 'user5' -DisplayName 'User Five' -RecipientType 'UserMailbox' -MailboxType 'DiscoveryMailbox' -LitigationHold $false -RetentionPolicy $null -ComplianceHolds @() -Licensing $true
    Assert-True -Condition ($row.MigrationStatus -eq 'Review' -and $row.BlockingReasons -match 'DiscoveryMailbox') -Name 'Invoke-MigrationReadinessCheck marks DiscoveryMailbox for review'

    # Ready when no issues.
    $row = Invoke-MigrationReadinessCheck -Identity 'user6' -DisplayName 'User Six' -RecipientType 'UserMailbox' -MailboxType 'UserMailbox' -LitigationHold $false -RetentionPolicy $null -ComplianceHolds @() -Licensing $true
    Assert-True -Condition ($row.MigrationStatus -eq 'Ready') -Name 'Invoke-MigrationReadinessCheck marks as Ready when no blockers'

    # License names are joined into the Licenses column.
    $row = Invoke-MigrationReadinessCheck -Identity 'user7' -DisplayName 'User Seven' -RecipientType 'UserMailbox' -MailboxType 'UserMailbox' -LitigationHold $false -RetentionPolicy $null -ComplianceHolds @() -Licensing $true -Licenses @('EMS', 'SPE_E5')
    Assert-True -Condition ($row.Licenses -eq 'EMS; SPE_E5') -Name 'Invoke-MigrationReadinessCheck joins license names into Licenses column'

    # Licenses column is empty when none are passed.
    $row = Invoke-MigrationReadinessCheck -Identity 'user8' -DisplayName 'User Eight' -RecipientType 'UserMailbox' -MailboxType 'UserMailbox' -LitigationHold $false -RetentionPolicy $null -ComplianceHolds @() -Licensing $true
    Assert-True -Condition ($row.Licenses -eq '') -Name 'Invoke-MigrationReadinessCheck emits empty Licenses when none provided'
}
Test-MigrationStatusLogic

function Test-ScriptStaticContract {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw
    Assert-True -Condition ($scriptText -match 'Errors-MigrationReadiness\.csv') -Name 'Script writes job-specific errors CSV'
    Assert-True -Condition ($scriptText -match 'MigrationReadiness\.csv') -Name 'Script writes MigrationReadiness.csv'
    Assert-True -Condition ($scriptText -match 'BlockedObjects\.csv') -Name 'Script writes BlockedObjects.csv'
    Assert-True -Condition (($scriptText -match "ParameterSetName.*=.*'All'" -and $scriptText -match "ParameterSetName.*=.*'Csv'" -and $scriptText -match "ParameterSetName.*=.*'Identity'")) -Name 'Script supports All, Csv, and Identity parameter sets'
    Assert-True -Condition ($scriptText -match "ParameterSetName.*=.*'ObjectsCsv'") -Name 'Script supports ObjectsCsv parameter (job input)'
    Assert-True -Condition ($scriptText -match 'MigrationStatus') -Name 'Script calculates MigrationStatus (Ready/Blocked/Risky/Review)'

    # Get-EXOMailbox only accepts real mailbox property names. LitigationHold,
    # ComplianceHold, and ComplianceTagHold are not valid; LitigationHoldEnabled is.
    $propertiesMatch = [regex]::Match($scriptText, 'Get-EXOMailbox[^\r\n]*-Properties\s+([^\r\n]+)')
    Assert-True -Condition $propertiesMatch.Success -Name 'Script requests explicit Get-EXOMailbox properties'
    $requestedProperties = @($propertiesMatch.Groups[1].Value -split ',' | ForEach-Object { ($_.Trim() -split '\s')[0] } | Where-Object { $_ })
    $validProperties = @('LitigationHoldEnabled', 'RetentionPolicy', 'InPlaceHolds')
    $invalidRequested = @($requestedProperties | Where-Object { $_ -notin $validProperties })
    Assert-True -Condition ($invalidRequested.Count -eq 0) -Name "Script requests only valid mailbox properties (found: $($requestedProperties -join ', '))"
    Assert-True -Condition ('LitigationHoldEnabled' -in $requestedProperties) -Name 'Script requests LitigationHoldEnabled (not the invalid LitigationHold)'

    # License names come from Graph licenseDetails (skuPartNumber); default Get-MgUser
    # output omits license properties, so the raw endpoint must be used.
    Assert-True -Condition ($scriptText -match 'licenseDetails') -Name 'Script reads license names from the Graph licenseDetails endpoint'
    Assert-True -Condition ($scriptText -match "'Licenses'") -Name 'Script exports a Licenses column'
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
