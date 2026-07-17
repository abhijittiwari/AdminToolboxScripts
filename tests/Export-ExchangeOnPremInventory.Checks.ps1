#Requires -Version 7
$ErrorActionPreference = 'Stop'

$script:ScriptPath = Join-Path $PSScriptRoot '..' 'Export-ExchangeOnPremInventory.ps1'
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
    foreach ($name in @('New-ExchangeServerRow','New-AcceptedDomainRow','New-SendConnectorRow','New-ReceiveConnectorRow','New-DagRow','Get-ExchangeOnPremInventory','Get-ExchangeOnPremFacts')) {
        Assert-True -Condition ($functionNames -contains $name) -Name "Function exists: $name"
    }
}
Test-RequiredFunctionsExist

function Test-ReadOnly {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw
    Assert-True -Condition ($scriptText -notmatch 'Set-ReceiveConnector\b' -and $scriptText -notmatch 'New-SendConnector\b' -and $scriptText -notmatch 'Set-AcceptedDomain\b' -and $scriptText -notmatch 'Enable-Mailbox\b') -Name 'Script performs no Exchange write operations'
    Assert-True -Condition ($scriptText -match 'Get-ExchangeServer' -and $scriptText -match 'Get-AcceptedDomain') -Name 'Exchange server and accepted-domain queries are present'
}
Test-ReadOnly

function New-MockExchangeFacts {
    return @{
        Servers = @(
            [pscustomobject]@{ Name = 'EXCH01'; ServerRole = 'Mailbox'; Edition = 'Enterprise'; AdminDisplayVersion = 'Version 15.2 (Build 1544.4)'; Site = 'Azure-AE' }
        )
        AcceptedDomains = @(
            [pscustomobject]@{ Name = 'contoso.com'; DomainName = 'contoso.com'; DomainType = 'Authoritative'; Default = $true },
            [pscustomobject]@{ Name = 'relay.contoso.com'; DomainName = 'relay.contoso.com'; DomainType = 'InternalRelay'; Default = $false }
        )
        SendConnectors = @(
            [pscustomobject]@{ Name = 'To Internet'; AddressSpaces = @('SMTP:*;1'); SmartHosts = @('mail.contoso.com'); Enabled = $true; SourceTransportServers = @('EXCH01') }
        )
        ReceiveConnectors = @(
            [pscustomobject]@{ Name = 'Default EXCH01'; Server = 'EXCH01'; Bindings = @('0.0.0.0:25'); PermissionGroups = @('ExchangeUsers','ExchangeServers'); AuthMechanism = @('Tls'); Enabled = $true },
            [pscustomobject]@{ Name = 'App Relay'; Server = 'EXCH01'; Bindings = @('0.0.0.0:25'); PermissionGroups = @('AnonymousUsers'); AuthMechanism = @('Tls','ExternalAuthoritative'); Enabled = $true }
        )
        Dags = @(
            [pscustomobject]@{ Name = 'DAG01'; Servers = @('EXCH01','EXCH02') }
        )
        HybridConfiguration = @(
            [pscustomobject]@{ Domains = @('contoso.com','contoso.mail.onmicrosoft.com') }
        )
        PublicFolderMailboxCount = 3
    }
}

function Test-InventoryShaping {
    $inventory = Get-ExchangeOnPremInventory -Facts (New-MockExchangeFacts)

    Assert-True -Condition ($inventory.HybridConfigured -eq $true -and ($inventory.HybridDomains -contains 'contoso.com')) -Name 'Hybrid configuration detected with domains'
    Assert-True -Condition (@($inventory.Servers).Count -eq 1 -and $inventory.Servers[0].IsHubTransport -eq $true) -Name 'Mailbox server flagged as transport-capable'

    $authoritative = $inventory.AcceptedDomains | Where-Object DomainName -eq 'contoso.com'
    Assert-True -Condition ($authoritative.DomainType -eq 'Authoritative' -and $authoritative.Default -eq $true) -Name 'Authoritative default accepted domain shaped'

    $appRelay = $inventory.ReceiveConnectors | Where-Object Name -eq 'App Relay'
    Assert-True -Condition ($appRelay.AllowsAnonymousRelay -eq $true) -Name 'Anonymous relay receive connector flagged'
    $default = $inventory.ReceiveConnectors | Where-Object Name -eq 'Default EXCH01'
    Assert-True -Condition ($default.AllowsAnonymousRelay -eq $false) -Name 'Non-anonymous receive connector not flagged'

    Assert-True -Condition ($inventory.SendConnectors[0].SmartHosts -eq 'mail.contoso.com') -Name 'Send connector smart host shaped'
    Assert-True -Condition (@($inventory.Dags).Count -eq 1 -and $inventory.Dags[0].MemberCount -eq 2) -Name 'DAG members counted'
    Assert-True -Condition ($inventory.PublicFolderMailboxCount -eq 3) -Name 'Public folder mailbox count carried through'
}
Test-InventoryShaping

function Test-NoHybridFacts {
    $facts = New-MockExchangeFacts
    $facts.HybridConfiguration = @()
    $inventory = Get-ExchangeOnPremInventory -Facts $facts
    Assert-True -Condition ($inventory.HybridConfigured -eq $false -and @($inventory.HybridDomains).Count -eq 0) -Name 'No hybrid configuration shapes to false/empty'
}
Test-NoHybridFacts

if ($script:Failures.Count -gt 0) {
    Write-Host "$($script:Failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host "All $($script:Passes) checks passed." -ForegroundColor Green
exit 0
