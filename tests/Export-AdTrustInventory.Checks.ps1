#Requires -Version 7
$ErrorActionPreference = 'Stop'

$script:ScriptPath = Join-Path $PSScriptRoot '..' 'Export-AdTrustInventory.ps1'
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
    foreach ($name in @('ConvertTo-TrustDirectionName','ConvertTo-TrustTypeName','ConvertTo-TrustAttributeInfo','New-TrustRow','Get-AdTrustFacts','Get-AdTrustInventory')) {
        Assert-True -Condition ($functionNames -contains $name) -Name "Function exists: $name"
    }
}
Test-RequiredFunctionsExist

function Test-ReadOnly {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw
    Assert-True -Condition ($scriptText -notmatch 'New-ADTrust' -and $scriptText -notmatch 'Set-ADTrust' -and $scriptText -notmatch 'Remove-ADTrust') -Name 'Script performs no trust write operations'
    Assert-True -Condition ($scriptText -match 'Get-ADTrust') -Name 'Trust query is present'
}
Test-ReadOnly

function Test-DirectionAndTypeDecoding {
    Assert-True -Condition ((ConvertTo-TrustDirectionName -Direction 1) -eq 'Inbound') -Name 'Direction 1 = Inbound'
    Assert-True -Condition ((ConvertTo-TrustDirectionName -Direction 2) -eq 'Outbound') -Name 'Direction 2 = Outbound'
    Assert-True -Condition ((ConvertTo-TrustDirectionName -Direction 3) -eq 'Bidirectional') -Name 'Direction 3 = Bidirectional'
    Assert-True -Condition ((ConvertTo-TrustTypeName -TrustType 2) -eq 'Uplevel (AD Domain)') -Name 'Type 2 = Uplevel'
    Assert-True -Condition ((ConvertTo-TrustTypeName -TrustType 3) -eq 'MIT (Kerberos realm)') -Name 'Type 3 = MIT realm'
}
Test-DirectionAndTypeDecoding

function Test-AttributeDecoding {
    # Forest-transitive (0x8) + quarantined/SID-filtering (0x4).
    $forestFiltered = ConvertTo-TrustAttributeInfo -TrustAttributes 0xC
    Assert-True -Condition ($forestFiltered.IsForestTransitive -eq $true -and $forestFiltered.SidFilteringEnabled -eq $true) -Name 'Forest-transitive + SID filtering decoded'

    # Selective auth (0x10) + non-transitive (0x1).
    $selective = ConvertTo-TrustAttributeInfo -TrustAttributes 0x11
    Assert-True -Condition ($selective.SelectiveAuth -eq $true -and $selective.IsNonTransitive -eq $true) -Name 'Selective auth + non-transitive decoded'

    # Within-forest (0x20).
    $withinForest = ConvertTo-TrustAttributeInfo -TrustAttributes 0x20
    Assert-True -Condition ($withinForest.WithinForest -eq $true) -Name 'Within-forest flag decoded'
}
Test-AttributeDecoding

function Test-TrustRowShaping {
    $facts = @{
        Trusts = @(
            # External trust, bidirectional, SID filtering ON (quarantined), no selective auth.
            [pscustomobject]@{ SourceDomain = 'contoso.com'; Target = 'legacy.example'; Direction = 3; TrustType = 2; TrustAttributes = 0x4 },
            # Forest trust, outbound, forest-transitive, selective auth ON.
            [pscustomobject]@{ SourceDomain = 'contoso.com'; Target = 'partner.com'; Direction = 2; TrustType = 2; TrustAttributes = 0x18 },
            # Intra-forest parent-child (within forest).
            [pscustomobject]@{ SourceDomain = 'contoso.com'; Target = 'emea.contoso.com'; Direction = 3; TrustType = 2; TrustAttributes = 0x20 }
        )
    }

    $rows = @(Get-AdTrustInventory -Facts $facts)
    Assert-True -Condition ($rows.Count -eq 3) -Name 'One row per trust'

    $external = $rows | Where-Object PartnerDomain -eq 'legacy.example'
    Assert-True -Condition ($external.TrustCategory -eq 'External' -and $external.Direction -eq 'Bidirectional' -and $external.SidFilteringEnabled -eq $true -and $external.SelectiveAuth -eq $false) -Name 'External trust with SID filtering shaped correctly'

    $forest = $rows | Where-Object PartnerDomain -eq 'partner.com'
    Assert-True -Condition ($forest.TrustCategory -eq 'Forest' -and $forest.ForestTransitive -eq $true -and $forest.SelectiveAuth -eq $true -and $forest.Direction -eq 'Outbound') -Name 'Forest trust with selective auth shaped correctly'

    $intra = $rows | Where-Object PartnerDomain -eq 'emea.contoso.com'
    Assert-True -Condition ($intra.TrustCategory -eq 'Intra-forest') -Name 'Within-forest trust categorised as intra-forest'
}
Test-TrustRowShaping

if ($script:Failures.Count -gt 0) {
    Write-Host "$($script:Failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host "All $($script:Passes) checks passed." -ForegroundColor Green
exit 0
