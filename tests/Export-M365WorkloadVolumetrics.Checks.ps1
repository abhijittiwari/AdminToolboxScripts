#Requires -Version 7
$ErrorActionPreference = 'Stop'

$script:ScriptPath = Join-Path $PSScriptRoot '..' 'Export-M365WorkloadVolumetrics.ps1'
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
    foreach ($name in @('ConvertFrom-SizeString','Format-Bytes','New-VolumetricRow','Get-ExchangeVolumetricRows','Get-SiteVolumetricRows','Get-M365WorkloadInventory','Get-M365WorkloadFacts')) {
        Assert-True -Condition ($functionNames -contains $name) -Name "Function exists: $name"
    }
}
Test-RequiredFunctionsExist

function Test-ReadOnly {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw
    Assert-True -Condition ($scriptText -notmatch 'Set-Mailbox\b' -and $scriptText -notmatch 'Remove-\w' -and $scriptText -notmatch 'New-MoveRequest') -Name 'Script performs no write/migration operations'
    Assert-True -Condition ($scriptText -match 'Get-EXOMailbox' -and $scriptText -match 'Get-SPOSite') -Name 'Exchange and SharePoint queries are present'
}
Test-ReadOnly

function Test-SizeParsing {
    Assert-True -Condition ((ConvertFrom-SizeString -Size '1.523 GB (1,635,608,853 bytes)') -eq 1635608853) -Name 'Exchange size string parsed to exact bytes'
    Assert-True -Condition ((ConvertFrom-SizeString -Size '0 B (0 bytes)') -eq 0) -Name 'Zero size parsed'
    Assert-True -Condition ((ConvertFrom-SizeString -Size $null) -eq 0) -Name 'Null size is zero bytes'
    Assert-True -Condition ((ConvertFrom-SizeString -Size '5 GB') -eq [int64](5 * 1GB)) -Name 'Plain unit string parsed via multiplier'
}
Test-SizeParsing

function Test-FormatBytes {
    Assert-True -Condition ((Format-Bytes -Bytes ([int64](2 * 1TB))) -match '2.00 TB') -Name 'TB formatting'
    Assert-True -Condition ((Format-Bytes -Bytes ([int64](3 * 1GB))) -match '3.00 GB') -Name 'GB formatting'
}
Test-FormatBytes

function Test-ExchangeAggregation {
    $stats = @(
        [pscustomobject]@{ IsArchive = $false; TotalItemSize = '1 GB (1,073,741,824 bytes)' },
        [pscustomobject]@{ IsArchive = $false; TotalItemSize = '2 GB (2,147,483,648 bytes)' },
        [pscustomobject]@{ IsArchive = $true;  TotalItemSize = '5 GB (5,368,709,120 bytes)' }
    )
    $rows = @(Get-ExchangeVolumetricRows -MailboxStats $stats)
    $primary = $rows | Where-Object Metric -eq 'Primary mailboxes'
    $archive = $rows | Where-Object Metric -eq 'Archive mailboxes'
    Assert-True -Condition ($primary.Count -eq 2 -and $primary.TotalBytes -eq 3221225472) -Name 'Primary mailbox count and byte total aggregated'
    Assert-True -Condition ($archive.Count -eq 1 -and $archive.TotalBytes -eq 5368709120) -Name 'Archive mailbox count and byte total aggregated'
}
Test-ExchangeAggregation

function Test-SiteAggregation {
    $sites = @(
        [pscustomobject]@{ Template = 'GROUP#0'; StorageUsageCurrent = 1024 },  # 1024 MB
        [pscustomobject]@{ Template = 'STS#3';   StorageUsageCurrent = 512 }
    )
    $allRow = Get-SiteVolumetricRows -Workload 'SharePoint' -Metric 'All sites' -Sites $sites
    Assert-True -Condition ($allRow.Count -eq 2 -and $allRow.TotalBytes -eq [int64](1536 * 1MB)) -Name 'Site storage (MB) aggregated to bytes'
}
Test-SiteAggregation

function Test-InventoryAssembly {
    $facts = @{
        MailboxStats = @(
            [pscustomobject]@{ IsArchive = $false; TotalItemSize = '1 GB (1,073,741,824 bytes)' },
            [pscustomobject]@{ IsArchive = $true;  TotalItemSize = '2 GB (2,147,483,648 bytes)' }
        )
        SharedMailboxCount = 12
        OneDriveSites = @([pscustomobject]@{ Template = 'SPSPERS'; StorageUsageCurrent = 2048 })
        SharePointSites = @(
            [pscustomobject]@{ Template = 'GROUP#0'; StorageUsageCurrent = 1024 },
            [pscustomobject]@{ Template = 'STS#3';   StorageUsageCurrent = 1024 }
        )
        M365GroupCount = 40
        DistributionListCount = 15
        PublicFolderStats = $null
    }
    $rows = @(Get-M365WorkloadInventory -Facts $facts)

    $shared = $rows | Where-Object Metric -eq 'Shared/room/equipment mailboxes'
    Assert-True -Condition ($shared.Count -eq 12) -Name 'Shared mailbox count row emitted'
    $onedrive = $rows | Where-Object Workload -eq 'OneDrive'
    Assert-True -Condition ($onedrive.Count -eq 1 -and $onedrive.TotalBytes -eq [int64](2048 * 1MB)) -Name 'OneDrive row aggregated'
    $teams = $rows | Where-Object Metric -eq 'Teams-connected sites'
    Assert-True -Condition ($teams.Count -eq 1) -Name 'Teams-connected sites split from all sites'
    $groups = $rows | Where-Object Metric -eq 'Microsoft 365 groups'
    Assert-True -Condition ($groups.Count -eq 40) -Name 'M365 group count row emitted'
    Assert-True -Condition (@($rows | Where-Object Workload -eq 'Public Folders').Count -eq 0) -Name 'Null public folder facts produce no rows'
}
Test-InventoryAssembly

if ($script:Failures.Count -gt 0) {
    Write-Host "$($script:Failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host "All $($script:Passes) checks passed." -ForegroundColor Green
exit 0
