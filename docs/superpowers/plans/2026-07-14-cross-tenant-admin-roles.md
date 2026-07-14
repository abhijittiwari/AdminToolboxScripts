# Cross-Tenant Admin Roles Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a read-only PowerShell report that uses `Admin.csv` mappings to export WAE and Fortescue Entra ID directory roles, immutable IDs, UPNs, display names, and assigned licenses.

**Architecture:** Add a standalone script that validates the CSV, connects to Microsoft Graph once per tenant, builds tenant-local role indexes, resolves mapped users, and exports one combined CSV. Keep live Graph calls inside the executable script and test only static parsing plus pure helper behavior so checks do not need tenant access.

**Tech Stack:** PowerShell 7-compatible syntax, Microsoft Graph PowerShell SDK, CSV import/export, existing lightweight PowerShell check scripts.

## Global Constraints

- The script is read-only and must not change users, roles, licenses, or tenant configuration.
- The input CSV must contain `WAEUPN`, `Prefix`, and `FortescueUPN`.
- The output CSV must contain `Environment`, `Prefix`, `InputUPN`, `UserPrincipalName`, `DisplayName`, `ImmutableId`, `ActiveDirectoryRoles`, `PimEligibleDirectoryRoles`, `AllDirectoryRoles`, `LicensesAssigned`, `LookupStatus`, and `Error`.
- Microsoft Graph delegated scopes are `Directory.Read.All` and `RoleManagement.Read.Directory`.
- Report Entra ID directory roles only, including active and PIM-eligible assignments.
- Do not report Exchange Online RBAC roles, MFA, mailbox details, general group memberships, app-only authentication, or scheduled execution.
- Do not create a git commit unless the user explicitly asks.

---

## File Structure

- Create `Get-CrossTenantAdminRoleAssignments.ps1`: user-facing script, parameter handling, Graph connection, role collection, user lookup, license lookup, output export.
- Create `tests/Get-CrossTenantAdminRoleAssignments.Checks.ps1`: static parse checks and pure helper checks that do not connect to Microsoft Graph.
- Modify `README.md`: add the new script to the script table and document requirements, parameters, usage, output, and notes.
- Modify `docs/script-reference.md`: add a concise reference section for the new script.

---

### Task 1: Script Core

**Files:**
- Create: `Get-CrossTenantAdminRoleAssignments.ps1`

**Interfaces:**
- Consumes: CSV file with `WAEUPN`, `Prefix`, `FortescueUPN` columns.
- Produces: `Get-CrossTenantAdminRoleAssignments.ps1` with helper functions `Test-RequiredCsvColumns`, `Join-ReportValues`, `Add-UserRole`, `Resolve-RolePrincipal`, `Get-TenantRoleIndex`, `Get-UserReportRow`, and `Invoke-TenantExport`.

- [ ] **Step 1: Write the script skeleton and parameter block**

Use `apply_patch` to add this complete file at `Get-CrossTenantAdminRoleAssignments.ps1`:

```powershell
<#
.SYNOPSIS
    Exports WAE and Fortescue admin account Entra ID directory role and license details.

.DESCRIPTION
    Reads a CSV containing WAEUPN, Prefix, and FortescueUPN columns. The script connects
    to Microsoft Graph for WAE, exports the WAE account details, disconnects, prompts for
    Fortescue sign-in, exports the Fortescue account details, and writes one combined CSV.

    The script is read-only. It reports Entra ID directory roles only, including active
    and PIM-eligible assignments when available.

.PARAMETER InputCsv
    Path to the admin mapping CSV. Defaults to .\Admin.csv.

.PARAMETER OutputPath
    Combined CSV output path. Defaults to .\CrossTenantAdminRoles_yyyyMMdd_HHmmss.csv.

.EXAMPLE
    .\Get-CrossTenantAdminRoleAssignments.ps1 -InputCsv .\Admin.csv
#>

[CmdletBinding()]
param(
    [string]$InputCsv = '.\Admin.csv',
    [string]$OutputPath = ".\CrossTenantAdminRoles_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

$ErrorActionPreference = 'Stop'

function Test-RequiredCsvColumns {
    param(
        [Parameter(Mandatory = $true)] [object[]]$Rows,
        [Parameter(Mandatory = $true)] [string[]]$RequiredColumns
    )

    if ($Rows.Count -eq 0) {
        throw 'Input CSV contains no data rows.'
    }

    $actualColumns = @($Rows[0].PSObject.Properties.Name)
    $missingColumns = @($RequiredColumns | Where-Object { $actualColumns -notcontains $_ })
    if ($missingColumns.Count -gt 0) {
        throw "Input CSV is missing required column(s): $($missingColumns -join ', ')"
    }
}

function Join-ReportValues {
    param([object[]]$Values)

    $cleanValues = @($Values | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Sort-Object -Unique)
    return ($cleanValues -join '; ')
}

function Add-UserRole {
    param(
        [Parameter(Mandatory = $true)] [hashtable]$RoleMap,
        [Parameter(Mandatory = $true)] [string]$UserId,
        [Parameter(Mandatory = $true)] [string]$RoleName
    )

    if (-not $RoleMap.ContainsKey($UserId)) {
        $RoleMap[$UserId] = [System.Collections.Generic.List[string]]::new()
    }

    if (-not $RoleMap[$UserId].Contains($RoleName)) {
        $RoleMap[$UserId].Add($RoleName) | Out-Null
    }
}

function Resolve-RolePrincipal {
    param(
        [Parameter(Mandatory = $true)] [string]$PrincipalId,
        [Parameter(Mandatory = $true)] [string]$RoleName,
        [Parameter(Mandatory = $true)] [hashtable]$RoleMap
    )

    $principal = Get-MgDirectoryObjectById -Ids $PrincipalId -ErrorAction SilentlyContinue
    if (-not $principal) { return }

    switch ($principal.AdditionalProperties.'@odata.type') {
        '#microsoft.graph.user' {
            Add-UserRole -RoleMap $RoleMap -UserId $PrincipalId -RoleName $RoleName
        }
        '#microsoft.graph.group' {
            $groupName = $principal.AdditionalProperties.displayName
            Get-MgGroupTransitiveMember -GroupId $PrincipalId -All -ErrorAction SilentlyContinue |
                Where-Object { $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.user' } |
                ForEach-Object {
                    Add-UserRole -RoleMap $RoleMap -UserId $_.Id -RoleName "$RoleName via group: $groupName"
                }
        }
    }
}
```

- [ ] **Step 2: Add tenant role collection functions**

Use `apply_patch` to append this code after `Resolve-RolePrincipal`:

```powershell
function Get-TenantRoleIndex {
    $roleDefinitions = @{}
    Get-MgRoleManagementDirectoryRoleDefinition -All | ForEach-Object {
        $roleDefinitions[$_.Id] = $_.DisplayName
    }

    $activeRoles = @{}
    $eligibleRoles = @{}
    $pimWarning = ''

    Write-Host 'Collecting active Entra ID role assignments...' -ForegroundColor Cyan
    $activeAssignments = @(Get-MgRoleManagementDirectoryRoleAssignment -All)
    foreach ($assignment in $activeAssignments) {
        $roleName = $roleDefinitions[$assignment.RoleDefinitionId]
        if ($roleName) {
            Resolve-RolePrincipal -PrincipalId $assignment.PrincipalId -RoleName "$roleName (Active)" -RoleMap $activeRoles
        }
    }

    Write-Host 'Collecting PIM-eligible Entra ID role assignments...' -ForegroundColor Cyan
    try {
        $eligibleAssignments = @(Get-MgRoleManagementDirectoryRoleEligibilityScheduleInstance -All)
        foreach ($assignment in $eligibleAssignments) {
            $roleName = $roleDefinitions[$assignment.RoleDefinitionId]
            if ($roleName) {
                Resolve-RolePrincipal -PrincipalId $assignment.PrincipalId -RoleName "$roleName (PIM Eligible)" -RoleMap $eligibleRoles
            }
        }
    }
    catch {
        $pimWarning = "Could not read PIM eligibility schedules: $($_.Exception.Message)"
        Write-Warning $pimWarning
    }

    return [pscustomobject]@{
        ActiveRoles  = $activeRoles
        EligibleRoles = $eligibleRoles
        PimWarning   = $pimWarning
    }
}

function Get-UserReportRow {
    param(
        [Parameter(Mandatory = $true)] [string]$Environment,
        [Parameter(Mandatory = $true)] [string]$Prefix,
        [Parameter(Mandatory = $true)] [string]$InputUPN,
        [Parameter(Mandatory = $true)] [object]$RoleIndex
    )

    $errorMessages = [System.Collections.Generic.List[string]]::new()
    if (-not [string]::IsNullOrWhiteSpace($RoleIndex.PimWarning)) {
        $errorMessages.Add($RoleIndex.PimWarning) | Out-Null
    }

    try {
        $user = Get-MgUser -UserId $InputUPN -Property 'id,displayName,userPrincipalName,onPremisesImmutableId' -ErrorAction Stop
    }
    catch {
        return [pscustomobject]@{
            Environment               = $Environment
            Prefix                    = $Prefix
            InputUPN                  = $InputUPN
            UserPrincipalName         = ''
            DisplayName               = ''
            ImmutableId               = ''
            ActiveDirectoryRoles      = ''
            PimEligibleDirectoryRoles = ''
            AllDirectoryRoles         = ''
            LicensesAssigned          = ''
            LookupStatus              = 'NotFound'
            Error                     = "User lookup failed: $($_.Exception.Message)"
        }
    }

    $licenses = ''
    try {
        $licenses = Join-ReportValues -Values @(Get-MgUserLicenseDetail -UserId $user.Id -ErrorAction Stop | Select-Object -ExpandProperty SkuPartNumber)
    }
    catch {
        $errorMessages.Add("License lookup failed: $($_.Exception.Message)") | Out-Null
    }

    $activeRoles = if ($RoleIndex.ActiveRoles.ContainsKey($user.Id)) { Join-ReportValues -Values $RoleIndex.ActiveRoles[$user.Id] } else { '' }
    $eligibleRoles = if ($RoleIndex.EligibleRoles.ContainsKey($user.Id)) { Join-ReportValues -Values $RoleIndex.EligibleRoles[$user.Id] } else { '' }
    $allRoleValues = [System.Collections.Generic.List[string]]::new()
    if (-not [string]::IsNullOrWhiteSpace($activeRoles)) {
        ($activeRoles -split '; ') | ForEach-Object { $allRoleValues.Add($_) | Out-Null }
    }
    if (-not [string]::IsNullOrWhiteSpace($eligibleRoles)) {
        ($eligibleRoles -split '; ') | ForEach-Object { $allRoleValues.Add($_) | Out-Null }
    }
    $allRoles = Join-ReportValues -Values $allRoleValues

    return [pscustomobject]@{
        Environment               = $Environment
        Prefix                    = $Prefix
        InputUPN                  = $InputUPN
        UserPrincipalName         = $user.UserPrincipalName
        DisplayName               = $user.DisplayName
        ImmutableId               = $user.OnPremisesImmutableId
        ActiveDirectoryRoles      = $activeRoles
        PimEligibleDirectoryRoles = $eligibleRoles
        AllDirectoryRoles         = $allRoles
        LicensesAssigned          = $licenses
        LookupStatus              = 'Found'
        Error                     = Join-ReportValues -Values $errorMessages
    }
}
```

- [ ] **Step 3: Add tenant execution and main flow**

Use `apply_patch` to append this code after `Get-UserReportRow`:

```powershell
function Invoke-TenantExport {
    param(
        [Parameter(Mandatory = $true)] [string]$Environment,
        [Parameter(Mandatory = $true)] [object[]]$Mappings,
        [Parameter(Mandatory = $true)] [string]$UpnColumn
    )

    Write-Host "Connecting to Microsoft Graph for $Environment..." -ForegroundColor Cyan
    Connect-MgGraph -Scopes @('Directory.Read.All', 'RoleManagement.Read.Directory') -NoWelcome

    try {
        $roleIndex = Get-TenantRoleIndex
        $rows = [System.Collections.Generic.List[object]]::new()
        $count = 0

        foreach ($mapping in $Mappings) {
            $count++
            $upn = [string]$mapping.$UpnColumn
            Write-Progress -Activity "Processing $Environment admin accounts" `
                -Status ("{0} of {1} - {2}" -f $count, $Mappings.Count, $upn) `
                -PercentComplete (($count / $Mappings.Count) * 100)

            $rows.Add((Get-UserReportRow -Environment $Environment -Prefix $mapping.Prefix -InputUPN $upn -RoleIndex $roleIndex)) | Out-Null
        }

        Write-Progress -Activity "Processing $Environment admin accounts" -Completed
        return $rows
    }
    finally {
        Disconnect-MgGraph | Out-Null
    }
}

if (-not (Test-Path -Path $InputCsv -PathType Leaf)) {
    throw "Input CSV not found: $InputCsv"
}

$mappings = @(Import-Csv -Path $InputCsv)
Test-RequiredCsvColumns -Rows $mappings -RequiredColumns @('WAEUPN', 'Prefix', 'FortescueUPN')

$reportRows = [System.Collections.Generic.List[object]]::new()
$reportRows.AddRange([object[]](Invoke-TenantExport -Environment 'WAE' -Mappings $mappings -UpnColumn 'WAEUPN'))

Write-Host ''
Write-Host 'Sign in to the Fortescue tenant when prompted.' -ForegroundColor Yellow
$reportRows.AddRange([object[]](Invoke-TenantExport -Environment 'Fortescue' -Mappings $mappings -UpnColumn 'FortescueUPN'))

$columns = 'Environment','Prefix','InputUPN','UserPrincipalName','DisplayName','ImmutableId',
           'ActiveDirectoryRoles','PimEligibleDirectoryRoles','AllDirectoryRoles',
           'LicensesAssigned','LookupStatus','Error'

$reportRows |
    Sort-Object Prefix, Environment |
    Select-Object $columns |
    Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

Write-Host "Report exported: $OutputPath ($($reportRows.Count) rows)" -ForegroundColor Green
```

- [ ] **Step 4: Run parser check manually**

Run: `pwsh -NoProfile -Command '$tokens=$null; $errors=$null; [System.Management.Automation.Language.Parser]::ParseFile("./Get-CrossTenantAdminRoleAssignments.ps1",[ref]$tokens,[ref]$errors) > $null; if ($errors.Count) { $errors | ForEach-Object Message; exit 1 }; "Parse OK"'`

Expected: `Parse OK`

---

### Task 2: Lightweight Checks

**Files:**
- Create: `tests/Get-CrossTenantAdminRoleAssignments.Checks.ps1`

**Interfaces:**
- Consumes: functions in `Get-CrossTenantAdminRoleAssignments.ps1` by AST-loading function definitions only.
- Produces: A local check script that validates parseability, required functions, CSV column validation, value joining, output columns, and Graph scope strings without calling Microsoft Graph.

- [ ] **Step 1: Add the check script**

Use `apply_patch` to add this complete file at `tests/Get-CrossTenantAdminRoleAssignments.Checks.ps1`:

```powershell
#Requires -Version 7
$ErrorActionPreference = 'Stop'

$script:ScriptPath = Join-Path $PSScriptRoot '..' 'Get-CrossTenantAdminRoleAssignments.ps1'
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
    foreach ($name in @('Test-RequiredCsvColumns','Join-ReportValues','Add-UserRole','Resolve-RolePrincipal','Get-TenantRoleIndex','Get-UserReportRow','Invoke-TenantExport')) {
        Assert-True -Condition ($functionNames -contains $name) -Name "Function exists: $name"
    }
}
Test-RequiredFunctionsExist

function Test-CsvValidation {
    $valid = @([pscustomobject]@{ WAEUPN = 'a@wae.com'; Prefix = 'adma'; FortescueUPN = 'a@fortescue.com' })
    Test-RequiredCsvColumns -Rows $valid -RequiredColumns @('WAEUPN','Prefix','FortescueUPN')
    Assert-True -Condition $true -Name 'CSV validation accepts required columns'

    $threw = $false
    try {
        Test-RequiredCsvColumns -Rows @([pscustomobject]@{ WAEUPN = 'a@wae.com'; Prefix = 'adma' }) -RequiredColumns @('WAEUPN','Prefix','FortescueUPN')
    }
    catch {
        $threw = $_.Exception.Message -match 'FortescueUPN'
    }
    Assert-True -Condition $threw -Name 'CSV validation reports missing FortescueUPN'
}
Test-CsvValidation

function Test-JoinReportValues {
    $joined = Join-ReportValues -Values @('Role B', '', $null, 'Role A', 'Role A')
    Assert-True -Condition ($joined -eq 'Role A; Role B') -Name 'Join-ReportValues sorts, deduplicates, and skips blanks'
}
Test-JoinReportValues

function Test-AddUserRole {
    $roleMap = @{}
    Add-UserRole -RoleMap $roleMap -UserId 'u1' -RoleName 'Global Reader (Active)'
    Add-UserRole -RoleMap $roleMap -UserId 'u1' -RoleName 'Global Reader (Active)'
    Assert-True -Condition ($roleMap['u1'].Count -eq 1 -and $roleMap['u1'][0] -eq 'Global Reader (Active)') -Name 'Add-UserRole deduplicates user roles'
}
Test-AddUserRole

function Test-StaticScriptContent {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw
    foreach ($column in @('Environment','Prefix','InputUPN','UserPrincipalName','DisplayName','ImmutableId','ActiveDirectoryRoles','PimEligibleDirectoryRoles','AllDirectoryRoles','LicensesAssigned','LookupStatus','Error')) {
        Assert-True -Condition ($scriptText -match [regex]::Escape($column)) -Name "Output column present: $column"
    }
    Assert-True -Condition ($scriptText -match 'Directory\.Read\.All' -and $scriptText -match 'RoleManagement\.Read\.Directory') -Name 'Required Graph scopes present'
    Assert-True -Condition ($scriptText -match 'Get-MgRoleManagementDirectoryRoleAssignment' -and $scriptText -match 'Get-MgRoleManagementDirectoryRoleEligibilityScheduleInstance') -Name 'Active and PIM role cmdlets present'
    Assert-True -Condition ($scriptText -notmatch 'Connect-ExchangeOnline' -and $scriptText -notmatch 'Get-ManagementRoleAssignment') -Name 'Exchange Online RBAC is out of scope'
}
Test-StaticScriptContent

if ($script:Failures.Count -gt 0) {
    Write-Host "$($script:Failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host "All $($script:Passes) checks passed." -ForegroundColor Green
exit 0
```

- [ ] **Step 2: Run the new check script and verify it passes**

Run: `pwsh -NoProfile -File ./tests/Get-CrossTenantAdminRoleAssignments.Checks.ps1`

Expected: output ends with `All 26 checks passed.` and exit code `0`.

---

### Task 3: Documentation

**Files:**
- Modify: `README.md`
- Modify: `docs/script-reference.md`

**Interfaces:**
- Consumes: implemented script name, parameters, scopes, input columns, and output columns from Tasks 1 and 2.
- Produces: user-facing usage instructions consistent with the script behavior.

- [ ] **Step 1: Update README script table**

Use `apply_patch` to add this table row after `Get-EntraAdminAccounts.ps1` in `README.md`:

```markdown
| `Get-CrossTenantAdminRoleAssignments.ps1` | PowerShell | Uses a WAE/Fortescue admin mapping CSV to export Entra ID active and PIM-eligible directory roles, immutable IDs, UPNs, display names, and licenses from both tenants. |
```

- [ ] **Step 2: Add README section**

Use `apply_patch` to add this section after the `Get-EntraAdminAccounts.ps1` section in `README.md` and before `Find-ADDuplicateEmailProxyAddresses.ps1`:

```markdown
## Get-CrossTenantAdminRoleAssignments.ps1

Exports WAE and Fortescue admin account details from Microsoft Graph using a mapping CSV with `WAEUPN`, `Prefix`, and `FortescueUPN` columns. The report includes Entra ID active directory roles, PIM-eligible directory roles when readable, immutable ID, actual user principal name, display name, assigned license SKU part numbers, lookup status, and errors.

### Requirements

- PowerShell.
- Microsoft Graph PowerShell SDK: `Install-Module Microsoft.Graph`.
- Recommended Entra role: Global Reader or equivalent in each tenant.
- Microsoft Graph permissions:
  - `Directory.Read.All`
  - `RoleManagement.Read.Directory`

### Parameters

| Parameter | Type | Description |
| --- | --- | --- |
| `-InputCsv` | String | Admin mapping CSV path. Defaults to `./Admin.csv`. |
| `-OutputPath` | String | Combined CSV output path. Defaults to `./CrossTenantAdminRoles_yyyyMMdd_HHmmss.csv`. |

### Usage

```powershell
./Get-CrossTenantAdminRoleAssignments.ps1 -InputCsv ./Admin.csv
```

```powershell
./Get-CrossTenantAdminRoleAssignments.ps1 -InputCsv ./Admin.csv -OutputPath ./WAE-Fortescue-AdminRoles.csv
```

### Notes

- The script connects to Microsoft Graph once for WAE and once for Fortescue; sign into the requested tenant at each prompt.
- Exchange Online RBAC roles are not included.
- If PIM eligibility cannot be read, active roles are still exported and the PIM warning is included in the `Error` column.
- Missing users are exported with `LookupStatus=NotFound`.
```

- [ ] **Step 3: Add script-reference section**

Use `apply_patch` to add this section after the `Get-EntraAdminAccounts.ps1` reference in `docs/script-reference.md` and before `Find-ADDuplicateEmailProxyAddresses.ps1`:

```markdown
## Get-CrossTenantAdminRoleAssignments.ps1

Uses a WAE/Fortescue admin mapping CSV to export matching admin account role and license details from both tenants.

### Common Uses

- Compare WAE and Fortescue admin accounts by shared `Prefix`.
- Review Entra ID active directory role assignments.
- Review PIM-eligible directory role assignments when permissions and licensing allow it.
- Export immutable IDs, actual UPNs, display names, and assigned license SKU part numbers.

### Requirements

- PowerShell.
- Microsoft Graph PowerShell SDK.
- Recommended role: Global Reader or equivalent in both tenants.
- Graph permissions: `Directory.Read.All` and `RoleManagement.Read.Directory`.

### Parameters

| Parameter | Description |
| --- | --- |
| `-InputCsv <path>` | Admin mapping CSV with `WAEUPN`, `Prefix`, and `FortescueUPN`. Defaults to `./Admin.csv`. |
| `-OutputPath <path>` | Destination CSV. Defaults to `./CrossTenantAdminRoles_yyyyMMdd_HHmmss.csv`. |

### Examples

```powershell
./Get-CrossTenantAdminRoleAssignments.ps1 -InputCsv ./Admin.csv
```

```powershell
./Get-CrossTenantAdminRoleAssignments.ps1 -InputCsv ./Admin.csv -OutputPath ./WAE-Fortescue-AdminRoles.csv
```

### Output

Writes one combined CSV with one row per mapped account per environment. Columns include environment, prefix, input UPN, actual UPN, display name, immutable ID, active roles, PIM-eligible roles, all roles, assigned license SKU part numbers, lookup status, and error text.

### Validation Notes

- Spot-check role assignments against the Entra admin center.
- Check `Error` for PIM eligibility warnings or license lookup failures.
- Confirm expected missing users appear with `LookupStatus=NotFound`.
```

- [ ] **Step 4: Run all local checks**

Run: `pwsh -NoProfile -File ./tests/Get-CrossTenantAdminRoleAssignments.Checks.ps1`

Expected: output ends with `All 26 checks passed.` and exit code `0`.

Run: `pwsh -NoProfile -File ./tests/Export-ExchangeObjectData.Checks.ps1`

Expected: output ends with `All 34 checks passed.` or another all-passing count from the existing test file, and exit code `0`. If this fails because of unrelated existing working-tree changes, record that exactly in the final summary and do not modify unrelated files.

---

## Plan Self-Review

Spec coverage:

- CSV input columns are covered in Task 1 and Task 2.
- Combined WAE and Fortescue output is covered in Task 1.
- Required output columns are covered in Task 1 and Task 2.
- Separate Graph connections per tenant are covered in Task 1.
- Active and PIM-eligible directory role collection is covered in Task 1 and Task 2.
- User display name, UPN, immutable ID, and licenses are covered in Task 1.
- NotFound and PIM warning handling are covered in Task 1 and Task 2.
- Documentation is covered in Task 3.

Placeholder scan:

- No deferred implementation placeholders are intentionally left in the executable steps.

Type consistency:

- Helper names in Task 2 match helper names produced by Task 1.
- Output column names in tests and documentation match the spec.
