#Requires -Version 7
$ErrorActionPreference = 'Stop'

$script:ScriptPath = Join-Path $PSScriptRoot '..' 'Export-AdcsHardeningAssessment.ps1'
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
    foreach ($name in @('New-ControlRow','Get-BroadPrincipalMatches','ConvertTo-TemplateFacts','Test-TemplateAllowsAuthentication','Test-AdcsControls','Get-AdcsFacts')) {
        Assert-True -Condition ($functionNames -contains $name) -Name "Function exists: $name"
    }
}
Test-RequiredFunctionsExist

function Test-StaticScriptContent {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw
    Assert-True -Condition ($scriptText -match 'ESC1') -Name 'ESC1 is referenced'
    Assert-True -Condition ($scriptText -match 'EDITF_ATTRIBUTESUBJECTALTNAME2') -Name 'ESC6 EDITF flag is referenced'
    Assert-True -Condition ($scriptText -match 'pKIEnrollmentService') -Name 'Enrollment services are queried'
    Assert-True -Condition ($scriptText -match 'msPKI-Certificate-Name-Flag') -Name 'Name flag (ENROLLEE_SUPPLIES_SUBJECT) is read'
    foreach ($controlId in 1..10 | ForEach-Object { 'ADCS-{0:d3}' -f $_ }) {
        Assert-True -Condition ($scriptText -match [regex]::Escape($controlId)) -Name "Control $controlId is present"
    }
    Assert-True -Condition ($scriptText -notmatch 'Set-ADObject' -and $scriptText -notmatch '-setreg') -Name 'Script performs no AD CS configuration changes'
}
Test-StaticScriptContent

function New-Sid { param([string]$Name) return [pscustomobject]@{ Value = $Name } }

function New-Ace {
    param([string]$Identity, [string]$Rights, [string]$ObjectType = '00000000-0000-0000-0000-000000000000', [string]$Type = 'Allow')
    return [pscustomobject]@{
        IdentityReference = $Identity
        ActiveDirectoryRights = $Rights
        ObjectType = $ObjectType
        AccessControlType = $Type
    }
}

function New-TemplateObject {
    param(
        [string]$Name,
        [int]$NameFlag = 0,
        [int]$EnrollmentFlag = 0,
        [int]$RaSignature = 0,
        [string[]]$Eku = @('1.3.6.1.5.5.7.3.2'),
        [object[]]$Aces = @()
    )
    return [pscustomobject]@{
        Name = $Name
        'msPKI-Certificate-Name-Flag' = $NameFlag
        'msPKI-Enrollment-Flag' = $EnrollmentFlag
        'msPKI-RA-Signature' = $RaSignature
        pKIExtendedKeyUsage = $Eku
        nTSecurityDescriptor = [pscustomobject]@{ Access = $Aces }
    }
}

function Test-TemplateConversion {
    $enrollAce = New-Ace -Identity 'CONTOSO\Domain Users' -Rights 'ExtendedRight' -ObjectType '0e10c968-78fb-11d2-90d4-00c04f79dc55'
    $writeAce = New-Ace -Identity 'CONTOSO\Helpdesk' -Rights 'WriteDacl'
    # ENROLLEE_SUPPLIES_SUBJECT = 0x1; enrollment flag without PEND_ALL_REQUESTS (approval) = 0
    $templateObject = New-TemplateObject -Name 'VulnUser' -NameFlag 0x1 -EnrollmentFlag 0 -Aces @($enrollAce, $writeAce)
    $facts = ConvertTo-TemplateFacts -TemplateObject $templateObject -PublishedTemplateNames @('VulnUser')

    Assert-True -Condition ($facts.EnrolleeSuppliesSubject -eq $true) -Name 'ENROLLEE_SUPPLIES_SUBJECT flag is decoded'
    Assert-True -Condition ($facts.ManagerApprovalRequired -eq $false) -Name 'Manager approval flag is decoded'
    Assert-True -Condition ($facts.Published -eq $true) -Name 'Published state is set from enrollment services'
    Assert-True -Condition ($facts.EnrollPrincipals -contains 'CONTOSO\Domain Users') -Name 'Enrollment principal is extracted from the ACL'
    Assert-True -Condition ($facts.WritePrincipals -contains 'CONTOSO\Helpdesk') -Name 'Write principal is extracted from the ACL'

    Assert-True -Condition ((Get-BroadPrincipalMatches -PrincipalNames @('CONTOSO\Domain Users','CONTOSO\PKI-Admins')).Count -eq 1) -Name 'Broad principal matcher flags Domain Users only'
    Assert-True -Condition ((Test-TemplateAllowsAuthentication -EkuList @('1.3.6.1.5.5.7.3.2')) -eq $true) -Name 'Client Authentication EKU counts as authentication'
    Assert-True -Condition ((Test-TemplateAllowsAuthentication -EkuList @()) -eq $true) -Name 'No EKU constraint counts as authentication-capable'
    Assert-True -Condition ((Test-TemplateAllowsAuthentication -EkuList @('1.3.6.1.5.5.7.3.4')) -eq $false) -Name 'Email-only EKU does not count as authentication'
}
Test-TemplateConversion

function New-HardenedAdcsFacts {
    $safeTemplate = @{
        Name = 'WebServer'; Published = $true; EnrolleeSuppliesSubject = $true; ManagerApprovalRequired = $false
        RaSignaturesRequired = 0; EkuList = @('1.3.6.1.5.5.7.3.1'); EnrollPrincipals = @('CONTOSO\Web-Admins'); WritePrincipals = @('CONTOSO\PKI-Admins')
    }
    $userTemplate = @{
        Name = 'User'; Published = $true; EnrolleeSuppliesSubject = $false; ManagerApprovalRequired = $false
        RaSignaturesRequired = 0; EkuList = @('1.3.6.1.5.5.7.3.2'); EnrollPrincipals = @('CONTOSO\Staff-Certificate-Users'); WritePrincipals = @('CONTOSO\PKI-Admins')
    }
    return @{
        ForestName = 'contoso.com'
        Templates = @($safeTemplate, $userTemplate)
        CaObjectWritePrincipals = @('CONTOSO\PKI-Admins')
        EnrollmentServices = @(@{
            Name = 'Contoso-Issuing-CA'; HostName = 'ca01.contoso.com'
            EditfAttributeSan = $false; AuditFilter = 127; KeyProvider = 'nCipher Security World Key Storage Provider'; WebEnrollmentDetected = $false
        })
    }
}

function Test-HardenedAdcs {
    $rows = @(Test-AdcsControls -Facts (New-HardenedAdcsFacts))
    $byId = @{}
    foreach ($row in $rows) { $byId[$row.ControlId] = $row }

    foreach ($controlId in @('ADCS-001','ADCS-002','ADCS-003','ADCS-004','ADCS-005','ADCS-008','ADCS-009')) {
        Assert-True -Condition ($byId[$controlId].Status -eq 'Pass') -Name "Hardened AD CS passes $controlId"
    }
    Assert-True -Condition ($byId['ADCS-006'].Status -eq 'Manual') -Name 'ADCS-006 (ESC7) is Manual review'
    Assert-True -Condition ($byId['ADCS-007'].Status -eq 'Pass') -Name 'No web enrollment passes ADCS-007'
    Assert-True -Condition ($byId['ADCS-010'].Status -eq 'Manual') -Name 'ADCS-010 CRL/OCSP is Manual review'
}
Test-HardenedAdcs

function Test-Esc1AndEsc6Fail {
    $facts = New-HardenedAdcsFacts
    $facts.Templates = @(
        @{ Name = 'ESC1Template'; Published = $true; EnrolleeSuppliesSubject = $true; ManagerApprovalRequired = $false
           RaSignaturesRequired = 0; EkuList = @('1.3.6.1.5.5.7.3.2'); EnrollPrincipals = @('CONTOSO\Domain Users'); WritePrincipals = @('CONTOSO\PKI-Admins') }
    )
    $facts.EnrollmentServices[0].EditfAttributeSan = $true
    $facts.EnrollmentServices[0].AuditFilter = 0
    $facts.EnrollmentServices[0].WebEnrollmentDetected = $true
    $facts.EnrollmentServices[0].KeyProvider = 'Microsoft Software Key Storage Provider'

    $rows = @(Test-AdcsControls -Facts $facts)
    $byId = @{}
    foreach ($row in $rows) { $byId[$row.ControlId] = $row }

    Assert-True -Condition ($byId['ADCS-002'].Status -eq 'Fail' -and $byId['ADCS-002'].Evidence -match 'ESC1Template') -Name 'ESC1-exploitable template fails ADCS-002'
    Assert-True -Condition ($byId['ADCS-005'].Status -eq 'Fail') -Name 'EDITF SAN flag fails ADCS-005 (ESC6)'
    Assert-True -Condition ($byId['ADCS-007'].Status -eq 'Warning') -Name 'HTTP web enrollment warns ADCS-007 (ESC8)'
    Assert-True -Condition ($byId['ADCS-008'].Status -eq 'Fail') -Name 'AuditFilter=0 fails ADCS-008'
    Assert-True -Condition ($byId['ADCS-009'].Status -eq 'Warning') -Name 'Software key provider warns ADCS-009'
}
Test-Esc1AndEsc6Fail

function Test-Esc4Detection {
    $facts = New-HardenedAdcsFacts
    $facts.Templates[0].WritePrincipals = @('CONTOSO\Domain Users')
    $rows = @(Test-AdcsControls -Facts $facts)
    $esc4Row = $rows | Where-Object ControlId -eq 'ADCS-004'
    Assert-True -Condition ($esc4Row.Status -eq 'Fail' -and $esc4Row.Evidence -match 'ESC4') -Name 'Broad write on template fails ADCS-004 (ESC4)'
}
Test-Esc4Detection

function Test-NoAdcsProducesNotApplicable {
    $facts = @{ ForestName = 'contoso.com'; Templates = @(); CaObjectWritePrincipals = @(); EnrollmentServices = @() }
    $rows = @(Test-AdcsControls -Facts $facts)
    Assert-True -Condition ($rows.Count -eq 1 -and $rows[0].ControlId -eq 'ADCS-000' -and $rows[0].Evidence -match 'not applicable') -Name 'No AD CS yields a single not-applicable row'
}
Test-NoAdcsProducesNotApplicable

if ($script:Failures.Count -gt 0) {
    Write-Host "$($script:Failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host "All $($script:Passes) checks passed." -ForegroundColor Green
exit 0
