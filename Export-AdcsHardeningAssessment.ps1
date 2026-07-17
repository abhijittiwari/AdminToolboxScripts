<#
.SYNOPSIS
    Assesses Active Directory Certificate Services hardening controls ADCS-001 to ADCS-010 (Appendix B baseline).

.DESCRIPTION
    Read-only assessment of AD CS against the security hardening baseline, including the
    well-known certificate attack paths (ESC1-ESC8): template enrollment permissions,
    requester-supplied subject (ESC1), enrollment-agent and dangerous EKUs (ESC2/ESC3),
    vulnerable template and CA object ACLs (ESC4/ESC5), the EDITF_ATTRIBUTESUBJECTALTNAME2
    flag (ESC6), CA administration permissions (ESC7), NTLM relay to enrollment endpoints
    (ESC8), CA audit logging, private key protection, and CRL/OCSP management.

    Certificate templates and enrollment services are read from the AD configuration
    partition; CA registry settings are read with certutil. Each control is emitted as one
    row with Status: Pass / Fail / Warning / Manual / Error. If no enrollment services are
    found, a single not-applicable row is emitted.

    The script makes no changes to any system.

.PARAMETER Server
    Optional Domain Controller / AD Web Services server for directory queries.

.PARAMETER SkipCaRegistryChecks
    Skip certutil-based CA registry checks (ESC6, audit filter, key provider), e.g. when
    the CA hosts are unreachable from this machine.

.PARAMETER OutputPath
    Optional CSV output path. If omitted, results are only written to the pipeline/table.

.EXAMPLE
    .\Export-AdcsHardeningAssessment.ps1 -OutputPath .\AdcsHardening.csv
#>

[CmdletBinding()]
param(
    [string]$Server = '',
    [switch]$SkipCaRegistryChecks,
    [string]$OutputPath = ''
)

$ErrorActionPreference = 'Stop'

function New-ControlRow {
    param(
        [Parameter(Mandatory = $true)] [string]$ControlId,
        [Parameter(Mandatory = $true)] [string]$Control,
        [Parameter(Mandatory = $true)] [string]$Target,
        [Parameter(Mandatory = $true)] [ValidateSet('Pass','Fail','Warning','Manual','Error')] [string]$Status,
        [Parameter(Mandatory = $false)] [string]$Evidence = ''
    )

    return [pscustomobject]@{
        ControlId = $ControlId
        Control   = $Control
        Target    = $Target
        Status    = $Status
        Evidence  = $Evidence
    }
}

function Get-BroadPrincipalMatches {
    param([string[]]$PrincipalNames)

    $broadPatterns = @('Everyone','Authenticated Users','Domain Users','Domain Computers','\\Users$','BUILTIN\\Users')
    return @($PrincipalNames | Where-Object {
        $principal = $_
        @($broadPatterns | Where-Object { $principal -match $_ }).Count -gt 0
    })
}

function ConvertTo-TemplateFacts {
    param(
        [Parameter(Mandatory = $true)] $TemplateObject,
        [Parameter(Mandatory = $true)] [string[]]$PublishedTemplateNames
    )

    $nameFlagValue = $TemplateObject.'msPKI-Certificate-Name-Flag'
    $nameFlag = if ($null -ne $nameFlagValue) { [int64]$nameFlagValue } else { 0 }
    $enrollmentFlagValue = $TemplateObject.'msPKI-Enrollment-Flag'
    $enrollmentFlag = if ($null -ne $enrollmentFlagValue) { [int64]$enrollmentFlagValue } else { 0 }
    $ekuList = @($TemplateObject.pKIExtendedKeyUsage)

    $enrollPrincipals = [System.Collections.Generic.List[string]]::new()
    $writePrincipals = [System.Collections.Generic.List[string]]::new()
    $securityDescriptor = $TemplateObject.nTSecurityDescriptor
    if ($securityDescriptor) {
        $enrollGuids = @('0e10c968-78fb-11d2-90d4-00c04f79dc55','a05b8cc2-17bc-4802-a710-e7c15ab866a2')
        foreach ($ace in @($securityDescriptor.Access)) {
            if ($ace.AccessControlType -ne 'Allow') { continue }
            $identity = [string]$ace.IdentityReference
            $rights = [string]$ace.ActiveDirectoryRights
            if ($rights -match 'ExtendedRight|GenericAll' -and ($enrollGuids -contains [string]$ace.ObjectType -or [string]$ace.ObjectType -eq '00000000-0000-0000-0000-000000000000')) {
                if ($enrollPrincipals -notcontains $identity) { $enrollPrincipals.Add($identity) }
            }
            if ($rights -match 'GenericAll|GenericWrite|WriteDacl|WriteOwner|WriteProperty') {
                if ($writePrincipals -notcontains $identity) { $writePrincipals.Add($identity) }
            }
        }
    }

    return @{
        Name                    = [string]$TemplateObject.Name
        Published               = ($PublishedTemplateNames -contains [string]$TemplateObject.Name)
        EnrolleeSuppliesSubject = (($nameFlag -band 0x1) -ne 0)
        ManagerApprovalRequired = (($enrollmentFlag -band 0x2) -ne 0)
        RaSignaturesRequired    = [int]($(if ($null -ne $TemplateObject.'msPKI-RA-Signature') { $TemplateObject.'msPKI-RA-Signature' } else { 0 }))
        EkuList                 = @($ekuList | ForEach-Object { [string]$_ })
        EnrollPrincipals        = $enrollPrincipals.ToArray()
        WritePrincipals         = $writePrincipals.ToArray()
    }
}

function Test-TemplateAllowsAuthentication {
    param([string[]]$EkuList)

    $authEkus = @('1.3.6.1.5.5.7.3.2','1.3.6.1.4.1.311.20.2.2','1.3.6.1.5.2.3.4','2.5.29.37.0')
    if ($EkuList.Count -eq 0) { return $true }  # no EKU constraint = usable for any purpose
    return (@($EkuList | Where-Object { $authEkus -contains $_ }).Count -gt 0)
}

function Test-AdcsControls {
    param(
        [Parameter(Mandatory = $true)] $Facts
    )

    $rows = [System.Collections.Generic.List[object]]::new()
    $forest = [string]$Facts.ForestName

    if (@($Facts.EnrollmentServices).Count -eq 0) {
        $rows.Add((New-ControlRow -ControlId 'ADCS-000' -Control 'AD CS presence' -Target $forest -Status 'Pass' -Evidence 'No AD CS enrollment services found in the configuration partition; ADCS-001..010 are not applicable'))
        return $rows.ToArray()
    }

    $templates = @($Facts.Templates)
    $publishedTemplates = @($templates | Where-Object { $_.Published })

    # ADCS-001 template enrollment least privilege
    $broadEnrollTemplates = @($publishedTemplates | Where-Object { (Get-BroadPrincipalMatches -PrincipalNames $_.EnrollPrincipals).Count -gt 0 })
    if ($broadEnrollTemplates.Count -eq 0) {
        $rows.Add((New-ControlRow -ControlId 'ADCS-001' -Control 'Certificate template enrollment permissions shall follow least privilege' -Target $forest -Status 'Pass' -Evidence ("No published template grants enrollment to broad principals ({0} published template(s))" -f $publishedTemplates.Count)))
    }
    else {
        $rows.Add((New-ControlRow -ControlId 'ADCS-001' -Control 'Certificate template enrollment permissions shall follow least privilege' -Target $forest -Status 'Warning' -Evidence ("Broad enrollment on: {0}" -f (($broadEnrollTemplates | ForEach-Object { "{0} ({1})" -f $_.Name, ((Get-BroadPrincipalMatches -PrincipalNames $_.EnrollPrincipals) -join '/') }) -join '; '))))
    }

    # ADCS-002 ESC1 requester-supplied subject
    $esc1Templates = @($publishedTemplates | Where-Object {
        $_.EnrolleeSuppliesSubject -and
        -not $_.ManagerApprovalRequired -and
        $_.RaSignaturesRequired -eq 0 -and
        (Test-TemplateAllowsAuthentication -EkuList $_.EkuList)
    })
    $esc1Exploitable = @($esc1Templates | Where-Object { (Get-BroadPrincipalMatches -PrincipalNames $_.EnrollPrincipals).Count -gt 0 })
    if ($esc1Exploitable.Count -gt 0) {
        $rows.Add((New-ControlRow -ControlId 'ADCS-002' -Control 'Templates allowing requester-supplied subject (SAN) shall be restricted (ESC1)' -Target $forest -Status 'Fail' -Evidence ("ESC1-exploitable templates (broad enroll + supplies subject + auth EKU, no approval): {0}" -f (($esc1Exploitable | ForEach-Object Name) -join ', '))))
    }
    elseif ($esc1Templates.Count -gt 0) {
        $rows.Add((New-ControlRow -ControlId 'ADCS-002' -Control 'Templates allowing requester-supplied subject (SAN) shall be restricted (ESC1)' -Target $forest -Status 'Warning' -Evidence ("Templates allow requester-supplied subject with auth EKU but enrollment is not broad: {0} - confirm enrollment principals are trusted" -f (($esc1Templates | ForEach-Object Name) -join ', '))))
    }
    else {
        $rows.Add((New-ControlRow -ControlId 'ADCS-002' -Control 'Templates allowing requester-supplied subject (SAN) shall be restricted (ESC1)' -Target $forest -Status 'Pass' -Evidence 'No published template combines requester-supplied subject, authentication EKU, and no approval'))
    }

    # ADCS-003 ESC2/ESC3 dangerous EKUs
    $anyPurposeTemplates = @($publishedTemplates | Where-Object { $_.EkuList -contains '2.5.29.37.0' -or $_.EkuList.Count -eq 0 })
    $agentTemplates = @($publishedTemplates | Where-Object { $_.EkuList -contains '1.3.6.1.4.1.311.20.2.1' -and -not $_.ManagerApprovalRequired })
    $esc23Issues = [System.Collections.Generic.List[string]]::new()
    foreach ($template in $anyPurposeTemplates) { $esc23Issues.Add(("{0} (Any Purpose / no EKU restriction - ESC2)" -f $template.Name)) }
    foreach ($template in $agentTemplates) { $esc23Issues.Add(("{0} (Certificate Request Agent without approval - ESC3)" -f $template.Name)) }
    if ($esc23Issues.Count -eq 0) {
        $rows.Add((New-ControlRow -ControlId 'ADCS-003' -Control 'Enrollment-agent and dangerous EKU configurations shall be controlled (ESC2/ESC3)' -Target $forest -Status 'Pass' -Evidence 'No published template with Any Purpose EKU or unapproved Certificate Request Agent EKU'))
    }
    else {
        $status = if (@($anyPurposeTemplates + $agentTemplates | Where-Object { (Get-BroadPrincipalMatches -PrincipalNames $_.EnrollPrincipals).Count -gt 0 }).Count -gt 0) { 'Fail' } else { 'Warning' }
        $rows.Add((New-ControlRow -ControlId 'ADCS-003' -Control 'Enrollment-agent and dangerous EKU configurations shall be controlled (ESC2/ESC3)' -Target $forest -Status $status -Evidence ($esc23Issues -join '; ')))
    }

    # ADCS-004 ESC4/ESC5 vulnerable ACLs
    $esc4Templates = @($templates | Where-Object { (Get-BroadPrincipalMatches -PrincipalNames $_.WritePrincipals).Count -gt 0 })
    $esc5Objects = @(Get-BroadPrincipalMatches -PrincipalNames @($Facts.CaObjectWritePrincipals))
    if ($esc4Templates.Count -eq 0 -and $esc5Objects.Count -eq 0) {
        $rows.Add((New-ControlRow -ControlId 'ADCS-004' -Control 'Vulnerable template and CA object ACLs shall be remediated (ESC4/ESC5)' -Target $forest -Status 'Pass' -Evidence 'No broad principals hold write permissions on templates or CA objects'))
    }
    else {
        $aclIssues = [System.Collections.Generic.List[string]]::new()
        foreach ($template in $esc4Templates) { $aclIssues.Add(("template {0} writable by {1} (ESC4)" -f $template.Name, ((Get-BroadPrincipalMatches -PrincipalNames $template.WritePrincipals) -join '/'))) }
        if ($esc5Objects.Count -gt 0) { $aclIssues.Add(("CA objects writable by {0} (ESC5)" -f ($esc5Objects -join '/'))) }
        $rows.Add((New-ControlRow -ControlId 'ADCS-004' -Control 'Vulnerable template and CA object ACLs shall be remediated (ESC4/ESC5)' -Target $forest -Status 'Fail' -Evidence ($aclIssues -join '; ')))
    }

    # Per-CA controls
    foreach ($ca in @($Facts.EnrollmentServices)) {
        $caTarget = "{0}\{1}" -f $ca.HostName, $ca.Name

        # ADCS-005 ESC6
        if ($null -eq $ca.EditfAttributeSan) {
            $rows.Add((New-ControlRow -ControlId 'ADCS-005' -Control 'The EDITF_ATTRIBUTESUBJECTALTNAME2 flag shall be disabled on the CA (ESC6)' -Target $caTarget -Status 'Manual' -Evidence 'CA policy EditFlags could not be read; run: certutil -config "<CA>" -getreg policy\EditFlags'))
        }
        elseif ($ca.EditfAttributeSan) {
            $rows.Add((New-ControlRow -ControlId 'ADCS-005' -Control 'The EDITF_ATTRIBUTESUBJECTALTNAME2 flag shall be disabled on the CA (ESC6)' -Target $caTarget -Status 'Fail' -Evidence 'EDITF_ATTRIBUTESUBJECTALTNAME2 is set - any request can specify an arbitrary SAN'))
        }
        else {
            $rows.Add((New-ControlRow -ControlId 'ADCS-005' -Control 'The EDITF_ATTRIBUTESUBJECTALTNAME2 flag shall be disabled on the CA (ESC6)' -Target $caTarget -Status 'Pass' -Evidence 'EDITF_ATTRIBUTESUBJECTALTNAME2 not set'))
        }

        # ADCS-006 ESC7
        $rows.Add((New-ControlRow -ControlId 'ADCS-006' -Control 'CA administration and management permissions shall be restricted (ESC7)' -Target $caTarget -Status 'Manual' -Evidence 'Review ManageCA/ManageCertificates holders: certutil -config "<CA>" -getreg CA\Security - only tier-0 PKI admins should hold these rights'))

        # ADCS-007 ESC8
        if ($ca.WebEnrollmentDetected -eq $true) {
            $rows.Add((New-ControlRow -ControlId 'ADCS-007' -Control 'NTLM relay to AD CS enrollment endpoints shall be mitigated with Extended Protection (ESC8)' -Target $caTarget -Status 'Warning' -Evidence 'HTTP enrollment endpoint (/certsrv) detected - require HTTPS with Extended Protection for Authentication, or disable web enrollment'))
        }
        elseif ($ca.WebEnrollmentDetected -eq $false) {
            $rows.Add((New-ControlRow -ControlId 'ADCS-007' -Control 'NTLM relay to AD CS enrollment endpoints shall be mitigated with Extended Protection (ESC8)' -Target $caTarget -Status 'Pass' -Evidence 'No HTTP web enrollment endpoint reachable from this host; also verify CES/NDES endpoints if deployed'))
        }
        else {
            $rows.Add((New-ControlRow -ControlId 'ADCS-007' -Control 'NTLM relay to AD CS enrollment endpoints shall be mitigated with Extended Protection (ESC8)' -Target $caTarget -Status 'Manual' -Evidence 'Web enrollment endpoint state could not be probed; verify IIS EPA settings on the CA'))
        }

        # ADCS-008 audit logging
        if ($null -eq $ca.AuditFilter) {
            $rows.Add((New-ControlRow -ControlId 'ADCS-008' -Control 'CA audit logging shall be enabled and regularly reviewed' -Target $caTarget -Status 'Manual' -Evidence 'AuditFilter could not be read; run: certutil -config "<CA>" -getreg CA\AuditFilter'))
        }
        elseif ([int]$ca.AuditFilter -eq 127) {
            $rows.Add((New-ControlRow -ControlId 'ADCS-008' -Control 'CA audit logging shall be enabled and regularly reviewed' -Target $caTarget -Status 'Pass' -Evidence 'AuditFilter=127 (all CA events audited); confirm events reach the SIEM'))
        }
        elseif ([int]$ca.AuditFilter -gt 0) {
            $rows.Add((New-ControlRow -ControlId 'ADCS-008' -Control 'CA audit logging shall be enabled and regularly reviewed' -Target $caTarget -Status 'Warning' -Evidence ("AuditFilter={0} (partial; recommended 127)" -f $ca.AuditFilter)))
        }
        else {
            $rows.Add((New-ControlRow -ControlId 'ADCS-008' -Control 'CA audit logging shall be enabled and regularly reviewed' -Target $caTarget -Status 'Fail' -Evidence 'AuditFilter=0 - CA auditing disabled'))
        }

        # ADCS-009 key protection
        if ([string]::IsNullOrWhiteSpace([string]$ca.KeyProvider)) {
            $rows.Add((New-ControlRow -ControlId 'ADCS-009' -Control 'CA private keys shall be protected (HSM or equivalent) and securely backed up' -Target $caTarget -Status 'Manual' -Evidence 'Key storage provider could not be read; verify HSM protection and key backup procedures'))
        }
        elseif ([string]$ca.KeyProvider -match 'nCipher|SafeNet|Luna|Thales|HSM|YubiHSM|Utimaco') {
            $rows.Add((New-ControlRow -ControlId 'ADCS-009' -Control 'CA private keys shall be protected (HSM or equivalent) and securely backed up' -Target $caTarget -Status 'Pass' -Evidence ("Provider: {0}; verify key backup and quorum procedures" -f $ca.KeyProvider)))
        }
        else {
            $rows.Add((New-ControlRow -ControlId 'ADCS-009' -Control 'CA private keys shall be protected (HSM or equivalent) and securely backed up' -Target $caTarget -Status 'Warning' -Evidence ("Software key storage provider in use ({0}); evaluate HSM protection and verify secure key backup" -f $ca.KeyProvider)))
        }

        # ADCS-010 issuance / CRL / OCSP operations
        $rows.Add((New-ControlRow -ControlId 'ADCS-010' -Control 'Certificate issuance, renewal, revocation, and CRL/OCSP availability shall be managed' -Target $caTarget -Status 'Manual' -Evidence 'Verify CRL publication is current (certutil -URL), OCSP responder availability, and documented issuance/revocation procedures'))
    }

    return $rows.ToArray()
}

function Get-AdcsFacts {
    param(
        [string]$Server = '',
        [switch]$SkipCaRegistryChecks
    )

    Import-Module ActiveDirectory -ErrorAction Stop
    $adParams = @{}
    if (-not [string]::IsNullOrWhiteSpace($Server)) { $adParams.Server = $Server }

    $rootDse = Get-ADRootDSE @adParams
    $configNc = $rootDse.configurationNamingContext
    $facts = @{ ForestName = ($rootDse.rootDomainNamingContext -replace ',?DC=', '.').Trim('.') }

    $pkiBase = "CN=Public Key Services,CN=Services,$configNc"
    $enrollmentServices = @()
    try {
        $enrollmentServices = @(Get-ADObject -SearchBase "CN=Enrollment Services,$pkiBase" -Filter { objectClass -eq 'pKIEnrollmentService' } -Properties dNSHostName, certificateTemplates, displayName @adParams)
    } catch { }

    $publishedTemplateNames = @($enrollmentServices | ForEach-Object { @($_.certificateTemplates) } | Where-Object { $_ } | Select-Object -Unique)

    $facts.Templates = @()
    try {
        $templateObjects = @(Get-ADObject -SearchBase "CN=Certificate Templates,$pkiBase" -Filter { objectClass -eq 'pKICertificateTemplate' } -Properties nTSecurityDescriptor, 'msPKI-Certificate-Name-Flag', 'msPKI-Enrollment-Flag', 'msPKI-RA-Signature', pKIExtendedKeyUsage @adParams)
        foreach ($templateObject in $templateObjects) {
            $facts.Templates += ConvertTo-TemplateFacts -TemplateObject $templateObject -PublishedTemplateNames $publishedTemplateNames
        }
    } catch { }

    $facts.CaObjectWritePrincipals = @()
    foreach ($service in $enrollmentServices) {
        try {
            $serviceObject = Get-ADObject -Identity $service.DistinguishedName -Properties nTSecurityDescriptor @adParams
            foreach ($ace in @($serviceObject.nTSecurityDescriptor.Access)) {
                if ($ace.AccessControlType -eq 'Allow' -and [string]$ace.ActiveDirectoryRights -match 'GenericAll|GenericWrite|WriteDacl|WriteOwner') {
                    $identity = [string]$ace.IdentityReference
                    if ($facts.CaObjectWritePrincipals -notcontains $identity) { $facts.CaObjectWritePrincipals += $identity }
                }
            }
        } catch { }
    }

    $facts.EnrollmentServices = @()
    foreach ($service in $enrollmentServices) {
        $caEntry = @{
            Name                 = [string]$service.Name
            HostName             = [string]$service.dNSHostName
            EditfAttributeSan    = $null
            AuditFilter          = $null
            KeyProvider          = $null
            WebEnrollmentDetected = $null
        }

        if (-not $SkipCaRegistryChecks) {
            $caConfig = "{0}\{1}" -f $caEntry.HostName, $caEntry.Name
            $editFlagsOutput = certutil -config $caConfig -getreg policy\EditFlags 2>$null
            if ($LASTEXITCODE -eq 0 -and $editFlagsOutput) {
                $caEntry.EditfAttributeSan = ([string]($editFlagsOutput -join ' ') -match 'EDITF_ATTRIBUTESUBJECTALTNAME2')
            }
            $auditOutput = certutil -config $caConfig -getreg CA\AuditFilter 2>$null
            if ($LASTEXITCODE -eq 0 -and $auditOutput) {
                $auditMatch = [regex]::Match(($auditOutput -join ' '), 'AuditFilter REG_DWORD = (?:0x)?([0-9a-fA-F]+)')
                if ($auditMatch.Success) { $caEntry.AuditFilter = [convert]::ToInt32($auditMatch.Groups[1].Value, 16) }
            }
            $providerOutput = certutil -config $caConfig -getreg 'CA\CSP\Provider' 2>$null
            if ($LASTEXITCODE -eq 0 -and $providerOutput) {
                $providerMatch = [regex]::Match(($providerOutput -join ' '), 'Provider REG_SZ = (.+?)(?:\s{2,}|$)')
                if ($providerMatch.Success) { $caEntry.KeyProvider = $providerMatch.Groups[1].Value.Trim() }
            }

            try {
                Invoke-WebRequest -Uri ("http://{0}/certsrv/" -f $caEntry.HostName) -Method Head -TimeoutSec 5 -UseBasicParsing -ErrorAction Stop | Out-Null
                $caEntry.WebEnrollmentDetected = $true
            } catch {
                $statusCode = $null
                if ($_.Exception.Response) { $statusCode = [int]$_.Exception.Response.StatusCode }
                # 401/403 still prove the endpoint exists; connection failures mean no endpoint reachable.
                $caEntry.WebEnrollmentDetected = ($statusCode -in @(200, 401, 403))
            }
        }

        $facts.EnrollmentServices += $caEntry
    }

    return $facts
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
Write-Host 'Collecting AD CS facts from the configuration partition...' -ForegroundColor Cyan
$facts = Get-AdcsFacts -Server $Server -SkipCaRegistryChecks:$SkipCaRegistryChecks
Write-Host ("Found {0} enrollment service(s), {1} certificate template(s)" -f @($facts.EnrollmentServices).Count, @($facts.Templates).Count) -ForegroundColor Cyan

$rows = @(Test-AdcsControls -Facts $facts)

$sorted = $rows | Sort-Object ControlId, Target
$sorted | Format-Table ControlId, Target, Status, Evidence -AutoSize

$summary = $sorted | Group-Object Status | ForEach-Object { "{0}={1}" -f $_.Name, $_.Count }
Write-Host ("Summary: {0}" -f ($summary -join ', ')) -ForegroundColor Yellow

if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
    $sorted | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "Report exported: $OutputPath" -ForegroundColor Green
}
