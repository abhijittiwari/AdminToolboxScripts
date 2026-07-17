<#
.SYNOPSIS
    Assesses Microsoft Entra ID tenant hardening controls AAD-001 to AAD-015 (Appendix B baseline).

.DESCRIPTION
    Read-only assessment of an Entra ID tenant against the security hardening baseline:
    MFA enforcement, phishing-resistant MFA for administrators, legacy authentication
    blocking, Conditional Access vs Security Defaults, per-user MFA retirement, PIM
    just-in-time activation, break-glass accounts, Global Administrator minimization,
    guest invitation restrictions, user consent restrictions, log retention/SIEM,
    password protection, compliant-device requirements for admin portals, Identity
    Protection risk policies, and periodic access reviews.

    Uses Microsoft Graph v1.0 (beta only for directory settings). Each control is
    emitted as one row with Status: Pass / Fail / Warning / Manual / Error.
    Controls that Graph cannot verify (diagnostic settings, per-user MFA at scale)
    are reported as Manual with the evidence that was collected.

    The script makes no changes to the tenant.

.PARAMETER BreakGlassUpn
    UPNs of the designated break-glass (emergency access) accounts. When provided,
    AAD-007 verifies the accounts exist, are enabled, and are excluded from every
    enabled Conditional Access policy.

.PARAMETER MaxGlobalAdmins
    Active Global Administrator assignment count above which AAD-008 warns. Default 5.

.PARAMETER OutputPath
    Optional CSV output path. If omitted, results are only written to the pipeline/table.

.EXAMPLE
    .\Export-EntraIdHardeningAssessment.ps1 -OutputPath .\EntraHardening.csv

.EXAMPLE
    .\Export-EntraIdHardeningAssessment.ps1 -BreakGlassUpn bg1@contoso.com, bg2@contoso.com
#>

[CmdletBinding()]
param(
    [string[]]$BreakGlassUpn = @(),
    [int]$MaxGlobalAdmins = 5,
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

function Connect-GraphIfNeeded {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    $requiredScopes = @('Policy.Read.All','Directory.Read.All','RoleManagement.Read.Directory','AccessReview.Read.All','User.Read.All')
    $context = Get-MgContext -ErrorAction SilentlyContinue
    if ($context) {
        Write-Host "Reusing existing Microsoft Graph context for tenant $($context.TenantId)." -ForegroundColor Cyan
        return
    }

    Connect-MgGraph -Scopes $requiredScopes -NoWelcome
}

function New-GraphRequestExecutor {
    return {
        param(
            [Parameter(Mandatory = $true)] [string]$Method,
            [Parameter(Mandatory = $true)] [string]$Uri,
            [Parameter(Mandatory = $false)] $Body
        )

        if ($null -ne $Body) {
            Invoke-MgGraphRequest -Method $Method -Uri $Uri -Body ($Body | ConvertTo-Json -Depth 10) -ContentType 'application/json'
        }
        else {
            Invoke-MgGraphRequest -Method $Method -Uri $Uri
        }
    }
}

function Get-ObjectPropertyValue {
    param(
        [Parameter(Mandatory = $false)] [AllowNull()] $InputObject,
        [Parameter(Mandatory = $true)] [string]$PropertyName
    )

    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [hashtable] -and $InputObject.ContainsKey($PropertyName)) {
        return $InputObject[$PropertyName]
    }

    $property = $InputObject.PSObject.Properties[$PropertyName]
    if ($property) { return $property.Value }

    $additionalProperties = $InputObject.PSObject.Properties['AdditionalProperties']
    if ($additionalProperties -and $additionalProperties.Value -is [hashtable] -and $additionalProperties.Value.ContainsKey($PropertyName)) {
        return $additionalProperties.Value[$PropertyName]
    }

    return $null
}

function Get-GraphCollection {
    param(
        [Parameter(Mandatory = $true)] [scriptblock]$GraphRequest,
        [Parameter(Mandatory = $true)] [string]$Uri
    )

    $items = [System.Collections.Generic.List[object]]::new()
    $nextUri = $Uri
    while ($nextUri) {
        $response = & $GraphRequest 'GET' $nextUri $null
        foreach ($item in @(Get-ObjectPropertyValue -InputObject $response -PropertyName 'value')) {
            if ($null -ne $item) { $items.Add($item) }
        }
        $nextUri = Get-ObjectPropertyValue -InputObject $response -PropertyName '@odata.nextLink'
    }
    return $items.ToArray()
}

function Get-EntraFacts {
    param(
        [Parameter(Mandatory = $true)] [scriptblock]$GraphRequest,
        [string[]]$BreakGlassUpn = @()
    )

    $globalAdminRoleId = '62e90394-69f5-4237-9190-012177145e10'
    $facts = @{}

    try {
        $organization = @(Get-GraphCollection -GraphRequest $GraphRequest -Uri 'https://graph.microsoft.com/v1.0/organization?$select=id,displayName')
        $facts.TenantId = [string](Get-ObjectPropertyValue -InputObject $organization[0] -PropertyName 'id')
        $facts.TenantName = [string](Get-ObjectPropertyValue -InputObject $organization[0] -PropertyName 'displayName')
    } catch {
        $facts.TenantId = ''
        $facts.TenantName = ''
    }

    try {
        $facts.CaPolicies = @(Get-GraphCollection -GraphRequest $GraphRequest -Uri 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies')
    } catch { $facts.CaPolicies = $null }

    try {
        $securityDefaults = & $GraphRequest 'GET' 'https://graph.microsoft.com/v1.0/policies/identitySecurityDefaultsEnforcementPolicy' $null
        $facts.SecurityDefaultsEnabled = [bool](Get-ObjectPropertyValue -InputObject $securityDefaults -PropertyName 'isEnabled')
    } catch { $facts.SecurityDefaultsEnabled = $null }

    try {
        $facts.AuthorizationPolicy = & $GraphRequest 'GET' 'https://graph.microsoft.com/v1.0/policies/authorizationPolicy' $null
    } catch { $facts.AuthorizationPolicy = $null }

    try {
        $activeGa = @(Get-GraphCollection -GraphRequest $GraphRequest -Uri ("https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments?`$filter=roleDefinitionId eq '{0}'" -f $globalAdminRoleId))
        $facts.GlobalAdminActiveCount = $activeGa.Count
    } catch { $facts.GlobalAdminActiveCount = $null }

    try {
        $eligibleGa = @(Get-GraphCollection -GraphRequest $GraphRequest -Uri ("https://graph.microsoft.com/v1.0/roleManagement/directory/roleEligibilityScheduleInstances?`$filter=roleDefinitionId eq '{0}'" -f $globalAdminRoleId))
        $facts.GlobalAdminEligibleCount = $eligibleGa.Count
    } catch { $facts.GlobalAdminEligibleCount = $null }

    $facts.PimGaApprovalRequired = $null
    $facts.PimGaEnablementRules = @()
    try {
        $policyAssignments = @(Get-GraphCollection -GraphRequest $GraphRequest -Uri ("https://graph.microsoft.com/v1.0/policies/roleManagementPolicyAssignments?`$filter=scopeId eq '/' and scopeType eq 'DirectoryRole' and roleDefinitionId eq '{0}'&`$expand=policy(`$expand=rules)" -f $globalAdminRoleId))
        if ($policyAssignments.Count -gt 0) {
            $policy = Get-ObjectPropertyValue -InputObject $policyAssignments[0] -PropertyName 'policy'
            foreach ($rule in @(Get-ObjectPropertyValue -InputObject $policy -PropertyName 'rules')) {
                $ruleId = [string](Get-ObjectPropertyValue -InputObject $rule -PropertyName 'id')
                if ($ruleId -eq 'Approval_EndUser_Assignment') {
                    $setting = Get-ObjectPropertyValue -InputObject $rule -PropertyName 'setting'
                    $facts.PimGaApprovalRequired = [bool](Get-ObjectPropertyValue -InputObject $setting -PropertyName 'isApprovalRequired')
                }
                if ($ruleId -eq 'Enablement_EndUser_Assignment') {
                    $facts.PimGaEnablementRules = @(Get-ObjectPropertyValue -InputObject $rule -PropertyName 'enabledRules')
                }
            }
        }
    } catch { }

    try {
        $facts.AccessReviewCount = @(Get-GraphCollection -GraphRequest $GraphRequest -Uri 'https://graph.microsoft.com/v1.0/identityGovernance/accessReviews/definitions?$top=100').Count
    } catch { $facts.AccessReviewCount = $null }

    $facts.PasswordRuleSettings = $null
    try {
        foreach ($setting in @(Get-GraphCollection -GraphRequest $GraphRequest -Uri 'https://graph.microsoft.com/v1.0/groupSettings')) {
            if ([string](Get-ObjectPropertyValue -InputObject $setting -PropertyName 'displayName') -eq 'Password Rule Settings') {
                $ruleValues = @{}
                foreach ($value in @(Get-ObjectPropertyValue -InputObject $setting -PropertyName 'values')) {
                    $ruleValues[[string](Get-ObjectPropertyValue -InputObject $value -PropertyName 'name')] = [string](Get-ObjectPropertyValue -InputObject $value -PropertyName 'value')
                }
                $facts.PasswordRuleSettings = $ruleValues
            }
        }
    } catch { }

    $facts.BreakGlassAccounts = @()
    foreach ($upn in $BreakGlassUpn) {
        try {
            $user = & $GraphRequest 'GET' ("https://graph.microsoft.com/v1.0/users/{0}?`$select=id,userPrincipalName,accountEnabled" -f [uri]::EscapeDataString($upn)) $null
            $facts.BreakGlassAccounts += @{
                Upn     = $upn
                Id      = [string](Get-ObjectPropertyValue -InputObject $user -PropertyName 'id')
                Enabled = [bool](Get-ObjectPropertyValue -InputObject $user -PropertyName 'accountEnabled')
                Found   = $true
            }
        } catch {
            $facts.BreakGlassAccounts += @{ Upn = $upn; Id = ''; Enabled = $false; Found = $false }
        }
    }

    return $facts
}

function Get-EnabledCaPolicies {
    param($Facts)
    return @(@($Facts.CaPolicies) | Where-Object { [string](Get-ObjectPropertyValue -InputObject $_ -PropertyName 'state') -eq 'enabled' })
}

function Test-CaPolicyProperty {
    param($Policy, [string[]]$Path)

    $current = $Policy
    foreach ($segment in $Path) {
        $current = Get-ObjectPropertyValue -InputObject $current -PropertyName $segment
        if ($null -eq $current) { return $null }
    }
    return $current
}

function Test-EntraControls {
    param(
        [Parameter(Mandatory = $true)] $Facts,
        [string[]]$BreakGlassUpn = @(),
        [int]$MaxGlobalAdmins = 5
    )

    $phishingResistantStrengthId = '00000000-0000-0000-0000-000000000004'
    $rows = [System.Collections.Generic.List[object]]::new()
    $tenant = if ($Facts.TenantName) { $Facts.TenantName } else { 'tenant' }
    $add = {
        param($Id, $Control, $Status, $Evidence)
        $rows.Add((New-ControlRow -ControlId $Id -Control $Control -Target $tenant -Status $Status -Evidence $Evidence))
    }

    $caAvailable = ($null -ne $Facts.CaPolicies)
    $enabledPolicies = Get-EnabledCaPolicies -Facts $Facts

    # AAD-001 MFA for all users
    if (-not $caAvailable) {
        & $add 'AAD-001' 'Multi-factor authentication shall be enforced for all users' 'Error' 'Conditional Access policies could not be read'
    }
    else {
        $allUserMfaPolicies = @($enabledPolicies | Where-Object {
            $includeUsers = @(Test-CaPolicyProperty -Policy $_ -Path @('conditions','users','includeUsers'))
            $builtIn = @(Test-CaPolicyProperty -Policy $_ -Path @('grantControls','builtInControls'))
            $strength = Test-CaPolicyProperty -Policy $_ -Path @('grantControls','authenticationStrength')
            ($includeUsers -contains 'All') -and (($builtIn -contains 'mfa') -or ($null -ne $strength))
        })
        if ($allUserMfaPolicies.Count -gt 0) {
            & $add 'AAD-001' 'Multi-factor authentication shall be enforced for all users' 'Pass' ("Enabled CA policy requiring MFA for all users: {0}" -f (($allUserMfaPolicies | ForEach-Object { Get-ObjectPropertyValue -InputObject $_ -PropertyName 'displayName' }) -join ', '))
        }
        elseif ($Facts.SecurityDefaultsEnabled -eq $true) {
            & $add 'AAD-001' 'Multi-factor authentication shall be enforced for all users' 'Warning' 'No CA policy requires MFA for all users; Security Defaults currently provide MFA prompts'
        }
        else {
            & $add 'AAD-001' 'Multi-factor authentication shall be enforced for all users' 'Fail' 'No enabled CA policy requires MFA for all users and Security Defaults are off'
        }
    }

    # AAD-002 phishing-resistant MFA for admins
    if ($caAvailable) {
        $adminStrengthPolicies = @($enabledPolicies | Where-Object {
            $includeRoles = @(Test-CaPolicyProperty -Policy $_ -Path @('conditions','users','includeRoles'))
            $includeUsers = @(Test-CaPolicyProperty -Policy $_ -Path @('conditions','users','includeUsers'))
            $strength = Test-CaPolicyProperty -Policy $_ -Path @('grantControls','authenticationStrength')
            $strengthId = [string](Get-ObjectPropertyValue -InputObject $strength -PropertyName 'id')
            $strengthName = [string](Get-ObjectPropertyValue -InputObject $strength -PropertyName 'displayName')
            ($includeRoles.Count -gt 0 -or $includeUsers -contains 'All') -and ($strengthId -eq $phishingResistantStrengthId -or $strengthName -match '[Pp]hishing')
        })
        if ($adminStrengthPolicies.Count -gt 0) {
            & $add 'AAD-002' 'Phishing-resistant MFA shall be required for all administrative roles' 'Pass' ("Policy: {0}" -f (($adminStrengthPolicies | ForEach-Object { Get-ObjectPropertyValue -InputObject $_ -PropertyName 'displayName' }) -join ', '))
        }
        else {
            & $add 'AAD-002' 'Phishing-resistant MFA shall be required for all administrative roles' 'Fail' 'No enabled CA policy applies a phishing-resistant authentication strength to directory roles'
        }
    }
    else {
        & $add 'AAD-002' 'Phishing-resistant MFA shall be required for all administrative roles' 'Error' 'Conditional Access policies could not be read'
    }

    # AAD-003 legacy auth blocked
    if ($caAvailable) {
        $legacyBlockPolicies = @($enabledPolicies | Where-Object {
            $clientApps = @(Test-CaPolicyProperty -Policy $_ -Path @('conditions','clientAppTypes'))
            $builtIn = @(Test-CaPolicyProperty -Policy $_ -Path @('grantControls','builtInControls'))
            (($clientApps -contains 'exchangeActiveSync') -or ($clientApps -contains 'other')) -and ($builtIn -contains 'block')
        })
        if ($legacyBlockPolicies.Count -gt 0) {
            & $add 'AAD-003' 'Legacy and basic authentication protocols shall be blocked' 'Pass' ("Blocking policy: {0}" -f (($legacyBlockPolicies | ForEach-Object { Get-ObjectPropertyValue -InputObject $_ -PropertyName 'displayName' }) -join ', '))
        }
        elseif ($Facts.SecurityDefaultsEnabled -eq $true) {
            & $add 'AAD-003' 'Legacy and basic authentication protocols shall be blocked' 'Warning' 'No CA block policy; Security Defaults currently block legacy authentication'
        }
        else {
            & $add 'AAD-003' 'Legacy and basic authentication protocols shall be blocked' 'Fail' 'No enabled CA policy blocks legacy authentication client apps'
        }
    }
    else {
        & $add 'AAD-003' 'Legacy and basic authentication protocols shall be blocked' 'Error' 'Conditional Access policies could not be read'
    }

    # AAD-004 CA replaces Security Defaults
    if ($null -eq $Facts.SecurityDefaultsEnabled) {
        & $add 'AAD-004' 'Conditional Access policies shall be used in place of Security Defaults for enterprise access control' 'Error' 'Security Defaults state could not be read'
    }
    elseif (-not $Facts.SecurityDefaultsEnabled -and $enabledPolicies.Count -gt 0) {
        & $add 'AAD-004' 'Conditional Access policies shall be used in place of Security Defaults for enterprise access control' 'Pass' ("Security Defaults off; {0} enabled CA policies" -f $enabledPolicies.Count)
    }
    elseif ($Facts.SecurityDefaultsEnabled) {
        & $add 'AAD-004' 'Conditional Access policies shall be used in place of Security Defaults for enterprise access control' 'Warning' 'Security Defaults are enabled; migrate to Conditional Access (requires Entra ID P1)'
    }
    else {
        & $add 'AAD-004' 'Conditional Access policies shall be used in place of Security Defaults for enterprise access control' 'Fail' 'Security Defaults off and no enabled CA policies - no tenant-wide access control'
    }

    # AAD-005 per-user MFA disabled
    & $add 'AAD-005' 'Per-user MFA shall be disabled in favor of Conditional Access' 'Manual' 'Review per-user MFA state in the Entra admin center (Users > Per-user MFA); all users should be Disabled once CA MFA policies cover them'

    # AAD-006 PIM JIT for privileged roles
    if ($null -eq $Facts.GlobalAdminEligibleCount) {
        & $add 'AAD-006' 'Privileged Identity Management (PIM) shall enforce just-in-time activation for privileged roles with MFA, approval, and justification' 'Error' 'PIM role eligibility could not be read (requires Entra ID P2)'
    }
    else {
        $enablementRules = @($Facts.PimGaEnablementRules)
        $pimEvidence = "Global Administrator: {0} eligible, {1} active; activation approval required={2}; enablement rules: {3}" -f $Facts.GlobalAdminEligibleCount, $Facts.GlobalAdminActiveCount, $Facts.PimGaApprovalRequired, ($enablementRules -join ', ')
        if ($Facts.GlobalAdminEligibleCount -gt 0 -and ($enablementRules -contains 'MultiFactorAuthentication') -and ($enablementRules -contains 'Justification')) {
            $status = if ($Facts.PimGaApprovalRequired) { 'Pass' } else { 'Warning' }
            $note = if ($status -eq 'Warning') { ' (activation approval not required)' } else { '' }
            & $add 'AAD-006' 'Privileged Identity Management (PIM) shall enforce just-in-time activation for privileged roles with MFA, approval, and justification' $status ($pimEvidence + $note)
        }
        else {
            & $add 'AAD-006' 'Privileged Identity Management (PIM) shall enforce just-in-time activation for privileged roles with MFA, approval, and justification' 'Fail' $pimEvidence
        }
    }

    # AAD-007 break-glass accounts
    if ($BreakGlassUpn.Count -eq 0) {
        & $add 'AAD-007' 'Break-glass emergency access accounts shall be configured and excluded from Conditional Access policies' 'Manual' 'Provide -BreakGlassUpn to verify existence and CA exclusions automatically'
    }
    else {
        $issues = [System.Collections.Generic.List[string]]::new()
        foreach ($account in @($Facts.BreakGlassAccounts)) {
            if (-not $account.Found) { $issues.Add(("{0} not found" -f $account.Upn)); continue }
            if (-not $account.Enabled) { $issues.Add(("{0} is disabled" -f $account.Upn)) }
            foreach ($policy in $enabledPolicies) {
                $excludeUsers = @(Test-CaPolicyProperty -Policy $policy -Path @('conditions','users','excludeUsers'))
                $includeUsers = @(Test-CaPolicyProperty -Policy $policy -Path @('conditions','users','includeUsers'))
                if (($includeUsers -contains 'All') -and ($excludeUsers -notcontains $account.Id)) {
                    $issues.Add(("{0} not excluded from policy '{1}'" -f $account.Upn, (Get-ObjectPropertyValue -InputObject $policy -PropertyName 'displayName')))
                }
            }
        }
        if ($issues.Count -eq 0) {
            & $add 'AAD-007' 'Break-glass emergency access accounts shall be configured and excluded from Conditional Access policies' 'Pass' ("{0} break-glass account(s) exist, enabled, and excluded from all-user CA policies" -f $BreakGlassUpn.Count)
        }
        else {
            & $add 'AAD-007' 'Break-glass emergency access accounts shall be configured and excluded from Conditional Access policies' 'Fail' ($issues -join '; ')
        }
    }

    # AAD-008 GA minimization
    if ($null -eq $Facts.GlobalAdminActiveCount) {
        & $add 'AAD-008' 'Global Administrator and other highly privileged role assignments shall be minimized (least privilege)' 'Error' 'Role assignments could not be read'
    }
    else {
        $gaCount = [int]$Facts.GlobalAdminActiveCount
        $gaEvidence = "{0} active Global Administrator assignment(s); {1} PIM-eligible" -f $gaCount, $Facts.GlobalAdminEligibleCount
        if ($gaCount -ge 2 -and $gaCount -le $MaxGlobalAdmins) {
            & $add 'AAD-008' 'Global Administrator and other highly privileged role assignments shall be minimized (least privilege)' 'Pass' $gaEvidence
        }
        elseif ($gaCount -lt 2) {
            & $add 'AAD-008' 'Global Administrator and other highly privileged role assignments shall be minimized (least privilege)' 'Warning' ($gaEvidence + ' (Microsoft recommends at least 2 for resilience)')
        }
        else {
            & $add 'AAD-008' 'Global Administrator and other highly privileged role assignments shall be minimized (least privilege)' 'Warning' ($gaEvidence + (" (threshold {0})" -f $MaxGlobalAdmins))
        }
    }

    # AAD-009 guest invitations
    $authzPolicy = $Facts.AuthorizationPolicy
    if ($null -eq $authzPolicy) {
        & $add 'AAD-009' 'Guest user invitations shall be restricted to administrators and guest access reviewed periodically' 'Error' 'Authorization policy could not be read'
    }
    else {
        $allowInvitesFrom = [string](Get-ObjectPropertyValue -InputObject $authzPolicy -PropertyName 'allowInvitesFrom')
        if ($allowInvitesFrom -in @('none','admins','adminsAndGuestInviters')) {
            & $add 'AAD-009' 'Guest user invitations shall be restricted to administrators and guest access reviewed periodically' 'Pass' ("allowInvitesFrom={0}; periodic guest access review requires manual confirmation" -f $allowInvitesFrom)
        }
        else {
            & $add 'AAD-009' 'Guest user invitations shall be restricted to administrators and guest access reviewed periodically' 'Fail' ("allowInvitesFrom={0} (should be adminsAndGuestInviters or stricter)" -f $allowInvitesFrom)
        }
    }

    # AAD-010 user consent restricted
    if ($null -eq $authzPolicy) {
        & $add 'AAD-010' 'User consent for third-party applications shall be restricted' 'Error' 'Authorization policy could not be read'
    }
    else {
        $defaultPermissions = Get-ObjectPropertyValue -InputObject $authzPolicy -PropertyName 'defaultUserRolePermissions'
        $grantPolicies = @(Get-ObjectPropertyValue -InputObject $defaultPermissions -PropertyName 'permissionGrantPoliciesAssigned')
        $grantEvidence = "permissionGrantPoliciesAssigned: {0}" -f ($(if ($grantPolicies.Count -gt 0) { $grantPolicies -join ', ' } else { '(none)' }))
        if ($grantPolicies.Count -eq 0) {
            & $add 'AAD-010' 'User consent for third-party applications shall be restricted' 'Pass' ($grantEvidence + ' - user consent disabled; admin consent workflow applies')
        }
        elseif (@($grantPolicies | Where-Object { $_ -match 'legacy' }).Count -gt 0) {
            & $add 'AAD-010' 'User consent for third-party applications shall be restricted' 'Fail' ($grantEvidence + ' - users can consent to any application')
        }
        else {
            & $add 'AAD-010' 'User consent for third-party applications shall be restricted' 'Warning' ($grantEvidence + ' - limited user consent allowed; confirm this matches policy')
        }
    }

    # AAD-011 log retention + SIEM
    & $add 'AAD-011' 'Sign-in and audit logs shall be retained (minimum 30 days) and forwarded to a SIEM' 'Manual' 'Verify Entra diagnostic settings export SignInLogs and AuditLogs to the SIEM/Log Analytics; native retention is 30 days with P1/P2, 7 days otherwise'

    # AAD-012 password protection
    $passwordSettings = $Facts.PasswordRuleSettings
    if ($null -eq $passwordSettings) {
        & $add 'AAD-012' 'Password protection with banned-password lists and smart lockout shall be enabled' 'Warning' 'No Password Rule Settings object found; Microsoft defaults apply in the cloud - verify on-premises password protection deployment'
    }
    else {
        $bannedCheck = [string]$passwordSettings['EnableBannedPasswordCheck']
        $onPremMode = [string]$passwordSettings['BannedPasswordCheckOnPremisesMode']
        $lockoutThreshold = [string]$passwordSettings['LockoutThreshold']
        if ($bannedCheck -eq 'True' -and $onPremMode -eq 'Enforce') {
            & $add 'AAD-012' 'Password protection with banned-password lists and smart lockout shall be enabled' 'Pass' ("EnableBannedPasswordCheck=True, on-premises mode=Enforce, LockoutThreshold={0}" -f $lockoutThreshold)
        }
        else {
            & $add 'AAD-012' 'Password protection with banned-password lists and smart lockout shall be enabled' 'Warning' ("EnableBannedPasswordCheck={0}, BannedPasswordCheckOnPremisesMode={1}, LockoutThreshold={2}" -f $bannedCheck, $onPremMode, $lockoutThreshold)
        }
    }

    # AAD-013 compliant device for admin portals
    if ($caAvailable) {
        $adminPortalPolicies = @($enabledPolicies | Where-Object {
            $includeApps = @(Test-CaPolicyProperty -Policy $_ -Path @('conditions','applications','includeApplications'))
            $builtIn = @(Test-CaPolicyProperty -Policy $_ -Path @('grantControls','builtInControls'))
            (($includeApps -contains 'MicrosoftAdminPortals') -or ($includeApps -contains 'All')) -and (($builtIn -contains 'compliantDevice') -or ($builtIn -contains 'domainJoinedDevice'))
        })
        if ($adminPortalPolicies.Count -gt 0) {
            & $add 'AAD-013' 'Access to administrative portals shall require compliant or hybrid-joined devices' 'Pass' ("Policy: {0}" -f (($adminPortalPolicies | ForEach-Object { Get-ObjectPropertyValue -InputObject $_ -PropertyName 'displayName' }) -join ', '))
        }
        else {
            & $add 'AAD-013' 'Access to administrative portals shall require compliant or hybrid-joined devices' 'Fail' 'No enabled CA policy requires compliant/hybrid-joined devices for Microsoft Admin Portals'
        }
    }
    else {
        & $add 'AAD-013' 'Access to administrative portals shall require compliant or hybrid-joined devices' 'Error' 'Conditional Access policies could not be read'
    }

    # AAD-014 risk-based policies
    if ($caAvailable) {
        $signInRiskPolicies = @($enabledPolicies | Where-Object { @(Test-CaPolicyProperty -Policy $_ -Path @('conditions','signInRiskLevels')).Count -gt 0 })
        $userRiskPolicies = @($enabledPolicies | Where-Object { @(Test-CaPolicyProperty -Policy $_ -Path @('conditions','userRiskLevels')).Count -gt 0 })
        if ($signInRiskPolicies.Count -gt 0 -and $userRiskPolicies.Count -gt 0) {
            & $add 'AAD-014' 'High-risk sign-ins and users shall be blocked or remediated via Identity Protection' 'Pass' ("Sign-in risk policies: {0}; user risk policies: {1}" -f $signInRiskPolicies.Count, $userRiskPolicies.Count)
        }
        elseif ($signInRiskPolicies.Count -gt 0 -or $userRiskPolicies.Count -gt 0) {
            & $add 'AAD-014' 'High-risk sign-ins and users shall be blocked or remediated via Identity Protection' 'Warning' ("Sign-in risk policies: {0}; user risk policies: {1} - both dimensions should be covered" -f $signInRiskPolicies.Count, $userRiskPolicies.Count)
        }
        else {
            & $add 'AAD-014' 'High-risk sign-ins and users shall be blocked or remediated via Identity Protection' 'Fail' 'No enabled CA policy conditions on sign-in or user risk (requires Entra ID P2)'
        }
    }
    else {
        & $add 'AAD-014' 'High-risk sign-ins and users shall be blocked or remediated via Identity Protection' 'Error' 'Conditional Access policies could not be read'
    }

    # AAD-015 access reviews
    if ($null -eq $Facts.AccessReviewCount) {
        & $add 'AAD-015' 'Periodic access reviews shall be performed for privileged roles and group memberships' 'Error' 'Access review definitions could not be read (requires Entra ID P2)'
    }
    elseif ([int]$Facts.AccessReviewCount -gt 0) {
        & $add 'AAD-015' 'Periodic access reviews shall be performed for privileged roles and group memberships' 'Pass' ("{0} access review definition(s) found; verify scope covers privileged roles" -f $Facts.AccessReviewCount)
    }
    else {
        & $add 'AAD-015' 'Periodic access reviews shall be performed for privileged roles and group memberships' 'Fail' 'No access review definitions exist'
    }

    return $rows.ToArray()
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
Connect-GraphIfNeeded
$graphRequest = New-GraphRequestExecutor

Write-Host 'Collecting Entra ID tenant facts...' -ForegroundColor Cyan
$facts = Get-EntraFacts -GraphRequest $graphRequest -BreakGlassUpn $BreakGlassUpn
$rows = @(Test-EntraControls -Facts $facts -BreakGlassUpn $BreakGlassUpn -MaxGlobalAdmins $MaxGlobalAdmins)

$rows | Format-Table ControlId, Status, Evidence -AutoSize

$summary = $rows | Group-Object Status | ForEach-Object { "{0}={1}" -f $_.Name, $_.Count }
Write-Host ("Summary: {0}" -f ($summary -join ', ')) -ForegroundColor Yellow

if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
    $rows | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "Report exported: $OutputPath" -ForegroundColor Green
}
