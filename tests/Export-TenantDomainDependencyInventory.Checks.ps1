#Requires -Version 7
$ErrorActionPreference = 'Stop'

$script:ScriptPath = Join-Path $PSScriptRoot '..' 'Export-TenantDomainDependencyInventory.ps1'
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

function Test-RequiredFunctionsExist {
    $ast = Get-ScriptAst
    $functionNames = @($ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true) | ForEach-Object Name)
    foreach ($name in @('Add-ExportError','Add-ExportWarnings','Initialize-OutputFolder','Export-CsvWithHeaders','Connect-ExchangeOnlineIfNeeded','Connect-GraphForGroupFallbackIfNeeded','New-GraphRequestExecutor','Normalize-Domain','Join-ReportValues','Get-ObjectPropertyValue','Get-GraphCollectionValues','ConvertFrom-GraphDirectoryObjectLink','ConvertTo-StringValues','Get-AddressValue','Test-ValueUsesDomain','Get-TargetDomainValues','Add-Reason','Get-ReviewDisposition','Get-TestingAndCommsNeeds','Get-RecipientObjectType','ConvertTo-DependencyRow','Get-DistributionGroupMembersSafe','Get-UnifiedGroupLinksSafe','Get-GraphUnifiedGroupLinksSafe','Get-MailboxPermissionValuesSafe','Get-CalendarProcessingSafe','Get-AddressRewriteRows','Get-DistributionGroupsSafe','Get-DynamicDistributionGroupsFromRecipientFallback','Get-DynamicDistributionGroupsSafe','Get-MailboxesSafe','Get-DependencyInventory')) {
        Assert-True -Condition ($functionNames -contains $name) -Name "Function exists: $name"
    }
}
Test-RequiredFunctionsExist

function Test-StaticScriptContract {
    $scriptText = Get-Content -Path $script:ScriptPath -Raw
    foreach ($cmdlet in @('Get-DistributionGroup','Get-DynamicDistributionGroup','Get-UnifiedGroup','Get-EXOMailbox','Get-MailboxPermission','Get-RecipientPermission','Get-CalendarProcessing','Get-AddressRewriteEntry')) {
        Assert-True -Condition ($scriptText -match [regex]::Escape($cmdlet)) -Name "Exchange cmdlet present: $cmdlet"
    }
    foreach ($file in @('DomainDependencyInventory.csv','AddressRewriteRules.csv','Errors-DomainDependencyInventory.csv','Summary.csv')) {
        Assert-True -Condition ($scriptText -match [regex]::Escape($file)) -Name "Output file present: $file"
    }
    foreach ($column in @('Owners','MembershipSource','AllowsExternalSenders','ModerationEnabled','ForwardingOrRoutingTargets','RecommendedDisposition','TestingAndCommsNeeds')) {
        Assert-True -Condition ($scriptText -match [regex]::Escape($column)) -Name "Inventory column present: $column"
    }
    Assert-True -Condition ($scriptText -match 'Microsoft.Graph.Authentication' -and $scriptText -match 'GroupMember.Read.All' -and $scriptText -match 'Group.Read.All') -Name 'Graph fallback scopes and module are present'
    Assert-True -Condition ($scriptText -match 'WarningAction SilentlyContinue' -and $scriptText -match 'Add-ExportWarnings') -Name 'Exchange validation warnings are captured without console noise'
    Assert-True -Condition ($scriptText -match 'Get-EXORecipient:DynamicDistributionGroup' -and $scriptText -match 'Get-DynamicDistributionGroup:Identity') -Name 'Dynamic distribution group fallback is present'
}
Test-StaticScriptContract

function Test-ExportWarningCapture {
    $script:Errors.Clear()
    $warningOutput = @(Add-ExportWarnings -Identity 'BrokenGroup' -Stage 'Groups' -Operation 'Get-DistributionGroup' -Warnings @('Validation warning') 3>&1)

    Assert-True -Condition ($script:Errors.Count -eq 1) -Name 'Add-ExportWarnings records warning as export error'
    Assert-True -Condition ($script:Errors[0].Identity -eq 'BrokenGroup' -and $script:Errors[0].Message -eq 'Validation warning') -Name 'Add-ExportWarnings preserves warning identity and message'
    Assert-True -Condition ($warningOutput.Count -eq 0) -Name 'Add-ExportWarnings does not write warning stream'
}
Test-ExportWarningCapture

function Test-DomainMatchingHelpers {
    Assert-True -Condition ((Normalize-Domain -Domain '@WAE.COM') -eq 'wae.com') -Name 'Normalize-Domain trims at sign and lowercases'
    Assert-True -Condition (Test-ValueUsesDomain -Value 'SMTP:group@wae.com' -Domain 'wae.com') -Name 'Test-ValueUsesDomain matches exact SMTP domain'
    Assert-True -Condition (Test-ValueUsesDomain -Value 'smtp:group@mail.wae.com' -Domain 'wae.com') -Name 'Test-ValueUsesDomain matches subdomain SMTP domain'
    Assert-True -Condition (-not (Test-ValueUsesDomain -Value 'smtp:group@notwae.com' -Domain 'wae.com')) -Name 'Test-ValueUsesDomain rejects lookalike domain'
    $targetValues = @(Get-TargetDomainValues -Values @('smtp:a@wae.com','smtp:b@contoso.com','route wae.com') -Domain 'wae.com')
    Assert-True -Condition ($targetValues.Count -eq 2) -Name 'Get-TargetDomainValues filters matching values'
}
Test-DomainMatchingHelpers

function Test-DispositionAndCommsNeeds {
    $reasons = @('TargetDomainPrimarySmtp','TargetDomainMember','ModerationEnabled')
    $disposition = Get-ReviewDisposition -ObjectType 'DistributionList' -Reasons $reasons
    $needs = Get-TestingAndCommsNeeds -ObjectType 'DistributionList' -Reasons $reasons -AllowsExternalSenders $true

    Assert-True -Condition ($disposition -match 'Migration' -and $disposition -match 'OwnerDecision') -Name 'Get-ReviewDisposition combines migration and owner decision'
    Assert-True -Condition ($needs -match 'inbound delivery' -and $needs -match 'external sender' -and $needs -match 'moderator') -Name 'Get-TestingAndCommsNeeds includes delivery, external, and moderation tests'
}
Test-DispositionAndCommsNeeds

function Test-ConvertToDependencyRowForImpactedGroup {
    $group = [pscustomobject]@{
        Identity                                = 'All WAE'
        DisplayName                             = 'All WAE'
        RecipientType                           = 'MailUniversalDistributionGroup'
        RecipientTypeDetails                    = 'MailUniversalDistributionGroup'
        PrimarySmtpAddress                      = 'all@wae.com'
        Alias                                   = 'allwae'
        ExternalDirectoryObjectId               = 'group-1'
        EmailAddresses                          = @('SMTP:all@wae.com','smtp:all.alias@wae.com','smtp:all@contoso.com')
        ManagedBy                               = @('owner@contoso.com')
        RequireSenderAuthenticationEnabled      = $false
        AcceptMessagesOnlyFromSendersOrMembers  = @('allowed@wae.com')
        RejectMessagesFromSendersOrMembers      = @()
        ModerationEnabled                       = $true
        ModeratedBy                             = @('mod@contoso.com')
        BypassModerationFromSendersOrMembers    = @()
        GrantSendOnBehalfTo                     = @()
    }
    $members = @([pscustomobject]@{ PrimarySmtpAddress = 'member@wae.com'; WindowsEmailAddress = ''; ExternalEmailAddress = ''; Name = 'Member' })

    $row = ConvertTo-DependencyRow -Object $group -ObjectType 'DistributionList' -Domain 'wae.com' -Members $members -Owners $group.ManagedBy -MembershipSource 'StaticDistributionGroup'

    Assert-True -Condition ($row.DependencyReasons -match 'TargetDomainPrimarySmtp' -and $row.DependencyReasons -match 'TargetDomainMember' -and $row.DependencyReasons -match 'TargetDomainSenderRestriction') -Name 'ConvertTo-DependencyRow identifies address, member, and sender dependencies'
    Assert-True -Condition ($row.TargetDomainMemberCount -eq 1 -and $row.AllowsExternalSenders -eq $true) -Name 'ConvertTo-DependencyRow counts target members and external sender flag'
    Assert-True -Condition ($row.RecommendedDisposition -match 'Migration' -and $row.RecommendedDisposition -match 'OwnerDecision') -Name 'ConvertTo-DependencyRow recommends migration and owner decision'
    Assert-True -Condition ($row.TestingAndCommsNeeds -match 'Communicate impact') -Name 'ConvertTo-DependencyRow includes comms need'
}
Test-ConvertToDependencyRowForImpactedGroup

function Test-GraphUnifiedGroupLinksFallbackPaging {
    $script:GraphUris = [System.Collections.Generic.List[string]]::new()
    $mockGraph = {
        param([string]$Method, [string]$Uri, $Body)
        $script:GraphUris.Add($Uri) | Out-Null
        if ($Uri -match 'page2') {
            return @{
                value = @(
                    @{ displayName = 'Second User'; mail = 'second@wae.com'; userPrincipalName = 'second@wae.com' }
                )
            }
        }
        return @{
            value = @(
                @{ displayName = 'First User'; mail = 'first@wae.com'; userPrincipalName = 'first@wae.com' }
            )
            '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/groups/group-1/members?page2'
        }
    }

    $group = [pscustomobject]@{ Identity = 'M365 Group'; ExternalDirectoryObjectId = 'group-1' }
    $members = @(Get-GraphUnifiedGroupLinksSafe -Group $group -LinkType Members -GraphRequest $mockGraph)

    Assert-True -Condition ($script:GraphUris[0] -match '/groups/group-1/members' -and $script:GraphUris.Count -eq 2) -Name 'Graph unified group fallback pages member results'
    Assert-True -Condition ($members.Count -eq 2 -and $members[1].PrimarySmtpAddress -eq 'second@wae.com') -Name 'Graph unified group fallback converts member objects'
}
Test-GraphUnifiedGroupLinksFallbackPaging

function Test-ConvertToDependencyRowUsesGraphOwnerObjects {
    $group = [pscustomobject]@{
        Identity                           = 'M365 WAE'
        DisplayName                        = 'M365 WAE'
        RecipientType                      = 'GroupMailbox'
        RecipientTypeDetails               = 'GroupMailbox'
        PrimarySmtpAddress                 = 'm365@contoso.com'
        Alias                              = 'm365wae'
        ExternalDirectoryObjectId          = 'group-2'
        EmailAddresses                     = @('smtp:m365@contoso.com')
        RequireSenderAuthenticationEnabled = $true
        ModerationEnabled                  = $false
    }
    $owners = @([pscustomobject]@{ PrimarySmtpAddress = 'owner@wae.com'; UserPrincipalName = 'owner@wae.com'; DisplayName = 'Owner User' })

    $row = ConvertTo-DependencyRow -Object $group -ObjectType 'Microsoft365Group' -Domain 'wae.com' -Owners $owners -MembershipSource 'Microsoft365GroupLinks'

    Assert-True -Condition ($row.DependencyReasons -match 'TargetDomainOwner') -Name 'ConvertTo-DependencyRow detects target-domain Graph owner objects'
    Assert-True -Condition ($row.Owners -match 'owner@wae.com') -Name 'ConvertTo-DependencyRow reports Graph owner address'
}
Test-ConvertToDependencyRowUsesGraphOwnerObjects

function Get-DynamicDistributionGroup {
    param(
        [string]$Identity,
        [object]$ResultSize,
        [System.Management.Automation.ActionPreference]$ErrorAction,
        [System.Management.Automation.ActionPreference]$WarningAction,
        [string]$WarningVariable
    )

    if ([string]::IsNullOrWhiteSpace($Identity)) {
        throw 'A server side error has occurred because of which the operation could not be completed.'
    }

    return [pscustomobject]@{
        Identity             = $Identity
        DisplayName          = 'Dynamic WAE'
        RecipientTypeDetails = 'DynamicDistributionGroup'
        PrimarySmtpAddress   = 'dynamic@wae.com'
        EmailAddresses       = @('SMTP:dynamic@wae.com')
        ManagedBy            = @('owner@contoso.com')
        RecipientFilter      = "Department -eq 'IT'"
    }
}

function Get-EXORecipient {
    param(
        [object]$ResultSize,
        [object]$RecipientTypeDetails,
        [string[]]$Properties,
        [System.Management.Automation.ActionPreference]$ErrorAction,
        [System.Management.Automation.ActionPreference]$WarningAction,
        [string]$WarningVariable
    )

    return @([pscustomobject]@{
        Identity             = 'Dynamic WAE'
        DisplayName          = 'Dynamic WAE'
        RecipientTypeDetails = 'DynamicDistributionGroup'
        PrimarySmtpAddress   = 'dynamic@wae.com'
        EmailAddresses       = @('SMTP:dynamic@wae.com')
    })
}

function Test-DynamicDistributionGroupServerErrorFallsBackToExoRecipient {
    $script:Errors.Clear()

    $groups = @(Get-DynamicDistributionGroupsSafe -RecipientProperties @('EmailAddresses','RecipientTypeDetails'))

    Assert-True -Condition ($groups.Count -eq 1 -and $groups[0].RecipientFilter -eq "Department -eq 'IT'") -Name 'Dynamic group server-side enumeration error falls back through EXO recipient and identity lookup'
    Assert-True -Condition (@($script:Errors | Where-Object { $_.Operation -eq 'Get-DynamicDistributionGroup' -and $_.Message -match 'server side error' }).Count -eq 1) -Name 'Dynamic group enumeration failure is logged instead of terminating'
}
Test-DynamicDistributionGroupServerErrorFallsBackToExoRecipient

if ($script:Failures.Count -gt 0) {
    Write-Host "$($script:Failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host "All $($script:Passes) checks passed." -ForegroundColor Green
exit 0
