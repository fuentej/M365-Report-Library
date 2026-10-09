#Requires -Version 7.0

BeforeAll {
    $script:Root = Join-Path $PSScriptRoot '../../..' | Resolve-Path | Select-Object -ExpandProperty Path
    $script:Collectors = Join-Path $script:Root 'reports/identity-posture/collectors'
    $script:Samples = Join-Path $script:Root 'reports/identity-posture/samples'

    . (Join-Path $script:Root 'shared/tests/TenantCmdletStubs.ps1')
    Import-Module (Join-Path $script:Root 'shared/M365ReportLibrary.psm1') -Force

    $script:Schema = Import-PowerShellDataFile -LiteralPath (Join-Path $script:Collectors 'IdentityPostureSchema.psd1')

    function New-TestFolder {
        $path = Join-Path ([System.IO.Path]::GetTempPath()) ('identity-posture-' + [guid]::NewGuid().ToString('N'))
        New-Item -Path $path -ItemType Directory -Force | Out-Null
        return $path
    }

    function Get-HeaderText {
        param([Parameter(Mandatory)][string]$Path)
        return ((Get-CsvHeaderColumn -Path $Path) -join ',')
    }

    function Get-LogText {
        param([Parameter(Mandatory)][string]$Folder)
        $log = Join-Path $Folder 'run.log'
        if (Test-Path -LiteralPath $log) { return (Get-Content -LiteralPath $log -Raw) }
        return ''
    }

    function Invoke-CollectorScript {
        param([Parameter(Mandatory)][string]$Name, [hashtable]$Arguments = @{})
        & (Join-Path $script:Collectors $Name) @Arguments
    }

    # userRegistrationDetails: https://learn.microsoft.com/graph/api/resources/userregistrationdetails
    # List: https://learn.microsoft.com/graph/api/authenticationmethodsroot-list-userregistrationdetails
    function New-MockRegistration {
        param([string]$Id = 'user-1')

        [pscustomobject]@{
            Id                                            = $Id
            UserPrincipalName                             = 'avery.abara@example.com'
            UserDisplayName                               = 'Avery Abara'
            UserType                                      = 'member'
            IsAdmin                                       = $true
            IsMfaRegistered                               = $true
            IsMfaCapable                                  = $true
            IsPasswordlessCapable                         = $false
            IsSsprEnabled                                 = $true
            IsSsprRegistered                              = $true
            IsSsprCapable                                 = $true
            IsSystemPreferredAuthenticationMethodEnabled  = $true
            DefaultMfaMethod                              = 'mobilePhone'
            UserPreferredMethodForSecondaryAuthentication = 'push'
            MethodsRegistered                             = @('microsoftAuthenticatorPush', 'mobilePhone')
            SystemPreferredAuthenticationMethods          = @('microsoftAuthenticatorPush')
            LastUpdatedDateTime                           = [datetime]'2026-08-11T10:15:00Z'
        }
    }

    # conditionalAccessPolicy: https://learn.microsoft.com/graph/api/resources/conditionalaccesspolicy
    # List: https://learn.microsoft.com/graph/api/conditionalaccessroot-list-policies
    function New-MockPolicy {
        param([string]$Id = 'policy-1', [string]$State = 'enabledForReportingButNotEnforced')

        [pscustomobject]@{
            Id               = $Id
            DisplayName      = 'Block legacy authentication'
            State            = $State
            CreatedDateTime  = [datetime]'2025-05-01T12:00:00Z'
            ModifiedDateTime = [datetime]'2026-07-15T12:00:00Z'
            Conditions       = [pscustomobject]@{
                ClientAppTypes   = @('exchangeActiveSync', 'other')
                SignInRiskLevels = @('high')
                UserRiskLevels   = @()
                Applications     = [pscustomobject]@{
                    IncludeApplications = @('All')
                    ExcludeApplications = @('00000002-0000-0ff1-ce00-000000000000')
                }
                Users            = [pscustomobject]@{
                    IncludeUsers  = @('All')
                    ExcludeUsers  = @('aaaaaaaa-0000-4000-8000-000000000004')
                    IncludeGroups = @()
                    ExcludeGroups = @('bbbbbbbb-0000-4000-8000-000000000001')
                    IncludeRoles  = @()
                    ExcludeRoles  = @()
                }
            }
            GrantControls    = [pscustomobject]@{
                Operator        = 'OR'
                BuiltInControls = @('block')
            }
        }
    }

    # unifiedRoleAssignmentScheduleInstance:
    # https://learn.microsoft.com/graph/api/resources/unifiedroleassignmentscheduleinstance
    function New-MockActiveInstance {
        param([string]$Id = 'active-1')

        [pscustomobject]@{
            Id                       = $Id
            PrincipalId              = 'user-1'
            RoleDefinitionId         = '62e90394-69f5-4237-9190-012177145e10'
            DirectoryScopeId         = '/'
            AppScopeId               = $null
            AssignmentType           = 'Activated'
            MemberType               = 'Direct'
            StartDateTime            = [datetime]'2026-06-01T00:00:00Z'
            EndDateTime              = $null
            RoleAssignmentOriginId   = 'origin-1'
            RoleAssignmentScheduleId = 'schedule-1'
        }
    }

    # unifiedRoleEligibilityScheduleInstance:
    # https://learn.microsoft.com/graph/api/resources/unifiedroleeligibilityscheduleinstance
    function New-MockEligibleInstance {
        param([string]$Id = 'eligible-1')

        [pscustomobject]@{
            Id                        = $Id
            PrincipalId               = 'user-1'
            RoleDefinitionId          = '62e90394-69f5-4237-9190-012177145e10'
            DirectoryScopeId          = '/'
            AppScopeId                = $null
            MemberType                = 'Direct'
            StartDateTime             = [datetime]'2026-06-01T00:00:00Z'
            EndDateTime               = [datetime]'2027-06-01T00:00:00Z'
            RoleEligibilityScheduleId = 'eligibility-1'
        }
    }

    # unifiedRoleAssignment: https://learn.microsoft.com/graph/api/resources/unifiedroleassignment
    function New-MockRoleAssignment {
        param([string]$Id = 'assignment-1')

        [pscustomobject]@{
            Id               = $Id
            PrincipalId      = 'user-1'
            RoleDefinitionId = '62e90394-69f5-4237-9190-012177145e10'
            DirectoryScopeId = '/'
            AppScopeId       = $null
        }
    }

    # List users with signInActivity (example 11):
    # https://learn.microsoft.com/graph/api/user-list#example-11-get-users-including-their-last-sign-in-time
    # signInActivity resource: https://learn.microsoft.com/graph/api/resources/signinactivity
    function New-MockUser {
        param([string]$Id = 'user-1', [object]$Activity = 'default')

        if ($Activity -is [string] -and $Activity -eq 'default') {
            $Activity = [pscustomobject]@{
                LastSignInDateTime               = [datetime]'2026-09-21T07:45:00Z'
                LastNonInteractiveSignInDateTime = [datetime]'2026-09-21T22:10:00Z'
                LastSuccessfulSignInDateTime     = [datetime]'2026-09-21T07:45:00Z'
            }
        }
        [pscustomobject]@{
            Id                = $Id
            UserPrincipalName = 'avery.abara@example.com'
            SignInActivity    = $Activity
        }
    }

    # riskyUser: https://learn.microsoft.com/graph/api/resources/riskyuser
    # List: https://learn.microsoft.com/graph/api/riskyuser-list
    function New-MockRiskyUser {
        param([string]$Id = 'risky-1')

        [pscustomobject]@{
            Id                      = $Id
            UserPrincipalName       = 'casey.chaudhry@example.com'
            UserDisplayName         = 'Casey Chaudhry'
            RiskLevel               = 'high'
            RiskState               = 'atRisk'
            RiskDetail              = 'none'
            RiskLastUpdatedDateTime = [datetime]'2026-08-28T14:00:00Z'
            IsDeleted               = $false
            IsProcessing            = $false
        }
    }

    # signIn: https://learn.microsoft.com/graph/api/resources/signin
    # List: https://learn.microsoft.com/graph/api/signin-list
    function New-MockSignIn {
        param(
            [string]$Id = 'signin-1',
            [string]$ClientApp = 'IMAP',
            [datetime]$At = [datetime]'2026-09-10T06:05:00Z'
        )

        [pscustomobject]@{
            Id                      = $Id
            CreatedDateTime         = $At
            UserId                  = 'user-1'
            UserPrincipalName       = 'casey.chaudhry@example.com'
            AppDisplayName          = 'Office 365 Exchange Online'
            ResourceDisplayName     = 'Office 365 Exchange Online'
            IPAddress               = '203.0.113.10'
            ClientAppUsed           = $ClientApp
            IsInteractive           = $true
            ConditionalAccessStatus = 'notApplied'
            Status                  = [pscustomobject]@{ ErrorCode = 53003; FailureReason = 'Blocked by Conditional Access' }
        }
    }

    # One row per state collector: the script, its CSV, the schema key, the cmdlet it
    # reads through, and the builder for a mock object.
    $script:StateCases = @(
        @{ Script = 'Get-AuthenticationMethods.ps1'; Csv = 'authentication-methods.csv'; Key = 'AuthenticationMethods'; Cmdlet = 'Get-MgReportAuthenticationMethodUserRegistrationDetail'; Builder = 'New-MockRegistration' }
        @{ Script = 'Get-ConditionalAccessPolicies.ps1'; Csv = 'conditional-access-policies.csv'; Key = 'ConditionalAccessPolicies'; Cmdlet = 'Get-MgIdentityConditionalAccessPolicy'; Builder = 'New-MockPolicy' }
        @{ Script = 'Get-ActiveRoleAssignments.ps1'; Csv = 'role-assignments-active.csv'; Key = 'ActiveRoleAssignments'; Cmdlet = 'Get-MgRoleManagementDirectoryRoleAssignmentScheduleInstance'; Builder = 'New-MockActiveInstance' }
        @{ Script = 'Get-EligibleRoleAssignments.ps1'; Csv = 'role-assignments-eligible.csv'; Key = 'EligibleRoleAssignments'; Cmdlet = 'Get-MgRoleManagementDirectoryRoleEligibilityScheduleInstance'; Builder = 'New-MockEligibleInstance' }
        @{ Script = 'Get-RoleAssignments.ps1'; Csv = 'role-assignments.csv'; Key = 'RoleAssignments'; Cmdlet = 'Get-MgRoleManagementDirectoryRoleAssignment'; Builder = 'New-MockRoleAssignment' }
        @{ Script = 'Get-UserSignInActivity.ps1'; Csv = 'user-signin-activity.csv'; Key = 'UserSignInActivity'; Cmdlet = 'Get-MgUser'; Builder = 'New-MockUser' }
        @{ Script = 'Get-RiskyUsers.ps1'; Csv = 'risky-users.csv'; Key = 'RiskyUsers'; Cmdlet = 'Get-MgRiskyUser'; Builder = 'New-MockRiskyUser' }
    )
}

AfterAll {
    Remove-Module M365ReportLibrary -Force -ErrorAction SilentlyContinue
}

Describe 'Each collector writes the columns of its sample file' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
    }
    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It '<Csv>' -ForEach @(
        @{ Script = 'Get-AuthenticationMethods.ps1'; Csv = 'authentication-methods.csv'; Key = 'AuthenticationMethods'; Cmdlet = 'Get-MgReportAuthenticationMethodUserRegistrationDetail'; Builder = 'New-MockRegistration' }
        @{ Script = 'Get-ConditionalAccessPolicies.ps1'; Csv = 'conditional-access-policies.csv'; Key = 'ConditionalAccessPolicies'; Cmdlet = 'Get-MgIdentityConditionalAccessPolicy'; Builder = 'New-MockPolicy' }
        @{ Script = 'Get-ActiveRoleAssignments.ps1'; Csv = 'role-assignments-active.csv'; Key = 'ActiveRoleAssignments'; Cmdlet = 'Get-MgRoleManagementDirectoryRoleAssignmentScheduleInstance'; Builder = 'New-MockActiveInstance' }
        @{ Script = 'Get-EligibleRoleAssignments.ps1'; Csv = 'role-assignments-eligible.csv'; Key = 'EligibleRoleAssignments'; Cmdlet = 'Get-MgRoleManagementDirectoryRoleEligibilityScheduleInstance'; Builder = 'New-MockEligibleInstance' }
        @{ Script = 'Get-RoleAssignments.ps1'; Csv = 'role-assignments.csv'; Key = 'RoleAssignments'; Cmdlet = 'Get-MgRoleManagementDirectoryRoleAssignment'; Builder = 'New-MockRoleAssignment' }
        @{ Script = 'Get-UserSignInActivity.ps1'; Csv = 'user-signin-activity.csv'; Key = 'UserSignInActivity'; Cmdlet = 'Get-MgUser'; Builder = 'New-MockUser' }
        @{ Script = 'Get-RiskyUsers.ps1'; Csv = 'risky-users.csv'; Key = 'RiskyUsers'; Cmdlet = 'Get-MgRiskyUser'; Builder = 'New-MockRiskyUser' }
    ) {
        $builder = $Builder
        Mock -CommandName $Cmdlet -MockWith { & $builder }

        Invoke-CollectorScript $Script @{ OutputPath = $script:folder }

        $produced = Join-Path $script:folder $Csv
        @(Import-Csv -LiteralPath $produced).Count | Should -Be 1
        Get-HeaderText -Path $produced | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples $Csv))
        Get-HeaderText -Path $produced | Should -Be ($script:Schema[$Key] -join ',')
    }

    It 'signins.csv' {
        Mock Get-MgAuditLogSignIn -MockWith { New-MockSignIn }

        Invoke-CollectorScript 'Get-LegacySignIns.ps1' @{ OutputPath = $script:folder; LookbackDays = 1 }

        $produced = Join-Path $script:folder 'signins.csv'
        @(Import-Csv -LiteralPath $produced).Count | Should -BeGreaterThan 0
        Get-HeaderText -Path $produced | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples 'signins.csv'))
        Get-HeaderText -Path $produced | Should -Be ($script:Schema.SignIns -join ',')
    }
}

Describe 'Each connection targets the endpoints of its -Environment' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-MgGraph -ModuleName M365ReportLibrary -MockWith { }
        foreach ($case in $script:StateCases) { Mock -CommandName $case.Cmdlet -MockWith { @() } }
        Mock Get-MgAuditLogSignIn -MockWith { @() }
    }
    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    # https://learn.microsoft.com/graph/deployments: GCC calls the global service;
    # GCC High is the USGov environment.
    It '<Script> signs in to <Graph> for <Environment>' -ForEach @(
        foreach ($script in 'Get-AuthenticationMethods.ps1', 'Get-ConditionalAccessPolicies.ps1', 'Get-ActiveRoleAssignments.ps1', 'Get-EligibleRoleAssignments.ps1', 'Get-RoleAssignments.ps1', 'Get-UserSignInActivity.ps1', 'Get-RiskyUsers.ps1', 'Get-LegacySignIns.ps1') {
            @{ Script = $script; Environment = 'Commercial'; Graph = 'Global' }
            @{ Script = $script; Environment = 'GCC'; Graph = 'Global' }
            @{ Script = $script; Environment = 'GCCHigh'; Graph = 'USGov' }
        }
    ) {
        $arguments = @{ OutputPath = $script:folder; Environment = $Environment }
        if ($Script -eq 'Get-LegacySignIns.ps1') { $arguments['LookbackDays'] = 1 }

        Invoke-CollectorScript $Script $arguments

        $expected = $Graph
        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter { $Environment -eq $expected }
    }

    It 'requests the permission each source adds to the shared scopes' -ForEach @(
        @{ Script = 'Get-ConditionalAccessPolicies.ps1'; Scope = 'Policy.Read.All' }
        @{ Script = 'Get-ActiveRoleAssignments.ps1'; Scope = 'RoleAssignmentSchedule.Read.Directory' }
        @{ Script = 'Get-EligibleRoleAssignments.ps1'; Scope = 'RoleEligibilitySchedule.Read.Directory' }
        @{ Script = 'Get-RoleAssignments.ps1'; Scope = 'RoleManagement.Read.Directory' }
        @{ Script = 'Get-RiskyUsers.ps1'; Scope = 'IdentityRiskyUser.Read.All' }
        @{ Script = 'Get-AuthenticationMethods.ps1'; Scope = 'AuditLog.Read.All' }
        @{ Script = 'Get-UserSignInActivity.ps1'; Scope = 'AuditLog.Read.All' }
    ) {
        Invoke-CollectorScript $Script @{ OutputPath = $script:folder }

        $wanted = $Scope
        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter { $Scopes -contains $wanted -and $Scopes -contains 'Directory.Read.All' }
    }

    It 'signs in app-only without scopes when given a certificate' {
        Invoke-CollectorScript 'Get-RiskyUsers.ps1' @{
            OutputPath = $script:folder; AppId = 'app-1'; CertificateThumbprint = 'AB12'; TenantId = 'tenant-1'
        }

        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter {
            $ClientId -eq 'app-1' -and $CertificateThumbprint -eq 'AB12' -and $TenantId -eq 'tenant-1' -and -not $Scopes
        }
    }
}

Describe 'Paging is followed' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
    }
    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It '<Cmdlet> keeps every object when the service returns more than one page' -ForEach @(
        @{ Script = 'Get-AuthenticationMethods.ps1'; Csv = 'authentication-methods.csv'; Cmdlet = 'Get-MgReportAuthenticationMethodUserRegistrationDetail'; Builder = 'New-MockRegistration' }
        @{ Script = 'Get-ConditionalAccessPolicies.ps1'; Csv = 'conditional-access-policies.csv'; Cmdlet = 'Get-MgIdentityConditionalAccessPolicy'; Builder = 'New-MockPolicy' }
        @{ Script = 'Get-ActiveRoleAssignments.ps1'; Csv = 'role-assignments-active.csv'; Cmdlet = 'Get-MgRoleManagementDirectoryRoleAssignmentScheduleInstance'; Builder = 'New-MockActiveInstance' }
        @{ Script = 'Get-EligibleRoleAssignments.ps1'; Csv = 'role-assignments-eligible.csv'; Cmdlet = 'Get-MgRoleManagementDirectoryRoleEligibilityScheduleInstance'; Builder = 'New-MockEligibleInstance' }
        @{ Script = 'Get-RoleAssignments.ps1'; Csv = 'role-assignments.csv'; Cmdlet = 'Get-MgRoleManagementDirectoryRoleAssignment'; Builder = 'New-MockRoleAssignment' }
        @{ Script = 'Get-UserSignInActivity.ps1'; Csv = 'user-signin-activity.csv'; Cmdlet = 'Get-MgUser'; Builder = 'New-MockUser' }
        @{ Script = 'Get-RiskyUsers.ps1'; Csv = 'risky-users.csv'; Cmdlet = 'Get-MgRiskyUser'; Builder = 'New-MockRiskyUser' }
    ) {
        # 1,250 objects is more than a 500-row page (signInActivity, riskyUsers) or a
        # 1,000-row page (signIns). Without -All the mock returns the first 100, which is
        # what a single request returns. https://learn.microsoft.com/graph/paging
        $builder = $Builder
        Mock -CommandName $Cmdlet -MockWith {
            $limit = if ($All) { 1250 } else { 100 }
            1..$limit | ForEach-Object { & $builder -Id "id-$_" }
        }

        Invoke-CollectorScript $Script @{ OutputPath = $script:folder }

        @(Import-Csv -LiteralPath (Join-Path $script:folder $Csv)).Count | Should -Be 1250
        Should -Invoke -CommandName $Cmdlet -Times 1 -Exactly -ParameterFilter { $All }
    }

    It 'requests signInActivity explicitly' {
        Mock Get-MgUser -MockWith { New-MockUser }

        Invoke-CollectorScript 'Get-UserSignInActivity.ps1' @{ OutputPath = $script:folder }

        Should -Invoke Get-MgUser -Times 1 -Exactly -ParameterFilter { $All -and $Property -contains 'signInActivity' }
    }

    It 'reads every page of sign-ins for every legacy client app in a window' {
        Mock Get-MgAuditLogSignIn -MockWith {
            $limit = if ($All) { 1200 } else { 100 }
            $app = if ($Filter -match "clientAppUsed eq '([^']+)'") { $Matches[1] } else { 'none' }
            1..$limit | ForEach-Object { New-MockSignIn -Id "$app-$_" -ClientApp $app }
        }

        Invoke-CollectorScript 'Get-LegacySignIns.ps1' @{
            OutputPath = $script:folder; StartDate = [datetime]'2026-09-10T00:00:00Z'; EndDate = [datetime]'2026-09-11T00:00:00Z'
        }

        @(Import-Csv -LiteralPath (Join-Path $script:folder 'signins.csv')).Count | Should -Be (1200 * 6)
        foreach ($app in $script:Schema.LegacyClientAppValues) {
            $expected = "clientAppUsed eq '$app'"
            Should -Invoke Get-MgAuditLogSignIn -Times 1 -Exactly -ParameterFilter { $All -and $Filter.Contains($expected) }
        }
    }
}

Describe 'Empty signInActivity values stay empty' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
    }
    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'writes no date for the values Graph uses to say "none"' {
        # Example 11 on the list-users page returns 0001-01-01T00:00:00Z for
        # lastNonInteractiveSignInDateTime and "" for lastSuccessfulSignInDateTime.
        Mock Get-MgUser -MockWith {
            @(
                New-MockUser -Id 'never' -Activity ([pscustomobject]@{
                        LastSignInDateTime               = $null
                        LastNonInteractiveSignInDateTime = [datetime]'0001-01-01T00:00:00Z'
                        LastSuccessfulSignInDateTime     = ''
                    })
                New-MockUser -Id 'no-property' -Activity $null
                New-MockUser -Id 'offset' -Activity ([pscustomobject]@{
                        LastSignInDateTime               = [datetimeoffset]'2026-09-21T07:45:00+00:00'
                        LastNonInteractiveSignInDateTime = [datetimeoffset]::MinValue
                        LastSuccessfulSignInDateTime     = $null
                    })
                New-MockUser -Id 'real'
            )
        }

        Invoke-CollectorScript 'Get-UserSignInActivity.ps1' @{ OutputPath = $script:folder }

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:folder 'user-signin-activity.csv'))
        $rows.Count | Should -Be 4
        foreach ($id in 'never', 'no-property') {
            $row = $rows | Where-Object UserId -EQ $id
            $row.LastSignInDateTime | Should -BeNullOrEmpty
            $row.LastNonInteractiveSignInDateTime | Should -BeNullOrEmpty
            $row.LastSuccessfulSignInDateTime | Should -BeNullOrEmpty
        }
        $offset = $rows | Where-Object UserId -EQ 'offset'
        $offset.LastSignInDateTime | Should -Be '2026-09-21T07:45:00Z'
        $offset.LastNonInteractiveSignInDateTime | Should -BeNullOrEmpty
        $offset.LastSuccessfulSignInDateTime | Should -BeNullOrEmpty

        $real = $rows | Where-Object UserId -EQ 'real'
        $real.LastSignInDateTime | Should -Be '2026-09-21T07:45:00Z'
        $real.LastNonInteractiveSignInDateTime | Should -Be '2026-09-21T22:10:00Z'
        $real.LastSuccessfulSignInDateTime | Should -Be '2026-09-21T07:45:00Z'
        (Get-Content -LiteralPath (Join-Path $script:folder 'user-signin-activity.csv') -Raw) | Should -Not -Match '0001-01-01'
    }

    It 'converts the "no value" forms by themselves' {
        . (Join-Path $script:Collectors 'IdentityPostureHelpers.ps1')
        ConvertTo-SignInTimestamp $null | Should -Be ''
        ConvertTo-SignInTimestamp '' | Should -Be ''
        ConvertTo-SignInTimestamp '0001-01-01T00:00:00Z' | Should -Be ''
        ConvertTo-SignInTimestamp ([datetime]::MinValue) | Should -Be ''
        ConvertTo-SignInTimestamp '2026-09-21T07:45:00Z' | Should -Be '2026-09-21T07:45:00Z'
    }
}

Describe 'What each collector keeps' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
    }
    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'flattens authentication methods and keeps the booleans' {
        Mock Get-MgReportAuthenticationMethodUserRegistrationDetail -MockWith { New-MockRegistration }

        Invoke-CollectorScript 'Get-AuthenticationMethods.ps1' @{ OutputPath = $script:folder }

        $row = Import-Csv -LiteralPath (Join-Path $script:folder 'authentication-methods.csv')
        $row.UserId | Should -Be 'user-1'
        $row.IsAdmin | Should -Be 'True'
        $row.IsPasswordlessCapable | Should -Be 'False'
        $row.MethodsRegistered | Should -Be 'microsoftAuthenticatorPush;mobilePhone'
        $row.LastUpdatedDateTime | Should -Be '2026-08-11T10:15:00Z'
    }

    It 'keeps the report-only state and flattens targets and controls' {
        Mock Get-MgIdentityConditionalAccessPolicy -MockWith {
            @(New-MockPolicy -Id 'p1' -State 'enabledForReportingButNotEnforced'; New-MockPolicy -Id 'p2' -State 'disabled')
        }

        Invoke-CollectorScript 'Get-ConditionalAccessPolicies.ps1' @{ OutputPath = $script:folder }

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:folder 'conditional-access-policies.csv'))
        ($rows | Where-Object Id -EQ 'p1').State | Should -Be 'enabledForReportingButNotEnforced'
        ($rows | Where-Object Id -EQ 'p2').State | Should -Be 'disabled'
        $rows[0].IncludeUsers | Should -Be 'All'
        $rows[0].ExcludeGroups | Should -Be 'bbbbbbbb-0000-4000-8000-000000000001'
        $rows[0].ClientAppTypes | Should -Be 'exchangeActiveSync;other'
        $rows[0].SignInRiskLevels | Should -Be 'high'
        $rows[0].GrantOperator | Should -Be 'OR'
        $rows[0].BuiltInControls | Should -Be 'block'
    }

    It 'keeps assignment type, member type and an empty end date for a permanent assignment' {
        Mock Get-MgRoleManagementDirectoryRoleAssignmentScheduleInstance -MockWith { New-MockActiveInstance }

        Invoke-CollectorScript 'Get-ActiveRoleAssignments.ps1' @{ OutputPath = $script:folder }

        $row = Import-Csv -LiteralPath (Join-Path $script:folder 'role-assignments-active.csv')
        $row.AssignmentType | Should -Be 'Activated'
        $row.MemberType | Should -Be 'Direct'
        $row.EndDateTime | Should -BeNullOrEmpty
        $row.StartDateTime | Should -Be '2026-06-01T00:00:00Z'
    }

    It 'keeps eligible assignments in their own file' {
        Mock Get-MgRoleManagementDirectoryRoleEligibilityScheduleInstance -MockWith { New-MockEligibleInstance }

        Invoke-CollectorScript 'Get-EligibleRoleAssignments.ps1' @{ OutputPath = $script:folder }

        $row = Import-Csv -LiteralPath (Join-Path $script:folder 'role-assignments-eligible.csv')
        $row.RoleEligibilityScheduleId | Should -Be 'eligibility-1'
        $row.EndDateTime | Should -Be '2027-06-01T00:00:00Z'
    }

    It 'keeps the documented riskLevel and riskState values as returned' {
        Mock Get-MgRiskyUser -MockWith { New-MockRiskyUser }

        Invoke-CollectorScript 'Get-RiskyUsers.ps1' @{ OutputPath = $script:folder }

        $row = Import-Csv -LiteralPath (Join-Path $script:folder 'risky-users.csv')
        $row.RiskLevel | Should -Be 'high'
        $row.RiskState | Should -Be 'atRisk'
        $row.RiskLastUpdatedDateTime | Should -Be '2026-08-28T14:00:00Z'
    }

    It 'appends a second snapshot without repeating the header, and a same-day re-run adds nothing' {
        Mock Get-MgRiskyUser -MockWith { New-MockRiskyUser }

        Invoke-CollectorScript 'Get-RiskyUsers.ps1' @{ OutputPath = $script:folder }
        Invoke-CollectorScript 'Get-RiskyUsers.ps1' @{ OutputPath = $script:folder }

        @(Get-Content -LiteralPath (Join-Path $script:folder 'risky-users.csv')).Count | Should -Be 2
    }
}

Describe 'A sign-in collector resumes from its watermark' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
    }
    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'starts at the latest CreatedDateTime already exported and does not repeat that sign-in' {
        $mock = { New-MockSignIn -Id 'signin-1' -At ([datetime]'2026-09-10T06:05:00Z') }
        Mock Get-MgAuditLogSignIn -MockWith $mock
        Invoke-CollectorScript 'Get-LegacySignIns.ps1' @{
            OutputPath = $script:folder; StartDate = [datetime]'2026-09-10T00:00:00Z'; EndDate = [datetime]'2026-09-10T12:00:00Z'
        }

        Mock Get-MgAuditLogSignIn -MockWith {
            @(New-MockSignIn -Id 'signin-1' -At ([datetime]'2026-09-10T06:05:00Z'); New-MockSignIn -Id 'signin-2' -At ([datetime]'2026-09-10T08:00:00Z'))
        }
        Invoke-CollectorScript 'Get-LegacySignIns.ps1' @{ OutputPath = $script:folder; EndDate = [datetime]'2026-09-10T12:00:00Z' }

        Should -Invoke Get-MgAuditLogSignIn -ParameterFilter { $Filter.StartsWith('createdDateTime ge 2026-09-10T06:05:00Z') }
        $ids = @(Import-Csv -LiteralPath (Join-Path $script:folder 'signins.csv')).Id | Sort-Object -Unique
        $ids | Should -Be @('signin-1', 'signin-2')
    }

    It 'keeps complete windows and stops when a later window fails' {
        Mock Get-MgAuditLogSignIn -MockWith {
            if ($Filter -match 'createdDateTime ge 2026-09-11') { throw 'Request timed out.' }
            New-MockSignIn -Id 'signin-1'
        }

        { Invoke-CollectorScript 'Get-LegacySignIns.ps1' @{
                OutputPath = $script:folder; StartDate = [datetime]'2026-09-10T00:00:00Z'; EndDate = [datetime]'2026-09-12T00:00:00Z'
            } } | Should -Throw '*part-way*'

        @(Import-Csv -LiteralPath (Join-Path $script:folder 'signins.csv')).Count | Should -Be 1
        Get-LogText -Folder $script:folder | Should -Match 'Request timed out'
    }

    It 'rejects an empty range that was asked for explicitly' {
        Mock Get-MgAuditLogSignIn -MockWith { @() }
        { Invoke-CollectorScript 'Get-LegacySignIns.ps1' @{
                OutputPath = $script:folder; StartDate = [datetime]'2026-09-12T00:00:00Z'; EndDate = [datetime]'2026-09-10T00:00:00Z'
            } } | Should -Throw '*range is empty*'
    }
}

Describe 'A source Microsoft documents as NotAvailable writes the header only' {
    BeforeEach {
        $script:folder = New-TestFolder
        # The contract marks no source NotAvailable today, so the tests point the
        # collectors at a copy of the schema that does.
        $text = Get-Content -LiteralPath (Join-Path $script:Collectors 'IdentityPostureSchema.psd1') -Raw
        $script:notAvailable = Join-Path $script:folder 'schema.psd1'
        $changed = $text -replace "Status = '(Available|Unverified)'", "Status = 'NotAvailable'"
        Set-Content -LiteralPath $script:notAvailable -Value $changed -Encoding utf8
        Mock Connect-MgGraph -ModuleName M365ReportLibrary -MockWith { }
        foreach ($case in $script:StateCases) { Mock -CommandName $case.Cmdlet -MockWith { throw "$($case.Cmdlet) must not be called" } }
        Mock Get-MgAuditLogSignIn -MockWith { throw 'Get-MgAuditLogSignIn must not be called' }
    }
    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It '<Script> skips, signs in to nothing and logs why' -ForEach @(
        @{ Script = 'Get-AuthenticationMethods.ps1'; Csv = 'authentication-methods.csv'; Key = 'AuthenticationMethods' }
        @{ Script = 'Get-ConditionalAccessPolicies.ps1'; Csv = 'conditional-access-policies.csv'; Key = 'ConditionalAccessPolicies' }
        @{ Script = 'Get-ActiveRoleAssignments.ps1'; Csv = 'role-assignments-active.csv'; Key = 'ActiveRoleAssignments' }
        @{ Script = 'Get-EligibleRoleAssignments.ps1'; Csv = 'role-assignments-eligible.csv'; Key = 'EligibleRoleAssignments' }
        @{ Script = 'Get-RoleAssignments.ps1'; Csv = 'role-assignments.csv'; Key = 'RoleAssignments' }
        @{ Script = 'Get-UserSignInActivity.ps1'; Csv = 'user-signin-activity.csv'; Key = 'UserSignInActivity' }
        @{ Script = 'Get-RiskyUsers.ps1'; Csv = 'risky-users.csv'; Key = 'RiskyUsers' }
        @{ Script = 'Get-LegacySignIns.ps1'; Csv = 'signins.csv'; Key = 'SignIns' }
    ) {
        Invoke-CollectorScript $Script @{ OutputPath = $script:folder; SchemaPath = $script:notAvailable } 3>$null

        $path = Join-Path $script:folder $Csv
        (Get-Content -LiteralPath $path).Count | Should -Be 1
        Get-HeaderText -Path $path | Should -Be ($script:Schema[$Key] -join ',')
        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 0 -Exactly
        Get-LogText -Folder $script:folder | Should -Match 'documented as unavailable'
    }
}

Describe 'An UNVERIFIED source is attempted with a warning' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-MgGraph -ModuleName M365ReportLibrary -MockWith { }
    }
    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'asks GCC High for signInActivity and warns' {
        Mock Get-MgUser -MockWith { New-MockUser }

        Invoke-CollectorScript 'Get-UserSignInActivity.ps1' @{ OutputPath = $script:folder; Environment = 'GCCHigh' } 3>$null

        Should -Invoke Get-MgUser -Times 1 -Exactly
        @(Import-Csv -LiteralPath (Join-Path $script:folder 'user-signin-activity.csv')).Count | Should -Be 1
        Get-LogText -Folder $script:folder | Should -Match 'UNVERIFIED'
    }

    It 'records a refusal, writes the header only and does not report success' {
        Mock Get-MgUser -MockWith { throw 'Property signInActivity is not supported in this cloud.' }

        { Invoke-CollectorScript 'Get-UserSignInActivity.ps1' @{ OutputPath = $script:folder; Environment = 'GCCHigh' } 3>$null } |
            Should -Throw '*not supported in this cloud*'

        (Get-Content -LiteralPath (Join-Path $script:folder 'user-signin-activity.csv')).Count | Should -Be 1
        Get-LogText -Folder $script:folder | Should -Match 'not supported in this cloud'
    }

    It 'does not warn about an Available source' {
        Mock Get-MgUser -MockWith { New-MockUser }

        Invoke-CollectorScript 'Get-UserSignInActivity.ps1' @{ OutputPath = $script:folder; Environment = 'Commercial' }

        Get-LogText -Folder $script:folder | Should -Not -Match 'UNVERIFIED'
    }
}

Describe 'A missing licence is a logged skip, not a failure' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
    }
    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It '<Cmdlet> writes the header only and does not throw' -ForEach @(
        @{ Script = 'Get-ActiveRoleAssignments.ps1'; Csv = 'role-assignments-active.csv'; Cmdlet = 'Get-MgRoleManagementDirectoryRoleAssignmentScheduleInstance' }
        @{ Script = 'Get-EligibleRoleAssignments.ps1'; Csv = 'role-assignments-eligible.csv'; Cmdlet = 'Get-MgRoleManagementDirectoryRoleEligibilityScheduleInstance' }
        @{ Script = 'Get-RiskyUsers.ps1'; Csv = 'risky-users.csv'; Cmdlet = 'Get-MgRiskyUser' }
        @{ Script = 'Get-AuthenticationMethods.ps1'; Csv = 'authentication-methods.csv'; Cmdlet = 'Get-MgReportAuthenticationMethodUserRegistrationDetail' }
    ) {
        Mock -CommandName $Cmdlet -MockWith { throw 'Authentication_RequestFromNonPremiumTenantOrB2CTenant: Neither tenant is B2C or tenant does not have premium license' }

        { Invoke-CollectorScript $Script @{ OutputPath = $script:folder } 3>$null } | Should -Not -Throw

        (Get-Content -LiteralPath (Join-Path $script:folder $Csv)).Count | Should -Be 1
        $log = Get-LogText -Folder $script:folder
        $log | Should -Match '\[Warning\]'
        $log | Should -Match 'not licensed'
        $log | Should -Not -Match '\[Error\]'
    }

    It 'throws a permission failure instead of reporting success' {
        Mock Get-MgRiskyUser -MockWith { throw 'Insufficient privileges to complete the operation.' }

        { Invoke-CollectorScript 'Get-RiskyUsers.ps1' @{ OutputPath = $script:folder } 3>$null } |
            Should -Throw '*Insufficient privileges*'

        (Get-Content -LiteralPath (Join-Path $script:folder 'risky-users.csv')).Count | Should -Be 1
        $log = Get-LogText -Folder $script:folder
        $log | Should -Match '\[Error\]'
        $log | Should -Not -Match 'not licensed'
    }

    It 'skips the sign-in log without an error when it is not licensed' {
        Mock Get-MgAuditLogSignIn -MockWith { throw 'Tenant does not have a premium license' }

        Invoke-CollectorScript 'Get-LegacySignIns.ps1' @{ OutputPath = $script:folder; LookbackDays = 1 } 3>$null

        (Get-Content -LiteralPath (Join-Path $script:folder 'signins.csv')).Count | Should -Be 1
        Get-LogText -Folder $script:folder | Should -Match 'not licensed'
    }

    It 'throws when the first sign-in window fails and nothing was written' {
        Mock Get-MgAuditLogSignIn -MockWith { throw 'Request timed out.' }

        { Invoke-CollectorScript 'Get-LegacySignIns.ps1' @{ OutputPath = $script:folder; LookbackDays = 1 } 3>$null } |
            Should -Throw '*Reading the sign-in log failed*'

        (Get-Content -LiteralPath (Join-Path $script:folder 'signins.csv')).Count | Should -Be 1
        $log = Get-LogText -Folder $script:folder
        $log | Should -Match '\[Error\]'
        $log | Should -Match 'Request timed out'
        $log | Should -Not -Match 'windows read before the failure were kept'
    }
}

Describe 'Run-All.ps1' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
        # Invoke-EntraUserCollector signs in from inside the module.
        Mock Connect-M365Service -ModuleName M365ReportLibrary -MockWith { }
        Mock Get-MgUser -MockWith { @() }
        foreach ($case in $script:StateCases) { Mock -CommandName $case.Cmdlet -MockWith { @() } }
        Mock Get-MgAuditLogSignIn -MockWith { @() }
    }
    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'writes every CSV, carries on past a failing collector and exits with an error' {
        Mock Get-MgRiskyUser -MockWith { throw 'boom' }

        { Invoke-CollectorScript 'Run-All.ps1' @{ OutputPath = $script:folder; StartDate = [datetime]'2026-09-10T00:00:00Z'; EndDate = [datetime]'2026-09-10T06:00:00Z' } 3>$null } |
            Should -Throw '*1 collector(s) stopped with an error*'

        foreach ($name in 'users', 'authentication-methods', 'conditional-access-policies', 'role-assignments-active', 'role-assignments-eligible', 'role-assignments', 'user-signin-activity', 'risky-users', 'signins') {
            Test-Path -LiteralPath (Join-Path $script:folder "$name.csv") | Should -BeTrue -Because $name
        }
        Get-LogText -Folder $script:folder | Should -Match 'risky-users collector stopped'
    }
}
