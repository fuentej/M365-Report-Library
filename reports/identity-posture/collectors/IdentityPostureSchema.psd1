@{
    # The column order of every CSV this report produces. The collectors, the
    # sample files and the tests all read the order from here, so a sample file
    # can never drift from its collector's output.
    # Contract: docs/candidates/identity-posture.md

    # Source 2. Properties of userRegistrationDetails (v1.0):
    # https://learn.microsoft.com/graph/api/resources/userregistrationdetails
    # The API does not return disabled users. defaultMfaMethod exists only on the beta
    # resource and is not collected.
    # https://learn.microsoft.com/graph/api/resources/userregistrationdetails?view=graph-rest-beta
    AuthenticationMethods = @(
        'RunDate'
        'UserId'
        'UserPrincipalName'
        'UserDisplayName'
        'UserType'
        'IsAdmin'
        'IsMfaRegistered'
        'IsMfaCapable'
        'IsPasswordlessCapable'
        'IsSsprEnabled'
        'IsSsprRegistered'
        'IsSsprCapable'
        'IsSystemPreferredAuthenticationMethodEnabled'
        'UserPreferredMethodForSecondaryAuthentication'
        'MethodsRegistered'
        'SystemPreferredAuthenticationMethods'
        'LastUpdatedDateTime'
    )

    # Source 3. State is enabled, disabled or enabledForReportingButNotEnforced.
    # https://learn.microsoft.com/graph/api/resources/conditionalaccesspolicy
    ConditionalAccessPolicies = @(
        'RunDate'
        'Id'
        'DisplayName'
        'State'
        'CreatedDateTime'
        'ModifiedDateTime'
        'IncludeUsers'
        'ExcludeUsers'
        'IncludeGroups'
        'ExcludeGroups'
        'IncludeRoles'
        'ExcludeRoles'
        'IncludeApplications'
        'ExcludeApplications'
        'ClientAppTypes'
        'SignInRiskLevels'
        'UserRiskLevels'
        'GrantOperator'
        'BuiltInControls'
        'CustomAuthenticationFactors'
        'TermsOfUse'
        # displayName of grantControls.authenticationStrength. A policy can set this
        # and leave builtInControls empty.
        # https://learn.microsoft.com/graph/api/resources/conditionalaccessgrantcontrols
        'AuthenticationStrength'
    )

    # Source 4a. AssignmentType is Assigned or Activated; MemberType is Inherited,
    # Direct or Group. An empty EndDateTime is a permanent assignment.
    # https://learn.microsoft.com/graph/api/resources/unifiedroleassignmentscheduleinstance
    ActiveRoleAssignments = @(
        'RunDate'
        'Id'
        'PrincipalId'
        'RoleDefinitionId'
        'DirectoryScopeId'
        'AppScopeId'
        'AssignmentType'
        'MemberType'
        'StartDateTime'
        'EndDateTime'
        'RoleAssignmentOriginId'
        'RoleAssignmentScheduleId'
    )

    # Source 4b.
    # https://learn.microsoft.com/graph/api/resources/unifiedroleeligibilityscheduleinstance
    EligibleRoleAssignments = @(
        'RunDate'
        'Id'
        'PrincipalId'
        'RoleDefinitionId'
        'DirectoryScopeId'
        'AppScopeId'
        'MemberType'
        'StartDateTime'
        'EndDateTime'
        'RoleEligibilityScheduleId'
    )

    # Source 4c. Active holders without PIM; never eligible assignments.
    # https://learn.microsoft.com/graph/api/resources/unifiedroleassignment
    RoleAssignments = @(
        'RunDate'
        'Id'
        'PrincipalId'
        'RoleDefinitionId'
        'DirectoryScopeId'
        'AppScopeId'
    )

    # Source 5. An empty cell means Graph returned no value. It never holds the
    # 0001-01-01T00:00:00Z that Graph returns for "none", and it is not proof the
    # account never signed in.
    # https://learn.microsoft.com/graph/api/resources/signinactivity
    UserSignInActivity = @(
        'RunDate'
        'UserId'
        'UserPrincipalName'
        'LastSignInDateTime'
        'LastNonInteractiveSignInDateTime'
        'LastSuccessfulSignInDateTime'
    )

    # Source 6. Interactive-in-nature sign-ins and successful federated sign-ins only
    # (v1.0), narrowed to the legacy clientAppUsed values below.
    # https://learn.microsoft.com/graph/api/resources/signin#properties
    SignIns = @(
        'CreatedDateTime'
        'Id'
        'UserId'
        'UserPrincipalName'
        'AppDisplayName'
        'ResourceDisplayName'
        'IpAddress'
        'ClientAppUsed'
        'IsInteractive'
        'ErrorCode'
        'ConditionalAccessStatus'
    )

    LegacyClientAppValues = @(
        'Exchange ActiveSync'
        'IMAP'
        'MAPI'
        'SMTP'
        'POP'
        'other clients'
    )

    # Source 7. https://learn.microsoft.com/graph/api/resources/riskyuser
    RiskyUsers = @(
        'RunDate'
        'Id'
        'UserPrincipalName'
        'UserDisplayName'
        'RiskLevel'
        'RiskState'
        'RiskDetail'
        'RiskLastUpdatedDateTime'
        'IsDeleted'
        'IsProcessing'
    )

    # Graph permissions each source needs beyond the shared sign-in's defaults.
    # Directory.Read.All (already a default) also reads the tenant licence, which
    # keeps AuditLog.Read.All from failing intermittently:
    # https://learn.microsoft.com/troubleshoot/entra/entra-id/users-groups-entra-apis/b2c-or-tenant-premium-license-sign-in-activities
    ExtraScopes = @{
        AuthenticationMethods     = @()
        ConditionalAccessPolicies = @('Policy.Read.All')
        ActiveRoleAssignments     = @('RoleAssignmentSchedule.Read.Directory')
        EligibleRoleAssignments   = @('RoleEligibilitySchedule.Read.Directory')
        RoleAssignments           = @('RoleManagement.Read.Directory')
        UserSignInActivity        = @()
        # conditionalAccessStatus is on the sign-in. Policy.Read.All is only for
        # appliedConditionalAccessPolicies, which this collector does not read.
        # https://learn.microsoft.com/graph/api/resources/signin
        SignIns                   = @()
        RiskyUsers                = @('IdentityRiskyUser.Read.All')
    }

    # Availability per source and cloud.
    #   Available    - documented as available; collect it.
    #   NotAvailable - documented as unavailable; skip it, write a header-only CSV.
    #   Unverified   - no Microsoft page says either way; attempt it and log a warning.
    # No source in the contract is NotAvailable in any cloud today; the skip path stays
    # so a future documented gap needs one word changed here.
    SourceAvailability = @{
        AuthenticationMethods = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/authenticationmethodsroot-list-userregistrationdetails' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/authenticationmethodsroot-list-userregistrationdetails' }
        }
        ConditionalAccessPolicies = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/conditionalaccessroot-list-policies' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/conditionalaccessroot-list-policies' }
        }
        ActiveRoleAssignments = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/rbacapplication-list-roleassignmentscheduleinstances' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/rbacapplication-list-roleassignmentscheduleinstances' }
        }
        EligibleRoleAssignments = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/rbacapplication-list-roleeligibilityscheduleinstances' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/rbacapplication-list-roleeligibilityscheduleinstances' }
        }
        RoleAssignments = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/rbacapplication-list-roleassignments' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/rbacapplication-list-roleassignments' }
        }
        # List users is available in US Government L4; no page states the signInActivity
        # property itself in that cloud.
        UserSignInActivity = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/user-list#example-11-get-users-including-their-last-sign-in-time' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/graph/api/user-list' }
        }
        SignIns = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/signin-list' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/signin-list' }
        }
        RiskyUsers = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/riskyuser-list' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/riskyuser-list' }
        }
    }
}
