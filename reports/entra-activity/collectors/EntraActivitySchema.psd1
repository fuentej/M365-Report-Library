@{
    # The column order of every CSV this report produces. The collectors, the sample
    # files and the tests all read the order from here, so a sample file can never
    # drift from its collector's output.
    # Contract: docs/candidates/entra-activity.md

    # Sources 1 and 2. Properties of signIn (v1.0 and beta):
    # https://learn.microsoft.com/graph/api/resources/signin#properties
    # errorCode is an Int32 and is kept as returned (0 is success; 1024 and other
    # codes that are not 5 or 6 digits are kept). userPrincipalName is always
    # lowercase and a guest is the home UPN, so group by UserId.
    SignIns = @(
        'CreatedDateTime'
        'Id'
        'UserId'
        'UserPrincipalName'
        'AppId'
        'AppDisplayName'
        'ResourceDisplayName'
        'IpAddress'
        'City'
        'State'
        'CountryOrRegion'
        'ClientAppUsed'
        'DeviceOperatingSystem'
        'DeviceBrowser'
        'DeviceIsCompliant'
        'DeviceIsManaged'
        'IsInteractive'
        'SignInEventTypes'
        'ErrorCode'
        'FailureReason'
        'AdditionalDetails'
        'ConditionalAccessStatus'
        'RiskDetail'
        'RiskLevelAggregated'
        'RiskLevelDuringSignIn'
        'RiskState'
    )

    # Source 3. One row per applied Conditional Access policy per sign-in, read from the
    # sign-in responses of sources 1 and 2. A sign-in with no policy detail gets one row
    # with the policy columns empty, so conditionalAccessStatus is never lost.
    # https://learn.microsoft.com/graph/api/resources/appliedconditionalaccesspolicy
    # PolicyDetailReadable says whether the run held a Conditional Access read
    # permission; without it appliedConditionalAccessPolicies is omitted without error.
    SignInConditionalAccess = @(
        'CreatedDateTime'
        'SignInId'
        'UserId'
        'IsInteractive'
        'ConditionalAccessStatus'
        'PolicyId'
        'PolicyDisplayName'
        'PolicyResult'
        'EnforcedGrantControls'
        'EnforcedSessionControls'
        'PolicyDetailReadable'
    )

    # Source 4. Properties of directoryAudit:
    # https://learn.microsoft.com/graph/api/resources/directoryaudit#properties
    # Every category, activityDisplayName, operationType and targetResources.type is
    # kept as returned. ModifiedProperties is the targetResources' modifiedProperties as
    # compact JSON.
    DirectoryAudits = @(
        'ActivityDateTime'
        'Id'
        'ActivityDisplayName'
        'Category'
        'OperationType'
        'Result'
        'ResultReason'
        'LoggedByService'
        'CorrelationId'
        'InitiatedByUserId'
        'InitiatedByUserPrincipalName'
        'InitiatedByAppId'
        'InitiatedByAppDisplayName'
        'TargetResourceTypes'
        'TargetResourceIds'
        'TargetResourceDisplayNames'
        'TargetResourceUserPrincipalNames'
        'ModifiedProperties'
    )

    # Source 5. Reference, not a tenant call. Retention by licence level:
    # https://learn.microsoft.com/entra/identity/monitoring-health/reference-reports-data-retention#how-long-does-microsoft-entra-id-store-the-data
    RetentionReference = @(
        'RunDate'
        'Environment'
        'LicenseLevel'
        'SignInRetentionDays'
        'AuditRetentionDays'
        'RiskySignInRetentionDays'
        'RetentionStatus'
        'Reference'
    )

    # The documented retention windows. The page names no cloud, so the contract marks
    # retention UNVERIFIED in all three clouds and the collector writes that status.
    RetentionLevels = @(
        @{ LicenseLevel = 'Free'; SignInRetentionDays = 7; AuditRetentionDays = 7; RiskySignInRetentionDays = 7 }
        @{ LicenseLevel = 'P1'; SignInRetentionDays = 30; AuditRetentionDays = 30; RiskySignInRetentionDays = 30 }
        @{ LicenseLevel = 'P2'; SignInRetentionDays = 30; AuditRetentionDays = 30; RiskySignInRetentionDays = 90 }
    )

    # signInEventTypes values for the beta filter. The list examples spell the
    # non-interactive value 'nonInteractiveUser'.
    # https://learn.microsoft.com/graph/api/signin-list?view=graph-rest-beta
    NonInteractiveEventType = 'nonInteractiveUser'

    # Graph permissions a source adds to the shared sign-in's scopes. Source 3 needs a
    # Conditional Access read permission to receive appliedConditionalAccessPolicies;
    # Policy.ReadWrite.ConditionalAccess is a write permission and is not requested.
    # https://learn.microsoft.com/graph/api/signin-list#permissions
    ExtraScopes = @{
        InteractiveSignIns      = @()
        NonInteractiveSignIns   = @()
        SignInConditionalAccess = @('Policy.Read.All')
        DirectoryAudits         = @()
    }

    # Any one of these in the session means appliedConditionalAccessPolicies is returned.
    ConditionalAccessReadPermissions = @(
        'Policy.Read.All'
        'Policy.Read.ConditionalAccess'
        'Policy.ReadWrite.ConditionalAccess'
    )

    # Availability per cloud, from the contract's source table. Only NotAvailable skips;
    # Unverified is attempted with a warning. The log retention window is UNVERIFIED in
    # GCC and GCC High, but that is not an availability question for the API.
    SourceAvailability = @{
        InteractiveSignIns = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/signin-list' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/signin-list' }
        }
        NonInteractiveSignIns = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/signin-list?view=graph-rest-beta' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/signin-list?view=graph-rest-beta' }
        }
        SignInConditionalAccess = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/signin-list' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/signin-list' }
        }
        DirectoryAudits = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/directoryaudit-list' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/directoryaudit-list' }
        }
        RetentionReference = @{
            Commercial = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/entra/identity/monitoring-health/reference-reports-data-retention' }
            GCC        = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/entra/identity/monitoring-health/reference-reports-data-retention' }
            GCCHigh    = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/entra/identity/monitoring-health/reference-reports-data-retention' }
        }
    }
}
