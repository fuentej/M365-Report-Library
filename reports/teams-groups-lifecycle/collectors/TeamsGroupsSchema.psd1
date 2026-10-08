@{
    # The column order of every CSV this report produces. The collectors, the
    # sample-data generator and the tests all read the order from here, so a
    # sample file can never drift from its collector's output.
    # Contract: docs/candidates/teams-groups-lifecycle.md

    Groups = @(
        'RunDate'
        'Id'
        'DisplayName'
        'Mail'
        'GroupTypes'
        'SecurityEnabled'
        'MailEnabled'
        'Visibility'
        'CreatedDateTime'
        'RenewedDateTime'
        'ExpirationDateTime'
        'DeletedDateTime'
        'OnPremisesSyncEnabled'
        'ResourceProvisioningOptions'
        'IsTeam'
    )

    # OwnerListStatus: Listed (owners returned), None (the call succeeded and
    # returned no owner), Unknown (Graph cannot return owners for this group, or
    # the call failed). Unknown is not the same as None.
    GroupOwners = @(
        'RunDate'
        'GroupId'
        'OwnerListStatus'
        'OwnerId'
        'OwnerType'
        'OwnerDisplayName'
        'OwnerUserPrincipalName'
    )

    DeletedGroups = @(
        'RunDate'
        'Id'
        'DisplayName'
        'GroupTypes'
        'SecurityEnabled'
        'MailEnabled'
        'CreatedDateTime'
        'DeletedDateTime'
        'PurgeDateTime'
        'ResourceProvisioningOptions'
        'IsTeam'
    )

    GroupLifecyclePolicies = @(
        'RunDate'
        'Id'
        'GroupLifetimeInDays'
        'ManagedGroupTypes'
        'AlternateNotificationEmails'
    )

    # CoverageStatus: Covered, NotCovered, or Unknown (the per-group call failed).
    GroupLifecycleCoverage = @(
        'RunDate'
        'GroupId'
        'PolicyId'
        'CoverageStatus'
    )

    TeamActivity = @(
        'RunDate'
        'ReportRefreshDate'
        'ReportPeriod'
        'TeamId'
        'TeamName'
        'TeamType'
        'LastActivityDate'
        'ActiveUsers'
        'ActiveChannels'
        'Guests'
        'Reactions'
        'MeetingsOrganized'
        'PostMessages'
        'ReplyMessages'
        'ChannelMessages'
        'UrgentMessages'
        'Mentions'
        'ActiveSharedChannels'
        'ActiveExternalUsers'
    )

    GroupActivity = @(
        'RunDate'
        'ReportRefreshDate'
        'ReportPeriod'
        'GroupId'
        'GroupDisplayName'
        'IsDeleted'
        'OwnerPrincipalName'
        'LastActivityDate'
        'GroupType'
        'MemberCount'
        'ExternalMemberCount'
        'ExchangeReceivedEmailCount'
        'SharePointActiveFileCount'
        'YammerPostedMessageCount'
        'YammerReadMessageCount'
        'YammerLikedMessageCount'
        'ExchangeMailboxTotalItemCount'
        'ExchangeMailboxStorageUsedByte'
        'SharePointTotalFileCount'
        'SharePointSiteStorageUsedByte'
    )

    TeamArchiveStatus = @(
        'RunDate'
        'TeamId'
        'DisplayName'
        'IsArchived'
    )

    # Copy the Operation column exactly: AddGroup has no trailing period.
    # https://learn.microsoft.com/purview/audit-log-activities
    GroupCreationEvents = @(
        'CreationTime'
        'Id'
        'Operation'
        'UserId'
        'Workload'
        'ObjectId'
        'TargetDisplayName'
        'TargetGroupId'
    )

    CreationOperations = @(
        'AddGroup'
        'TeamCreated'
    )

    # Header names accepted for a report column. The Learn pages spell some
    # headers two ways, so the collector reads whichever is present.
    #   Team type / Team Type
    #   External Member Count / Guest Count
    # https://learn.microsoft.com/graph/api/reportroot-getteamsteamactivitydetail
    # https://learn.microsoft.com/graph/api/reportroot-getoffice365groupsactivitydetail
    ReportHeaderAliases = @{
        TeamType            = @('Team Type', 'Team type')
        ExternalMemberCount = @('External Member Count', 'Guest Count')
    }

    # Availability per source and cloud.
    #   Available    - documented as available; collect it.
    #   NotAvailable - documented as unavailable; skip it, write a header-only CSV.
    #   Unverified   - no Microsoft page says either way; attempt it and log a warning.
    SourceAvailability = @{
        Groups = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/group-list' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/group-list' }
        }
        GroupOwners = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/group-list-owners' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/group-list-owners' }
        }
        DeletedGroups = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/directory-deleteditems-list' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/directory-deleteditems-list' }
        }
        GroupLifecyclePolicies = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/grouplifecyclepolicy-list' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/grouplifecyclepolicy-list' }
        }
        TeamActivity = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getteamsteamactivitydetail' }
            GCC        = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getteamsteamactivitydetail' }
            GCCHigh    = @{ Status = 'NotAvailable'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getteamsteamactivitydetail' }
        }
        GroupActivity = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getoffice365groupsactivitydetail' }
            GCC        = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getoffice365groupsactivitydetail' }
            GCCHigh    = @{ Status = 'NotAvailable'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getoffice365groupsactivitydetail' }
        }
        TeamArchiveStatus = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/team-get' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/team-get' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/team-get' }
        }
        # Audit (Standard) is Available in all three clouds. Whether AddGroup and
        # TeamCreated records are written in GCC and GCC High is UNVERIFIED.
        GroupCreationEvents = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/purview/audit-solutions-overview#comparison-of-key-capabilities' }
            GCC        = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-deployments' }
            GCCHigh    = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-high-deployments' }
        }
    }
}
