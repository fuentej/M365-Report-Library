@{
    # The column order of every CSV this report produces. The collectors, the sample data and the
    # tests all read the order from here, so a sample file can never drift from its collector's output.
    # Contract: docs/candidates/teams-activity.md (source numbers below are that table's).
    # Source 4 (team activity) is collected by reports/teams-groups-lifecycle (Get-TeamActivity.ps1), which
    # already writes every column the contract lists. Source 7 (message content) is not built: it reads
    # message bodies. Source 9 (the admin center export) is a manual fallback, not a collector.

    # Source 1. getTeamsUserActivityUserDetail. One row per user per run and period. QueryDate is the day
    # asked for with -Date, or empty for a period report. Exactly one of the two is sent. Tenant Display
    # Name and Shared Channel Tenant Display Names are not kept.
    TeamsUserActivityUserDetail = @(
        'RunDate'
        'QueryDate'
        'ReportRefreshDate'
        'UserId'
        'UserPrincipalName'
        'LastActivityDate'
        'IsDeleted'
        'DeletedDate'
        'AssignedProducts'
        'TeamChatMessageCount'
        'PrivateChatMessageCount'
        'CallCount'
        'MeetingCount'
        'PostMessages'
        'ReplyMessages'
        'UrgentMessages'
        'MeetingsOrganizedCount'
        'MeetingsAttendedCount'
        'AdHocMeetingsOrganizedCount'
        'AdHocMeetingsAttendedCount'
        'ScheduledOneTimeMeetingsOrganizedCount'
        'ScheduledOneTimeMeetingsAttendedCount'
        'ScheduledRecurringMeetingsOrganizedCount'
        'ScheduledRecurringMeetingsAttendedCount'
        'AudioDuration'
        'VideoDuration'
        'ScreenShareDuration'
        'AudioDurationInSeconds'
        'VideoDurationInSeconds'
        'ScreenShareDurationInSeconds'
        'HasOtherAction'
        'IsLicensed'
        'ReportPeriod'
    )

    # Source 2. getTeamsUserActivityCounts. One row per Report Date per run and period (tenant totals for
    # Teams licensed users).
    TeamsUserActivityCounts = @(
        'RunDate'
        'ReportRefreshDate'
        'ReportDate'
        'TeamChatMessages'
        'PostMessages'
        'ReplyMessages'
        'PrivateChatMessages'
        'Calls'
        'Meetings'
        'AudioDuration'
        'VideoDuration'
        'ScreenShareDuration'
        'MeetingsOrganized'
        'MeetingsAttended'
        'ReportPeriod'
    )

    # Source 3. getTeamsDeviceUsageUserDetail. Yes/no per platform for the period, not counts. QueryDate
    # as in source 1.
    TeamsDeviceUsageUserDetail = @(
        'RunDate'
        'QueryDate'
        'ReportRefreshDate'
        'UserId'
        'UserPrincipalName'
        'LastActivityDate'
        'IsDeleted'
        'DeletedDate'
        'UsedWeb'
        'UsedWindowsPhone'
        'UsediOS'
        'UsedMac'
        'UsedAndroidPhone'
        'UsedWindows'
        'UsedChromeOS'
        'UsedLinux'
        'IsLicensed'
        'ReportPeriod'
    )

    # Source 5. GET /admin/reportSettings. DisplayConcealedNames is True when usage reports hold concealed
    # identifiers, so sources 1 and 3 cannot be joined on UserPrincipalName.
    ReportSettings = @(
        'RunDate'
        'DisplayConcealedNames'
    )

    # Source 6. GET /communications/callRecords, then each record with sessions expanded. Event source.
    # One row per session (one row with empty session columns when a record has none). Version is part of
    # the key: a later version of a record appends new rows, and a reader keeps the highest Version per
    # CallRecordId. Platform is the callRecords userAgent platform of the caller and callee endpoint.
    CallRecords = @(
        'CallRecordId'
        'Version'
        'Type'
        'Modalities'
        'StartDateTime'
        'EndDateTime'
        'LastModifiedDateTime'
        'SessionId'
        'SessionStartDateTime'
        'SessionEndDateTime'
        'CallerUserId'
        'CallerPlatform'
        'CalleeUserId'
        'CalleePlatform'
    )

    # Source 8. Search-UnifiedAuditLog, counted per UTC day, workload, user and operation. Event source;
    # resumes the day after the latest Date.
    TeamsAuditEvents = @(
        'Date'
        'Workload'
        'UserId'
        'Operation'
        'EventCount'
    )

    # Per-cloud availability of each source, from the contract table. Only NotAvailable skips;
    # Unverified is attempted with a warning.
    SourceAvailability = @{
        TeamsUserActivityUserDetail = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getteamsuseractivityuserdetail' }
            GCC        = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getteamsuseractivityuserdetail' }
            GCCHigh    = @{ Status = 'NotAvailable'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getteamsuseractivityuserdetail' }
        }
        TeamsUserActivityCounts = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getteamsuseractivitycounts' }
            GCC        = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getteamsuseractivitycounts' }
            GCCHigh    = @{ Status = 'NotAvailable'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getteamsuseractivitycounts' }
        }
        TeamsDeviceUsageUserDetail = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getteamsdeviceusageuserdetail' }
            GCC        = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getteamsdeviceusageuserdetail' }
            GCCHigh    = @{ Status = 'NotAvailable'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getteamsdeviceusageuserdetail' }
        }
        ReportSettings = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/adminreportsettings-get' }
            GCC        = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/graph/api/adminreportsettings-get' }
            GCCHigh    = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/graph/api/adminreportsettings-get' }
        }
        CallRecords = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/callrecords-cloudcommunications-list-callrecords' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/callrecords-cloudcommunications-list-callrecords' }
        }
        # Audit (Standard) is Available in all three clouds. MeetingDetail and MeetingParticipantDetail are
        # Available everywhere; CallParticipantDetail, MessageSent and ChatCreated are UNVERIFIED in GCC and GCC High.
        TeamsAuditEvents = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/purview/audit-log-activities#teams-activities' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-deployments' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-high-deployments' }
        }
    }

    # Operations searched for source 8, from https://learn.microsoft.com/purview/audit-log-activities#teams-activities.
    # MeetingParticipantDetail already includes recorded or transcribed calls, so it is not added to CallParticipantDetail.
    # Operations whose cloud availability is UNVERIFIED are listed in UnverifiedAuditOperations.
    TeamsAuditOperations = @(
        'MeetingDetail'
        'MeetingParticipantDetail'
        'CallParticipantDetail'
        'MessageSent'
        'ChatCreated'
    )
    UnverifiedAuditOperations = @(
        'CallParticipantDetail'
        'MessageSent'
        'ChatCreated'
    )
}
