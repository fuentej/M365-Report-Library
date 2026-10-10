@{
    # The column order of every CSV this report produces. The collectors, the sample data and the
    # tests all read the order from here, so a sample file can never drift from its collector's output.
    # Contract: docs/candidates/exchange-activity.md (source numbers below are that table's).
    # Source 10 (the admin center export) is a manual fallback, not a collector.
    # users.csv is not part of this report; the shared Invoke-EntraUserCollector writes it where a report wants it.

    # Source 1. getMailboxUsageDetail. One row per mailbox per run and period. QuotaStatus is
    # derived with the admin center's four categories: at or above a quota is the next category.
    # Good, Warning, CantSend, CantSendReceive; empty when a quota or the storage is missing.
    MailboxUsageDetail = @(
        'RunDate'
        'ReportRefreshDate'
        'UserPrincipalName'
        'DisplayName'
        'IsDeleted'
        'DeletedDate'
        'CreatedDate'
        'LastActivityDate'
        'ItemCount'
        'StorageUsedByte'
        'IssueWarningQuotaByte'
        'ProhibitSendQuotaByte'
        'ProhibitSendReceiveQuotaByte'
        'DeletedItemCount'
        'DeletedItemSizeByte'
        'DeletedItemQuotaByte'
        'HasArchive'
        'ReportPeriod'
        'QuotaStatus'
    )

    # Source 2. getMailboxUsageStorage. One row per report date per run and period.
    MailboxUsageStorage = @(
        'RunDate'
        'ReportRefreshDate'
        'StorageUsedByte'
        'ReportDate'
        'ReportPeriod'
    )

    # Source 3. getEmailActivityUserDetail. QueryDate is the day asked for with -Date, or empty
    # for a period report. Exactly one of the two is sent.
    EmailActivityUserDetail = @(
        'RunDate'
        'QueryDate'
        'ReportRefreshDate'
        'UserPrincipalName'
        'DisplayName'
        'IsDeleted'
        'DeletedDate'
        'LastActivityDate'
        'SendCount'
        'ReceiveCount'
        'ReadCount'
        'MeetingCreatedCount'
        'MeetingInteractedCount'
        'AssignedProducts'
        'ReportPeriod'
    )

    # Source 4. getEmailAppUsageUserDetail. QueryDate as in source 3.
    EmailAppUsageUserDetail = @(
        'RunDate'
        'QueryDate'
        'ReportRefreshDate'
        'UserPrincipalName'
        'DisplayName'
        'IsDeleted'
        'DeletedDate'
        'LastActivityDate'
        'MailForMac'
        'OutlookForMac'
        'OutlookForWindows'
        'OutlookForMobile'
        'OtherForMobile'
        'OutlookForWeb'
        'POP3App'
        'IMAP4App'
        'SMTPApp'
        'ReportPeriod'
    )

    # Source 5. GET /admin/reportSettings. DisplayConcealedNames is True when usage reports
    # hold concealed identifiers, so sources 1, 3 and 4 cannot be joined on UserPrincipalName.
    ReportSettings = @(
        'RunDate'
        'DisplayConcealedNames'
    )

    # Source 6. Get-EXOMailbox -PropertySets Minimum, Quota.
    Mailboxes = @(
        'RunDate'
        'ExternalDirectoryObjectId'
        'UserPrincipalName'
        'PrimarySmtpAddress'
        'RecipientType'
        'RecipientTypeDetails'
        'IssueWarningQuota'
        'ProhibitSendQuota'
        'ProhibitSendReceiveQuota'
        'RecoverableItemsQuota'
        'ArchiveQuota'
        'UseDatabaseQuotaDefaults'
    )

    # Source 7. Get-EXOMailboxStatistics -PropertySets All, one call per mailbox. The size columns
    # keep the cmdlet's text; the Bytes columns hold the byte count read from it.
    MailboxStatistics = @(
        'RunDate'
        'ExternalDirectoryObjectId'
        'UserPrincipalName'
        'ItemCount'
        'TotalItemSize'
        'TotalItemSizeBytes'
        'DeletedItemCount'
        'TotalDeletedItemSize'
        'TotalDeletedItemSizeBytes'
        'StorageLimitStatus'
        'LastLogonTime'
        'LastLogoffTime'
        'LastLoggedOnUserAccount'
        'IsArchiveMailbox'
        'DatabaseIssueWarningQuota'
        'DatabaseProhibitSendQuota'
        'DatabaseProhibitSendReceiveQuota'
    )

    # Source 8. Get-MessageTraceV2. Event source. One row per recipient, not per distinct message.
    # Sent is SenderAddress = a mailbox; received is RecipientAddress = a mailbox. There is no read count.
    MessageTrace = @(
        'Received'
        'MessageTraceId'
        'SenderAddress'
        'RecipientAddress'
        'Status'
    )

    # Source 9. Get-EXOMobileDeviceStatistics -Mailbox. The columns are the properties Learn shows
    # in its troubleshooting output (DeviceType, DeviceOS, DeviceAccessState, DeviceUserAgent) and DeviceId.
    MobileDevices = @(
        'RunDate'
        'MailboxUserPrincipalName'
        'DeviceId'
        'DeviceType'
        'DeviceOS'
        'DeviceAccessState'
        'DeviceUserAgent'
    )

    # Source 11. GET /beta/admin/exchange/tracing/messageTraces. Event source. One row per recipient.
    GraphMessageTrace = @(
        'ReceivedDateTime'
        'Id'
        'SenderAddress'
        'RecipientAddress'
        'Status'
        'Size'
    )

    # Per-cloud availability of each source.
    # Availability comes from the contract table. Only NotAvailable skips; Unverified is attempted with a warning.
    SourceAvailability = @{
        MailboxUsageDetail = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getmailboxusagedetail' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'NotAvailable'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getmailboxusagedetail' }
        }
        MailboxUsageStorage = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getmailboxusagestorage' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'NotAvailable'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getmailboxusagestorage' }
        }
        EmailActivityUserDetail = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getemailactivityuserdetail' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'NotAvailable'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getemailactivityuserdetail' }
        }
        EmailAppUsageUserDetail = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getemailappusageuserdetail' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'NotAvailable'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getemailappusageuserdetail' }
        }
        ReportSettings = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/adminreportsettings-get' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/graph/api/adminreportsettings-get' }
        }
        Mailboxes = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/powershell/module/exchangepowershell/get-exomailbox' }
            GCC        = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/office-365-platform-service-description/office-365-us-government/exchange-online-for-us-government-environments' }
            GCCHigh    = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/office-365-platform-service-description/office-365-us-government/exchange-online-for-us-government-environments' }
        }
        MailboxStatistics = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/powershell/module/exchangepowershell/get-exomailboxstatistics' }
            GCC        = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/office-365-platform-service-description/office-365-us-government/exchange-online-for-us-government-environments' }
            GCCHigh    = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/office-365-platform-service-description/office-365-us-government/exchange-online-for-us-government-environments' }
        }
        MessageTrace = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/powershell/module/exchangepowershell/get-messagetracev2' }
            GCC        = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/office-365-platform-service-description/office-365-us-government/exchange-online-for-us-government-environments' }
            GCCHigh    = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/office-365-platform-service-description/office-365-us-government/exchange-online-for-us-government-environments' }
        }
        MobileDevices = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/powershell/module/exchangepowershell/get-exomobiledevicestatistics' }
            GCC        = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/office-365-platform-service-description/office-365-us-government/exchange-online-for-us-government-environments' }
            GCCHigh    = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/office-365-platform-service-description/office-365-us-government/exchange-online-for-us-government-environments' }
        }
        GraphMessageTrace = @{
            Commercial = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/exchange/monitoring/trace-an-email-message/graph-api-message-trace' }
            GCC        = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/exchange/monitoring/trace-an-email-message/graph-api-message-trace' }
            GCCHigh    = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/exchange/monitoring/trace-an-email-message/graph-api-message-trace' }
        }
    }
}
