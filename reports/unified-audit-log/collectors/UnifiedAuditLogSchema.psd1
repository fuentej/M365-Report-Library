@{
    # The column order of every CSV this report produces. The collectors, the sample data and the
    # tests all read the order from here, so a sample file can never drift from its collector's output.
    # Contract: docs/candidates/unified-audit-log.md (source numbers below are that table's).
    # users.csv is not part of this report; the shared Invoke-EntraUserCollector writes it where a report wants it.

    # Source 1. Search-UnifiedAuditLog. Event source: one row per audit record, appended from the latest
    # CreationTime already in the file. RecordType is the name the cmdlet returns (for example
    # SharePointFileOperation). AuditData is the record's JSON as returned, on one line.
    AuditSearchCmdlet = @(
        'CreationTime'
        'RecordId'
        'RecordType'
        'Operation'
        'Workload'
        'UserId'
        'ObjectId'
        'ClientIP'
        'ResultStatus'
        'OrganizationId'
        'AuditData'
    )

    # Source 2. Graph Audit Search API records. Event source. RecordType is the auditLogRecordType name
    # (for example sharePointFileOperation); Workload is the record's service property. QueryId is the
    # auditLogQuery that returned the row.
    AuditGraphRecords = @(
        'CreationTime'
        'RecordId'
        'RecordType'
        'Operation'
        'Workload'
        'UserId'
        'ObjectId'
        'ClientIP'
        'ResultStatus'
        'OrganizationId'
        'QueryId'
        'AuditData'
    )

    # Source 3. Office 365 Management Activity API. Event source: one row per event inside a content blob.
    # CreationTime is the event time; ContentCreated is when the blob became available, which is what the
    # API's startTime and endTime select on, so it is the column the next run resumes from. RecordType is
    # the number the schema page uses (for example 6 for SharePointFileOperation).
    AuditActivityFeed = @(
        'ContentCreated'
        'ContentType'
        'ContentId'
        'CreationTime'
        'RecordId'
        'RecordType'
        'Operation'
        'Workload'
        'UserId'
        'ObjectId'
        'ClientIP'
        'ResultStatus'
        'OrganizationId'
        'AuditData'
    )

    # Source 4. Get-AdminAuditLogConfig in Exchange Online PowerShell. State source: one row per run.
    AuditIngestion = @(
        'RunDate'
        'UnifiedAuditLogIngestionEnabled'
    )

    # Source 5. Get-UnifiedAuditLogRetentionPolicy in Security & Compliance PowerShell. State source: one
    # row per policy per run. The cmdlet does not return the default policy, so no rows is not "no
    # one-year retention". List values are joined with a semicolon.
    AuditRetentionPolicies = @(
        'RunDate'
        'Priority'
        'Name'
        'RecordTypes'
        'Operations'
        'UserIds'
        'RetentionDuration'
    )

    # Availability per cloud, from the contract's source table. Only NotAvailable skips a collector.
    # Unverified is attempted and logs a warning.
    SourceAvailability = @{
        AuditSearchCmdlet = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/purview/audit-solutions-overview' }
            GCC        = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-deployments' }
            GCCHigh    = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-high-deployments' }
        }
        AuditGraphRecords = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/security-auditcoreroot-list-auditlogqueries' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'NotAvailable'; Reference = 'https://learn.microsoft.com/graph/api/security-auditcoreroot-list-auditlogqueries' }
        }
        AuditActivityFeed = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/office/office-365-management-api/office-365-management-activity-api-reference#activity-api-operations' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/office/office-365-management-api/office-365-management-activity-api-reference#activity-api-operations' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/office/office-365-management-api/office-365-management-activity-api-reference#activity-api-operations' }
        }
        AuditIngestion = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/purview/audit-log-search-script' }
            GCC        = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/purview/audit-log-enable-disable' }
            GCCHigh    = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/purview/audit-log-enable-disable' }
        }
        AuditRetentionPolicies = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/powershell/module/exchangepowershell/get-unifiedauditlogretentionpolicy' }
            GCC        = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-deployments' }
            GCCHigh    = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-high-deployments' }
        }
    }

    # Management Activity API root per cloud and the OAuth resource the token is issued for.
    # https://learn.microsoft.com/office/office-365-management-api/office-365-management-activity-api-reference#activity-api-operations
    ActivityFeedRoot = @{
        Commercial = 'https://manage.office.com'
        GCC        = 'https://manage-gcc.office.com'
        GCCHigh    = 'https://manage.office365.us'
    }
}
