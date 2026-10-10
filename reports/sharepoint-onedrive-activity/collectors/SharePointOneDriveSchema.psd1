@{
    # The column order of every CSV this report produces. The collectors, the sample data and the
    # tests all read the order from here, so a sample file can never drift from its collector's output.
    # Contract: docs/candidates/sharepoint-onedrive-activity.md (source numbers below are that table's).
    # Source 8 (the admin center export) is a manual fallback, not a collector. Source 9 route (a), the
    # SharePoint Storage page in the admin center, is manual too; route (b), Get-SPOTenant, is a collector.
    # users.csv is not part of this report; the shared Invoke-EntraUserCollector writes it where a report wants it.

    # Source 1. getSharePointSiteUsageDetail. One row per site per run and period. Storage is bytes in
    # the Graph CSV; the admin center shows MB. Do not add the two.
    SharePointSiteUsageDetail = @(
        'RunDate'
        'ReportRefreshDate'
        'SiteId'
        'SiteUrl'
        'OwnerDisplayName'
        'IsDeleted'
        'LastActivityDate'
        'FileCount'
        'ActiveFileCount'
        'PageViewCount'
        'VisitedPageCount'
        'StorageUsedByte'
        'StorageAllocatedByte'
        'RootWebTemplate'
        'OwnerPrincipalName'
        'ReportPeriod'
    )

    # Source 2. getOneDriveUsageAccountDetail. One row per account per run and period. The Graph CSV
    # has no Site Id column.
    OneDriveUsageAccountDetail = @(
        'RunDate'
        'ReportRefreshDate'
        'SiteUrl'
        'OwnerDisplayName'
        'IsDeleted'
        'LastActivityDate'
        'FileCount'
        'ActiveFileCount'
        'StorageUsedByte'
        'StorageAllocatedByte'
        'OwnerPrincipalName'
        'ReportPeriod'
    )

    # Source 3. getSharePointSiteUsageStorage. One row per report date per run and period.
    SharePointSiteUsageStorage = @(
        'RunDate'
        'ReportRefreshDate'
        'SiteType'
        'StorageUsedByte'
        'ReportDate'
        'ReportPeriod'
    )

    # Source 4. getOneDriveUsageStorage. Same columns as source 3.
    OneDriveUsageStorage = @(
        'RunDate'
        'ReportRefreshDate'
        'SiteType'
        'StorageUsedByte'
        'ReportDate'
        'ReportPeriod'
    )

    # Source 5. getSharePointActivityUserDetail. QueryDate is the day asked for with -Date, or empty
    # for a period report. Exactly one of the two is sent.
    SharePointActivityUserDetail = @(
        'RunDate'
        'QueryDate'
        'ReportRefreshDate'
        'UserPrincipalName'
        'IsDeleted'
        'DeletedDate'
        'LastActivityDate'
        'ViewedOrEditedFileCount'
        'SyncedFileCount'
        'SharedInternallyFileCount'
        'SharedExternallyFileCount'
        'VisitedPageCount'
        'AssignedProducts'
        'ReportPeriod'
    )

    # Source 6. getOneDriveActivityUserDetail. As source 5 without VisitedPageCount.
    OneDriveActivityUserDetail = @(
        'RunDate'
        'QueryDate'
        'ReportRefreshDate'
        'UserPrincipalName'
        'IsDeleted'
        'DeletedDate'
        'LastActivityDate'
        'ViewedOrEditedFileCount'
        'SyncedFileCount'
        'SharedInternallyFileCount'
        'SharedExternallyFileCount'
        'AssignedProducts'
        'ReportPeriod'
    )

    # Source 7. GET /admin/reportSettings. DisplayConcealedNames is True when usage reports hold
    # concealed identifiers, so sources 1, 2, 5 and 6 cannot be joined to the site list or users.
    ReportSettings = @(
        'RunDate'
        'DisplayConcealedNames'
    )

    # Source 9, route (b). Get-SPOTenant. The properties the cmdlet page names; one that the
    # cmdlet does not return is left empty.
    TenantStorage = @(
        'RunDate'
        'StorageQuota'
        'StorageQuotaAllocated'
        'ResourceQuota'
        'ResourceQuotaAllocated'
        'OneDriveStorageQuota'
    )

    # Source 10. Get-SPOSite -Limit ALL -IncludePersonalSite $true, without -Detailed. The storage
    # columns are empty when the cmdlet does not return them without -Detailed.
    SpoSites = @(
        'RunDate'
        'Url'
        'Title'
        'Template'
        'StorageUsageCurrent'
        'ResourceUsageCurrent'
        'WebsCount'
    )

    # Source 11. GET /sites/getAllSites, then GET /sites/{id}/drives. One row per drive per run.
    # The quota object has no file count. LastModifiedDateTime is when the drive was modified, not
    # the usage report's Last Activity Date.
    DriveQuota = @(
        'RunDate'
        'SiteId'
        'SiteWebUrl'
        'IsPersonalSite'
        'DriveId'
        'DriveName'
        'DriveType'
        'DriveWebUrl'
        'QuotaState'
        'QuotaUsed'
        'QuotaTotal'
        'QuotaRemaining'
        'QuotaDeleted'
        'LastModifiedDateTime'
    )

    # Source 12. Search-UnifiedAuditLog. Event source. Counted by the library, one row per UTC day,
    # workload, user and operation; the counts will not match the usage reports. Only complete days
    # are written, so Date is the resume point.
    FileEvents = @(
        'Date'
        'Workload'
        'UserId'
        'Operation'
        'EventCount'
    )

    # Source 13. getActivitiesByInterval on a site, interval day. Not file counts and not per-user
    # counts. An action the interval does not carry is empty, not zero. IncompleteData is True when
    # the interval carries the incompleteData facet.
    SiteActivity = @(
        'RunDate'
        'SiteId'
        'SiteWebUrl'
        'IntervalStart'
        'IntervalEnd'
        'AccessActionCount'
        'AccessActorCount'
        'CreateActionCount'
        'CreateActorCount'
        'EditActionCount'
        'EditActorCount'
        'DeleteActionCount'
        'DeleteActorCount'
        'MoveActionCount'
        'MoveActorCount'
        'IncompleteData'
        'MissingDataBeforeDateTime'
        'WasThrottled'
    )

    # SharePoint and OneDrive operations counted by source 12, as the contract lists them on
    # https://learn.microsoft.com/purview/audit-log-activities
    FileOperations = @(
        'FileAccessed'
        'FileModified'
        'FileDownloaded'
        'FileUploaded'
        'FileSyncDownloadedFull'
        'FileSyncUploadedFull'
        'PageViewed'
        'SharingSet'
        'AnonymousLinkCreated'
        'SecureLinkCreated'
    )

    # Per-cloud availability of each source.
    # Availability comes from the contract table. Only NotAvailable skips; Unverified is attempted with a warning.
    SourceAvailability = @{
        SharePointSiteUsageDetail = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getsharepointsiteusagedetail' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'NotAvailable'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getsharepointsiteusagedetail' }
        }
        OneDriveUsageAccountDetail = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getonedriveusageaccountdetail' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'NotAvailable'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getonedriveusageaccountdetail' }
        }
        SharePointSiteUsageStorage = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getsharepointsiteusagestorage' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'NotAvailable'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getsharepointsiteusagestorage' }
        }
        OneDriveUsageStorage = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getonedriveusagestorage' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'NotAvailable'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getonedriveusagestorage' }
        }
        SharePointActivityUserDetail = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getsharepointactivityuserdetail' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'NotAvailable'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getsharepointactivityuserdetail' }
        }
        OneDriveActivityUserDetail = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getonedriveactivityuserdetail' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'NotAvailable'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getonedriveactivityuserdetail' }
        }
        ReportSettings = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/adminreportsettings-get' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/graph/api/adminreportsettings-get' }
        }
        TenantStorage = @{
            Commercial = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/powershell/module/microsoft.online.sharepoint.powershell/get-spotenant' }
            GCC        = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/powershell/module/microsoft.online.sharepoint.powershell/get-spotenant' }
            GCCHigh    = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/powershell/module/microsoft.online.sharepoint.powershell/get-spotenant' }
        }
        SpoSites = @{
            Commercial = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/powershell/module/microsoft.online.sharepoint.powershell/get-sposite' }
            GCC        = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/powershell/module/microsoft.online.sharepoint.powershell/get-sposite' }
            GCCHigh    = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/powershell/module/microsoft.online.sharepoint.powershell/connect-sposervice' }
        }
        DriveQuota = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/site-getallsites' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/site-getallsites' }
        }
        FileEvents = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/purview/audit-solutions-overview' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-deployments' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-high-deployments' }
        }
        SiteActivity = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/itemactivitystat-getactivitybyinterval' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/graph/api/itemactivitystat-getactivitybyinterval' }
        }
    }
}
