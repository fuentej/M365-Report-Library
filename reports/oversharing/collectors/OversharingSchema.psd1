@{
    # The column order of every CSV this report produces. The collectors, the
    # sample-data generator and the tests all read the order from here, so a
    # sample file can never drift from its collector's output.
    # Contract: docs/candidates/oversharing.md

    # Source 1. https://learn.microsoft.com/graph/api/site-getallsites
    # One row per site, OneDrive sites included (IsPersonalSite).
    Sites = @(
        'RunDate'
        'SiteId'
        'Name'
        'WebUrl'
        'IsPersonalSite'
        'HostName'
        'DataLocationCode'
    )

    # Source 2. https://learn.microsoft.com/graph/api/driveitem-list-permissions
    # One row per permission on an item. LinkScope is anonymous, organization, users or
    # existingAccess. IsInherited is True when inheritedFrom is set. The sharing link's own
    # webUrl is a secret the page says is returned only to callers who can create the
    # permission; it is never written.
    ItemPermissions = @(
        'RunDate'
        'SiteId'
        'DriveId'
        'ItemId'
        'ItemName'
        'ItemWebUrl'
        'PermissionId'
        'Roles'
        'LinkScope'
        'LinkType'
        'LinkPreventsDownload'
        'HasPassword'
        'ExpirationDateTime'
        'IsInherited'
        'InheritedFromItemId'
        'GrantedTo'
    )

    # Source 3a. Columns of the downloaded site permissions report, renamed:
    # https://learn.microsoft.com/sharepoint/data-access-governance-site-permissions-report#download-the-site-permissions-for-your-organization-reports
    # ReportId is the report the row came from, so a report that is reused on a later run is
    # appended once.
    SitePermissionBreadth = @(
        'RunDate'
        'Workload'
        'ReportId'
        'ReportDate'
        'SiteId'
        'SiteName'
        'SiteUrl'
        'SiteTemplate'
        'PrimaryAdmin'
        'PrimaryAdminEmail'
        'ExternalSharing'
        'SitePrivacy'
        'SiteSensitivity'
        'UsersWithAccess'
        'GuestUserPermissions'
        'ExternalParticipantPermissions'
        'EntraGroupPermissions'
        'FileCount'
        'ItemsWithUniquePermissions'
        'PeopleInYourOrgLinks'
        'AnyoneLinks'
        'EeeuPermissions'
        'EveryonePermissions'
    )

    # Source 3b. Columns of the downloaded special-groups report, renamed:
    # https://learn.microsoft.com/sharepoint/data-access-governance-detailed-eeeu-everyone-permissions-report#download-the-report
    # ReportEntity is EveryoneExceptExternalUsers or Everyone. The report's TenantId and its
    # UserPrincipalName column (the same text as Recipient) are not kept.
    EveryoneItemExposure = @(
        'RunDate'
        'ReportEntity'
        'ReportId'
        'ReportDate'
        'SiteId'
        'WebId'
        'ListId'
        'ScopeId'
        'UniqueId'
        'ListItemId'
        'ItemType'
        'ItemUrl'
        'RoleDefinition'
        'LinkId'
        'LinkScope'
        'Recipient'
        'ParentObjectId'
        'ParentGroupName'
        'ParentGroupEmail'
        'ParentGroupType'
        'TotalUserCount'
    )

    # Source 3c. Learn does not list the columns of the sharing links activity CSV, so the
    # exported row is kept whole as JSON in ReportRow; SiteId and SiteUrl are copied out when
    # the export has a column of that name. ReportEntity is SharingLinks_Anyone,
    # SharingLinks_PeopleInYourOrg or SharingLinks_Guests. ReportStartTime and ReportEndTime
    # are the period the report covers, from Get-SPODataAccessGovernanceInsight.
    SharingLinkActivity = @(
        'RunDate'
        'ReportEntity'
        'Workload'
        'ReportId'
        'ReportStartTime'
        'ReportEndTime'
        'SiteId'
        'SiteUrl'
        'ReportRow'
    )

    # Source 3d. As 3c. ReportEntity is EveryoneExceptExternalUsersAtSite or
    # EveryoneExceptExternalUsersForItems.
    EeeuActivity = @(
        'RunDate'
        'ReportEntity'
        'Workload'
        'ReportId'
        'ReportStartTime'
        'ReportEndTime'
        'SiteId'
        'SiteUrl'
        'ReportRow'
    )

    # Source 3e. As 3c, one label per report. A snapshot, so there is no period.
    LabeledFileSites = @(
        'RunDate'
        'LabelGuid'
        'LabelName'
        'Workload'
        'ReportId'
        'ReportCreatedDateTime'
        'SiteId'
        'SiteUrl'
        'ReportRow'
    )

    # Source 4. https://learn.microsoft.com/powershell/module/microsoft.online.sharepoint.powershell/get-sposite
    # Scope is Tenant (from Get-SPOTenant, UNVERIFIED) or Site. SharingCapability is
    # ExternalUserAndGuestSharing, Disabled, ExternalUserSharingOnly or
    # ExistingExternalUserSharingOnly. DefaultSharingLinkType is None (the widest scope the
    # other settings allow, not "no link"), Direct, Internal or AnonymousAccess.
    SiteSharingSettings = @(
        'RunDate'
        'Scope'
        'Url'
        'Title'
        'Template'
        'SharingCapability'
        'DefaultSharingLinkType'
        'DisableCompanyWideSharingLinks'
        'SensitivityLabel'
    )

    # Sources 5a and 5b share these columns. The SharePoint and sharing fields come from the
    # audit record's AuditData:
    # https://learn.microsoft.com/office/office-365-management-api/office-365-management-activity-api-schema#sharepoint-sharing-schema
    AuditEvents = @(
        'CreationTime'
        'Id'
        'RecordType'
        'Operation'
        'UserId'
        'Workload'
        'ObjectId'
        'ItemType'
        'SiteUrl'
        'SourceRelativeUrl'
        'SourceFileName'
        'TargetUserOrGroupName'
        'TargetUserOrGroupType'
        'ClientIP'
    )

    # Source 5c. Read through Exchange Online: the same property is always False in
    # Security & Compliance PowerShell.
    # https://learn.microsoft.com/purview/audit-log-enable-disable
    AuditLogStatus = @(
        'RunDate'
        'UnifiedAuditLogIngestionEnabled'
    )

    # Source 5a, copied exactly from the contract doc.
    AnonymousLinkOperations = @(
        'AnonymousLinkCreated'
        'AnonymousLinkUpdated'
        'AnonymousLinkUsed'
        'AnonymousLinkRemoved'
    )

    # Source 5b, copied exactly from the contract doc.
    SharingOperations = @(
        'CompanyLinkCreated'
        'CompanyLinkUsed'
        'CompanyLinkRemoved'
        'SecureLinkCreated'
        'SecureLinkUsed'
        'SecureLinkDeleted'
        'AddedToSecureLink'
        'RemovedFromSecureLink'
        'SharingSet'
        'SharingRevoked'
        'SharingInvitationCreated'
        'SharingInvitationAccepted'
        'SharingInvitationBlocked'
        'SharingInvitationUpdated'
        'SharingInvitationRevoked'
        'AddedToGroup'
        'SharingInheritanceBroken'
        'PermissionLevelsInheritanceBroken'
        'SharingInheritanceReset'
    )

    # Available / NotAvailable / Unverified, from the source table in the contract doc.
    # Only NotAvailable skips a source. No source in the doc is NotAvailable in any cloud.
    # The 3x rows rest on a feature row of the SharePoint Advanced Management table, not on
    # a cmdlet; whether Connect-SPOService and the Data access governance cmdlets connect in
    # GCC and GCC High is an open item in the doc, so those clouds log a warning on connect.
    SourceAvailability = @{
        Sites = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/site-getallsites' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments#microsoft-graph-and-graph-explorer-service-root-endpoints' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/site-getallsites' }
        }
        ItemPermissions = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/driveitem-list-permissions' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments#microsoft-graph-and-graph-explorer-service-root-endpoints' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/driveitem-list-permissions' }
        }
        SitePermissionBreadth = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/sharepoint/sharepoint-advanced-management-features-copilot-license#sam-features-in-copilot-environments' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/sharepoint/sharepoint-advanced-management-features-copilot-license#sam-features-in-copilot-environments' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/sharepoint/sharepoint-advanced-management-features-copilot-license#sam-features-in-copilot-environments' }
        }
        EveryoneItemExposure = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/sharepoint/sharepoint-advanced-management-features-copilot-license#sam-features-in-copilot-environments' }
            GCC        = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/sharepoint/sharepoint-advanced-management-features-copilot-license#sam-features-in-copilot-environments' }
            GCCHigh    = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/sharepoint/sharepoint-advanced-management-features-copilot-license#sam-features-in-copilot-environments' }
        }
        SharingLinkActivity = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/sharepoint/sharepoint-advanced-management-features-copilot-license#sam-features-in-copilot-environments' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/sharepoint/sharepoint-advanced-management-features-copilot-license#sam-features-in-copilot-environments' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/sharepoint/sharepoint-advanced-management-features-copilot-license#sam-features-in-copilot-environments' }
        }
        EeeuActivity = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/sharepoint/sharepoint-advanced-management-features-copilot-license#sam-features-in-copilot-environments' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/sharepoint/sharepoint-advanced-management-features-copilot-license#sam-features-in-copilot-environments' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/sharepoint/sharepoint-advanced-management-features-copilot-license#sam-features-in-copilot-environments' }
        }
        LabeledFileSites = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/sharepoint/sharepoint-advanced-management-features-copilot-license#sam-features-in-copilot-environments' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/sharepoint/sharepoint-advanced-management-features-copilot-license#sam-features-in-copilot-environments' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/sharepoint/sharepoint-advanced-management-features-copilot-license#sam-features-in-copilot-environments' }
        }
        SiteSharingSettings = @{
            Commercial = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/powershell/module/microsoft.online.sharepoint.powershell/get-sposite' }
            GCC        = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/powershell/module/microsoft.online.sharepoint.powershell/get-sposite' }
            GCCHigh    = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/powershell/module/microsoft.online.sharepoint.powershell/get-sposite' }
        }
        AnonymousLinkEvents = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/purview/audit-solutions-overview#audit-standard' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-deployments' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-high-deployments' }
        }
        SharingEvents = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/purview/audit-solutions-overview#audit-standard' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-deployments' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-high-deployments' }
        }
        AuditLogStatus = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/purview/audit-log-enable-disable' }
            GCC        = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/purview/audit-log-enable-disable' }
            GCCHigh    = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/purview/audit-log-enable-disable' }
        }
    }
}
