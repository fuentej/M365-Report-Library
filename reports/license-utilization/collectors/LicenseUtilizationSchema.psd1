@{
    # The column order of every CSV this report produces. The collectors, the
    # sample files and the tests all read the order from here, so a sample file
    # can never drift from its collector's output.
    # Contract: docs/candidates/license-utilization.md
    #
    # Every source here is state: a snapshot stamped with RunDate and appended each
    # run. No source is an event stream, so there is no watermark in this report.

    # Source 1. subscribedSku: https://learn.microsoft.com/graph/api/resources/subscribedsku
    # FreeUnits is derived (prepaidUnits.enabled minus consumedUnits); it is not a
    # field. Suspended, warning and locked-out units are separate and are not folded
    # into it.
    SubscribedSkus = @(
        'RunDate'
        'SkuId'
        'SkuPartNumber'
        'AppliesTo'
        'CapabilityStatus'
        'ConsumedUnits'
        'PrepaidEnabled'
        'PrepaidSuspended'
        'PrepaidWarning'
        'PrepaidLockedOut'
        'FreeUnits'
    )

    # Source 1, service plans of each SKU (servicePlanInfo). Written by the same
    # collector, from the same response.
    # https://learn.microsoft.com/graph/api/resources/serviceplaninfo
    SkuServicePlans = @(
        'RunDate'
        'SkuId'
        'ServicePlanId'
        'ServicePlanName'
        'ProvisioningStatus'
        'AppliesTo'
    )

    # Source 2. One row per licence assignment state, so a user holding a licence
    # directly and through a group has two rows. AssignedByGroup is empty for a direct
    # assignment. State is Active, ActiveWithError, Disabled or Error. Department,
    # job title, city and country are not repeated here: join to users.csv, which the
    # shared Entra users collector writes.
    # https://learn.microsoft.com/graph/api/resources/licenseassignmentstate
    UserLicenses = @(
        'RunDate'
        'UserId'
        'UserPrincipalName'
        'AccountEnabled'
        'UsageLocation'
        'OfficeLocation'
        'SkuId'
        'AssignedByGroup'
        'AssignmentType'
        'State'
        'Error'
        'DisabledPlans'
    )

    # Source 3 (optional). One row per service plan of each licence a user holds.
    # https://learn.microsoft.com/graph/api/resources/licensedetails
    LicenseDetails = @(
        'RunDate'
        'UserId'
        'SkuId'
        'SkuPartNumber'
        'ServicePlanId'
        'ServicePlanName'
        'ProvisioningStatus'
    )

    # Source 4. An empty cell means Graph returned no value; it never holds
    # 0001-01-01T00:00:00Z. Same shape as identity-posture's user-signin-activity.csv.
    # https://learn.microsoft.com/graph/api/resources/signinactivity
    UserSignInActivity = @(
        'RunDate'
        'UserId'
        'UserPrincipalName'
        'LastSignInDateTime'
        'LastNonInteractiveSignInDateTime'
        'LastSuccessfulSignInDateTime'
    )

    # Source 6. https://learn.microsoft.com/graph/api/resources/adminreportsettings
    ReportSettings = @(
        'RunDate'
        'DisplayConcealedNames'
    )

    # Sources 5a to 5g. The headers of each usage report, copied from the Learn page's
    # CSV schema. A CSV column is RunDate, then each header with the characters that
    # are not letters or digits removed and each word capitalised (Outlook (Windows)
    # becomes OutlookWindows), in this order. A header is matched ignoring case,
    # spaces and punctuation, so a report that spells one differently still lands in
    # the right column. A header the report does not return is an empty cell.
    UsageReports = @{
        # https://learn.microsoft.com/graph/api/reportroot-getoffice365activeuserdetail
        # That page's CSV schema ends at Assigned Products and does not list Report Period.
        # The collector still writes ReportPeriod, filled from -Period when the download omits it.
        ActiveUserUsage = @(
            'Report Refresh Date', 'User Principal Name', 'Display Name', 'Is Deleted', 'Deleted Date'
            'Has Exchange License', 'Has OneDrive License', 'Has SharePoint License'
            'Has Skype For Business License', 'Has Yammer License', 'Has Teams License'
            'Exchange Last Activity Date', 'OneDrive Last Activity Date', 'SharePoint Last Activity Date'
            'Skype For Business Last Activity Date', 'Yammer Last Activity Date', 'Teams Last Activity Date'
            'Exchange License Assign Date', 'OneDrive License Assign Date', 'SharePoint License Assign Date'
            'Skype For Business License Assign Date', 'Yammer License Assign Date', 'Teams License Assign Date'
            'Assigned Products', 'Report Period'
        )
        # https://learn.microsoft.com/graph/api/reportroot-getemailactivityuserdetail
        EmailActivityUsage = @(
            'Report Refresh Date', 'User Principal Name', 'Display Name', 'Is Deleted', 'Deleted Date'
            'Last Activity Date', 'Send Count', 'Receive Count', 'Read Count'
            'Meeting Created Count', 'Meeting Interacted Count', 'Assigned Products', 'Report Period'
        )
        # https://learn.microsoft.com/graph/api/reportroot-getteamsuseractivityuserdetail
        TeamsActivityUsage = @(
            'Report Refresh Date', 'Tenant Display Name', 'Shared Channel Tenant Display Names', 'User Id'
            'User Principal Name', 'Last Activity Date', 'Is Deleted', 'Deleted Date', 'Assigned Products'
            'Team Chat Message Count', 'Private Chat Message Count', 'Call Count', 'Meeting Count'
            'Post Messages', 'Reply Messages', 'Urgent Messages', 'Meetings Organized Count'
            'Meetings Attended Count', 'Ad Hoc Meetings Organized Count', 'Ad Hoc Meetings Attended Count'
            'Scheduled One-time Meetings Organized Count', 'Scheduled One-time Meetings Attended Count'
            'Scheduled Recurring Meetings Organized Count', 'Scheduled Recurring Meetings Attended Count'
            'Audio Duration', 'Video Duration', 'Screen Share Duration', 'Audio Duration In Seconds'
            'Video Duration In Seconds', 'Screen Share Duration In Seconds', 'Has Other Action'
            'Is Licensed', 'Report Period'
        )
        # https://learn.microsoft.com/graph/api/reportroot-getsharepointactivityuserdetail
        SharePointActivityUsage = @(
            'Report Refresh Date', 'User Principal Name', 'Is Deleted', 'Deleted Date', 'Last Activity Date'
            'Viewed Or Edited File Count', 'Synced File Count', 'Shared Internally File Count'
            'Shared Externally File Count', 'Visited Page Count', 'Assigned Products', 'Report Period'
        )
        # https://learn.microsoft.com/graph/api/reportroot-getonedriveactivityuserdetail
        OneDriveActivityUsage = @(
            'Report Refresh Date', 'User Principal Name', 'Is Deleted', 'Deleted Date', 'Last Activity Date'
            'Viewed Or Edited File Count', 'Synced File Count', 'Shared Internally File Count'
            'Shared Externally File Count', 'Assigned Products', 'Report Period'
        )
        # https://learn.microsoft.com/graph/api/reportroot-getm365appuserdetail
        M365AppUsage = @(
            'Report Refresh Date', 'User Principal Name', 'Last Activation Date', 'Last Activity Date', 'Report Period'
            'Windows', 'Mac', 'Mobile', 'Web'
            'Outlook', 'Word', 'Excel', 'PowerPoint', 'OneNote', 'Teams'
            'Outlook (Windows)', 'Word (Windows)', 'Excel (Windows)', 'PowerPoint (Windows)', 'OneNote (Windows)', 'Teams (Windows)'
            'Outlook (Mac)', 'Word (Mac)', 'Excel (Mac)', 'PowerPoint (Mac)', 'OneNote (Mac)', 'Teams (Mac)'
            'Outlook (Mobile)', 'Word (Mobile)', 'Excel (Mobile)', 'PowerPoint (Mobile)', 'OneNote (Mobile)', 'Teams (Mobile)'
            'Outlook (Web)', 'Word (Web)', 'Excel (Web)', 'PowerPoint (Web)', 'OneNote (Web)', 'Teams (Web)'
        )
        # https://learn.microsoft.com/microsoft-365-copilot/extensibility/api/admin-settings/reports/copilotreportroot-getmicrosoft365copilotusageuserdetail
        # The first thirteen are the v1 CSV header printed on that page. The last nine are
        # the version 2 additions, which the page names in prose and does not print as a
        # CSV header, so their spelling is UNVERIFIED; a header that does not match leaves
        # the cell empty.
        CopilotUsage = @(
            'Report Refresh Date', 'User Principal Name', 'Display Name', 'Last Activity Date'
            'Copilot Chat Last Activity Date', 'Microsoft Teams Copilot Last Activity Date'
            'Word Copilot Last Activity Date', 'Excel Copilot Last Activity Date'
            'PowerPoint Copilot Last Activity Date', 'Outlook Copilot Last Activity Date'
            'OneNote Copilot Last Activity Date', 'Loop Copilot Last Activity Date', 'Report Period'
            'Prompts submitted for all apps', 'Prompts submitted for Copilot Chat (work)'
            'Prompts submitted for Copilot Chat (web)', 'Active Usage Days for all apps'
            'Copilot Chat (work) Last Activity Date', 'Copilot Chat (web) Last Activity Date'
            'Microsoft 365 Copilot Last Activity Date', 'Edge Last Activity Date'
            'Copilot Agent Last Activity Date'
        )
    }

    # Graph permissions each source needs beyond the shared sign-in's defaults.
    # The shared sign-in already asks for Directory.Read.All, which the subscribedSku
    # and user pages list as a higher privileged alternative to LicenseAssignment.Read.All.
    ExtraScopes = @{
        SubscribedSkus          = @('LicenseAssignment.Read.All')
        UserLicenses            = @('LicenseAssignment.Read.All')
        LicenseDetails          = @('LicenseAssignment.Read.All')
        UserSignInActivity      = @()
        ActiveUserUsage         = @('Reports.Read.All')
        EmailActivityUsage      = @('Reports.Read.All')
        TeamsActivityUsage      = @('Reports.Read.All')
        SharePointActivityUsage = @('Reports.Read.All')
        OneDriveActivityUsage   = @('Reports.Read.All')
        M365AppUsage            = @('Reports.Read.All')
        CopilotUsage            = @('Reports.Read.All')
        ReportSettings          = @('ReportSettings.Read.All')
    }

    # Availability per source and cloud, from the contract's source table.
    #   Available    - documented as available; collect it.
    #   NotAvailable - documented as unavailable; skip it, write a header-only CSV.
    #   Unverified   - no Microsoft page says either way; attempt it and log a warning.
    # GCC calls the global service: https://learn.microsoft.com/graph/deployments
    SourceAvailability = @{
        SubscribedSkus = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/subscribedsku-list' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/subscribedsku-list' }
        }
        UserLicenses = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/user-list' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/user-list' }
        }
        LicenseDetails = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/user-list-licensedetails' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/user-list-licensedetails' }
        }
        # List users is available in US Government L4; no page states the signInActivity
        # property itself in that cloud (BRO-243).
        UserSignInActivity = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/resources/user' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/graph/api/user-list' }
        }
        ActiveUserUsage = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getoffice365activeuserdetail' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'NotAvailable'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getoffice365activeuserdetail' }
        }
        EmailActivityUsage = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getemailactivityuserdetail' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'NotAvailable'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getemailactivityuserdetail' }
        }
        TeamsActivityUsage = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getteamsuseractivityuserdetail' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'NotAvailable'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getteamsuseractivityuserdetail' }
        }
        SharePointActivityUsage = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getsharepointactivityuserdetail' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'NotAvailable'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getsharepointactivityuserdetail' }
        }
        OneDriveActivityUsage = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getonedriveactivityuserdetail' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'NotAvailable'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getonedriveactivityuserdetail' }
        }
        M365AppUsage = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getm365appuserdetail' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'NotAvailable'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getm365appuserdetail' }
        }
        CopilotUsage = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/reportroot-getmicrosoft365copilotusageuserdetail?view=graph-rest-beta' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'NotAvailable'; Reference = 'https://learn.microsoft.com/microsoft-365-copilot/extensibility/api/admin-settings/reports/copilotreportroot-getmicrosoft365copilotusageuserdetail' }
        }
        # The Graph page marks US Government L4 unavailable; the activity-reports page
        # says the API works in all environments. Attempted, with a warning.
        ReportSettings = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/adminreportsettings-get' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/microsoft-365/admin/activity-reports/activity-reports' }
        }
    }
}
