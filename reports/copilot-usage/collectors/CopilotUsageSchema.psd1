@{
    # The column order of every CSV this report produces, and the report header each
    # column is read from. The collectors, the sample files and the tests all read the
    # order from here, so a sample file can never drift from its collector's output.
    # Contract: docs/candidates/copilot-usage.md

    # Source 2. Headers are named as on
    # https://learn.microsoft.com/microsoft-365-copilot/extensibility/api/admin-settings/reports/copilotreportroot-getmicrosoft365copilotusageuserdetail
    # (v1.0 returns a CSV stream; version defaults to v2, the only supported value).
    # With concealed names (displayConcealedNames = true, the default) the user principal
    # name and display name are hashes, so the file cannot be joined to a user list.
    UsageUserDetailMap = @(
        @{ Column = 'ReportRefreshDate'; Header = 'Report Refresh Date' }
        @{ Column = 'UserPrincipalName'; Header = 'User Principal Name' }
        @{ Column = 'DisplayName'; Header = 'Display Name' }
        @{ Column = 'LastActivityDate'; Header = 'Last Activity Date' }
        @{ Column = 'CopilotChatLastActivityDate'; Header = 'Copilot Chat Last Activity Date' }
        @{ Column = 'MicrosoftTeamsCopilotLastActivityDate'; Header = 'Microsoft Teams Copilot Last Activity Date' }
        @{ Column = 'WordCopilotLastActivityDate'; Header = 'Word Copilot Last Activity Date' }
        @{ Column = 'ExcelCopilotLastActivityDate'; Header = 'Excel Copilot Last Activity Date' }
        @{ Column = 'PowerPointCopilotLastActivityDate'; Header = 'PowerPoint Copilot Last Activity Date' }
        @{ Column = 'OutlookCopilotLastActivityDate'; Header = 'Outlook Copilot Last Activity Date' }
        @{ Column = 'OneNoteCopilotLastActivityDate'; Header = 'OneNote Copilot Last Activity Date' }
        @{ Column = 'LoopCopilotLastActivityDate'; Header = 'Loop Copilot Last Activity Date' }
        @{ Column = 'ReportPeriod'; Header = 'Report Period' }
        @{ Column = 'PromptsSubmittedAllApps'; Header = 'Prompts submitted for all apps' }
        @{ Column = 'PromptsSubmittedCopilotChatWork'; Header = 'Prompts submitted for Copilot Chat (work)' }
        @{ Column = 'PromptsSubmittedCopilotChatWeb'; Header = 'Prompts submitted for Copilot Chat (web)' }
        @{ Column = 'ActiveUsageDaysAllApps'; Header = 'Active Usage Days for all apps' }
        @{ Column = 'CopilotChatWorkLastActivityDate'; Header = 'Copilot Chat (work) Last Activity Date' }
        @{ Column = 'CopilotChatWebLastActivityDate'; Header = 'Copilot Chat (web) Last Activity Date' }
        @{ Column = 'Microsoft365CopilotLastActivityDate'; Header = 'Microsoft 365 Copilot Last Activity Date' }
        @{ Column = 'EdgeLastActivityDate'; Header = 'Edge Last Activity Date' }
        @{ Column = 'CopilotAgentLastActivityDate'; Header = 'Copilot Agent Last Activity Date' }
    )

    # Source 3. One Enabled Users and one Active Users header per app. v1 apps first, then
    # the v2 additions, then the two prompt columns.
    # https://learn.microsoft.com/microsoft-365-copilot/extensibility/api/admin-settings/reports/copilotreportroot-getmicrosoft365copilotusercountsummary
    UserCountSummaryMap = @(
        @{ Column = 'ReportRefreshDate'; Header = 'Report Refresh Date' }
        @{ Column = 'ReportPeriod'; Header = 'Report Period' }
        @{ Column = 'MicrosoftTeamsEnabledUsers'; Header = 'Microsoft Teams Enabled Users' }
        @{ Column = 'MicrosoftTeamsActiveUsers'; Header = 'Microsoft Teams Active Users' }
        @{ Column = 'WordEnabledUsers'; Header = 'Word Enabled Users' }
        @{ Column = 'WordActiveUsers'; Header = 'Word Active Users' }
        @{ Column = 'PowerPointEnabledUsers'; Header = 'PowerPoint Enabled Users' }
        @{ Column = 'PowerPointActiveUsers'; Header = 'PowerPoint Active Users' }
        @{ Column = 'OutlookEnabledUsers'; Header = 'Outlook Enabled Users' }
        @{ Column = 'OutlookActiveUsers'; Header = 'Outlook Active Users' }
        @{ Column = 'ExcelEnabledUsers'; Header = 'Excel Enabled Users' }
        @{ Column = 'ExcelActiveUsers'; Header = 'Excel Active Users' }
        @{ Column = 'OneNoteEnabledUsers'; Header = 'OneNote Enabled Users' }
        @{ Column = 'OneNoteActiveUsers'; Header = 'OneNote Active Users' }
        @{ Column = 'LoopEnabledUsers'; Header = 'Loop Enabled Users' }
        @{ Column = 'LoopActiveUsers'; Header = 'Loop Active Users' }
        @{ Column = 'AnyAppEnabledUsers'; Header = 'Any App Enabled Users' }
        @{ Column = 'AnyAppActiveUsers'; Header = 'Any App Active Users' }
        @{ Column = 'CopilotChatEnabledUsers'; Header = 'Copilot Chat Enabled Users' }
        @{ Column = 'CopilotChatActiveUsers'; Header = 'Copilot Chat Active Users' }
        @{ Column = 'EdgeEnabledUsers'; Header = 'Edge Enabled Users' }
        @{ Column = 'EdgeActiveUsers'; Header = 'Edge Active Users' }
        @{ Column = 'Microsoft365CopilotEnabledUsers'; Header = 'Microsoft 365 Copilot Enabled Users' }
        @{ Column = 'Microsoft365CopilotActiveUsers'; Header = 'Microsoft 365 Copilot Active Users' }
        @{ Column = 'CopilotChatWorkEnabledUsers'; Header = 'Copilot Chat (work) Enabled Users' }
        @{ Column = 'CopilotChatWorkActiveUsers'; Header = 'Copilot Chat (work) Active Users' }
        @{ Column = 'CopilotChatWebEnabledUsers'; Header = 'Copilot Chat (web) Enabled Users' }
        @{ Column = 'CopilotChatWebActiveUsers'; Header = 'Copilot Chat (web) Active Users' }
        @{ Column = 'TotalPromptsSubmitted'; Header = 'Total prompts submitted' }
        @{ Column = 'AveragePromptsSubmitted'; Header = 'Average prompts submitted' }
    )

    # Source 4. The same pairs as source 3, one row per Report Date, and the daily
    # Prompts submitted column (not the summary's total or average).
    # https://learn.microsoft.com/microsoft-365-copilot/extensibility/api/admin-settings/reports/copilotreportroot-getmicrosoft365copilotusercounttrend
    UserCountTrendMap = @(
        @{ Column = 'ReportRefreshDate'; Header = 'Report Refresh Date' }
        @{ Column = 'ReportDate'; Header = 'Report Date' }
        @{ Column = 'ReportPeriod'; Header = 'Report Period' }
        @{ Column = 'MicrosoftTeamsEnabledUsers'; Header = 'Microsoft Teams Enabled Users' }
        @{ Column = 'MicrosoftTeamsActiveUsers'; Header = 'Microsoft Teams Active Users' }
        @{ Column = 'WordEnabledUsers'; Header = 'Word Enabled Users' }
        @{ Column = 'WordActiveUsers'; Header = 'Word Active Users' }
        @{ Column = 'PowerPointEnabledUsers'; Header = 'PowerPoint Enabled Users' }
        @{ Column = 'PowerPointActiveUsers'; Header = 'PowerPoint Active Users' }
        @{ Column = 'OutlookEnabledUsers'; Header = 'Outlook Enabled Users' }
        @{ Column = 'OutlookActiveUsers'; Header = 'Outlook Active Users' }
        @{ Column = 'ExcelEnabledUsers'; Header = 'Excel Enabled Users' }
        @{ Column = 'ExcelActiveUsers'; Header = 'Excel Active Users' }
        @{ Column = 'OneNoteEnabledUsers'; Header = 'OneNote Enabled Users' }
        @{ Column = 'OneNoteActiveUsers'; Header = 'OneNote Active Users' }
        @{ Column = 'LoopEnabledUsers'; Header = 'Loop Enabled Users' }
        @{ Column = 'LoopActiveUsers'; Header = 'Loop Active Users' }
        @{ Column = 'AnyAppEnabledUsers'; Header = 'Any App Enabled Users' }
        @{ Column = 'AnyAppActiveUsers'; Header = 'Any App Active Users' }
        @{ Column = 'CopilotChatEnabledUsers'; Header = 'Copilot Chat Enabled Users' }
        @{ Column = 'CopilotChatActiveUsers'; Header = 'Copilot Chat Active Users' }
        @{ Column = 'EdgeEnabledUsers'; Header = 'Edge Enabled Users' }
        @{ Column = 'EdgeActiveUsers'; Header = 'Edge Active Users' }
        @{ Column = 'Microsoft365CopilotEnabledUsers'; Header = 'Microsoft 365 Copilot Enabled Users' }
        @{ Column = 'Microsoft365CopilotActiveUsers'; Header = 'Microsoft 365 Copilot Active Users' }
        @{ Column = 'CopilotChatWorkEnabledUsers'; Header = 'Copilot Chat (work) Enabled Users' }
        @{ Column = 'CopilotChatWorkActiveUsers'; Header = 'Copilot Chat (work) Active Users' }
        @{ Column = 'CopilotChatWebEnabledUsers'; Header = 'Copilot Chat (web) Enabled Users' }
        @{ Column = 'CopilotChatWebActiveUsers'; Header = 'Copilot Chat (web) Active Users' }
        @{ Column = 'PromptsSubmitted'; Header = 'Prompts submitted' }
    )

    # Source 6. One row per CopilotInteraction audit record. A record is NOT one prompt: it
    # typically holds a prompt and a response, and can hold one prompt with several
    # responses. MessageCount and PromptMessageCount say how many of each. The message
    # text is not in the record (Messages hold an id and isPrompt) and is not collected.
    # https://learn.microsoft.com/purview/audit-copilot
    CopilotAuditEvents = @(
        'CreationTime'
        'Id'
        'UserId'
        'Operation'
        'RecordType'
        'Workload'
        'AppHost'
        'AppIdentity'
        'AgentId'
        'AgentName'
        'MessageCount'
        'PromptMessageCount'
        'AccessedResourceCount'
        'PluginIds'
    )

    # Source 8. One row per interaction: metadata only. body, attachments, links and
    # mentions carry prompt and response content and are never read into the file.
    # https://learn.microsoft.com/microsoft-365-copilot/extensibility/api/ai-services/interaction-export/aiinteractionhistory-getallenterpriseinteractions
    CopilotInteractions = @(
        'CreatedDateTime'
        'Id'
        'UserId'
        'SessionId'
        'RequestId'
        'AppClass'
        'InteractionType'
        'ConversationType'
        'Locale'
        'ContextCount'
        'ContextTypes'
    )

    # Source 10. Reference, not a tenant call. A dated copy of the "feature availability"
    # table of the Microsoft 365 Copilot service description. Only what the contract
    # (docs/candidates/copilot-usage.md, source 10) records is here; NotStated means the
    # contract does not say, not that the feature is absent.
    CopilotFeatureAvailability = @(
        'RunDate'
        'Feature'
        'Commercial'
        'GCC'
        'GCCHigh'
        'Note'
        'PageReadDate'
        'Reference'
    )

    FeaturePageReadDate = '2026-10-10'
    FeaturePage = 'https://learn.microsoft.com/office365/servicedescriptions/office-365-platform-service-description/microsoft-365-copilot#feature-availability'

    # Status is Yes, NotCurrentlyAvailable, Limited or NotStated.
    FeatureRows = @(
        @{ Feature = 'Microsoft Copilot'; Commercial = 'Yes'; GCC = 'Yes'; GCCHigh = 'Yes'; Note = 'Also listed for DoD.' }
        @{ Feature = 'Copilot in Outlook'; Commercial = 'Yes'; GCC = 'Yes'; GCCHigh = 'Yes'; Note = '' }
        @{ Feature = 'Copilot in Teams (chat, channel, meetings)'; Commercial = 'NotStated'; GCC = 'Yes'; GCCHigh = 'NotCurrentlyAvailable'; Note = 'A GCC High tenant with no Teams Copilot activity is expected, not a collection gap.' }
        @{ Feature = 'Copilot in SharePoint'; Commercial = 'NotStated'; GCC = 'NotStated'; GCCHigh = 'NotCurrentlyAvailable'; Note = '' }
        @{ Feature = 'SharePoint agents'; Commercial = 'NotStated'; GCC = 'NotStated'; GCCHigh = 'NotCurrentlyAvailable'; Note = '' }
        @{ Feature = 'Schedule with Copilot'; Commercial = 'NotStated'; GCC = 'NotCurrentlyAvailable'; GCCHigh = 'NotCurrentlyAvailable'; Note = 'Not yet available in GCC and GCC High.' }
        @{ Feature = 'Themes by Copilot'; Commercial = 'NotStated'; GCC = 'NotCurrentlyAvailable'; GCCHigh = 'NotCurrentlyAvailable'; Note = 'Not yet available in GCC and GCC High.' }
        @{ Feature = 'Word Agent'; Commercial = 'NotStated'; GCC = 'NotCurrentlyAvailable'; GCCHigh = 'NotCurrentlyAvailable'; Note = '' }
        @{ Feature = 'Excel Agent'; Commercial = 'NotStated'; GCC = 'NotCurrentlyAvailable'; GCCHigh = 'NotCurrentlyAvailable'; Note = '' }
        @{ Feature = 'PowerPoint Agent'; Commercial = 'NotStated'; GCC = 'NotCurrentlyAvailable'; GCCHigh = 'NotCurrentlyAvailable'; Note = '' }
        @{ Feature = 'Edit with Copilot in Word'; Commercial = 'NotStated'; GCC = 'NotCurrentlyAvailable'; GCCHigh = 'NotCurrentlyAvailable'; Note = '' }
        @{ Feature = 'Edit with Copilot in Excel'; Commercial = 'NotStated'; GCC = 'NotCurrentlyAvailable'; GCCHigh = 'NotCurrentlyAvailable'; Note = '' }
        @{ Feature = 'Edit with Copilot in PowerPoint'; Commercial = 'NotStated'; GCC = 'NotCurrentlyAvailable'; GCCHigh = 'NotCurrentlyAvailable'; Note = '' }
        @{ Feature = 'Edit with Copilot in OneNote'; Commercial = 'NotStated'; GCC = 'NotCurrentlyAvailable'; GCCHigh = 'NotCurrentlyAvailable'; Note = '' }
        @{ Feature = 'Scheduled Prompts'; Commercial = 'NotStated'; GCC = 'NotCurrentlyAvailable'; GCCHigh = 'NotCurrentlyAvailable'; Note = '' }
        @{ Feature = 'Power Platform connectors'; Commercial = 'NotStated'; GCC = 'NotStated'; GCCHigh = 'NotCurrentlyAvailable'; Note = '' }
        @{ Feature = 'Copilot Chat in the Edge sidebar'; Commercial = 'NotStated'; GCC = 'NotCurrentlyAvailable'; GCCHigh = 'NotCurrentlyAvailable'; Note = 'Also not available in DoD.' }
        @{ Feature = 'Copilot Chat in web apps for Microsoft 365 Copilot (Basic) users'; Commercial = 'NotStated'; GCC = 'NotCurrentlyAvailable'; GCCHigh = 'NotCurrentlyAvailable'; Note = 'Also not available in DoD.' }
        @{ Feature = 'Copilot Analytics in Viva Insights'; Commercial = 'NotStated'; GCC = 'Limited'; GCCHigh = 'NotCurrentlyAvailable'; Note = 'GCC: Yes, limited core metrics.' }
        @{ Feature = 'Microsoft Purview controls for Copilot'; Commercial = 'NotStated'; GCC = 'Yes'; GCCHigh = 'Yes'; Note = 'This is the feature, not the Search-UnifiedAuditLog cmdlet. Not every Purview feature: see Insider Risk Management.' }
        @{ Feature = 'Insider Risk Management'; Commercial = 'NotStated'; GCC = 'NotStated'; GCCHigh = 'NotCurrentlyAvailable'; Note = '' }
    )

    # Availability per cloud, from the contract's source table. Only NotAvailable skips;
    # Unverified is attempted with a warning.
    SourceAvailability = @{
        UsageUserDetail = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/microsoft-365-copilot/extensibility/api/admin-settings/reports/copilotreportroot-getmicrosoft365copilotusageuserdetail' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'NotAvailable'; Reference = 'https://learn.microsoft.com/microsoft-365-copilot/extensibility/api/admin-settings/reports/copilotreportroot-getmicrosoft365copilotusageuserdetail' }
        }
        UserCountSummary = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/microsoft-365-copilot/extensibility/api/admin-settings/reports/copilotreportroot-getmicrosoft365copilotusercountsummary' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'NotAvailable'; Reference = 'https://learn.microsoft.com/microsoft-365-copilot/extensibility/api/admin-settings/reports/copilotreportroot-getmicrosoft365copilotusercountsummary' }
        }
        UserCountTrend = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/microsoft-365-copilot/extensibility/api/admin-settings/reports/copilotreportroot-getmicrosoft365copilotusercounttrend' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'NotAvailable'; Reference = 'https://learn.microsoft.com/microsoft-365-copilot/extensibility/api/admin-settings/reports/copilotreportroot-getmicrosoft365copilotusercounttrend' }
        }
        CopilotAuditEvents = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/purview/audit-copilot' }
            GCC        = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/office-365-platform-service-description/microsoft-365-copilot#feature-availability' }
            GCCHigh    = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/office-365-platform-service-description/microsoft-365-copilot#feature-availability' }
        }
        CopilotInteractions = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/microsoft-365-copilot/extensibility/api/ai-services/interaction-export/aiinteractionhistory-getallenterpriseinteractions' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/microsoft-365-copilot/extensibility/api/ai-services/interaction-export/aiinteractionhistory-getallenterpriseinteractions' }
        }
        CopilotFeatureAvailability = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/office-365-platform-service-description/microsoft-365-copilot#feature-availability' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/office-365-platform-service-description/microsoft-365-copilot#feature-availability' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/office-365-platform-service-description/microsoft-365-copilot#feature-availability' }
        }
    }

    # appClass values the interaction export supports. The rest are not supported, so
    # Outlook, PowerPoint, OneNote and Loop are not in this export.
    # https://learn.microsoft.com/microsoftteams/export-teams-content-copilot#supported-appclass-filters
    SupportedAppClasses = @(
        'IPM.SkypeTeams.Message.Copilot.Word'
        'IPM.SkypeTeams.Message.Copilot.Excel'
        'IPM.SkypeTeams.Message.Copilot.Teams'
        'IPM.SkypeTeams.Message.Copilot.BizChat'
        'IPM.SkypeTeams.Message.Copilot.WebChat'
        'IPM.SkypeTeams.Message.Copilot.CoworkChat'
    )
}
