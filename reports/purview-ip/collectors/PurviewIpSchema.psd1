@{
    # The column order of every CSV this report produces, other than users.csv
    # (written by the shared module's Invoke-EntraUserCollector). The
    # collectors, the sample-data generator and the tests all read the order
    # from here, so a sample file can never drift from its collector's output.

    Policies                = @(
        'RunDate'
        'ObjectType'
        'ObjectId'
        'Name'
        'DisplayName'
        'ParentId'
        'ParentName'
        'Priority'
        'Mode'
        'Enabled'
        'Workload'
        'Locations'
        'AppliesToCopilot'        # derived: see Get-PurviewCopilotScope
        'LabelIds'
        'LabelNames'
        'RetentionAction'
        'RetentionDuration'
        'RetentionType'
        'IsRecordLabel'
        'SensitiveInformationTypes'
        'BlockAccess'
        'BlockAccessScope'
        'NotifyUser'
        'AllowOverride'
        'RequireJustification'
        'GenerateIncidentReport'
        'IsDefaultLabel'
        'EncryptionEnabled'
        'Comment'
        'WhenCreatedUtc'
        'WhenChangedUtc'
    )

    ActivityExplorerEvents  = @(
        'RecordIdentity'
        'Happened'
        'EventDate'                    # derived: UTC date part of Happened
        'Activity'                     # normalised to the filter-enum name
        'ActivityRaw'                  # exactly what the cmdlet returned
        'ActivityCategory'             # derived: grouping of Activity
        'Workload'
        'Application'
        'User'
        'UserType'
        'ItemName'
        'FilePath'
        'FullUrl'
        'FileExtension'
        'DeviceName'
        'Platform'
        'SensitivityLabel'
        'OldSensitivityLabel'
        'SensitivityLabelPolicyId'
        'LabelEventType'
        'IsLabelDowngrade'             # derived: LabelEventType -eq 'LabelDowngraded'
        'HowApplied'
        'Justification'
        'IsProtected'
        'IsProtectedBefore'
        'ProtectionEventType'
        'ProtectionType'
        'ProtectionOwner'
        'RMSEncrypted'
        'RetentionLabel'
        'OldRetentionLabel'
        'SensitiveInfoTypeName'        # derived from SensitiveInfoTypeData
        'SensitiveInfoTypeCount'       # derived from SensitiveInfoTypeData
        'SensitiveInfoTypeConfidence'  # derived from SensitiveInfoTypeData
        'PolicyId'
        'PolicyName'
        'PolicyMode'
        'RuleId'
        'RuleName'
        'RuleActions'
        'EnforcementMode'
        'FalsePositive'
        'DlpPolicyMatchId'
    )

    ContentExplorerSnapshot = @(
        'RunDate'
        'TagType'
        'TagName'
        'Workload'
        'TotalCount'
    )

    CopilotAccessedResources = @(
        'RecordId'
        'CreationTime'
        'EventDate'                    # derived: UTC date part of CreationTime
        'Operation'
        'RecordType'
        'Workload'
        'UserId'
        'UserKey'
        'UserType'
        'AppHost'
        'AppIdentity'
        'AgentId'
        'AgentName'
        'ThreadId'
        'ResourceId'
        'ResourceName'
        'ResourceType'
        'ResourceAction'
        'SiteUrl'
        'ListItemUniqueId'
        'SensitivityLabelId'
        'Status'
        'AccessBlocked'                # derived: Status/PolicyDetails say access was denied
        'XpiaDetected'
        'PolicyId'                     # from AccessedResources[].PolicyDetails
        'PolicyName'                   # from AccessedResources[].PolicyDetails
        'PolicyRules'                  # from AccessedResources[].PolicyDetails
    )

    # The ObjectType values Policies.csv uses, in the order Get-Policies.ps1
    # writes them.
    PolicyObjectTypes       = @(
        'SensitivityLabel'
        'LabelPolicy'
        'AutoLabelingPolicy'
        'DlpPolicy'
        'DlpRule'
        'RetentionPolicy'
        'RetentionLabel'
    )

    # Activity Explorer activity values, grouped for reporting. The activity
    # names are the ones Export-ActivityExplorerData accepts in its Activity
    # filter:
    # https://learn.microsoft.com/powershell/module/exchangepowershell/export-activityexplorerdata
    # Plain hashtable, not [ordered]: Import-PowerShellDataFile's restricted
    # data language does not allow the [ordered] type accelerator. Nothing
    # here depends on enumeration order - each category is checked by
    # membership, not by position.
    ActivityCategories      = @{
        Labeling   = @(
            'LabelApplied', 'LabelChanged', 'LabelRemoved',
            'LabelRecommended', 'LabelRecommendedAndDismissed', 'AutoLabelingSimulation'
        )
        Protection = @('NewProtection', 'ChangeProtection', 'RemoveProtection')
        Dlp        = @('DLPRuleMatch', 'DLPRuleEnforce', 'DLPRuleUndo', 'DLPInfo', 'DlpClassification')
        Retention  = @('ClassificationAdded', 'ClassificationUpdated', 'ClassificationDeleted')
        Ai         = @('CopilotInteraction', 'AIAppInteraction')
        Discovery  = @('FileDiscovered')
        Endpoint   = @(
            'ArchiveCreated', 'DownloadFile', 'DownloadText', 'FileAccessedByUnallowedApp',
            'FileArchived', 'FileCopiedToClipboard', 'FileCopiedToNetworkShare',
            'FileCopiedToRemoteDesktopSession', 'FileCopiedToRemovableMedia', 'FileCreated',
            'FileCreatedOnNetworkShare', 'FileCreatedOnRemovableMedia', 'FileDeleted',
            'FileModified', 'FilePrinted', 'FileRead', 'FileRenamed',
            'FileTransferredByBluetooth', 'FileUploadedToCloud', 'PastedToBrowser',
            'ScreenCapture', 'UploadFile', 'UploadText', 'WebpageCopiedToClipboard',
            'WebpagePrinted', 'WebpageSavedToLocal'
        )
    }

    # Status is one of:
    #   Available    - documented as available in this cloud; collect it.
    #   NotAvailable - documented as unavailable; skip it, write a header-only CSV.
    #   Unverified   - no first-party statement found. The collector still
    #                  attempts the source and reports whatever the tenant
    #                  returns, rather than silently dropping data on a guess.
    SourceAvailability      = @{
        ActivityExplorer      = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/purview/data-classification-activity-explorer' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-deployments' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-high-deployments' }
        }
        ContentExplorer       = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/purview/data-classification-content-explorer' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-deployments' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-high-deployments' }
        }
        # Content Explorer covers Teams in Commercial, but the government
        # clouds list "Content explorer includes Teams data" as still in
        # development, so asking for the Teams workload there returns nothing
        # useful.
        ContentExplorerTeams  = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/purview/data-classification-content-explorer' }
            GCC        = @{ Status = 'NotAvailable'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-deployments' }
            GCCHigh    = @{ Status = 'NotAvailable'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-high-deployments' }
        }
        CopilotAuditRecords   = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/purview/audit-copilot' }
            GCC        = @{ Status = 'Unverified'; Reference = $null }
            GCCHigh    = @{ Status = 'Unverified'; Reference = $null }
        }
        Policies              = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/powershell/exchange/connect-to-scc-powershell' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/powershell/exchange/connect-to-scc-powershell' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/powershell/exchange/connect-to-scc-powershell' }
        }
    }

    # The Microsoft 365 Copilot location GUID used in a DLP policy's Locations
    # JSON, and the enforcement plane that scopes a policy to Copilot
    # experiences.
    # https://learn.microsoft.com/powershell/module/exchangepowershell/new-dlpcompliancepolicy
    CopilotLocationId       = '470f2276-e011-4e9d-a6ec-20768be3a4b0'
    CopilotEnforcementPlane = 'CopilotExperiences'
}
