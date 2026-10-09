@{
    # The column order of every CSV this report produces. The collectors, the
    # sample-data generator and the tests all read the order from here, so a
    # sample file can never drift from its collector's output.
    # Contract: docs/candidates/copilot-studio-agents.md

    # Source 1. https://learn.microsoft.com/power-platform/admin/inventory-schema#environments
    Environments = @(
        'RunDate'
        'EnvironmentId'
        'DisplayName'
        'Location'
        'EnvironmentType'
        'IsManaged'
        'EnvironmentGroup'
        'LastModifiedAt'
    )

    # Source 2. https://learn.microsoft.com/microsoft-copilot-studio/admin-agent-inventory
    # The Listed* columns count what the inventory returned for the agent. The inventory
    # returns a random 200 of one resource type when an agent has more, so the
    # Capabilities* columns carry capabilitiesCounts, the complete count. A Listed value
    # below its Capabilities value means the connector rows for that agent are partial.
    # Viewer and editor identities are not collected: only counts are documented.
    Agents = @(
        'RunDate'
        'EnvironmentId'
        'AgentId'
        'DisplayName'
        'Harness'
        'CreatedIn'
        'CreatedAt'
        'CreatedBy'
        'OwnerId'
        'LastPublishedAt'
        'IsPublished'
        'IsQuarantined'
        'IsManaged'
        'Orchestration'
        'Model'
        'Authentication'
        'Channels'
        'ViewerUserCount'
        'ViewerGroupCount'
        'ViewerEntireTenant'
        'EditorUserCount'
        'EditorGroupCount'
        'ListedConnectorCount'
        'ListedConnectorOperationCount'
        'CapabilitiesDistinctConnectors'
        'CapabilitiesDistinctConnectorOperations'
        'IsWebSearchEnabledForKnowledge'
    )

    # Source 3. https://learn.microsoft.com/microsoft-copilot-studio/admin-agent-inventory#connector-properties
    # UsedAs: Tool, Topic Tool or Knowledge.
    AgentConnectors = @(
        'RunDate'
        'EnvironmentId'
        'AgentId'
        'ConnectorId'
        'OperationId'
        'UsedAs'
        'IsEnabled'
        'RequiresEndUserConsent'
        'WhenCanBeUsed'
        'ConnectionProvider'
    )

    # Source 4. Category is one of KnowledgeSource, Tool, HttpRequest, Prompt, Mcp. A
    # component that carries two markers produces two rows.
    # https://learn.microsoft.com/microsoft-copilot-studio/guidance/kit-agent-inventory-data-source
    AgentComponents = @(
        'RunDate'
        'DataverseUrl'
        'AgentId'
        'ComponentId'
        'ComponentName'
        'ComponentType'
        'Category'
    )

    # Source 5. The only place a last-modified date and last modifier are available.
    AgentModifications = @(
        'RunDate'
        'DataverseUrl'
        'AgentId'
        'Name'
        'CreatedOn'
        'ModifiedOn'
        'ModifiedBy'
        'PublishedOn'
    )

    # Source 6. Field names are the Copilot Studio audit fields:
    # https://learn.microsoft.com/microsoft-copilot-studio/admin-logging-copilot-studio#schema-audit-fields
    AgentAuditEvents = @(
        'CreationTime'
        'Id'
        'Operation'
        'UserId'
        'ResultStatus'
        'BotId'
        'BotSchemaName'
        'BotComponentId'
        'BotComponentType'
    )

    # The authoring event labels, copied exactly from
    # https://learn.microsoft.com/microsoft-copilot-studio/admin-logging-copilot-studio#see-audited-events-agent-authoring
    # CopilotInteraction (agent usage) is a record type, not one of these, and is not collected.
    AuditOperations = @(
        'BotDeleteCleanup'
        'BotUpdateOperation-BotNameUpdate'
        'BotCreate'
        'BotDelete'
        'BotUpdateOperation-BotAuthUpdate'
        'BotUpdateOperation-BotIconUpdate'
        'BotUpdateOperation-BotPublish'
        'BotUpdateOperation-BotShare'
        'BotAppInsightsUpdate'
        'BotComponentCreate'
        'BotComponentUpdate'
        'BotComponentDelete'
        'BotComponentCollectionCreate'
        'BotComponentCollectionDelete'
        'BotComponentCollectionUpdate'
        'AIPluginOperationCreate'
        'AIPluginOperationUpdate'
        'AIPluginOperationDelete'
        'EnvironmentVariableCreate'
        'EnvironmentVariableUpdate'
        'EnvironmentVariableDelete'
    )

    # Markers in botcomponent.data, from the Agent Inventory data source page. Each
    # is matched as a plain substring.
    ComponentMarkers = @(
        @{ Category = 'KnowledgeSource'; Marker = 'KnowledgeSourceConfiguration' }
        @{ Category = 'KnowledgeSource'; Marker = 'FileDataName' }
        @{ Category = 'Tool'; Marker = 'TaskDialog' }
        @{ Category = 'HttpRequest'; Marker = 'HttpRequestAction' }
        @{ Category = 'Prompt'; Marker = 'InvokeAIBuilderModelAction' }
        @{ Category = 'Mcp'; Marker = 'InvokeExternalAgentTaskAction' }
    )

    # Available / NotAvailable / Unverified, from the source table in the contract doc.
    # Only NotAvailable skips a source. Sources 1 and 2 are Available in all three clouds,
    # but the Power Platform API host for GCC and GCC High is UNVERIFIED, so those clouds
    # need -ApiHost.
    SourceAvailability = @{
        Environments = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/power-platform/admin/power-platform-inventory' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/power-platform/admin/power-platform-inventory#sovereign-clouds' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/power-platform/admin/power-platform-inventory#sovereign-clouds' }
        }
        Agents = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/power-platform/admin/power-platform-inventory' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/power-platform/admin/power-platform-inventory#sovereign-clouds' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/power-platform/admin/power-platform-inventory#sovereign-clouds' }
        }
        AgentConnectors = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/power-platform/admin/power-platform-inventory#connector-inventory-preview' }
            GCC        = @{ Status = 'NotAvailable'; Reference = 'https://learn.microsoft.com/power-platform/admin/power-platform-inventory#known-limitations' }
            GCCHigh    = @{ Status = 'NotAvailable'; Reference = 'https://learn.microsoft.com/power-platform/admin/power-platform-inventory#known-limitations' }
        }
        AgentComponents = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/microsoft-copilot-studio/guidance/custom-analytics-strategy#copilot-studio,-dataverse,-and-analytics' }
            GCC        = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/power-platform/admin/microsoft-dynamics-365-government#dynamics-365-us-government-urls' }
            GCCHigh    = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/power-platform/admin/microsoft-dynamics-365-government#dynamics-365-us-government-urls' }
        }
        AgentModifications = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/microsoft-copilot-studio/guidance/custom-analytics-strategy#copilot-studio,-dataverse,-and-analytics' }
            GCC        = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/power-platform/admin/microsoft-dynamics-365-government#dynamics-365-us-government-urls' }
            GCCHigh    = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/power-platform/admin/microsoft-dynamics-365-government#dynamics-365-us-government-urls' }
        }
        AgentAuditEvents = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/microsoft-copilot-studio/admin-logging-copilot-studio' }
            GCC        = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/microsoft-copilot-studio/admin-logging-copilot-studio#prerequisites' }
            GCCHigh    = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/microsoft-copilot-studio/admin-logging-copilot-studio#prerequisites' }
        }
    }
}
