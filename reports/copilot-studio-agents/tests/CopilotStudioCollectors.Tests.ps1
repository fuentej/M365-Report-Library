#Requires -Version 7.0

BeforeAll {
    $script:Root = Join-Path $PSScriptRoot '../../..' | Resolve-Path | Select-Object -ExpandProperty Path
    $script:Collectors = Join-Path $script:Root 'reports/copilot-studio-agents/collectors'
    $script:Samples = Join-Path $script:Root 'reports/copilot-studio-agents/samples'

    . (Join-Path $script:Root 'shared/tests/TenantCmdletStubs.ps1')
    . (Join-Path $PSScriptRoot 'CopilotStudioStubs.ps1')
    Import-Module (Join-Path $script:Root 'shared/M365ReportLibrary.psm1') -Force

    $script:Schema = Import-PowerShellDataFile -LiteralPath (Join-Path $script:Collectors 'CopilotStudioSchema.psd1')

    function New-TestFolder {
        $path = Join-Path ([System.IO.Path]::GetTempPath()) ('copilot-studio-' + [guid]::NewGuid().ToString('N'))
        New-Item -Path $path -ItemType Directory -Force | Out-Null
        return $path
    }

    function Get-HeaderText {
        param([Parameter(Mandatory)][string]$Path)
        return ((Get-CsvHeaderColumn -Path $Path) -join ',')
    }

    function Invoke-CollectorScript {
        param([Parameter(Mandatory)][string]$Name, [hashtable]$Arguments = @{})
        & (Join-Path $script:Collectors $Name) @Arguments
    }

    # Inventory response: https://learn.microsoft.com/power-platform/admin/inventory-api#response-format
    # (totalRecords, count, resultTruncated, skipToken, data) and the ResourceItem fields at
    # https://learn.microsoft.com/rest/api/power-platform/resourcequery/resource-query/query-resources
    # (id, name, type, location, tenantId, properties). The Invoke-RestMethod result is the
    # parsed JSON, so each response is a PSCustomObject.
    function New-InventoryResponse {
        param([object[]]$Data = @(), [string]$SkipToken = '', [object]$ResultTruncated = 0)

        $response = [ordered]@{
            totalRecords    = $Data.Count
            count           = $Data.Count
            resultTruncated = $ResultTruncated
            data            = @($Data)
        }
        if ($SkipToken) { $response['skipToken'] = $SkipToken }
        return [pscustomobject]$response
    }

    # Environment fields: https://learn.microsoft.com/power-platform/admin/inventory-schema#environments
    function New-MockEnvironmentItem {
        param([string]$Id = 'env-1')

        [pscustomobject]@{
            id         = "/providers/Microsoft.PowerPlatform/environments/$Id"
            name       = $Id
            type       = 'microsoft.powerplatform/environments'
            location   = 'unitedstates'
            tenantId   = '00000000-0000-0000-0000-000000000000'
            properties = [pscustomobject]@{
                displayName      = 'Contoso Production'
                environmentType  = 'Production'
                isManaged        = $true
                environmentGroup = 'Finance'
                lastModifiedAt   = '2026-01-15T10:30:00Z'
            }
        }
    }

    # Agent fields: https://learn.microsoft.com/microsoft-copilot-studio/admin-agent-inventory
    # (core, entra, configuration, connector properties; the connector example shows
    # operationId, usedAs, isEnabled, requiresEndUserConsent, whenCanBeUsed,
    # connectionProvider, connectionIdSharedByMaker, createdBy).
    function New-MockAgentItem {
        param(
            [string]$Name = 'agent-1',
            [string]$EnvironmentId = 'env-1',
            [object]$LastPublishedAt = '2026-01-15T10:30:00Z',
            [int]$ListedConnectors = 1,
            [int]$CapabilityConnectors = 1
        )

        $connectors = foreach ($n in 1..$ListedConnectors) {
            [pscustomobject]@{
                connectorId = "shared_connector$n"
                operations  = @(
                    [pscustomobject]@{
                        operationId               = 'RunScriptProd'
                        usedAs                    = 'Tool'
                        isEnabled                 = $true
                        requiresEndUserConsent    = $false
                        whenCanBeUsed             = 'ViaDirectReferenceOnly'
                        connectionProvider        = 'Maker'
                        connectionIdSharedByMaker = '11112222-3333-4444-5555-666677778888'
                        createdBy                 = '52bff06b-5db5-42cd-9919-28f95e3c07af'
                    }
                )
            }
        }

        [pscustomobject]@{
            id         = "/providers/Microsoft.PowerPlatform/agents/$Name"
            name       = $Name
            type       = 'microsoft.copilotstudio/agents'
            location   = 'unitedstates'
            tenantId   = '00000000-0000-0000-0000-000000000000'
            properties = [pscustomobject]@{
                harness                 = 'Standard'
                displayName             = 'Customer support agent'
                name                    = $Name
                botId                   = $Name
                createdAt               = '2024-12-13T04:00:00Z'
                createdBy               = 'aaaa0000-bb11-2222-33cc-444444dddddd'
                ownerId                 = 'aaaa0000-bb11-2222-33cc-444444dddddd'
                environmentId           = $EnvironmentId
                lastPublishedAt         = $LastPublishedAt
                createdIn               = 'Copilot Studio'
                isQuarantined           = $false
                isManaged               = $false
                orchestration           = 'Generative'
                model                   = 'gpt-4o'
                authentication          = 'Microsoft Entra'
                channels                = @('Teams', 'SharePoint')
                sharedWithViewers       = [pscustomobject]@{ groupCount = 2; userCount = 7; entireTenant = $false }
                sharedWithEditors       = [pscustomobject]@{ groupCount = 0; userCount = 3 }
                capabilitiesCounts      = [pscustomobject]@{
                    distinctPowerPlatformConnectors           = $CapabilityConnectors
                    distinctPowerPlatformConnectorsOperations = $CapabilityConnectors
                }
                powerPlatformConnectors = @($connectors)
                IsWebSearchEnabledForKnowledge = $true
            }
        }
    }

    # Dataverse Web API collection: { "value": [ ... ], "@odata.nextLink": "..." }
    # https://learn.microsoft.com/power-apps/developer/data-platform/webapi/query/page-results
    # Row columns are those named on the Agent Inventory data source page:
    # https://learn.microsoft.com/microsoft-copilot-studio/guidance/kit-agent-inventory-data-source
    function New-DataverseResponse {
        param([object[]]$Value = @(), [string]$NextLink = '')

        $response = [ordered]@{ value = @($Value) }
        if ($NextLink) { $response['@odata.nextLink'] = $NextLink }
        return [pscustomobject]$response
    }

    function New-MockBot {
        [pscustomobject]@{
            botid             = 'bot-1'
            name              = 'Customer support agent'
            createdon         = '2024-12-13T04:00:00Z'
            modifiedon        = '2026-02-01T08:00:00Z'
            _modifiedby_value = 'aaaa0000-bb11-2222-33cc-444444dddddd'
            published         = '2026-01-15T10:30:00Z'
        }
    }

    function New-MockBotComponent {
        param([string]$Id = 'component-1', [string]$Data = 'kind: TaskDialog')

        [pscustomobject]@{
            botcomponentid     = $Id
            name               = 'Lookup order'
            componenttype      = 9
            data               = $Data
            _parentbotid_value = 'bot-1'
        }
    }

    # Unified audit log: https://learn.microsoft.com/microsoft-copilot-studio/admin-logging-copilot-studio
    # (Operation labels and the common and Copilot Studio audit fields). The detail comes
    # back as a JSON string in AuditData.
    function New-MockAuditRecord {
        param([string]$Id = 'event-1', [string]$Operation = 'BotUpdateOperation-BotPublish', [object]$ResultCount = $null)

        $auditData = @{
            CreationTime     = '2026-08-10T12:00:00'
            Id               = $Id
            Operation        = $Operation
            UserKey          = 'aaaa0000-bb11-2222-33cc-444444dddddd'
            ResultStatus     = 'Success'
            BotId            = 'bot-1'
            BotSchemaName    = 'cr5e3_agentName'
            BotComponentId   = 'component-1'
            BotComponentType = 'Topic'
        } | ConvertTo-Json -Compress

        $record = [ordered]@{ RecordType = 'PowerPlatformAdministratorActivity'; AuditData = $auditData }
        if ($null -ne $ResultCount) { $record['ResultCount'] = $ResultCount }
        return [pscustomobject]$record
    }
}

AfterAll {
    Remove-Variable -Name CsCalls, CsSessions, CsItem -Scope Global -ErrorAction SilentlyContinue
    Remove-Module M365ReportLibrary -Force -ErrorAction SilentlyContinue
}

Describe 'Collector output matches the committed sample files' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
        Mock Connect-AzAccount -MockWith { }
        Mock Get-AzAccessToken -MockWith { [pscustomobject]@{ Token = (ConvertTo-SecureString 'fake-token' -AsPlainText -Force) } }
        Mock Disconnect-ExchangeOnline -MockWith { }
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It '<Csv> columns match the sample and the schema' -ForEach @(
        @{ Csv = 'environments.csv'; Script = 'Get-PowerPlatformEnvironments.ps1'; Key = 'Environments' }
        @{ Csv = 'agents.csv'; Script = 'Get-CopilotStudioAgents.ps1'; Key = 'Agents' }
        @{ Csv = 'agent-connectors.csv'; Script = 'Get-AgentConnectors.ps1'; Key = 'AgentConnectors' }
    ) {
        Mock Invoke-RestMethod -MockWith {
            New-InventoryResponse -Data @((New-MockEnvironmentItem), (New-MockAgentItem))
        }

        Invoke-CollectorScript $Script @{ OutputPath = $script:folder }

        $produced = Join-Path $script:folder $Csv
        @(Import-Csv -LiteralPath $produced).Count | Should -BeGreaterThan 0
        Get-HeaderText -Path $produced | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples $Csv))
        Get-HeaderText -Path $produced | Should -Be ($script:Schema[$Key] -join ',')
    }

    It '<Csv> columns match the sample and the schema' -ForEach @(
        @{ Csv = 'agent-components.csv'; Script = 'Get-AgentComponents.ps1'; Key = 'AgentComponents'; Value = { New-MockBotComponent } }
        @{ Csv = 'agent-modifications.csv'; Script = 'Get-AgentModifications.ps1'; Key = 'AgentModifications'; Value = { New-MockBot } }
    ) {
        $global:CsItem = & $Value
        Mock Invoke-RestMethod -MockWith { New-DataverseResponse -Value @($global:CsItem) }

        Invoke-CollectorScript $Script @{ OutputPath = $script:folder; DataverseUrl = 'https://org.crm.example.com' }

        $produced = Join-Path $script:folder $Csv
        @(Import-Csv -LiteralPath $produced).Count | Should -BeGreaterThan 0
        Get-HeaderText -Path $produced | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples $Csv))
        Get-HeaderText -Path $produced | Should -Be ($script:Schema[$Key] -join ',')
    }

    It 'agent-audit-events.csv' {
        Mock Search-UnifiedAuditLog -MockWith { New-MockAuditRecord }

        Invoke-CollectorScript 'Get-AgentAuditEvents.ps1' @{
            OutputPath = $script:folder
            StartDate  = [datetime]'2026-08-10T00:00:00Z'
            EndDate    = [datetime]'2026-08-11T00:00:00Z'
        }

        $produced = Join-Path $script:folder 'agent-audit-events.csv'
        @(Import-Csv -LiteralPath $produced).Count | Should -BeGreaterThan 0
        Get-HeaderText -Path $produced | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples 'agent-audit-events.csv'))
        Get-HeaderText -Path $produced | Should -Be ($script:Schema.AgentAuditEvents -join ',')
    }

    It 'every sample file has the schema columns, and the header-only samples are header-only' {
        foreach ($pair in @(
                @{ Csv = 'environments.csv'; Key = 'Environments' }
                @{ Csv = 'agents.csv'; Key = 'Agents' }
                @{ Csv = 'agent-connectors.csv'; Key = 'AgentConnectors' }
                @{ Csv = 'agent-components.csv'; Key = 'AgentComponents' }
                @{ Csv = 'agent-modifications.csv'; Key = 'AgentModifications' }
                @{ Csv = 'agent-audit-events.csv'; Key = 'AgentAuditEvents' }
            )) {
            Get-HeaderText -Path (Join-Path $script:Samples $pair.Csv) | Should -Be ($script:Schema[$pair.Key] -join ',')
        }
        foreach ($cloud in 'gcc', 'gcchigh') {
            $path = Join-Path $script:Samples "$cloud/agent-connectors.csv"
            Get-HeaderText -Path $path | Should -Be ($script:Schema.AgentConnectors -join ',')
            @(Get-Content -LiteralPath $path).Count | Should -Be 1
        }
    }
}

Describe 'Power Platform inventory collectors' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-AzAccount -MockWith { }
        Mock Get-AzAccessToken -MockWith { [pscustomobject]@{ Token = (ConvertTo-SecureString 'fake-token' -AsPlainText -Force) } }
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'signs in as a user with Connect-AzAccount and takes a token for the Power Platform API' {
        Mock Invoke-RestMethod -MockWith { New-InventoryResponse -Data @((New-MockAgentItem)) }

        Invoke-CollectorScript 'Get-CopilotStudioAgents.ps1' @{ OutputPath = $script:folder }

        Should -Invoke Connect-AzAccount -Times 1 -Exactly -ParameterFilter { $Environment -eq 'AzureCloud' }
        Should -Invoke Get-AzAccessToken -Times 1 -Exactly -ParameterFilter { $ResourceUrl -eq 'https://api.powerplatform.com' }
    }

    It 'signs in to the Azure Government cloud for GCC High and does not guess an API host' {
        Mock Invoke-RestMethod -MockWith { New-InventoryResponse -Data @((New-MockAgentItem)) }

        Invoke-CollectorScript 'Get-CopilotStudioAgents.ps1' @{
            OutputPath = $script:folder; Environment = 'GCCHigh'; ApiHost = 'https://api.example.us/'
            WarningAction = 'SilentlyContinue'
        }

        Should -Invoke Connect-AzAccount -Times 1 -Exactly -ParameterFilter { $Environment -eq 'AzureUSGovernment' }
        Should -Invoke Invoke-RestMethod -Times 1 -Exactly -ParameterFilter { $Uri -like 'https://api.example.us/resourcequery/*' }
    }

    It 'writes the header only in GCC when no -ApiHost is given, because no host is documented' {
        Mock Invoke-RestMethod -MockWith { throw 'must not be called' }

        Invoke-CollectorScript 'Get-CopilotStudioAgents.ps1' @{
            OutputPath = $script:folder; Environment = 'GCC'; WarningAction = 'SilentlyContinue'
        }

        Should -Not -Invoke Invoke-RestMethod
        @(Get-Content -LiteralPath (Join-Path $script:folder 'agents.csv')).Count | Should -Be 1
        Get-Content -LiteralPath (Join-Path $script:folder 'run.log') -Raw | Should -Match 'ApiHost'
    }

    It 'POSTs a PowerPlatformResources query for the agents type, with api-version 2024-10-01' {
        Mock Invoke-RestMethod -MockWith { New-InventoryResponse -Data @((New-MockAgentItem)) }

        Invoke-CollectorScript 'Get-CopilotStudioAgents.ps1' @{ OutputPath = $script:folder }

        Should -Invoke Invoke-RestMethod -Times 1 -Exactly -ParameterFilter {
            $Method -eq 'Post' -and
            $Uri -eq 'https://api.powerplatform.com/resourcequery/resources/query?api-version=2024-10-01' -and
            $Headers.Authorization -eq 'Bearer fake-token' -and
            $Body -match '"TableName":\s*"PowerPlatformResources"' -and
            $Body -match "microsoft.copilotstudio/agents"
        }
    }

    It 'follows SkipToken while the inventory returns one, and sends it back in Options' {
        $global:CsCalls = 0
        Mock Invoke-RestMethod -MockWith {
            $global:CsCalls++
            switch ($global:CsCalls) {
                1 { New-InventoryResponse -Data @((New-MockAgentItem -Name 'agent-1')) -SkipToken 'token-1' -ResultTruncated 1 }
                2 { New-InventoryResponse -Data @((New-MockAgentItem -Name 'agent-2')) -SkipToken 'token-2' -ResultTruncated 1 }
                default { New-InventoryResponse -Data @((New-MockAgentItem -Name 'agent-3')) -ResultTruncated 0 }
            }
        }

        Invoke-CollectorScript 'Get-CopilotStudioAgents.ps1' @{ OutputPath = $script:folder }

        Should -Invoke Invoke-RestMethod -Times 3 -Exactly
        Should -Invoke Invoke-RestMethod -Times 1 -Exactly -ParameterFilter { $Body -match '"SkipToken":\s*"token-1"' }
        Should -Invoke Invoke-RestMethod -Times 1 -Exactly -ParameterFilter { $Body -match '"SkipToken":\s*"token-2"' }
        @(Import-Csv -LiteralPath (Join-Path $script:folder 'agents.csv')).AgentId | Should -Be @('agent-1', 'agent-2', 'agent-3')
    }

    It 'stops when the same SkipToken comes back twice' {
        Mock Invoke-RestMethod -MockWith { New-InventoryResponse -Data @((New-MockAgentItem)) -SkipToken 'stuck' -ResultTruncated 1 }

        Invoke-CollectorScript 'Get-CopilotStudioAgents.ps1' @{ OutputPath = $script:folder; WarningAction = 'SilentlyContinue' }

        Should -Invoke Invoke-RestMethod -Times 2 -Exactly
        @(Get-Content -LiteralPath (Join-Path $script:folder 'agents.csv')).Count | Should -Be 1
        Get-Content -LiteralPath (Join-Path $script:folder 'run.log') -Raw | Should -Match 'same skipToken twice'
    }

    It 'writes both the listed connector count and capabilitiesCounts, and warns when they differ' {
        Mock Invoke-RestMethod -MockWith {
            New-InventoryResponse -Data @((New-MockAgentItem -ListedConnectors 2 -CapabilityConnectors 450))
        }

        Invoke-CollectorScript 'Get-CopilotStudioAgents.ps1' @{ OutputPath = $script:folder; WarningAction = 'SilentlyContinue' }

        $row = Import-Csv -LiteralPath (Join-Path $script:folder 'agents.csv')
        $row.ListedConnectorCount | Should -Be '2'
        $row.ListedConnectorOperationCount | Should -Be '2'
        $row.CapabilitiesDistinctConnectors | Should -Be '450'
        $row.CapabilitiesDistinctConnectorOperations | Should -Be '450'
        Get-Content -LiteralPath (Join-Path $script:folder 'run.log') -Raw | Should -Match 'random 200'
    }

    It 'reads ownership, sharing, authentication and publish state from the inventory fields' {
        Mock Invoke-RestMethod -MockWith { New-InventoryResponse -Data @((New-MockAgentItem -Name 'bot-9' -EnvironmentId 'env-7')) }

        Invoke-CollectorScript 'Get-CopilotStudioAgents.ps1' @{ OutputPath = $script:folder }

        $row = Import-Csv -LiteralPath (Join-Path $script:folder 'agents.csv')
        $row.EnvironmentId | Should -Be 'env-7'
        $row.AgentId | Should -Be 'bot-9'
        $row.OwnerId | Should -Be 'aaaa0000-bb11-2222-33cc-444444dddddd'
        $row.Authentication | Should -Be 'Microsoft Entra'
        $row.Channels | Should -Be 'Teams;SharePoint'
        $row.ViewerUserCount | Should -Be '7'
        $row.ViewerGroupCount | Should -Be '2'
        $row.ViewerEntireTenant | Should -Be 'False'
        $row.EditorUserCount | Should -Be '3'
        $row.LastPublishedAt | Should -Be '2026-01-15T10:30:00Z'
        $row.IsPublished | Should -Be 'True'
    }

    It 'leaves LastPublishedAt empty and IsPublished False for a draft agent' {
        Mock Invoke-RestMethod -MockWith { New-InventoryResponse -Data @((New-MockAgentItem -LastPublishedAt $null)) }

        Invoke-CollectorScript 'Get-CopilotStudioAgents.ps1' @{ OutputPath = $script:folder }

        $row = Import-Csv -LiteralPath (Join-Path $script:folder 'agents.csv')
        $row.LastPublishedAt | Should -Be ''
        $row.IsPublished | Should -Be 'False'
    }

    It 'writes the header only and logs the AI Reader role when the inventory returns 403' {
        Mock Invoke-RestMethod -MockWith { throw 'Response status code does not indicate success: 403 (Forbidden).' }

        Invoke-CollectorScript 'Get-CopilotStudioAgents.ps1' @{ OutputPath = $script:folder; WarningAction = 'SilentlyContinue' }

        @(Get-Content -LiteralPath (Join-Path $script:folder 'agents.csv')).Count | Should -Be 1
        Get-Content -LiteralPath (Join-Path $script:folder 'run.log') -Raw | Should -Match 'AI Reader'
    }

    It 'writes environments with their type and group' {
        Mock Invoke-RestMethod -MockWith { New-InventoryResponse -Data @((New-MockEnvironmentItem -Id 'env-5')) }

        Invoke-CollectorScript 'Get-PowerPlatformEnvironments.ps1' @{ OutputPath = $script:folder }

        $row = Import-Csv -LiteralPath (Join-Path $script:folder 'environments.csv')
        $row.EnvironmentId | Should -Be 'env-5'
        $row.EnvironmentType | Should -Be 'Production'
        $row.IsManaged | Should -Be 'True'
        $row.EnvironmentGroup | Should -Be 'Finance'
    }

    It 'flattens one row per connector operation, with usedAs' {
        Mock Invoke-RestMethod -MockWith { New-InventoryResponse -Data @((New-MockAgentItem -ListedConnectors 2)) }

        Invoke-CollectorScript 'Get-AgentConnectors.ps1' @{ OutputPath = $script:folder }

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:folder 'agent-connectors.csv'))
        $rows.Count | Should -Be 2
        $rows[0].ConnectorId | Should -Be 'shared_connector1'
        $rows[0].OperationId | Should -Be 'RunScriptProd'
        $rows[0].UsedAs | Should -Be 'Tool'
        $rows[0].ConnectionProvider | Should -Be 'Maker'
        $rows[0].WhenCanBeUsed | Should -Be 'ViaDirectReferenceOnly'
    }

    It 'skips the connector source in <Cloud>: header only, no sign-in, no query, and the reason is logged' -ForEach @(
        @{ Cloud = 'GCC' }
        @{ Cloud = 'GCCHigh' }
    ) {
        Mock Invoke-RestMethod -MockWith { throw 'must not be called' }

        Invoke-CollectorScript 'Get-AgentConnectors.ps1' @{
            OutputPath = $script:folder; Environment = $Cloud; ApiHost = 'https://api.example.us'
            WarningAction = 'SilentlyContinue'
        }

        Should -Not -Invoke Invoke-RestMethod
        Should -Not -Invoke Connect-AzAccount
        $path = Join-Path $script:folder 'agent-connectors.csv'
        @(Get-Content -LiteralPath $path).Count | Should -Be 1
        Get-HeaderText -Path $path | Should -Be ($script:Schema.AgentConnectors -join ',')
        Get-Content -LiteralPath (Join-Path $script:folder 'run.log') -Raw | Should -Match 'unavailable in'
    }

    It 'does not skip the agents source in <Cloud>' -ForEach @(
        @{ Cloud = 'GCC' }
        @{ Cloud = 'GCCHigh' }
    ) {
        Mock Invoke-RestMethod -MockWith { New-InventoryResponse -Data @((New-MockAgentItem)) }

        Invoke-CollectorScript 'Get-CopilotStudioAgents.ps1' @{
            OutputPath = $script:folder; Environment = $Cloud; ApiHost = 'https://api.example.us'
            WarningAction = 'SilentlyContinue'
        }

        @(Import-Csv -LiteralPath (Join-Path $script:folder 'agents.csv')).Count | Should -Be 1
    }

    It 'appends a second snapshot with the same columns' {
        Mock Invoke-RestMethod -MockWith { New-InventoryResponse -Data @((New-MockAgentItem)) }

        Invoke-CollectorScript 'Get-CopilotStudioAgents.ps1' @{ OutputPath = $script:folder }
        Invoke-CollectorScript 'Get-CopilotStudioAgents.ps1' @{ OutputPath = $script:folder }

        # Same RunDate and key, so the second run is skipped rather than duplicated.
        @(Import-Csv -LiteralPath (Join-Path $script:folder 'agents.csv')).Count | Should -Be 1
    }
}

Describe 'Dataverse collectors' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-AzAccount -MockWith { }
        Mock Get-AzAccessToken -MockWith { [pscustomobject]@{ Token = (ConvertTo-SecureString 'fake-token' -AsPlainText -Force) } }
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'writes the header only and says so when no -DataverseUrl is given' {
        Mock Invoke-RestMethod -MockWith { throw 'must not be called' }

        Invoke-CollectorScript 'Get-AgentComponents.ps1' @{ OutputPath = $script:folder; WarningAction = 'SilentlyContinue' }

        Should -Not -Invoke Invoke-RestMethod
        @(Get-Content -LiteralPath (Join-Path $script:folder 'agent-components.csv')).Count | Should -Be 1
        Get-Content -LiteralPath (Join-Path $script:folder 'run.log') -Raw | Should -Match 'DataverseUrl'
    }

    It 'GETs from the environment with a token scoped to that environment' {
        Mock Invoke-RestMethod -MockWith { New-DataverseResponse -Value @((New-MockBot)) }

        Invoke-CollectorScript 'Get-AgentModifications.ps1' @{
            OutputPath = $script:folder; DataverseUrl = 'https://org.crm.example.com/'
        }

        Should -Invoke Get-AzAccessToken -Times 1 -Exactly -ParameterFilter { $ResourceUrl -eq 'https://org.crm.example.com/' }
        Should -Invoke Invoke-RestMethod -Times 1 -Exactly -ParameterFilter {
            $Method -eq 'Get' -and $Uri -like 'https://org.crm.example.com/api/data/v9.1/bots?*modifiedon*' -and
            $Headers.Authorization -eq 'Bearer fake-token'
        }
    }

    It 'reads the last modified date, the modifier and the published date from bot' {
        Mock Invoke-RestMethod -MockWith { New-DataverseResponse -Value @((New-MockBot)) }

        Invoke-CollectorScript 'Get-AgentModifications.ps1' @{
            OutputPath = $script:folder; DataverseUrl = 'https://org.crm.example.com'
        }

        $row = Import-Csv -LiteralPath (Join-Path $script:folder 'agent-modifications.csv')
        $row.AgentId | Should -Be 'bot-1'
        $row.ModifiedOn | Should -Be '2026-02-01T08:00:00Z'
        $row.ModifiedBy | Should -Be 'aaaa0000-bb11-2222-33cc-444444dddddd'
        $row.PublishedOn | Should -Be '2026-01-15T10:30:00Z'
    }

    It 'follows @odata.nextLink' {
        $global:CsCalls = 0
        Mock Invoke-RestMethod -MockWith {
            $global:CsCalls++
            if ($global:CsCalls -eq 1) {
                New-DataverseResponse -Value @((New-MockBotComponent -Id 'c1')) -NextLink 'https://org.crm.example.com/api/data/v9.1/botcomponents?$skiptoken=2'
            }
            else {
                New-DataverseResponse -Value @((New-MockBotComponent -Id 'c2'))
            }
        }

        Invoke-CollectorScript 'Get-AgentComponents.ps1' @{
            OutputPath = $script:folder; DataverseUrl = 'https://org.crm.example.com'
        }

        Should -Invoke Invoke-RestMethod -Times 2 -Exactly
        @(Import-Csv -LiteralPath (Join-Path $script:folder 'agent-components.csv')).ComponentId | Should -Be @('c1', 'c2')
    }

    It 'classifies components by the documented markers and writes a row per category' {
        Mock Invoke-RestMethod -MockWith {
            New-DataverseResponse -Value @(
                (New-MockBotComponent -Id 'tool' -Data 'kind: TaskDialog')
                (New-MockBotComponent -Id 'both' -Data "HttpRequestAction`nInvokeAIBuilderModelAction")
                (New-MockBotComponent -Id 'mcp' -Data 'kind: InvokeExternalAgentTaskAction')
                (New-MockBotComponent -Id 'file' -Data 'FileDataName: policy.pdf')
                (New-MockBotComponent -Id 'plain' -Data 'kind: AdaptiveDialog')
            )
        }

        Invoke-CollectorScript 'Get-AgentComponents.ps1' @{
            OutputPath = $script:folder; DataverseUrl = 'https://org.crm.example.com'
        }

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:folder 'agent-components.csv'))
        ($rows | Where-Object ComponentId -eq 'tool').Category | Should -Be 'Tool'
        @(($rows | Where-Object ComponentId -eq 'both').Category | Sort-Object) | Should -Be @('HttpRequest', 'Prompt')
        ($rows | Where-Object ComponentId -eq 'mcp').Category | Should -Be 'Mcp'
        ($rows | Where-Object ComponentId -eq 'file').Category | Should -Be 'KnowledgeSource'
        $rows.ComponentId | Should -Not -Contain 'plain'
        ($rows | Where-Object ComponentId -eq 'tool').AgentId | Should -Be 'bot-1'
    }

    It 'logs the unverified GCC High tables, attempts the read, and warns' {
        Mock Invoke-RestMethod -MockWith { New-DataverseResponse -Value @((New-MockBot)) }

        Invoke-CollectorScript 'Get-AgentModifications.ps1' @{
            OutputPath = $script:folder; Environment = 'GCCHigh'; DataverseUrl = 'https://org.crm.example.us'
            WarningAction = 'SilentlyContinue'
        }

        Should -Invoke Invoke-RestMethod -Times 1 -Exactly
        Get-Content -LiteralPath (Join-Path $script:folder 'run.log') -Raw | Should -Match 'UNVERIFIED'
        @(Import-Csv -LiteralPath (Join-Path $script:folder 'agent-modifications.csv')).Count | Should -Be 1
    }

    It 'skips an environment it cannot read, keeps the others, and logs it' {
        $global:CsCalls = 0
        Mock Invoke-RestMethod -MockWith {
            $global:CsCalls++
            if ($global:CsCalls -eq 1) { throw 'Forbidden: 403' }
            New-DataverseResponse -Value @((New-MockBot))
        }

        Invoke-CollectorScript 'Get-AgentModifications.ps1' @{
            OutputPath = $script:folder
            DataverseUrl = @('https://one.crm.example.com', 'https://two.crm.example.com')
            WarningAction = 'SilentlyContinue'
        }

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:folder 'agent-modifications.csv'))
        $rows.Count | Should -Be 1
        $rows[0].DataverseUrl | Should -Be 'https://two.crm.example.com'
        Get-Content -LiteralPath (Join-Path $script:folder 'run.log') -Raw | Should -Match 'one.crm.example.com'
    }

    It 'writes the header only when every environment fails' {
        Mock Invoke-RestMethod -MockWith { throw 'Forbidden: 403' }

        Invoke-CollectorScript 'Get-AgentComponents.ps1' @{
            OutputPath = $script:folder; DataverseUrl = 'https://one.crm.example.com'; WarningAction = 'SilentlyContinue'
        }

        @(Get-Content -LiteralPath (Join-Path $script:folder 'agent-components.csv')).Count | Should -Be 1
    }
}

Describe 'Get-AgentAuditEvents.ps1' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
        Mock Disconnect-ExchangeOnline -MockWith { }
        $script:range = @{
            StartDate = [datetime]'2026-08-10T00:00:00Z'
            EndDate   = [datetime]'2026-08-11T00:00:00Z'
        }
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'searches by -Operations with the authoring labels, never by RecordType, in a ReturnLargeSet session' {
        Mock Search-UnifiedAuditLog -MockWith { New-MockAuditRecord }

        Invoke-CollectorScript 'Get-AgentAuditEvents.ps1' (@{ OutputPath = $script:folder } + $script:range)

        Should -Invoke Search-UnifiedAuditLog -Times 1 -Exactly -ParameterFilter {
            $SessionCommand -eq 'ReturnLargeSet' -and -not [string]::IsNullOrEmpty($SessionId) -and
            $ResultSize -eq 5000 -and -not $PSBoundParameters.ContainsKey('RecordType') -and
            $Operations -ccontains 'BotCreate' -and $Operations -ccontains 'BotUpdateOperation-BotPublish' -and
            $Operations -ccontains 'BotUpdateOperation-BotShare' -and $Operations.Count -eq 21
        }
    }

    It 'does not use the connection pattern for delegated sign-in: it uses Exchange Online' {
        Mock Search-UnifiedAuditLog -MockWith { New-MockAuditRecord }

        Invoke-CollectorScript 'Get-AgentAuditEvents.ps1' (@{ OutputPath = $script:folder } + $script:range)

        Should -Invoke Connect-M365Service -Times 1 -Exactly -ParameterFilter { $Service -eq 'ExchangeOnline' }
        Should -Invoke Disconnect-ExchangeOnline -Times 1 -Exactly
    }

    It 'reads the operation, actor and agent out of AuditData' {
        Mock Search-UnifiedAuditLog -MockWith { New-MockAuditRecord -Id 'event-9' -Operation 'BotCreate' }

        Invoke-CollectorScript 'Get-AgentAuditEvents.ps1' (@{ OutputPath = $script:folder } + $script:range)

        $row = Import-Csv -LiteralPath (Join-Path $script:folder 'agent-audit-events.csv')
        $row.Id | Should -Be 'event-9'
        $row.Operation | Should -Be 'BotCreate'
        $row.UserId | Should -Be 'aaaa0000-bb11-2222-33cc-444444dddddd'
        $row.BotId | Should -Be 'bot-1'
        $row.BotSchemaName | Should -Be 'cr5e3_agentName'
        $row.CreationTime | Should -Be '2026-08-10T12:00:00Z'
    }

    It 'keeps paging in the same session until a page comes back empty' {
        $global:CsCalls = 0
        Mock Search-UnifiedAuditLog -MockWith {
            $global:CsCalls++
            switch ($global:CsCalls) {
                1 { New-MockAuditRecord -Id 'a' -ResultCount 10000 }
                2 { New-MockAuditRecord -Id 'b' -ResultCount 10000 }
                default { $null }
            }
        }

        Invoke-CollectorScript 'Get-AgentAuditEvents.ps1' (@{ OutputPath = $script:folder } + $script:range)

        Should -Invoke Search-UnifiedAuditLog -Times 3 -Exactly
        @(Import-Csv -LiteralPath (Join-Path $script:folder 'agent-audit-events.csv')).Id | Should -Be @('a', 'b')
    }

    It 'repeats the same SessionId across pages' {
        $global:CsSessions = [System.Collections.Generic.List[string]]::new()
        $global:CsCalls = 0
        Mock Search-UnifiedAuditLog -MockWith {
            $global:CsSessions.Add($SessionId)
            $global:CsCalls++
            if ($global:CsCalls -le 2) { New-MockAuditRecord -Id "p$($global:CsCalls)" -ResultCount 10000 } else { $null }
        }

        Invoke-CollectorScript 'Get-AgentAuditEvents.ps1' (@{ OutputPath = $script:folder } + $script:range)

        @($global:CsSessions | Select-Object -Unique).Count | Should -Be 1
    }

    It 'stops at the 50,000-record session cap, writes nothing from that window, and names the window to re-run' {
        Mock Search-UnifiedAuditLog -MockWith { New-MockAuditRecord -ResultCount 60000 }

        { Invoke-CollectorScript 'Get-AgentAuditEvents.ps1' (@{ OutputPath = $script:folder; WarningAction = 'SilentlyContinue' } + $script:range) } |
            Should -Throw '*50,000*'

        @(Get-Content -LiteralPath (Join-Path $script:folder 'agent-audit-events.csv')).Count | Should -Be 1
        $log = Get-Content -LiteralPath (Join-Path $script:folder 'run.log') -Raw
        $log | Should -Match '-StartDate 2026-08-10T00:00:00Z -EndDate 2026-08-11T00:00:00Z'
    }

    It 'stops paging once 50,000 records have been read in a session' {
        $global:CsCalls = 0
        Mock Search-UnifiedAuditLog -MockWith {
            $global:CsCalls++
            1..5000 | ForEach-Object { New-MockAuditRecord -Id "r$($global:CsCalls)-$_" }
        }

        { Invoke-CollectorScript 'Get-AgentAuditEvents.ps1' (@{ OutputPath = $script:folder; WarningAction = 'SilentlyContinue' } + $script:range) } |
            Should -Throw '*50,000*'

        Should -Invoke Search-UnifiedAuditLog -Times 10 -Exactly
    }

    It 'starts at the latest CreationTime already collected' {
        Export-AppendCsv -Path (Join-Path $script:folder 'agent-audit-events.csv') -Column $script:Schema.AgentAuditEvents -Rows @(
            [pscustomobject]@{
                CreationTime = '2026-08-09T05:00:00Z'; Id = 'old'; Operation = 'BotCreate'; UserId = 'u'
                ResultStatus = 'Success'; BotId = 'b'; BotSchemaName = ''; BotComponentId = ''; BotComponentType = ''
            }
        )
        Mock Search-UnifiedAuditLog -MockWith { $null }

        Invoke-CollectorScript 'Get-AgentAuditEvents.ps1' @{ OutputPath = $script:folder; EndDate = [datetime]'2026-08-09T06:00:00Z' }

        Should -Invoke Search-UnifiedAuditLog -ParameterFilter { $StartDate.ToUniversalTime() -eq [datetime]'2026-08-09T05:00:00Z' }
    }

    It 'writes the header only and names the Audit Reader role when the search is refused' {
        Mock Search-UnifiedAuditLog -MockWith { throw 'The term is not recognized' }

        Invoke-CollectorScript 'Get-AgentAuditEvents.ps1' (@{ OutputPath = $script:folder; WarningAction = 'SilentlyContinue' } + $script:range)

        @(Get-Content -LiteralPath (Join-Path $script:folder 'agent-audit-events.csv')).Count | Should -Be 1
        Get-Content -LiteralPath (Join-Path $script:folder 'run.log') -Raw | Should -Match 'Audit Reader'
    }

    It 'attempts GCC High with an UNVERIFIED warning instead of skipping' {
        Mock Search-UnifiedAuditLog -MockWith { New-MockAuditRecord }

        Invoke-CollectorScript 'Get-AgentAuditEvents.ps1' (@{ OutputPath = $script:folder; Environment = 'GCCHigh'; WarningAction = 'SilentlyContinue' } + $script:range)

        Should -Invoke Search-UnifiedAuditLog -Times 1 -Exactly
        Get-Content -LiteralPath (Join-Path $script:folder 'run.log') -Raw | Should -Match 'UNVERIFIED'
    }
}

Describe 'Source availability' {
    It 'marks only the connector source NotAvailable, and only in GCC and GCC High' {
        foreach ($source in $script:Schema.SourceAvailability.Keys) {
            foreach ($cloud in 'Commercial', 'GCC', 'GCCHigh') {
                $status = $script:Schema.SourceAvailability[$source][$cloud].Status
                $status | Should -BeIn @('Available', 'NotAvailable', 'Unverified')
                $script:Schema.SourceAvailability[$source][$cloud].Reference | Should -Match '^https://learn\.microsoft\.com/'
                if ($status -eq 'NotAvailable') {
                    $source | Should -Be 'AgentConnectors'
                    $cloud | Should -BeIn @('GCC', 'GCCHigh')
                }
            }
        }
    }

    It 'marks the Dataverse and audit sources UNVERIFIED in GCC and GCC High, as the contract doc does' {
        foreach ($source in 'AgentComponents', 'AgentModifications', 'AgentAuditEvents') {
            $script:Schema.SourceAvailability[$source].GCC.Status | Should -Be 'Unverified'
            $script:Schema.SourceAvailability[$source].GCCHigh.Status | Should -Be 'Unverified'
        }
    }
}

Describe 'Run-All.ps1' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
        Mock Connect-AzAccount -MockWith { }
        Mock Get-AzAccessToken -MockWith { [pscustomobject]@{ Token = (ConvertTo-SecureString 'fake-token' -AsPlainText -Force) } }
        Mock Disconnect-ExchangeOnline -MockWith { }
        Mock Search-UnifiedAuditLog -MockWith { New-MockAuditRecord }
        Mock Invoke-RestMethod -MockWith {
            if ($Uri -like '*resourcequery*') { New-InventoryResponse -Data @((New-MockEnvironmentItem), (New-MockAgentItem)) }
            elseif ($Uri -like '*botcomponents*') { New-DataverseResponse -Value @((New-MockBotComponent)) }
            else { New-DataverseResponse -Value @((New-MockBot)) }
        }
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'produces all six CSVs' {
        Invoke-CollectorScript 'Run-All.ps1' @{
            OutputPath = $script:folder; DataverseUrl = 'https://org.crm.example.com'
            StartDate = [datetime]'2026-08-10T00:00:00Z'; EndDate = [datetime]'2026-08-11T00:00:00Z'
            WarningAction = 'SilentlyContinue'
        }

        foreach ($csv in 'environments', 'agents', 'agent-connectors', 'agent-components', 'agent-modifications', 'agent-audit-events') {
            Test-Path -LiteralPath (Join-Path $script:folder "$csv.csv") | Should -BeTrue
        }
    }
}
