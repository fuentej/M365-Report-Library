#Requires -Version 7.0

<#
    .SYNOPSIS
        Generates the fake Copilot Studio agents data in ./samples.

    .DESCRIPTION
        The sample set lets the later Power BI report, and anyone reading this repository,
        work against realistic shapes without a tenant. Nothing here comes from a real
        directory: every address and URL is under example.com, which RFC 2606 reserves
        for documentation.

        The generator is deterministic. The same -Seed and -EndDate always produce the
        same files.

        Every file is written through Export-AppendCsv with the column list the collectors
        use, so a sample file cannot drift from its collector's output. ./samples/gcc and
        ./samples/gcchigh hold the header-only agent-connectors.csv that the collector
        writes there, where Microsoft documents connector inventory as not available.

    .PARAMETER EndDate
        The "now" the sample set is generated around. Fixed by default to keep the
        committed files stable.

    .EXAMPLE
        ./New-SampleData.ps1
#>
[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path $PSScriptRoot 'samples'),

    [datetime]$EndDate = [datetime]::new(2026, 9, 1, 0, 0, 0, [System.DateTimeKind]::Utc),

    [int]$Seed = 20260901
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot '../../shared/M365ReportLibrary.psm1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'collectors/CopilotStudioSchema.psd1')

$EndDate = $EndDate.ToUniversalTime()
$random = [System.Random]::new($Seed)

# Three snapshots a month apart, the last one on EndDate.
$snapshotDates = 2..0 | ForEach-Object { $EndDate.AddMonths(-$_) }

function New-DeterministicGuid {
    $bytes = [byte[]]::new(16)
    $random.NextBytes($bytes)
    return [guid]::new($bytes).ToString()
}

function Get-RandomItem {
    param([Parameter(Mandatory)][object[]]$Items)
    return $Items[$random.Next(0, $Items.Count)]
}

function Format-Stamp {
    param([AllowNull()][object]$Value)
    return ConvertTo-CsvTimestamp $Value
}

if (Test-Path -LiteralPath $OutputPath) {
    Get-ChildItem -LiteralPath $OutputPath -Recurse -File -Filter '*.csv' | Remove-Item -Force
}
New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null

$environmentNames = @(
    @{ Name = 'Contoso (default)'; Type = 'Default'; Group = '' }
    @{ Name = 'Contoso Production'; Type = 'Production'; Group = 'Production' }
    @{ Name = 'Contoso Finance'; Type = 'Production'; Group = 'Finance' }
    @{ Name = 'Contoso Sandbox'; Type = 'Sandbox'; Group = '' }
    @{ Name = 'Avery Abara personal'; Type = 'Developer'; Group = '' }
    @{ Name = 'Pilot Trial'; Type = 'Trial'; Group = '' }
)
$environments = foreach ($definition in $environmentNames) {
    [pscustomobject]@{
        Id       = New-DeterministicGuid
        Name     = $definition.Name
        Type     = $definition.Type
        Group    = $definition.Group
        Modified = $EndDate.AddDays(-$random.Next(1, 120))
    }
}
$environments = @($environments)

$owners = 1..10 | ForEach-Object { New-DeterministicGuid }
$agentNames = @(
    'HR Policy Helper', 'IT Service Desk', 'Expense Assistant', 'Sales Playbook Coach', 'Facilities Booking'
    'Onboarding Guide', 'Legal Intake', 'Procurement Advisor', 'Field Service Triage', 'Travel Concierge'
    'Benefits FAQ', 'Compliance Checker'
)
$channelSets = @('Teams', 'Teams;SharePoint', 'Microsoft Copilot', 'Teams;Microsoft Copilot', 'Website', '')
$connectorCatalog = @(
    @{ Id = 'shared_sharepointonline'; Operations = @('GetItems', 'PostItem') }
    @{ Id = 'shared_office365users'; Operations = @('MyProfile_V2') }
    @{ Id = 'shared_excelonlinebusiness'; Operations = @('AddRowV2', 'RunScriptProd') }
    @{ Id = 'shared_servicenow'; Operations = @('CreateRecord') }
)

$agents = foreach ($i in 0..($agentNames.Count - 1)) {
    $published = ($i % 4) -ne 3
    $connectorSet = @($connectorCatalog | Select-Object -First ($i % 4 + 1))
    [pscustomobject]@{
        Id            = New-DeterministicGuid
        EnvironmentId = $environments[$i % $environments.Count].Id
        Name          = $agentNames[$i]
        Created       = $EndDate.AddDays(-$random.Next(60, 500))
        Published     = if ($published) { $EndDate.AddDays(-$random.Next(1, 50)) } else { $null }
        Owner         = Get-RandomItem $owners
        Channels      = $channelSets[$i % $channelSets.Count]
        Authentication = @('Microsoft Entra', 'None', 'Generic OAuth 2.0')[$i % 3]
        Orchestration = @('Generative', 'Classic')[$i % 2]
        Connectors    = $connectorSet
        EntireTenant  = (($i % 5) -eq 4)
        Viewers       = $random.Next(0, 40)
        Editors       = $random.Next(0, 5)
    }
}
$agents = @($agents)

foreach ($snapshot in $snapshotDates) {
    $runDate = $snapshot.ToString('yyyy-MM-dd')

    $environmentRows = foreach ($environment in $environments) {
        [pscustomobject]@{
            RunDate          = $runDate
            EnvironmentId    = $environment.Id
            DisplayName      = $environment.Name
            Location         = 'unitedstates'
            EnvironmentType  = $environment.Type
            IsManaged        = ($environment.Type -eq 'Production').ToString()
            EnvironmentGroup = $environment.Group
            LastModifiedAt   = Format-Stamp $environment.Modified
        }
    }
    Export-AppendCsv -Path (Join-Path $OutputPath 'environments.csv') -Rows @($environmentRows) `
        -Column $schema.Environments -KeyColumn @('RunDate', 'EnvironmentId')

    # An agent created after this snapshot does not exist in it yet.
    $present = @($agents | Where-Object { $_.Created -le $snapshot })

    $agentRows = foreach ($agent in $present) {
        $operationCount = ($agent.Connectors | ForEach-Object { $_.Operations.Count } | Measure-Object -Sum).Sum
        [pscustomobject]@{
            RunDate                                 = $runDate
            EnvironmentId                           = $agent.EnvironmentId
            AgentId                                 = $agent.Id
            DisplayName                             = $agent.Name
            Harness                                 = 'Standard'
            CreatedIn                               = 'Copilot Studio'
            CreatedAt                               = Format-Stamp $agent.Created
            CreatedBy                               = $agent.Owner
            OwnerId                                 = $agent.Owner
            LastPublishedAt                         = if ($agent.Published -and $agent.Published -le $snapshot) { Format-Stamp $agent.Published } else { '' }
            IsPublished                             = [bool]($agent.Published -and $agent.Published -le $snapshot) -as [string]
            IsQuarantined                           = 'False'
            IsManaged                               = 'False'
            Orchestration                           = $agent.Orchestration
            Model                                   = 'gpt-4o'
            Authentication                          = $agent.Authentication
            Channels                                = $agent.Channels
            ViewerUserCount                         = [string]$agent.Viewers
            ViewerGroupCount                        = [string]($agent.Viewers % 3)
            ViewerEntireTenant                      = $agent.EntireTenant.ToString()
            EditorUserCount                         = [string]$agent.Editors
            EditorGroupCount                        = '0'
            ListedConnectorCount                    = [string]$agent.Connectors.Count
            ListedConnectorOperationCount           = [string]$operationCount
            CapabilitiesDistinctConnectors          = [string]$agent.Connectors.Count
            CapabilitiesDistinctConnectorOperations = [string]$operationCount
            IsWebSearchEnabledForKnowledge          = ($agent.Orchestration -eq 'Generative').ToString()
        }
    }
    Export-AppendCsv -Path (Join-Path $OutputPath 'agents.csv') -Rows @($agentRows) `
        -Column $schema.Agents -KeyColumn @('RunDate', 'EnvironmentId', 'AgentId')

    $connectorRows = foreach ($agent in $present) {
        foreach ($connector in $agent.Connectors) {
            foreach ($operationId in $connector.Operations) {
                [pscustomobject]@{
                    RunDate                = $runDate
                    EnvironmentId          = $agent.EnvironmentId
                    AgentId                = $agent.Id
                    ConnectorId            = $connector.Id
                    OperationId            = $operationId
                    UsedAs                 = @('Tool', 'Topic Tool', 'Knowledge')[$operationId.Length % 3]
                    IsEnabled              = 'True'
                    RequiresEndUserConsent = 'False'
                    WhenCanBeUsed          = 'Anytime'
                    ConnectionProvider     = @('Maker', 'User')[$operationId.Length % 2]
                }
            }
        }
    }
    Export-AppendCsv -Path (Join-Path $OutputPath 'agent-connectors.csv') -Rows @($connectorRows) `
        -Column $schema.AgentConnectors -KeyColumn @('RunDate', 'EnvironmentId', 'AgentId', 'ConnectorId', 'OperationId', 'UsedAs')

    $dataverseUrl = 'https://org12345.crm.example.com'

    $categories = @('KnowledgeSource', 'Tool', 'HttpRequest', 'Prompt', 'Mcp')
    $componentRows = foreach ($agent in $present) {
        foreach ($category in ($categories | Select-Object -First ($agent.Name.Length % 5 + 1))) {
            [pscustomobject]@{
                RunDate       = $runDate
                DataverseUrl  = $dataverseUrl
                AgentId       = $agent.Id
                ComponentId   = New-DeterministicGuid
                ComponentName = "$category for $($agent.Name)"
                ComponentType = '9'
                Category      = $category
            }
        }
    }
    Export-AppendCsv -Path (Join-Path $OutputPath 'agent-components.csv') -Rows @($componentRows) `
        -Column $schema.AgentComponents -KeyColumn @('RunDate', 'DataverseUrl', 'ComponentId', 'Category')

    $modificationRows = foreach ($agent in $present) {
        $modified = $agent.Created.AddDays($agent.Name.Length * 3)
        if ($modified -gt $snapshot) { $modified = $snapshot.AddDays(-1) }
        [pscustomobject]@{
            RunDate      = $runDate
            DataverseUrl = $dataverseUrl
            AgentId      = $agent.Id
            Name         = $agent.Name
            CreatedOn    = Format-Stamp $agent.Created
            ModifiedOn   = Format-Stamp $modified
            ModifiedBy   = $agent.Owner
            PublishedOn  = if ($agent.Published -and $agent.Published -le $snapshot) { Format-Stamp $agent.Published } else { '' }
        }
    }
    Export-AppendCsv -Path (Join-Path $OutputPath 'agent-modifications.csv') -Rows @($modificationRows) `
        -Column $schema.AgentModifications -KeyColumn @('RunDate', 'DataverseUrl', 'AgentId')
}

# Audit events: a create, some edits and a publish per agent, inside the last 90 days.
$operationsForSample = @('BotCreate', 'BotComponentUpdate', 'BotUpdateOperation-BotPublish', 'BotUpdateOperation-BotShare', 'BotUpdateOperation-BotAuthUpdate')
$auditRows = foreach ($agent in $agents) {
    foreach ($operation in ($operationsForSample | Select-Object -First ($agent.Name.Length % 5 + 1))) {
        [pscustomobject]@{
            CreationTime     = Format-Stamp $EndDate.AddMinutes(-$random.Next(60, 60 * 24 * 90))
            Id               = New-DeterministicGuid
            Operation        = $operation
            UserId           = $agent.Owner
            ResultStatus     = 'Success'
            BotId            = $agent.Id
            BotSchemaName    = 'cr5e3_' + ($agent.Name -replace '[^A-Za-z0-9]', '')
            BotComponentId   = if ($operation -eq 'BotComponentUpdate') { New-DeterministicGuid } else { '' }
            BotComponentType = if ($operation -eq 'BotComponentUpdate') { 'Topic' } else { '' }
        }
    }
}
$auditRows = @($auditRows | Sort-Object CreationTime)
Export-AppendCsv -Path (Join-Path $OutputPath 'agent-audit-events.csv') -Rows $auditRows `
    -Column $schema.AgentAuditEvents -KeyColumn 'Id'

# GCC and GCC High: Microsoft documents connector inventory as not available, so the
# collector writes the header only.
foreach ($cloud in 'gcc', 'gcchigh') {
    Export-AppendCsv -Path (Join-Path $OutputPath "$cloud/agent-connectors.csv") -Column $schema.AgentConnectors
}
