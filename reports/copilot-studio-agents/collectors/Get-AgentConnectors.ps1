#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes agent-connectors.csv: the connectors and operations each Copilot Studio agent
        uses, as a snapshot appended each run.

    .DESCRIPTION
        Source: properties.powerPlatformConnectors on the agent inventory records
        (https://learn.microsoft.com/microsoft-copilot-studio/admin-agent-inventory#connector-properties).
        UsedAs is Tool, Topic Tool or Knowledge. The connector fields are Preview.

        Connector inventory is documented as not available in GCC, GCC High or DoD
        (https://learn.microsoft.com/power-platform/admin/power-platform-inventory#known-limitations),
        so there this collector writes the header only and logs why.

        When an agent has more than 200 connectors the inventory returns a random 200;
        agents.csv holds the complete count in the Capabilities* columns.

        SIGN-IN IS DELEGATED AND INTERACTIVE (AI Reader role); service principals get
        HTTP 403. See Get-CopilotStudioAgents.ps1.

    .EXAMPLE
        ./Get-AgentConnectors.ps1 -OutputPath ./out
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$OutputPath,

    [ValidateSet('Commercial', 'GCC', 'GCCHigh')]
    [string]$Environment = 'Commercial',

    [string]$TenantId,
    [string]$ApiHost,

    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Imported by path without -Force: a forced re-import builds a fresh module session
# state, which would discard anything the caller has arranged around the module - and
# that is exactly how the tests put mocks in front of the tenant cmdlets.
Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

. (Join-Path $PSScriptRoot 'CopilotStudioHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'CopilotStudioSchema.psd1')
$columns = $schema.AgentConnectors
$source = 'agent-connectors'
$csvPath = Join-Path $OutputPath 'agent-connectors.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-CopilotStudioSourceAvailability -Source 'AgentConnectors' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping agent-connectors.csv. $($availability.Reason) $($availability.Reference)")
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}
if ($availability.Status -eq 'Unverified') {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message $availability.Reason
}

$apiBase = Get-PowerPlatformApiHost -Environment $Environment -ApiHost $ApiHost
if ($null -eq $apiBase) {
    Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
        "No Power Platform API base URL is documented for $Environment (UNVERIFIED). Pass -ApiHost. Writing the header only.")
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}

try {
    $token = Get-DelegatedAccessToken -ResourceUrl $apiBase -Environment $Environment -TenantId $TenantId -SkipConnect:$SkipConnect
    $read = Invoke-InventoryRead -ApiHost $apiBase -Token $token -OutputPath $OutputPath -Source $source `
        -Clauses (New-InventoryTypeClause -Type 'microsoft.copilotstudio/agents')
}
catch {
    Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
        'The Power Platform inventory is unavailable to this sign-in ({0}). It needs a signed-in user with the AI Reader role; service principals get HTTP 403. Writing the header only.' -f $_.Exception.Message)
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}

$items = @($read.Items)

$rows = foreach ($item in $items) {
    $agentId = Get-InventoryAgentId -Item $item
    $environmentId = [string](Get-JsonValue $item 'properties.environmentId')

    foreach ($connector in @(Get-JsonValue $item 'properties.powerPlatformConnectors' | Where-Object { $null -ne $_ })) {
        $connectorId = [string](Get-JsonValue $connector 'connectorId')
        # Tabular connectors (SharePoint, Dataverse, SQL, Excel) are returned with an
        # empty operations array. Dropping them would omit the connector id.
        # https://learn.microsoft.com/power-platform/admin/inventory-schema#known-limitations
        $operations = @(Get-JsonValue $connector 'operations' | Where-Object { $null -ne $_ })
        if ($operations.Count -eq 0) { $operations = @($null) }

        foreach ($operation in $operations) {
            [pscustomobject]@{
                RunDate                = $runDate
                EnvironmentId          = $environmentId
                AgentId                = [string]$agentId
                ConnectorId            = $connectorId
                OperationId            = [string](Get-JsonValue $operation 'operationId')
                UsedAs                 = [string](Get-JsonValue $operation 'usedAs')
                IsEnabled              = Get-CsvBoolean (Get-JsonValue $operation 'isEnabled')
                RequiresEndUserConsent = Get-CsvBoolean (Get-JsonValue $operation 'requiresEndUserConsent')
                WhenCanBeUsed          = [string](Get-JsonValue $operation 'whenCanBeUsed')
                ConnectionProvider     = [string](Get-JsonValue $operation 'connectionProvider')
            }
        }
    }
}

$result = Export-AppendCsv -Path $csvPath -Rows @($rows) -Column $columns `
    -KeyColumn @('RunDate', 'EnvironmentId', 'AgentId', 'ConnectorId', 'OperationId', 'UsedAs') -PassThru
if (-not [string]::IsNullOrWhiteSpace($read.PagingError)) {
    Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
        'agent-connectors.csv: wrote {0} row(s) already read, then stopped. The snapshot is incomplete. {1}' -f $result.Written, $read.PagingError)
    throw $read.PagingError
}
Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'agent-connectors.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
