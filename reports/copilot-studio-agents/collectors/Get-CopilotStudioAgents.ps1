#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes agents.csv: a snapshot of every Copilot Studio agent in the Power Platform
        inventory, appended each run.

    .DESCRIPTION
        Source: the Power Platform inventory API, type microsoft.copilotstudio/agents
        (https://learn.microsoft.com/power-platform/admin/inventory-api). Fields:
        https://learn.microsoft.com/microsoft-copilot-studio/admin-agent-inventory.
        The query follows SkipToken until no page returns one.

        SIGN-IN IS DELEGATED AND INTERACTIVE, as in Get-PowerPlatformEnvironments.ps1: the
        inventory returns HTTP 403 to service principals. Needs the AI Reader role.

        What the inventory does and does not show:
        * Draft agents are included. LastPublishedAt is empty for an agent never published.
        * For a published agent the inventory shows the published version; newer
          unpublished edits are withheld until it is published again.
        * V1 (Power Virtual Agents classic) agents are excluded.
        * When an agent has more than 200 resources of one type the inventory returns a
          random 200 of that type. Listed* columns count what came back; Capabilities*
          columns carry capabilitiesCounts, the complete count. Both are written.
        * channels, authentication, sharing counts, orchestration, model and
          capabilitiesCounts are Preview fields.
        * Viewer and editor identities are not available; only counts are written.
        * Changes appear after about 15 to 20 minutes.

        The Power Platform API base URL is documented for Commercial only. For GCC and
        GCC High pass -ApiHost (UNVERIFIED, never guessed). In GCC and GCC High only
        agents with the Standard harness appear.

    .PARAMETER ApiHost
        The Power Platform API base URL. Required for GCC and GCC High.

    .EXAMPLE
        ./Get-CopilotStudioAgents.ps1 -OutputPath ./out
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
$columns = $schema.Agents
$source = 'agents'
$csvPath = Join-Path $OutputPath 'agents.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-CopilotStudioSourceAvailability -Source 'Agents' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping agents.csv. $($availability.Reason) $($availability.Reference)")
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
    $connectors = @(Get-JsonValue $item 'properties.powerPlatformConnectors')
    $connectors = @($connectors | Where-Object { $null -ne $_ })
    $operationCount = 0
    foreach ($connector in $connectors) {
        $operationCount += @(Get-JsonValue $connector 'operations' | Where-Object { $null -ne $_ }).Count
    }

    $lastPublished = Get-JsonValue $item 'properties.lastPublishedAt'
    $agentId = Get-InventoryAgentId -Item $item

    [pscustomobject]@{
        RunDate                                 = $runDate
        EnvironmentId                           = [string](Get-JsonValue $item 'properties.environmentId')
        AgentId                                 = [string]$agentId
        DisplayName                             = [string](Get-JsonValue $item 'properties.displayName')
        Harness                                 = [string](Get-JsonValue $item 'properties.harness')
        CreatedIn                               = [string](Get-JsonValue $item 'properties.createdIn')
        CreatedAt                               = ConvertTo-CsvTimestamp (Get-JsonValue $item 'properties.createdAt')
        CreatedBy                               = [string](Get-JsonValue $item 'properties.createdBy')
        OwnerId                                 = [string](Get-JsonValue $item 'properties.ownerId')
        LastPublishedAt                         = ConvertTo-CsvTimestamp $lastPublished
        IsPublished                             = (-not [string]::IsNullOrWhiteSpace([string]$lastPublished)).ToString()
        IsQuarantined                           = Get-CsvBoolean (Get-JsonValue $item 'properties.isQuarantined')
        IsManaged                               = Get-CsvBoolean (Get-JsonValue $item 'properties.isManaged')
        Orchestration                           = [string](Get-JsonValue $item 'properties.orchestration')
        Model                                   = [string](Get-JsonValue $item 'properties.model')
        Authentication                          = [string](Get-JsonValue $item 'properties.authentication')
        Channels                                = Join-ListValue (Get-JsonValue $item 'properties.channels')
        ViewerUserCount                         = [string](Get-JsonValue $item 'properties.sharedWithViewers.userCount')
        ViewerGroupCount                        = [string](Get-JsonValue $item 'properties.sharedWithViewers.groupCount')
        ViewerEntireTenant                      = Get-CsvBoolean (Get-JsonValue $item 'properties.sharedWithViewers.entireTenant')
        EditorUserCount                         = [string](Get-JsonValue $item 'properties.sharedWithEditors.userCount')
        EditorGroupCount                        = [string](Get-JsonValue $item 'properties.sharedWithEditors.groupCount')
        ListedConnectorCount                    = [string]$connectors.Count
        ListedConnectorOperationCount           = [string]$operationCount
        CapabilitiesDistinctConnectors          = [string](Get-JsonValue $item 'properties.capabilitiesCounts.distinctPowerPlatformConnectors')
        CapabilitiesDistinctConnectorOperations = [string](Get-JsonValue $item 'properties.capabilitiesCounts.distinctPowerPlatformConnectorsOperations')
        IsWebSearchEnabledForKnowledge          = Get-CsvBoolean (Get-JsonValue $item 'properties.IsWebSearchEnabledForKnowledge')
    }
}

$rowList = @($rows)
$result = Export-AppendCsv -Path $csvPath -Rows $rowList -Column $columns -KeyColumn @('RunDate', 'EnvironmentId', 'AgentId') -PassThru
if (-not [string]::IsNullOrWhiteSpace($read.PagingError)) {
    Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
        'agents.csv: wrote {0} row(s) already read, then stopped. The snapshot is incomplete. {1}' -f $result.Written, $read.PagingError)
    throw $read.PagingError
}
Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'agents.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)

$partial = @($rowList | Where-Object {
        $_.CapabilitiesDistinctConnectors -ne '' -and [int]$_.ListedConnectorCount -lt [int]$_.CapabilitiesDistinctConnectors
    })
if ($partial.Count -gt 0) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        '{0} agent(s) list fewer connectors than capabilitiesCounts reports; the inventory returned a random 200 of that type for them.' -f $partial.Count)
}
