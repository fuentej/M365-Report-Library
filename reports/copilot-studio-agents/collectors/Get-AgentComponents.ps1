#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes agent-components.csv: the knowledge sources, tools, HTTP request actions,
        prompts and MCP actions in each Copilot Studio agent, read from Dataverse.

    .DESCRIPTION
        Source: the bot and botcomponent Dataverse tables in each environment. The
        markers searched for in botcomponent.data come from the Copilot Agent Kit page
        (https://learn.microsoft.com/microsoft-copilot-studio/guidance/kit-agent-inventory-data-source#agent-details-table):
        KnowledgeSourceConfiguration and FileDataName (KnowledgeSource), TaskDialog (Tool),
        HttpRequestAction (HttpRequest), InvokeAIBuilderModelAction (Prompt) and
        InvokeExternalAgentTaskAction (Mcp). A component with two markers gets two rows.

        Read with the Dataverse Web API (GET only), signed in as a user. The least
        privileged read-only Dataverse role for bot and botcomponent is UNVERIFIED;
        System Administrator and System Customizer are documented as able to read them.
        App-only access through an application user is not covered by any Learn page found.

        In GCC and GCC High it is UNVERIFIED that these tables hold the same columns, so
        the collector attempts them and logs a warning.

        The inventory does not expose each environment's Dataverse URL, so pass every URL
        to read with -DataverseUrl.

        Web API names this collector relies on that the cited pages do not spell out: the
        entity set names bots and botcomponents, and the lookup column
        _parentbotid_value. Treat a 404 or a missing-property error as that gap.

    .PARAMETER DataverseUrl
        One or more environment URLs, such as https://org.crm.dynamics.com.

    .EXAMPLE
        ./Get-AgentComponents.ps1 -OutputPath ./out -DataverseUrl https://org12345.crm.dynamics.com
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$OutputPath,

    [ValidateSet('Commercial', 'GCC', 'GCCHigh')]
    [string]$Environment = 'Commercial',

    [string[]]$DataverseUrl,
    [string]$TenantId,

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
$columns = $schema.AgentComponents
$source = 'agent-components'
$csvPath = Join-Path $OutputPath 'agent-components.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-CopilotStudioSourceAvailability -Source 'AgentComponents' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping agent-components.csv. $($availability.Reason) $($availability.Reference)")
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}
if ($availability.Status -eq 'Unverified') {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message $availability.Reason
}

if (-not $DataverseUrl) {
    Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
        'No -DataverseUrl was given. The inventory does not expose each environment''s Dataverse URL, so pass the URL of every environment to read. Writing the header only.')
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}

$rows = [System.Collections.Generic.List[object]]::new()
$failed = 0
$connected = $SkipConnect.IsPresent
foreach ($url in $DataverseUrl) {
    try {
        # The first sign-in is interactive; later environments reuse the session.
        $token = Get-DelegatedAccessToken -ResourceUrl $url -Environment $Environment -TenantId $TenantId -SkipConnect:$connected
        $connected = $true
        # Order by the primary key so a page boundary cannot repeat or skip a row.
        # https://learn.microsoft.com/power-apps/developer/data-platform/webapi/query/page-results
        $components = @(Invoke-DataverseQuery -DataverseUrl $url -Token $token `
                -Path 'botcomponents?$select=botcomponentid,name,componenttype,data,_parentbotid_value&$orderby=botcomponentid')

        foreach ($component in $components) {
            $data = [string](Get-JsonValue $component 'data')
            $categories = @($schema.ComponentMarkers |
                    Where-Object { $data.Contains($_.Marker) } |
                    ForEach-Object { $_.Category } | Select-Object -Unique)
            foreach ($category in $categories) {
                $rows.Add([pscustomobject]@{
                        RunDate       = $runDate
                        DataverseUrl  = $url
                        AgentId       = [string](Get-JsonValue $component '_parentbotid_value')
                        ComponentId   = [string](Get-JsonValue $component 'botcomponentid')
                        ComponentName = [string](Get-JsonValue $component 'name')
                        ComponentType = [string](Get-JsonValue $component 'componenttype')
                        Category      = $category
                    })
            }
        }
    }
    catch {
        # An off-host nextLink is not an environment the caller asked to skip.
        if ($_.Exception.Message -like '*different host*') { throw }
        $failed++
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            'Dataverse at {0} is unavailable to this sign-in ({1}). It needs a role that can read bot and botcomponent. Skipping this environment.' -f $url, $_.Exception.Message)
    }
}

if ($failed -eq @($DataverseUrl).Count) {
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}

$result = Export-AppendCsv -Path $csvPath -Rows $rows.ToArray() -Column $columns `
    -KeyColumn @('RunDate', 'DataverseUrl', 'ComponentId', 'Category') -PassThru
Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'agent-components.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
