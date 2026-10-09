#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes agent-modifications.csv: when each Copilot Studio agent was last modified,
        by whom, and when it was last published, read from Dataverse.

    .DESCRIPTION
        Source: bot columns modifiedon, modifiedby and published
        (https://learn.microsoft.com/microsoft-copilot-studio/guidance/kit-agent-inventory-data-source#agent-details-table).
        The inventory exposes createdAt and lastPublishedAt but no modified field, so
        this is the only source of a last-modified date.

        Read with the Dataverse Web API (GET only), signed in as a user. The least
        privileged read-only Dataverse role is UNVERIFIED. In GCC and GCC High it is
        UNVERIFIED that the bot table holds the same columns; the collector attempts it
        and logs a warning.

        Web API names relied on that the cited page does not spell out: the entity set
        name bots and the lookup column _modifiedby_value.

    .PARAMETER DataverseUrl
        One or more environment URLs, such as https://org.crm.dynamics.com.

    .EXAMPLE
        ./Get-AgentModifications.ps1 -OutputPath ./out -DataverseUrl https://org12345.crm.dynamics.com
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
$columns = $schema.AgentModifications
$source = 'agent-modifications'
$csvPath = Join-Path $OutputPath 'agent-modifications.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-CopilotStudioSourceAvailability -Source 'AgentModifications' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping agent-modifications.csv. $($availability.Reason) $($availability.Reference)")
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
        $bots = @(Invoke-DataverseQuery -DataverseUrl $url -Token $token `
                -Path 'bots?$select=botid,name,createdon,modifiedon,_modifiedby_value,published&$orderby=botid')

        foreach ($bot in $bots) {
            $rows.Add([pscustomobject]@{
                    RunDate      = $runDate
                    DataverseUrl = $url
                    AgentId      = [string](Get-JsonValue $bot 'botid')
                    Name         = [string](Get-JsonValue $bot 'name')
                    CreatedOn    = ConvertTo-CsvTimestamp (Get-JsonValue $bot 'createdon')
                    ModifiedOn   = ConvertTo-CsvTimestamp (Get-JsonValue $bot 'modifiedon')
                    ModifiedBy   = [string](Get-JsonValue $bot '_modifiedby_value')
                    PublishedOn  = ConvertTo-CsvTimestamp (Get-JsonValue $bot 'published')
                })
        }
    }
    catch {
        # An off-host nextLink is not an environment the caller asked to skip.
        if ($_.Exception.Message -like '*different host*') { throw }
        $failed++
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            'Dataverse at {0} is unavailable to this sign-in ({1}). It needs a role that can read bot. Skipping this environment.' -f $url, $_.Exception.Message)
    }
}

if ($failed -eq @($DataverseUrl).Count) {
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}

$result = Export-AppendCsv -Path $csvPath -Rows $rows.ToArray() -Column $columns `
    -KeyColumn @('RunDate', 'DataverseUrl', 'AgentId') -PassThru
Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'agent-modifications.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
