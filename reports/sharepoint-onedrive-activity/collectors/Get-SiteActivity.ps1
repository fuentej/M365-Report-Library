#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes site-activity.csv: access, create, edit, delete and move counts per site per day, appended each run.

    .DESCRIPTION
        Source 13: Microsoft Graph
        GET /sites/{site-id}/getActivitiesByInterval(startDateTime,endDateTime,interval='day')
        (https://learn.microsoft.com/graph/api/itemactivitystat-getactivitybyinterval). The site
        list is GET /sites/getAllSites, followed by @odata.nextLink as returned. Each interval
        carries access, create, delete, edit and move with actionCount and actorCount
        (https://learn.microsoft.com/graph/api/resources/itemactivitystat). These are site action
        counts, not file counts and not per-user counts. Aggregates might not exist for every
        action type, so an action an interval does not carry is left empty, not zero. When an
        interval carries the incompleteData facet (https://learn.microsoft.com/graph/api/resources/incompletedata),
        IncompleteData is True and a zero is not "no activity".

        The API supports a range of less than 90 days for daily counts, so -LookbackDays is at most
        89. Each run asks for the last -LookbackDays whole days and appends them stamped with the run
        date. GET /drives/{id}/activities is not used: it supports no query parameters and states
        no retention.

        Needs the application permission Files.Read.All (delegated least privileged is Files.Read).
        Available in Commercial and GCC. In GCC High the Graph page marks the API available but says
        itemAnalytics "is not yet available in all national deployments", so it is UNVERIFIED there:
        it is attempted with a warning, and a site the service refuses is skipped.

    .PARAMETER SiteLimit
        Read only the first N sites. For a trial run.

    .EXAMPLE
        ./Get-SiteActivity.ps1 -OutputPath ./out
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$OutputPath,

    [ValidateSet('Commercial', 'GCC', 'GCCHigh')]
    [string]$Environment = 'Commercial',

    [string]$AppId,
    [string]$CertificateThumbprint,
    [string]$TenantId,
    [string]$Organization,

    # Fewer than 90 days: the API's limit for daily counts.
    [ValidateRange(1, 89)]
    [int]$LookbackDays = 30,

    [int]$SiteLimit = 0,

    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Imported by path without -Force: a forced re-import builds a fresh module session
# state, which would discard anything the caller has arranged around the module - and
# that is exactly how the tests put mocks in front of the tenant cmdlets.
Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

. (Join-Path $PSScriptRoot 'SharePointOneDriveHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'SharePointOneDriveSchema.psd1')
$columns = $schema.SiteActivity
$source = 'site-activity'
$csvPath = Join-Path $OutputPath 'site-activity.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

if (Test-SharePointSourceSkipped -Source 'SiteActivity' -LogSource $source -CsvPath $csvPath -Column $columns `
        -OutputPath $OutputPath -Environment $Environment -Schema $schema) {
    return
}

if (-not $SkipConnect) {
    Connect-M365Service -Service Graph -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization `
        -Scopes (@(Get-DefaultGraphScope) + 'Sites.Read.All', 'Files.Read.All')
}

# Whole UTC days: from LookbackDays ago up to today 00:00, so the range is under 90 days.
$end = [datetime]::UtcNow.Date
$start = $end.AddDays(-$LookbackDays)
$startText = $start.ToString('yyyy-MM-dd')
$endText = $end.ToString('yyyy-MM-dd')

$rows = [System.Collections.Generic.List[object]]::new()
$skipped = 0
try {
    $sites = @(Get-SiteRow -OutputPath $OutputPath -LogSource $source)
}
catch {
    if (Test-GraphThrottleStatus -ErrorRecord $_) { throw }
    Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
        'The site list is unavailable to this sign-in ({0}). getAllSites needs the application permission Sites.Read.All. Writing the header only.' -f $_.Exception.Message)
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}

$count = 0
foreach ($site in $sites) {
    if ($SiteLimit -gt 0 -and $count -ge $SiteLimit) { break }
    $count++
    $siteId = [string](Get-GraphJsonValue -Object $site -Name 'id')
    $siteUrl = [string](Get-GraphJsonValue -Object $site -Name 'webUrl')
    $uri = "/v1.0/sites/{0}/getActivitiesByInterval(startDateTime='{1}',endDateTime='{2}',interval='day')" -f [uri]::EscapeDataString($siteId), $startText, $endText
    try {
        foreach ($interval in Get-GraphPagedValue -Uri $uri -OutputPath $OutputPath -LogSource $source) {
            $rows.Add((ConvertTo-SiteActivityRow -Interval $interval -SiteId $siteId -SiteWebUrl $siteUrl -RunDate $runDate))
        }
    }
    catch {
        # A 429 or 503 that is still failing after the wait is not "this site has no activity".
        if (Test-GraphThrottleStatus -ErrorRecord $_) { throw }
        $skipped++
        Write-Verbose ('Activity of site {0} not read: {1}' -f $siteId, $_.Exception.Message)
    }
}

if ($skipped -gt 0) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        '{0} site(s) returned no activity statistics and were skipped. The call needs Files.Read.All, and itemAnalytics is not yet available in all national deployments.' -f $skipped)
}
if ($rows.Count -gt 0 -and @($rows | Where-Object { $_.IncompleteData -eq 'True' }).Count -gt 0) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        'Some intervals are based on incomplete data (IncompleteData is True). A zero there is not "no activity".')
}

$result = Export-AppendCsv -Path $csvPath -Rows $rows.ToArray() -Column $columns -KeyColumn @('RunDate', 'SiteId', 'IntervalStart') -PassThru
Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'site-activity.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
