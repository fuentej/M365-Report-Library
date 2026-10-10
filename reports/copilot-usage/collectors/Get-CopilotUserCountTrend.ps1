#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes copilot-user-count-trend.csv: the daily trend of Copilot enabled and active users per app, one row per day, re-read and stamped with the run date each run.

    .DESCRIPTION
        Source 4 of docs/candidates/copilot-usage.md: GET /copilot/reports/getMicrosoft365CopilotUserCountTrend(period='D28')
        (https://learn.microsoft.com/microsoft-365-copilot/extensibility/api/admin-settings/reports/copilotreportroot-getmicrosoft365copilotusercounttrend).
        Re-reading a period appends the same Report Date again under a new RunDate, because a
        recent day can change between runs. The daily Prompts submitted column is not the
        summary's total or average.

        Period is one of D7, D28, D90, D180 or ALL (the four periods in one response, not
        lifetime). D30 is a v1 value and is refused on the default v2. The report is a
        window, not an event stream, so each run appends a snapshot stamped with the run date.

        NotAvailable in GCC High (the Learn page marks US Government L4 as not supported),
        where this collector writes the header only and logs why. Needs Reports.Read.All;
        a delegated sign-in also needs a role such as Reports Reader or AI Administrator.
        The report typically becomes available within 48 to 72 hours of the end of the UTC
        day (the usage page says 48, the Copilot reports overview says 72), so the newest
        days are incomplete and a later run can change them.

    .EXAMPLE
        ./Get-CopilotUserCountTrend.ps1 -OutputPath ./out
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

    [ValidateSet('D7', 'D28', 'D90', 'D180', 'ALL')]
    [string]$Period = 'D28',

    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Imported by path without -Force: a forced re-import builds a fresh module session
# state, which would discard anything the caller has arranged around the module - and
# that is exactly how the tests put mocks in front of the tenant cmdlets.
Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

. (Join-Path $PSScriptRoot 'CopilotUsageHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'CopilotUsageSchema.psd1')
$columns = Get-CopilotColumn -Schema $schema -MapKey 'UserCountTrendMap'
$source = 'copilot-user-count-trend'
$csvPath = Join-Path $OutputPath 'copilot-user-count-trend.csv'

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

if (Test-CopilotSourceSkipped -Source 'UserCountTrend' -LogSource $source -CsvPath $csvPath -Column $columns `
        -OutputPath $OutputPath -Environment $Environment -Schema $schema) {
    return
}

if (-not $SkipConnect) {
    Connect-M365Service -Service Graph -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization `
        -Scopes @('Reports.Read.All')
}

Invoke-CopilotReportCollector -OutputPath $OutputPath -LogSource $source -CsvName 'copilot-user-count-trend.csv' -Schema $schema `
    -MapKey 'UserCountTrendMap' -KeyColumn @('RunDate', 'ReportDate', 'ReportPeriod') `
    -ReportName 'Copilot user count trend' -Uri "v1.0/copilot/reports/getMicrosoft365CopilotUserCountTrend(period='$Period')"
