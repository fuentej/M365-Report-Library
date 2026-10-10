#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes copilot-user-count-summary.csv: Copilot enabled and active users per app over a period, one row per period, appended each run.

    .DESCRIPTION
        Source 3 of docs/candidates/copilot-usage.md: GET /copilot/reports/getMicrosoft365CopilotUserCountSummary(period='D28')
        (https://learn.microsoft.com/microsoft-365-copilot/extensibility/api/admin-settings/reports/copilotreportroot-getmicrosoft365copilotusercountsummary).
        Enabled users are the unique users who held a Copilot licence over the period, not
        the licences assigned on the run date. Total prompts submitted and Average prompts
        submitted are Copilot Chat prompts.

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
        ./Get-CopilotUserCountSummary.ps1 -OutputPath ./out
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
$columns = Get-CopilotColumn -Schema $schema -MapKey 'UserCountSummaryMap'
$source = 'copilot-user-count-summary'
$csvPath = Join-Path $OutputPath 'copilot-user-count-summary.csv'

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

if (Test-CopilotSourceSkipped -Source 'UserCountSummary' -LogSource $source -CsvPath $csvPath -Column $columns `
        -OutputPath $OutputPath -Environment $Environment -Schema $schema) {
    return
}

if (-not $SkipConnect) {
    Connect-M365Service -Service Graph -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization `
        -Scopes @('Reports.Read.All')
}

Invoke-CopilotReportCollector -OutputPath $OutputPath -LogSource $source -CsvName 'copilot-user-count-summary.csv' -Schema $schema `
    -MapKey 'UserCountSummaryMap' -KeyColumn @('RunDate', 'ReportPeriod') `
    -ReportName 'Copilot user count summary' -Uri "v1.0/copilot/reports/getMicrosoft365CopilotUserCountSummary(period='$Period')"
