#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes usage-copilot.csv: the last Microsoft 365 Copilot activity per user and app
        (Teams, Word, Excel, PowerPoint, Outlook, OneNote, Loop, Copilot Chat). A snapshot
        of a rolling period, appended each run.

    .DESCRIPTION
        Source 5g: the Microsoft 365 Copilot usage report
        (https://learn.microsoft.com/microsoft-365-copilot/extensibility/api/admin-settings/reports/copilotreportroot-getmicrosoft365copilotusageuserdetail),
        read with a GET to
        /v1.0/copilot/reports/getMicrosoft365CopilotUsageUserDetail(period='D28',version='v2').
        The beta /reports form is not supported for production applications and is not
        used. Version 2 is the default and the supported version; its periods are D7, D28,
        D90, D180 and ALL, and D30 is only in the v1 list, so it is not offered. ALL is the
        four periods in one response, not lifetime, and gives a row per user per period.
        v1.0 answers 200 with the CSV in the body, not a redirect, and is not paged.

        Only users with a Microsoft 365 Copilot licence are returned. Unlicensed Copilot
        Chat use is not in this API: use the admin center Copilot Chat Usage report or
        Purview audit.

        The page prints the v1 CSV header and names the version 2 additions (prompts
        submitted, active usage days, Copilot Chat work and web, Microsoft 365 Copilot,
        Edge and Copilot Agent last activity) in prose only. Those columns are matched by
        name and are UNVERIFIED: a header spelled differently leaves its cell empty.

        User names are replaced by hashes when the organization setting that conceals
        user, group and site names is on; see report-settings.csv and the README.

        Not available in GCC High (US Government L4): the header is written and the
        reason logged. Needs Reports.Read.All; a delegated caller also needs one of
        Company Administrator, AI Administrator, Exchange Administrator, SharePoint
        Administrator, Lync Administrator, Teams Service Administrator, Teams
        Communications Administrator or Reports Reader.

    .PARAMETER Period
        The report period. Defaults to D28.

    .EXAMPLE
        ./Get-CopilotUsage.ps1 -OutputPath ./out
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

    # An alternate schema file. The tests use it to exercise the NotAvailable path.
    [string]$SchemaPath = (Join-Path $PSScriptRoot 'LicenseUtilizationSchema.psd1'),

    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Imported by path without -Force: a forced re-import builds a fresh module session
# state, which would discard anything the caller has arranged around the module - and
# that is exactly how the tests put mocks in front of the tenant cmdlets.
Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

. (Join-Path $PSScriptRoot 'LicenseUtilizationHelpers.ps1')


Invoke-LicenseUsageReport -Source 'CopilotUsage' -CsvName 'usage-copilot.csv' `
    -Description 'Microsoft 365 Copilot usage report' `
    -ReportPath "copilot/reports/getMicrosoft365CopilotUsageUserDetail(period='$Period',version='v2')" `
    -OutputPath $OutputPath -Environment $Environment -AppId $AppId `
    -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization `
    -SchemaPath $SchemaPath -SkipConnect:$SkipConnect
