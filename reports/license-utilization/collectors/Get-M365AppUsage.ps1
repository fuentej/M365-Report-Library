#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes usage-m365-apps.csv: which Microsoft 365 apps and platforms each user has
        used (Windows, Mac, mobile and web; Outlook, Word, Excel, PowerPoint, OneNote and
        Teams). A snapshot of a rolling period, appended each run.

    .DESCRIPTION
        Source 5f: Microsoft Graph getM365AppUserDetail
        (https://learn.microsoft.com/graph/api/reportroot-getm365appuserdetail), read with
        a GET to /v1.0/reports/getM365AppUserDetail(period='D30')?$format=application/json.
        The JSON form is paged: the default page is 200 and one page is not the full set,
        so @odata.nextLink is followed until it is absent. A next link on another host is
        refused. Each user's details entry is written as one row, so the rows match the
        report's CSV columns.

        The columns hold true or false for whether the user used that app or platform in
        the period. Last Activity Date is the most recent activity whatever the period.

        User names are replaced by hashes when the organization setting that conceals
        user, group and site names is on; see report-settings.csv and the README.

        Not available in GCC High (US Government L4): the header is written and the
        reason logged. Needs Reports.Read.All; a delegated caller also needs a limited
        admin role such as Reports Reader.

    .PARAMETER Period
        The report period. Defaults to D30.

    .EXAMPLE
        ./Get-M365AppUsage.ps1 -OutputPath ./out
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

    [ValidateSet('D7', 'D30', 'D90', 'D180')]
    [string]$Period = 'D30',

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


Invoke-LicenseUsageReport -Source 'M365AppUsage' -CsvName 'usage-m365-apps.csv' -Json `
    -Description 'Microsoft 365 apps usage report' -ReportPath "reports/getM365AppUserDetail(period='$Period')" `
    -OutputPath $OutputPath -Environment $Environment -AppId $AppId `
    -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization `
    -SchemaPath $SchemaPath -SkipConnect:$SkipConnect
