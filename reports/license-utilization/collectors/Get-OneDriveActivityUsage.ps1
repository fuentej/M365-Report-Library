#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes usage-onedrive-activity.csv: OneDrive activity per user (files viewed or edited, synced, shared internally and externally, last activity). A snapshot of a rolling period, appended each run.

    .DESCRIPTION
        Source 5e: Microsoft Graph getOneDriveActivityUserDetail
        (https://learn.microsoft.com/graph/api/reportroot-getonedriveactivityuserdetail),
        read with a GET to /v1.0/reports/getOneDriveActivityUserDetail(period='D30'). Graph answers with a 302 to a
        preauthenticated CSV download that is valid for a few minutes, so the file is
        read at once, and removed.

        The counts aggregate the period (D7, D30, D90 or D180). Last Activity Date is the
        most recent intentional activity whatever the period, so a snapshot can show
        inactivity older than the period. Reports usually appear within 24 to 72 hours
        and sometimes take days, and a deleted user's row leaves the report within 30
        days: a missing row is not a zero-activity row
        (https://learn.microsoft.com/microsoft-365/admin/activity-reports/activity-reports).

        User names are replaced by hashes when the organization setting that conceals
        user, group and site names is on; see report-settings.csv and the README.

        Not available in GCC High (US Government L4): the header is written and the
        reason logged. Needs Reports.Read.All; a delegated caller also needs a limited
        admin role such as Reports Reader. Global Reader and Usage Summary Reports Reader
        do not receive the detail rows.

    .PARAMETER Period
        The report period. Defaults to D30.

    .EXAMPLE
        ./Get-OneDriveActivityUsage.ps1 -OutputPath ./out
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


Invoke-LicenseUsageReport -Source 'OneDriveActivityUsage' -CsvName 'usage-onedrive-activity.csv' `
    -Description 'OneDrive activity report' -ReportPath "reports/getOneDriveActivityUserDetail(period='$Period')" -ReportPeriod $Period `
    -OutputPath $OutputPath -Environment $Environment -AppId $AppId `
    -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization `
    -SchemaPath $SchemaPath -SkipConnect:$SkipConnect
