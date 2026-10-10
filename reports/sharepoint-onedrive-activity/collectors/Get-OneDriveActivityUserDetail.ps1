#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes onedrive-activity-user-detail.csv: files viewed or edited, synced and shared per OneDrive user, appended each run.

    .DESCRIPTION
        Source 6: Microsoft Graph getOneDriveActivityUserDetail
        (https://learn.microsoft.com/graph/api/reportroot-getonedriveactivityuserdetail), read with
        Get-MgReportOneDriveActivityUserDetail -OutFile.
        Exactly one of -Period and -Date is sent; see Get-SharePointActivityUserDetail.ps1 for the date
        window. The columns are those of the SharePoint report without Visited Page Count. Last
        Activity Date is for the selected date range
        (https://learn.microsoft.com/microsoft-365/admin/activity-reports/onedrive-for-business-activity).

        Not available in GCC High (US Government L4 is marked unsupported on the API page). A GCC High
        run writes the header only and logs why. Needs Reports.Read.All; a delegated caller also needs
        a limited admin role such as Reports Reader. Global Reader and Usage Summary Reports Reader do
        not receive the detail rows. When the tenant conceals names (see report-settings.csv), the
        identifier columns hold concealed values and cannot be joined to drive-quota.csv or users.csv.

    .EXAMPLE
        ./Get-OneDriveActivityUserDetail.ps1 -OutputPath ./out
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

    # A single day, within the last 30 days. When set it is sent instead of -Period.
    [datetime]$Date,

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
$columns = $schema.OneDriveActivityUserDetail
$source = 'onedrive-activity-user-detail'
$csvPath = Join-Path $OutputPath 'onedrive-activity-user-detail.csv'

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

if (Test-SharePointSourceSkipped -Source 'OneDriveActivityUserDetail' -LogSource $source -CsvPath $csvPath -Column $columns `
        -OutputPath $OutputPath -Environment $Environment -Schema $schema) {
    return
}

# The day asked for, or empty when the period form is used. Exactly one of -Date and -Period is sent.
# The date form is the past 30 days. An out-of-range day throws here, before a failed call is
# logged as a refused report.
$activityDay = $null
if ($PSBoundParameters.ContainsKey('Date')) { $activityDay = Get-ActivityReportDay -Value $Date }
$queryDate = if ($null -ne $activityDay) { $activityDay.ToString('yyyy-MM-dd') } else { '' }

if (-not $SkipConnect) {
    Connect-M365Service -Service Graph -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization
}

function Read-Column { param($Row, [string[]]$Header) Get-ReportColumnValue -Row $Row -Header $Header }

$map = {
    param($row, $runDate)
    [pscustomobject]@{
        RunDate                    = $runDate
        QueryDate                  = $queryDate
        ReportRefreshDate          = Read-Column $row 'Report Refresh Date'
        UserPrincipalName          = Read-Column $row 'User Principal Name'
        IsDeleted                  = Read-Column $row 'Is Deleted'
        DeletedDate                = Read-Column $row 'Deleted Date'
        LastActivityDate           = Read-Column $row 'Last Activity Date'
        ViewedOrEditedFileCount    = Read-Column $row 'Viewed Or Edited File Count'
        SyncedFileCount            = Read-Column $row 'Synced File Count'
        SharedInternallyFileCount  = Read-Column $row 'Shared Internally File Count'
        SharedExternallyFileCount  = Read-Column $row 'Shared Externally File Count'
        AssignedProducts           = Read-Column $row 'Assigned Products'
        ReportPeriod               = Read-Column $row 'Report Period'
    }
}

Invoke-UsageReportCollector -OutputPath $OutputPath -LogSource $source -CsvName 'onedrive-activity-user-detail.csv' -Column $columns `
    -KeyColumn @('RunDate', 'QueryDate', 'UserPrincipalName', 'ReportPeriod') -ReportName 'OneDrive activity' -MapRow $map -Fetch {
        param($file)
        if ($queryDate) {
            Get-MgReportOneDriveActivityUserDetail -Date $activityDay -OutFile $file -ErrorAction Stop
        }
        else {
            Get-MgReportOneDriveActivityUserDetail -Period $Period -OutFile $file -ErrorAction Stop
        }
    }
