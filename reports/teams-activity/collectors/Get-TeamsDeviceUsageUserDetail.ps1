#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes teams-device-usage-user-detail.csv: which platforms each user used Teams on for a rolling period, appended each run.

    .DESCRIPTION
        Source 3: Microsoft Graph getTeamsDeviceUsageUserDetail
        (https://learn.microsoft.com/graph/api/reportroot-getteamsdeviceusageuserdetail), read with
        Get-MgReportTeamDeviceUsageUserDetail -OutFile. The Used columns are yes/no for the period, not counts.
        Exactly one of -Period and -Date is sent. The date form is documented as available for the past 28 days on this API page.

        Not available in GCC High (US Government L4 is marked unsupported on the API page); a GCC
        High run writes the header only and logs why. GCC is UNVERIFIED and is attempted with a
        warning. Needs Reports.Read.All; a delegated caller also needs a limited admin role such as
        Reports Reader. Global Reader and Usage Summary Reports Reader do not receive the detail
        rows. When the tenant conceals names (see report-settings.csv), the user columns hold
        concealed identifiers and cannot be joined to the users collector. The CSV is a snapshot of
        a rolling period stamped with the run date, so a missing user is not a user with no activity.

    .EXAMPLE
        ./Get-TeamsDeviceUsageUserDetail.ps1 -OutputPath ./out
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

    # A single day. When set it is sent instead of -Period.
    [datetime]$Date,

    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Imported by path without -Force: a forced re-import builds a fresh module session
# state, which would discard anything the caller has arranged around the module - and
# that is exactly how the tests put mocks in front of the tenant cmdlets.
Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

. (Join-Path $PSScriptRoot 'TeamsActivityHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'TeamsActivitySchema.psd1')
$columns = $schema.TeamsDeviceUsageUserDetail
$source = 'teams-device-usage-user-detail'
$csvPath = Join-Path $OutputPath 'teams-device-usage-user-detail.csv'

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

if (Test-TeamsSourceSkipped -Source 'TeamsDeviceUsageUserDetail' -LogSource $source -CsvPath $csvPath -Column $columns `
        -OutputPath $OutputPath -Environment $Environment -Schema $schema) {
    return
}

if (-not $SkipConnect) {
    Connect-M365Service -Service Graph -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization
}

# The day asked for, or empty when the period form is used. Exactly one of -Date and -Period is sent.
$queryDate = if ($PSBoundParameters.ContainsKey('Date')) { $Date.ToString('yyyy-MM-dd') } else { '' }

function Read-Column { param($Row, [string[]]$Header) Get-ReportColumnValue -Row $Row -Header $Header }

$map = {
    param($row, $runDate)
    [pscustomobject]@{
        RunDate           = $runDate
        QueryDate         = $queryDate
        ReportRefreshDate = Read-Column $row 'Report Refresh Date'
        UserId            = Read-Column $row 'User Id'
        UserPrincipalName = Read-Column $row 'User Principal Name'
        LastActivityDate  = Read-Column $row 'Last Activity Date'
        IsDeleted         = Read-Column $row 'Is Deleted'
        DeletedDate       = Read-Column $row 'Deleted Date'
        UsedWeb           = Read-Column $row 'Used Web'
        UsedWindowsPhone  = Read-Column $row 'Used Windows Phone'
        UsediOS           = Read-Column $row 'Used iOS'
        UsedMac           = Read-Column $row 'Used Mac'
        UsedAndroidPhone  = Read-Column $row 'Used Android Phone'
        UsedWindows       = Read-Column $row 'Used Windows'
        UsedChromeOS      = Read-Column $row 'Used Chrome OS'
        UsedLinux         = Read-Column $row 'Used Linux'
        IsLicensed        = Read-Column $row 'Is Licensed'
        ReportPeriod      = Read-Column $row 'Report Period'
    }
}

Invoke-UsageReportCollector -OutputPath $OutputPath -LogSource $source -CsvName 'teams-device-usage-user-detail.csv' -Column $columns `
    -KeyColumn @('RunDate', 'QueryDate', 'UserPrincipalName', 'ReportPeriod') -ReportName 'Teams device usage' -MapRow $map -Fetch {
        param($file)
        if ($queryDate) {
            Get-MgReportTeamDeviceUsageUserDetail -Date $Date -OutFile $file -ErrorAction Stop
        }
        else {
            Get-MgReportTeamDeviceUsageUserDetail -Period $Period -OutFile $file -ErrorAction Stop
        }
    }
