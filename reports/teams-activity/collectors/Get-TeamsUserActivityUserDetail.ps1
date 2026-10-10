#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes teams-user-activity-user-detail.csv: chat, call and meeting counts per user for a rolling period, appended each run.

    .DESCRIPTION
        Source 1: Microsoft Graph getTeamsUserActivityUserDetail
        (https://learn.microsoft.com/graph/api/reportroot-getteamsuseractivityuserdetail), read with
        Get-MgReportTeamUserActivityUserDetail -OutFile. Team Chat Message Count already includes posts and
        replies; Meeting Count equals Meetings Attended Count and is being phased out. Do not add them.
        Exactly one of -Period and -Date is sent. The date form is documented as available for the past 30 days; the page does not say whether the counts are single-day values, so use the period form for totals.

        Not available in GCC High (US Government L4 is marked unsupported on the API page); a GCC
        High run writes the header only and logs why. GCC is UNVERIFIED and is attempted with a
        warning. Needs Reports.Read.All; a delegated caller also needs a limited admin role such as
        Reports Reader. Global Reader and Usage Summary Reports Reader do not receive the detail
        rows. When the tenant conceals names (see report-settings.csv), the user columns hold
        concealed identifiers and cannot be joined to the users collector. The CSV is a snapshot of
        a rolling period stamped with the run date, so a missing user is not a user with no activity.

    .EXAMPLE
        ./Get-TeamsUserActivityUserDetail.ps1 -OutputPath ./out
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
$columns = $schema.TeamsUserActivityUserDetail
$source = 'teams-user-activity-user-detail'
$csvPath = Join-Path $OutputPath 'teams-user-activity-user-detail.csv'

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

if (Test-TeamsSourceSkipped -Source 'TeamsUserActivityUserDetail' -LogSource $source -CsvPath $csvPath -Column $columns `
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
        RunDate                                  = $runDate
        QueryDate                                = $queryDate
        ReportRefreshDate                        = Read-Column $row 'Report Refresh Date'
        UserId                                   = Read-Column $row 'User Id'
        UserPrincipalName                        = Read-Column $row 'User Principal Name'
        LastActivityDate                         = Read-Column $row 'Last Activity Date'
        IsDeleted                                = Read-Column $row 'Is Deleted'
        DeletedDate                              = Read-Column $row 'Deleted Date'
        AssignedProducts                         = Read-Column $row 'Assigned Products'
        TeamChatMessageCount                     = Read-Column $row 'Team Chat Message Count'
        PrivateChatMessageCount                  = Read-Column $row 'Private Chat Message Count'
        CallCount                                = Read-Column $row 'Call Count'
        MeetingCount                             = Read-Column $row 'Meeting Count'
        PostMessages                             = Read-Column $row 'Post Messages'
        ReplyMessages                            = Read-Column $row 'Reply Messages'
        UrgentMessages                           = Read-Column $row 'Urgent Messages'
        MeetingsOrganizedCount                   = Read-Column $row 'Meetings Organized Count'
        MeetingsAttendedCount                    = Read-Column $row 'Meetings Attended Count'
        AdHocMeetingsOrganizedCount              = Read-Column $row 'Ad Hoc Meetings Organized Count'
        AdHocMeetingsAttendedCount               = Read-Column $row 'Ad Hoc Meetings Attended Count'
        ScheduledOneTimeMeetingsOrganizedCount   = Read-Column $row 'Scheduled One-time Meetings Organized Count'
        ScheduledOneTimeMeetingsAttendedCount    = Read-Column $row 'Scheduled One-time Meetings Attended Count'
        ScheduledRecurringMeetingsOrganizedCount = Read-Column $row 'Scheduled Recurring Meetings Organized Count'
        ScheduledRecurringMeetingsAttendedCount  = Read-Column $row 'Scheduled Recurring Meetings Attended Count'
        AudioDuration                            = Read-Column $row 'Audio Duration'
        VideoDuration                            = Read-Column $row 'Video Duration'
        ScreenShareDuration                      = Read-Column $row 'Screen Share Duration'
        AudioDurationInSeconds                   = Read-Column $row 'Audio Duration In Seconds'
        VideoDurationInSeconds                   = Read-Column $row 'Video Duration In Seconds'
        ScreenShareDurationInSeconds             = Read-Column $row 'Screen Share Duration In Seconds'
        HasOtherAction                           = Read-Column $row 'Has Other Action'
        IsLicensed                               = Read-Column $row 'Is Licensed'
        ReportPeriod                             = Read-Column $row 'Report Period'
    }
}

Invoke-UsageReportCollector -OutputPath $OutputPath -LogSource $source -CsvName 'teams-user-activity-user-detail.csv' -Column $columns `
    -KeyColumn @('RunDate', 'QueryDate', 'UserPrincipalName', 'ReportPeriod') -ReportName 'Teams user activity' -MapRow $map -Fetch {
        param($file)
        if ($queryDate) {
            Get-MgReportTeamUserActivityUserDetail -Date $Date -OutFile $file -ErrorAction Stop
        }
        else {
            Get-MgReportTeamUserActivityUserDetail -Period $Period -OutFile $file -ErrorAction Stop
        }
    }
