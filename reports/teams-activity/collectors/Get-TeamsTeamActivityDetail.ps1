#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes teams-team-activity-detail.csv: channel messages, replies, posts, meetings, active users and last activity per team for a rolling period, appended each run.

    .DESCRIPTION
        Source: Microsoft Graph getTeamsTeamActivityDetail
        (https://learn.microsoft.com/graph/api/reportroot-getteamsteamactivitydetail), read with
        Get-MgReportTeamActivityDetail -OutFile. The numbers include licensed and non-licensed users. Last
        Activity Date is the latest activity whatever the period; the counts are for the period. The report
        is a snapshot stamped with the run date, so a team that is missing from it is not a team with no activity.

        The page spells the header list "Team type" and "Active users" and its schema example "Team Type" and
        "Active Users"; both are accepted.

        Not available in US Government L4 or L5 (the API page marks the global service only), so a GCC High run
        writes the header only and logs why. GCC is UNVERIFIED and is attempted with a warning. Needs
        Reports.Read.All; a delegated caller also needs a limited admin role such as Reports Reader. Team names
        are concealed or blank when the organization setting that conceals user, group and site names is on
        (see report-settings.csv).

    .EXAMPLE
        ./Get-TeamsTeamActivityDetail.ps1 -OutputPath ./out
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
$columns = $schema.TeamsTeamActivityDetail
$source = 'teams-team-activity-detail'
$csvPath = Join-Path $OutputPath 'teams-team-activity-detail.csv'

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

if (Test-TeamsSourceSkipped -Source 'TeamsTeamActivityDetail' -LogSource $source -CsvPath $csvPath -Column $columns `
        -OutputPath $OutputPath -Environment $Environment -Schema $schema) {
    return
}

if (-not $SkipConnect) {
    Connect-M365Service -Service Graph -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization
}

function Read-Column { param($Row, [string[]]$Header) Get-ReportColumnValue -Row $Row -Header $Header }

$map = {
    param($row, $runDate)
    $teamId = Read-Column $row 'Team Id'
    if ([string]::IsNullOrWhiteSpace($teamId)) { return $null }
    [pscustomobject]@{
        RunDate              = $runDate
        ReportRefreshDate    = Read-Column $row 'Report Refresh Date'
        ReportPeriod         = Read-Column $row 'Report Period'
        TeamId               = $teamId
        TeamName             = Read-Column $row 'Team Name'
        TeamType             = Read-Column $row 'Team type', 'Team Type'
        LastActivityDate     = Read-Column $row 'Last Activity Date'
        ActiveUsers          = Read-Column $row 'Active users', 'Active Users'
        ActiveChannels       = Read-Column $row 'Active Channels'
        Guests               = Read-Column $row 'Guests'
        Reactions            = Read-Column $row 'Reactions'
        MeetingsOrganized    = Read-Column $row 'Meetings Organized'
        PostMessages         = Read-Column $row 'Post Messages'
        ReplyMessages        = Read-Column $row 'Reply Messages'
        ChannelMessages      = Read-Column $row 'Channel Messages'
        UrgentMessages       = Read-Column $row 'Urgent Messages'
        Mentions             = Read-Column $row 'Mentions'
        ActiveSharedChannels = Read-Column $row 'Active Shared Channels'
        ActiveExternalUsers  = Read-Column $row 'Active External Users'
    }
}

Invoke-UsageReportCollector -OutputPath $OutputPath -LogSource $source -CsvName 'teams-team-activity-detail.csv' -Column $columns `
    -KeyColumn @('RunDate', 'TeamId', 'ReportPeriod') -ReportName 'Teams team activity' -MapRow $map -Fetch {
        param($file)
        Get-MgReportTeamActivityDetail -Period $Period -OutFile $file -ErrorAction Stop
    }
