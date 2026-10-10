#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes teams-user-activity-counts.csv: tenant-wide Teams activity per day, appended each run.

    .DESCRIPTION
        Source 2: Microsoft Graph getTeamsUserActivityCounts
        (https://learn.microsoft.com/graph/api/reportroot-getteamsuseractivitycounts), read with
        Get-MgReportTeamUserActivityCount -OutFile. One row per Report Date inside the period, for
        Teams licensed users. Period only: the API has no date form. Re-reading a period appends
        the same Report Date again under a new RunDate, because a recent day can change between runs.

        Not available in GCC High; GCC is UNVERIFIED and is attempted with a warning. Needs
        Reports.Read.All. Global Reader and Usage Summary Reports Reader see this report, which holds
        tenant-level data only.

    .EXAMPLE
        ./Get-TeamsUserActivityCounts.ps1 -OutputPath ./out
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
$columns = $schema.TeamsUserActivityCounts
$source = 'teams-user-activity-counts'
$csvPath = Join-Path $OutputPath 'teams-user-activity-counts.csv'

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

if (Test-TeamsSourceSkipped -Source 'TeamsUserActivityCounts' -LogSource $source -CsvPath $csvPath -Column $columns `
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
    [pscustomobject]@{
        RunDate             = $runDate
        ReportRefreshDate   = Read-Column $row 'Report Refresh Date'
        ReportDate          = Read-Column $row 'Report Date'
        TeamChatMessages    = Read-Column $row 'Team Chat Messages'
        PostMessages        = Read-Column $row 'Post Messages'
        ReplyMessages       = Read-Column $row 'Reply Messages'
        PrivateChatMessages = Read-Column $row 'Private Chat Messages'
        Calls               = Read-Column $row 'Calls'
        Meetings            = Read-Column $row 'Meetings'
        AudioDuration       = Read-Column $row 'Audio Duration'
        VideoDuration       = Read-Column $row 'Video Duration'
        ScreenShareDuration = Read-Column $row 'Screen Share Duration'
        MeetingsOrganized   = Read-Column $row 'Meetings Organized'
        MeetingsAttended    = Read-Column $row 'Meetings Attended'
        ReportPeriod        = Read-Column $row 'Report Period'
    }
}

Invoke-UsageReportCollector -OutputPath $OutputPath -LogSource $source -CsvName 'teams-user-activity-counts.csv' -Column $columns `
    -KeyColumn @('RunDate', 'ReportDate', 'ReportPeriod') -ReportName 'Teams user activity counts' -MapRow $map -Fetch {
        param($file)
        Get-MgReportTeamUserActivityCount -Period $Period -OutFile $file -ErrorAction Stop
    }
