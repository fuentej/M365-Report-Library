#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes team-activity.csv: Microsoft Teams activity by team, appended each run.

    .DESCRIPTION
        Source: Microsoft Graph getTeamsTeamActivityDetail
        (https://learn.microsoft.com/graph/api/reportroot-getteamsteamactivitydetail),
        read with Get-MgReportTeamActivityDetail -OutFile. Graph answers with a 302 to a
        preauthenticated CSV download that is valid for a few minutes; the cmdlet saves it
        to a temporary file, which this script reads and removes.

        It is a snapshot of a rolling period (D7, D30, D90 or D180). Last Activity Date is
        the latest activity whatever the period; the counts are for the period.

        The team-type header is spelled "Team type" in the header list and "Team Type" in
        the schema example; both are accepted. Team names are blank when the organization
        setting that conceals user, group and site names is on.

        Not available in GCC High (US Government L4). GCC is UNVERIFIED and is attempted
        with a warning. Needs Reports.Read.All; a delegated caller also needs a limited
        admin role such as Reports Reader. Global Reader and Usage Summary Reports Reader
        do not receive the detail rows.

    .PARAMETER Period
        The report period. Defaults to D180, the longest.

    .EXAMPLE
        ./Get-TeamActivity.ps1 -OutputPath ./out
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
    [string]$Period = 'D180',

    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Imported by path without -Force: a forced re-import builds a fresh module session
# state, which would discard anything the caller has arranged around the module - and
# that is exactly how the tests put mocks in front of the tenant cmdlets.
Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

. (Join-Path $PSScriptRoot 'TeamsGroupsHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'TeamsGroupsSchema.psd1')
$columns = $schema.TeamActivity
$aliases = $schema.ReportHeaderAliases
$source = 'team-activity'
$csvPath = Join-Path $OutputPath 'team-activity.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-TeamsGroupsSourceAvailability -Source 'TeamActivity' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping team-activity.csv. $($availability.Reason) $($availability.Reference)")
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}
if ($availability.Status -eq 'Unverified') {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message $availability.Reason
}

if (-not $SkipConnect) {
    Connect-M365Service -Service Graph -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization
}

$download = Join-Path ([System.IO.Path]::GetTempPath()) ('team-activity-' + [guid]::NewGuid().ToString('N') + '.csv')
$report = $null
try {
    try {
        Get-MgReportTeamActivityDetail -Period $Period -OutFile $download -ErrorAction Stop
        $report = if (Test-Path -LiteralPath $download) { @(Import-Csv -LiteralPath $download) } else { @() }
    }
    catch {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            'The Teams team activity report is unavailable to this sign-in or cloud ({0}). It needs Reports.Read.All and, for a delegated sign-in, a role such as Reports Reader. Writing the header only.' -f $_.Exception.Message)
        Export-AppendCsv -Path $csvPath -Column $columns
        return
    }
}
finally {
    Remove-Item -LiteralPath $download -Force -ErrorAction SilentlyContinue
}

function Read-Column { param($Row, [string[]]$Header) Get-ReportColumnValue -Row $Row -Header $Header }

$rows = foreach ($row in $report) {
    [pscustomobject]@{
        RunDate              = $runDate
        ReportRefreshDate    = Read-Column $row 'Report Refresh Date'
        ReportPeriod         = Read-Column $row 'Report Period'
        TeamId               = Read-Column $row 'Team Id'
        TeamName             = Read-Column $row 'Team Name'
        TeamType             = Read-Column $row $aliases.TeamType
        LastActivityDate     = Read-Column $row 'Last Activity Date'
        ActiveUsers          = Read-Column $row 'Active Users', 'Active users'
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

$rows = @($rows | Where-Object { -not [string]::IsNullOrWhiteSpace($_.TeamId) })

$result = Export-AppendCsv -Path $csvPath -Rows $rows -Column $columns -KeyColumn @('RunDate', 'TeamId') -PassThru
Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'team-activity.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
