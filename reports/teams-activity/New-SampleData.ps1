#Requires -Version 7.0

<#
    .SYNOPSIS
        Regenerates samples/: fake CSVs with exactly the columns the collectors write.

    .DESCRIPTION
        Every name is invented and every address is on example.com. Columns come from
        collectors/TeamsActivitySchema.psd1, so a sample cannot drift from its collector.
        samples/gcchigh/ holds the header-only files a GCC High run leaves for the three Graph
        usage reports, which are not available in that cloud.

    .EXAMPLE
        ./New-SampleData.ps1
#>
[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path $PSScriptRoot 'samples')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot '../../shared/M365ReportLibrary.psm1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'collectors/TeamsActivitySchema.psd1')

function Write-Sample {
    param([string]$Folder, [string]$Name, [string[]]$Column, [object[]]$Row)
    $path = Join-Path $Folder $Name
    if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
    Export-AppendCsv -Path $path -Rows $Row -Column $Column
}

if (-not (Test-Path -LiteralPath $OutputPath)) { New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null }
$gccHigh = Join-Path $OutputPath 'gcchigh'
if (-not (Test-Path -LiteralPath $gccHigh)) { New-Item -Path $gccHigh -ItemType Directory -Force | Out-Null }

$people = @(
    @{ Id = '22222222-0000-0000-0000-000000000001'; Upn = 'avery.abara@example.com'; Last = '2026-10-07'; Chat = 41; Priv = 120; Calls = 6; Org = 4; Att = 12; Licensed = 'True' }
    @{ Id = '22222222-0000-0000-0000-000000000002'; Upn = 'blake.bishop@example.com'; Last = '2026-10-06'; Chat = 8; Priv = 33; Calls = 2; Org = 0; Att = 5; Licensed = 'True' }
    @{ Id = '22222222-0000-0000-0000-000000000003'; Upn = 'casey.cho@example.com'; Last = '2026-09-02'; Chat = 0; Priv = 0; Calls = 0; Org = 0; Att = 0; Licensed = 'True' }
    @{ Id = '22222222-0000-0000-0000-000000000004'; Upn = 'devon.diaz@example.com'; Last = ''; Chat = 0; Priv = 0; Calls = 0; Org = 0; Att = 0; Licensed = 'False' }
)
$runDates = '2026-10-01', '2026-10-08'

# Source 1
$rows = foreach ($run in $runDates) {
    foreach ($p in $people) {
        $scale = if ($run -eq '2026-10-08') { 1 } else { 0 }
        [pscustomobject]@{
            RunDate = $run; QueryDate = ''; ReportRefreshDate = ([datetime]$run).AddDays(-3).ToString('yyyy-MM-dd')
            UserId = $p.Id; UserPrincipalName = $p.Upn; LastActivityDate = $p.Last; IsDeleted = 'False'; DeletedDate = ''
            AssignedProducts = 'MICROSOFT 365 E3'
            TeamChatMessageCount = $p.Chat + $scale; PrivateChatMessageCount = $p.Priv + $scale; CallCount = $p.Calls; MeetingCount = $p.Att
            PostMessages = [int]($p.Chat * 0.6); ReplyMessages = [int]($p.Chat * 0.4); UrgentMessages = 0
            MeetingsOrganizedCount = $p.Org; MeetingsAttendedCount = $p.Att
            AdHocMeetingsOrganizedCount = 0; AdHocMeetingsAttendedCount = 1
            ScheduledOneTimeMeetingsOrganizedCount = $p.Org; ScheduledOneTimeMeetingsAttendedCount = [int]($p.Att / 2)
            ScheduledRecurringMeetingsOrganizedCount = 0; ScheduledRecurringMeetingsAttendedCount = [int]($p.Att / 2)
            AudioDuration = 'PT1H30M'; VideoDuration = 'PT20M'; ScreenShareDuration = 'PT5M'
            AudioDurationInSeconds = 5400; VideoDurationInSeconds = 1200; ScreenShareDurationInSeconds = 300
            HasOtherAction = 'Yes'; IsLicensed = $p.Licensed; ReportPeriod = 30
        }
    }
}
Write-Sample $OutputPath 'teams-user-activity-user-detail.csv' $schema.TeamsUserActivityUserDetail $rows

# Source 2
$rows = foreach ($run in $runDates) {
    foreach ($offset in 2, 1, 0) {
        $day = ([datetime]$run).AddDays(-3 - $offset).ToString('yyyy-MM-dd')
        [pscustomobject]@{
            RunDate = $run; ReportRefreshDate = ([datetime]$run).AddDays(-3).ToString('yyyy-MM-dd'); ReportDate = $day
            TeamChatMessages = 210 + $offset; PostMessages = 120; ReplyMessages = 90; PrivateChatMessages = 540 - $offset
            Calls = 31; Meetings = 77; AudioDuration = 'PT40H'; VideoDuration = 'PT12H'; ScreenShareDuration = 'PT3H'
            MeetingsOrganized = 25; MeetingsAttended = 77; ReportPeriod = 30
        }
    }
}
Write-Sample $OutputPath 'teams-user-activity-counts.csv' $schema.TeamsUserActivityCounts $rows

# Source 3
$platforms = @{
    '22222222-0000-0000-0000-000000000001' = 'Yes', 'No', 'Yes', 'No', 'Yes', 'No', 'No'
    '22222222-0000-0000-0000-000000000002' = 'Yes', 'No', 'No', 'No', 'No', 'Yes', 'No'
    '22222222-0000-0000-0000-000000000003' = 'No', 'No', 'No', 'Yes', 'No', 'No', 'Yes'
    '22222222-0000-0000-0000-000000000004' = 'No', 'No', 'No', 'No', 'No', 'No', 'No'
}
$rows = foreach ($run in $runDates) {
    foreach ($p in $people) {
        $u = $platforms[$p.Id]
        [pscustomobject]@{
            RunDate = $run; QueryDate = ''; ReportRefreshDate = ([datetime]$run).AddDays(-3).ToString('yyyy-MM-dd')
            UserId = $p.Id; UserPrincipalName = $p.Upn; LastActivityDate = $p.Last; IsDeleted = 'False'; DeletedDate = ''
            UsedWeb = $u[0]; UsedWindowsPhone = 'No'; UsediOS = $u[1]; UsedMac = $u[2]; UsedAndroidPhone = $u[3]
            UsedWindows = $u[4]; UsedChromeOS = $u[5]; UsedLinux = $u[6]; IsLicensed = $p.Licensed; ReportPeriod = 30
        }
    }
}
Write-Sample $OutputPath 'teams-device-usage-user-detail.csv' $schema.TeamsDeviceUsageUserDetail $rows

# Source 5
$rows = $runDates | ForEach-Object { [pscustomobject]@{ RunDate = $_; DisplayConcealedNames = 'False' } }
Write-Sample $OutputPath 'report-settings.csv' $schema.ReportSettings $rows

# Source 6
$rows = @(
    [pscustomobject]@{
        CallRecordId = '5a1f0001-0000-0000-0000-000000000001'; Version = 1; Type = 'groupCall'; Modalities = 'audio;video'
        StartDateTime = '2026-10-07T14:00:00Z'; EndDateTime = '2026-10-07T14:45:00Z'; LastModifiedDateTime = '2026-10-07T14:47:00Z'
        SessionId = '6b2e0001-0000-0000-0000-000000000001'; SessionStartDateTime = '2026-10-07T14:00:00Z'; SessionEndDateTime = '2026-10-07T14:45:00Z'
        CallerUserId = $people[0].Id; CallerPlatform = 'windows'; CalleeUserId = ''; CalleePlatform = ''
    }
    [pscustomobject]@{
        CallRecordId = '5a1f0001-0000-0000-0000-000000000001'; Version = 1; Type = 'groupCall'; Modalities = 'audio;video'
        StartDateTime = '2026-10-07T14:00:00Z'; EndDateTime = '2026-10-07T14:45:00Z'; LastModifiedDateTime = '2026-10-07T14:47:00Z'
        SessionId = '6b2e0001-0000-0000-0000-000000000002'; SessionStartDateTime = '2026-10-07T14:02:00Z'; SessionEndDateTime = '2026-10-07T14:44:00Z'
        CallerUserId = $people[1].Id; CallerPlatform = 'iOS'; CalleeUserId = ''; CalleePlatform = ''
    }
    # The same session id again: a transfer can involve more than one service identity.
    # https://learn.microsoft.com/graph/callrecords-api-faq
    [pscustomobject]@{
        CallRecordId = '5a1f0001-0000-0000-0000-000000000001'; Version = 1; Type = 'groupCall'; Modalities = 'audio;video'
        StartDateTime = '2026-10-07T14:00:00Z'; EndDateTime = '2026-10-07T14:45:00Z'; LastModifiedDateTime = '2026-10-07T14:47:00Z'
        SessionId = '6b2e0001-0000-0000-0000-000000000002'; SessionStartDateTime = '2026-10-07T14:02:00Z'; SessionEndDateTime = '2026-10-07T14:20:00Z'
        CallerUserId = $people[1].Id; CallerPlatform = 'iOS'; CalleeUserId = '33333333-0000-0000-0000-000000000099'; CalleePlatform = 'unknown'
    }
    [pscustomobject]@{
        CallRecordId = '5a1f0002-0000-0000-0000-000000000002'; Version = 2; Type = 'peerToPeer'; Modalities = 'audio'
        StartDateTime = '2026-10-08T09:30:00Z'; EndDateTime = '2026-10-08T09:40:00Z'; LastModifiedDateTime = '2026-10-08T09:43:00Z'
        SessionId = '6b2e0002-0000-0000-0000-000000000001'; SessionStartDateTime = '2026-10-08T09:30:00Z'; SessionEndDateTime = '2026-10-08T09:40:00Z'
        CallerUserId = $people[0].Id; CallerPlatform = 'macOS'; CalleeUserId = $people[1].Id; CalleePlatform = 'web'
    }
)
Write-Sample $OutputPath 'call-records.csv' $schema.CallRecords $rows

# Source 8
$rows = @(
    [pscustomobject]@{ Date = '2026-10-06'; Workload = 'MicrosoftTeams'; UserId = $people[0].Upn; Operation = 'MeetingDetail'; EventCount = 3 }
    [pscustomobject]@{ Date = '2026-10-06'; Workload = 'MicrosoftTeams'; UserId = $people[1].Upn; Operation = 'MeetingParticipantDetail'; EventCount = 2 }
    [pscustomobject]@{ Date = '2026-10-07'; Workload = 'MicrosoftTeams'; UserId = $people[0].Upn; Operation = 'CallParticipantDetail'; EventCount = 1 }
    [pscustomobject]@{ Date = '2026-10-07'; Workload = 'MicrosoftTeams'; UserId = $people[1].Upn; Operation = 'MessageSent'; EventCount = 4 }
    [pscustomobject]@{ Date = '2026-10-07'; Workload = 'MicrosoftTeams'; UserId = $people[0].Upn; Operation = 'ChatCreated'; EventCount = 1 }
)
Write-Sample $OutputPath 'teams-audit-events.csv' $schema.TeamsAuditEvents $rows

# Header-only files a GCC High run leaves for sources 1 to 3
$notAvailable = [ordered]@{
    'teams-user-activity-user-detail.csv' = 'TeamsUserActivityUserDetail'
    'teams-user-activity-counts.csv'      = 'TeamsUserActivityCounts'
    'teams-device-usage-user-detail.csv'  = 'TeamsDeviceUsageUserDetail'
}
foreach ($name in $notAvailable.Keys) {
    Write-Sample $gccHigh $name $schema[$notAvailable[$name]] @()
}
