#Requires -Version 7.0

<#
    .SYNOPSIS
        Regenerates samples/: fake CSVs with the exact columns each collector writes.

    .DESCRIPTION
        Every name, address and Id is invented; addresses use example.com. The column order
        comes from collectors/CopilotUsageSchema.psd1, so a sample cannot drift from its
        collector. No sign-in happens and no tenant is called. The usage-detail sample shows
        readable names, as when displayConcealedNames is false; with the default (true) the
        user principal name and display name are 32-character hashes.

        The usage report sources are NotAvailable in GCC High, so the header-only case is
        samples/gcchigh/: the same three files with no rows.

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
. (Join-Path $PSScriptRoot 'collectors/CopilotUsageHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'collectors/CopilotUsageSchema.psd1')

if (Test-Path -LiteralPath $OutputPath) {
    Get-ChildItem -LiteralPath $OutputPath -Recurse -Filter '*.csv' | Remove-Item -Force
}

function New-MapRow {
    param([string]$MapKey, [hashtable]$Values)

    $row = [ordered]@{ RunDate = '2026-09-30' }
    foreach ($entry in $schema[$MapKey]) {
        $row[$entry.Column] = if ($Values.ContainsKey($entry.Column)) { $Values[$entry.Column] } else { '' }
    }
    [pscustomobject]$row
}

# Source 2 and 3. One seeded random drives the per-person file, and the user-count summary is
# counted from that file, so the two cannot disagree. Same seed, same files.
$random = [System.Random]::new(20260929)
$refreshDate = [datetime]'2026-09-29'
$firstNames = 'Avery', 'Casey', 'Devon', 'Emerson', 'Finley', 'Gray', 'Harper', 'Indigo', 'Jordan', 'Kai', 'Logan', 'Morgan', 'Nico', 'Oakley', 'Parker', 'Quinn', 'Reese', 'Sage', 'Tatum', 'Wren'
$lastNames = 'Abara', 'Chaudhry', 'Dube', 'Eriksen', 'Fontaine', 'Gupta', 'Hollis', 'Iyer', 'Jansen', 'Kowalski', 'Lindqvist', 'Moreau'
$personCount = 60

# Per app: the share of people who use it, and the share of its monthly (28-day) users who were
# also active in the last 7 days. The second number is what makes the weekly-to-monthly ratio
# differ by app: Teams and Outlook are daily habits, PowerPoint and Loop are occasional.
$appProfile = [ordered]@{
    CopilotChat = @{ Column = 'CopilotChatLastActivityDate'; Adopt = 0.85; Weekly = 0.78 }
    MicrosoftTeams = @{ Column = 'MicrosoftTeamsCopilotLastActivityDate'; Adopt = 0.70; Weekly = 0.65 }
    Outlook = @{ Column = 'OutlookCopilotLastActivityDate'; Adopt = 0.65; Weekly = 0.55 }
    Word = @{ Column = 'WordCopilotLastActivityDate'; Adopt = 0.55; Weekly = 0.50 }
    Excel = @{ Column = 'ExcelCopilotLastActivityDate'; Adopt = 0.40; Weekly = 0.35 }
    PowerPoint = @{ Column = 'PowerPointCopilotLastActivityDate'; Adopt = 0.35; Weekly = 0.18 }
    OneNote = @{ Column = 'OneNoteCopilotLastActivityDate'; Adopt = 0.30; Weekly = 0.20 }
    Loop = @{ Column = 'LoopCopilotLastActivityDate'; Adopt = 0.25; Weekly = 0.10 }
}

$people = foreach ($index in 0..($personCount - 1)) {
    $first = $firstNames[$index % $firstNames.Count]
    $last = $lastNames[$index % $lastNames.Count]
    $values = @{
        ReportRefreshDate = $refreshDate.ToString('yyyy-MM-dd')
        UserPrincipalName = ('{0}.{1}@example.com' -f $first, $last).ToLowerInvariant()
        DisplayName = "$first $last"
        ReportPeriod = '28'
    }
    $latest = $null
    $appDates = @{}
    foreach ($app in $appProfile.Keys) {
        $profile = $appProfile[$app]
        # Adoption is random; weekly or not is spread evenly over people (not random) so each
        # app's ratio lands near its Weekly value instead of wandering with a small count.
        $appNumber = @($appProfile.Keys).IndexOf($app)
        $spread = (($index * 37 + $appNumber * 11) % 100) / 100
        if ($random.NextDouble() -ge $profile.Adopt) { continue }
        # Some users tried the app weeks ago and stopped: past the 28-day window.
        $daysAgo = if ($spread -lt $profile.Weekly) { $random.Next(1, 7) }
        elseif ($random.NextDouble() -lt 0.85) { $random.Next(7, 28) }
        else { $random.Next(28, 75) }
        $date = $refreshDate.AddDays(-$daysAgo)
        $appDates[$app] = $date
        $values[$profile.Column] = $date.ToString('yyyy-MM-dd')
        if ($null -eq $latest -or $date -gt $latest) { $latest = $date }
    }
    if ($appDates.ContainsKey('CopilotChat')) {
        $values.CopilotChatWorkLastActivityDate = $values.CopilotChatLastActivityDate
    }
    $workApps = @($appDates.Keys | Where-Object { $_ -ne 'CopilotChat' })
    if ($workApps.Count -gt 0) {
        $values.Microsoft365CopilotLastActivityDate = ($workApps | ForEach-Object { $appDates[$_] } | Sort-Object -Descending | Select-Object -First 1).ToString('yyyy-MM-dd')
    }
    if ($latest) { $values.LastActivityDate = $latest.ToString('yyyy-MM-dd') }

    $inWindow = @($appDates.Values | Where-Object { $_ -ge $refreshDate.AddDays(-27) })
    $values.ActiveUsageDaysAllApps = [string]$(if ($inWindow.Count -eq 0) { 0 } else { [math]::Min(28, $random.Next($inWindow.Count, 4 * $inWindow.Count + 4)) })
    $chatWork = if ($appDates.ContainsKey('CopilotChat') -and $appDates.CopilotChat -ge $refreshDate.AddDays(-27)) { $random.Next(1, 40) } else { 0 }
    $chatWeb = if ($chatWork -gt 0) { $random.Next(0, 6) } else { 0 }
    $otherPrompts = 0
    foreach ($app in $appDates.Keys) { if ($app -ne 'CopilotChat' -and $appDates[$app] -ge $refreshDate.AddDays(-27)) { $otherPrompts += $random.Next(1, 25) } }
    $values.PromptsSubmittedCopilotChatWork = [string]$chatWork
    $values.PromptsSubmittedCopilotChatWeb = [string]$chatWeb
    $values.PromptsSubmittedAllApps = [string]($chatWork + $chatWeb + $otherPrompts)

    [pscustomobject]@{ Row = (New-MapRow 'UsageUserDetailMap' $values); Dates = $appDates; Latest = $latest; Prompts = ($chatWork + $chatWeb + $otherPrompts) }
}
$detail = @($people.Row)
Export-AppendCsv -Path (Join-Path $OutputPath 'copilot-usage-user-detail.csv') -Rows $detail -Column (Get-CopilotColumn -Schema $schema -MapKey 'UsageUserDetailMap') -KeyColumn 'RunDate', 'UserPrincipalName', 'ReportPeriod'

function Get-ActiveCount {
    param([string]$App, [int]$Days)
    $cutoff = $refreshDate.AddDays(1 - $Days)
    if ($App -eq 'AnyApp') { return @($people | Where-Object { $_.Latest -and $_.Latest -ge $cutoff }).Count }
    @($people | Where-Object { $_.Dates.ContainsKey($App) -and $_.Dates[$App] -ge $cutoff }).Count
}

$totalPrompts28 = ($people | Measure-Object -Property Prompts -Sum).Sum
$summary = foreach ($period in '7', '28') {
    $days = [int]$period
    $values = @{ ReportRefreshDate = $refreshDate.ToString('yyyy-MM-dd'); ReportPeriod = $period }
    foreach ($app in 'MicrosoftTeams', 'Word', 'PowerPoint', 'Outlook', 'Excel', 'OneNote', 'Loop', 'CopilotChat', 'AnyApp') {
        $values["$($app)EnabledUsers"] = [string]$personCount
        $values["$($app)ActiveUsers"] = [string](Get-ActiveCount -App $app -Days $days)
    }
    $total = if ($days -eq 28) { $totalPrompts28 } else { [int][math]::Round($totalPrompts28 * 0.3) }
    $values.TotalPromptsSubmitted = [string]$total
    $values.AveragePromptsSubmitted = ($total / $personCount).ToString('0.0', [cultureinfo]::InvariantCulture)
    New-MapRow 'UserCountSummaryMap' $values
}
Export-AppendCsv -Path (Join-Path $OutputPath 'copilot-user-count-summary.csv') -Rows @($summary) -Column (Get-CopilotColumn -Schema $schema -MapKey 'UserCountSummaryMap') -KeyColumn 'RunDate', 'ReportPeriod'

# Source 4
$trend = foreach ($day in 0..2) {
    $date = ([datetime]'2026-09-27').AddDays($day).ToString('yyyy-MM-dd')
    New-MapRow 'UserCountTrendMap' @{
        ReportRefreshDate = '2026-09-29'; ReportDate = $date; ReportPeriod = '7'
        MicrosoftTeamsEnabledUsers = '60'; MicrosoftTeamsActiveUsers = (20 + $day).ToString()
        OutlookEnabledUsers = '60'; OutlookActiveUsers = (25 + $day).ToString()
        AnyAppEnabledUsers = '60'; AnyAppActiveUsers = (40 + $day).ToString()
        CopilotChatEnabledUsers = '60'; CopilotChatActiveUsers = (35 + $day).ToString()
        PromptsSubmitted = (130 + 10 * $day).ToString()
    }
}
Export-AppendCsv -Path (Join-Path $OutputPath 'copilot-user-count-trend.csv') -Rows @($trend) -Column (Get-CopilotColumn -Schema $schema -MapKey 'UserCountTrendMap') -KeyColumn 'RunDate', 'ReportDate', 'ReportPeriod'

# GCC High: the three usage reports are NotAvailable, so the files hold the header only.
foreach ($pair in @(
        @{ Csv = 'copilot-usage-user-detail.csv'; Map = 'UsageUserDetailMap' }
        @{ Csv = 'copilot-user-count-summary.csv'; Map = 'UserCountSummaryMap' }
        @{ Csv = 'copilot-user-count-trend.csv'; Map = 'UserCountTrendMap' }
    )) {
    Export-AppendCsv -Path (Join-Path $OutputPath "gcchigh/$($pair.Csv)") -Column (Get-CopilotColumn -Schema $schema -MapKey $pair.Map)
}

# Source 6
function New-AuditRow {
    param([string]$Id, [string]$At, [string]$User, [string]$Host_, [string]$Identity, [string]$AgentId, [string]$AgentName, [int]$Messages, [int]$Prompts, [int]$Resources, [string]$Plugins)
    [pscustomobject]@{
        CreationTime = $At; Id = $Id; UserId = $User; Operation = 'CopilotInteraction'; RecordType = 'CopilotInteraction'; Workload = 'Copilot'
        AppHost = $Host_; AppIdentity = $Identity; AgentId = $AgentId; AgentName = $AgentName
        MessageCount = $Messages; PromptMessageCount = $Prompts; AccessedResourceCount = $Resources; PluginIds = $Plugins
    }
}
$audit = @(
    New-AuditRow 'bbbbbbbb-0000-4000-8000-000000000001' '2026-09-28T14:02:11Z' 'avery.abara@example.com' 'Teams' 'Copilot.MicrosoftCopilot.BizChat' '' '' 2 1 3 'BingWebSearch'
    New-AuditRow 'bbbbbbbb-0000-4000-8000-000000000002' '2026-09-28T15:30:40Z' 'avery.abara@example.com' 'Word' 'Copilot.MicrosoftCopilot.Microsoft365Copilot' '' '' 3 1 1 ''
    New-AuditRow 'bbbbbbbb-0000-4000-8000-000000000003' '2026-09-29T08:12:05Z' 'devon.dube@example.com' 'Teams' 'Copilot.Studio.f4d97b45-1deb-40ce-9004-b473b79eab85' 'CopilotStudio.Declarative.8ad83f3e-b424-4d54-8ddb-15dc19247088' 'SalesAgent' 2 1 0 ''
)
Export-AppendCsv -Path (Join-Path $OutputPath 'copilot-audit-events.csv') -Rows $audit -Column $schema.CopilotAuditEvents -KeyColumn 'Id'

# Source 8: metadata only
function New-InteractionRow {
    param([string]$Id, [string]$At, [string]$User, [string]$Request, [string]$Type, [string]$AppClass, [string]$Conversation, [int]$Contexts, [string]$ContextTypes)
    [pscustomobject]@{
        CreatedDateTime = $At; Id = $Id; UserId = $User; SessionId = '19:sample-session@thread.v2'; RequestId = $Request
        AppClass = $AppClass; InteractionType = $Type; ConversationType = $Conversation; Locale = 'en-us'
        ContextCount = $Contexts; ContextTypes = $ContextTypes
    }
}
$interactions = @(
    New-InteractionRow '1764008994427' '2026-09-24T18:29:54Z' 'aaaaaaaa-0000-4000-8000-000000000001' '79699122-d834-6cc2-c1df-0332a0bd982d' 'userPrompt' 'IPM.SkypeTeams.Message.Copilot.BizChat' 'bizchat' 0 ''
    New-InteractionRow '1764008996100' '2026-09-24T18:29:56Z' 'aaaaaaaa-0000-4000-8000-000000000001' '79699122-d834-6cc2-c1df-0332a0bd982d' 'aiResponse' 'IPM.SkypeTeams.Message.Copilot.BizChat' 'bizchat' 0 ''
    New-InteractionRow '1764095000000' '2026-09-25T10:00:00Z' 'aaaaaaaa-0000-4000-8000-000000000002' '7336770c-fb25-48ac-8303-4493ad11ed71' 'aiResponse' 'IPM.SkypeTeams.Message.Copilot.Teams' 'appchat' 1 'TeamsMeeting'
)
Export-AppendCsv -Path (Join-Path $OutputPath 'copilot-interactions.csv') -Rows $interactions -Column $schema.CopilotInteractions -KeyColumn 'UserId', 'Id'

# Source 10
$features = foreach ($feature in $schema.FeatureRows) {
    [pscustomobject]@{
        RunDate = '2026-09-30'; Feature = $feature.Feature; Commercial = $feature.Commercial; GCC = $feature.GCC
        GCCHigh = $feature.GCCHigh; Note = $feature.Note; PageReadDate = $schema.FeaturePageReadDate; Reference = $schema.FeaturePage
    }
}
Export-AppendCsv -Path (Join-Path $OutputPath 'copilot-feature-availability.csv') -Rows @($features) -Column $schema.CopilotFeatureAvailability -KeyColumn 'RunDate', 'Feature'
