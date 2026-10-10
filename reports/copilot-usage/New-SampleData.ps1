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

# Source 2
$detail = @(
    New-MapRow 'UsageUserDetailMap' @{
        ReportRefreshDate = '2026-09-29'; UserPrincipalName = 'avery.abara@example.com'; DisplayName = 'Avery Abara'
        LastActivityDate = '2026-09-28'; CopilotChatLastActivityDate = '2026-09-28'; MicrosoftTeamsCopilotLastActivityDate = '2026-09-25'
        WordCopilotLastActivityDate = '2026-09-10'; OutlookCopilotLastActivityDate = '2026-09-22'; ReportPeriod = '28'
        PromptsSubmittedAllApps = '41'; PromptsSubmittedCopilotChatWork = '30'; PromptsSubmittedCopilotChatWeb = '3'
        ActiveUsageDaysAllApps = '12'; CopilotChatWorkLastActivityDate = '2026-09-28'; Microsoft365CopilotLastActivityDate = '2026-09-25'
    }
    New-MapRow 'UsageUserDetailMap' @{
        ReportRefreshDate = '2026-09-29'; UserPrincipalName = 'casey.chaudhry@example.com'; DisplayName = 'Casey Chaudhry'
        ReportPeriod = '28'; PromptsSubmittedAllApps = '0'; PromptsSubmittedCopilotChatWork = '0'; PromptsSubmittedCopilotChatWeb = '0'
        ActiveUsageDaysAllApps = '0'
    }
    New-MapRow 'UsageUserDetailMap' @{
        ReportRefreshDate = '2026-09-29'; UserPrincipalName = 'devon.dube@example.com'; DisplayName = 'Devon Dube'
        LastActivityDate = '2026-09-02'; ExcelCopilotLastActivityDate = '2026-09-02'; PowerPointCopilotLastActivityDate = '2026-08-14'
        ReportPeriod = '28'; PromptsSubmittedAllApps = '5'; PromptsSubmittedCopilotChatWork = '0'; PromptsSubmittedCopilotChatWeb = '0'
        ActiveUsageDaysAllApps = '2'; Microsoft365CopilotLastActivityDate = '2026-09-02'
    }
)
Export-AppendCsv -Path (Join-Path $OutputPath 'copilot-usage-user-detail.csv') -Rows $detail -Column (Get-CopilotColumn -Schema $schema -MapKey 'UsageUserDetailMap') -KeyColumn 'RunDate', 'UserPrincipalName', 'ReportPeriod'

# Source 3
$summary = foreach ($period in '7', '28') {
    $scale = if ($period -eq '28') { 2 } else { 1 }
    New-MapRow 'UserCountSummaryMap' @{
        ReportRefreshDate = '2026-09-29'; ReportPeriod = $period
        MicrosoftTeamsEnabledUsers = '120'; MicrosoftTeamsActiveUsers = (30 * $scale).ToString()
        WordEnabledUsers = '120'; WordActiveUsers = (20 * $scale).ToString()
        PowerPointEnabledUsers = '120'; PowerPointActiveUsers = (10 * $scale).ToString()
        OutlookEnabledUsers = '120'; OutlookActiveUsers = (40 * $scale).ToString()
        ExcelEnabledUsers = '120'; ExcelActiveUsers = (12 * $scale).ToString()
        OneNoteEnabledUsers = '120'; OneNoteActiveUsers = '2'
        LoopEnabledUsers = '120'; LoopActiveUsers = '1'
        AnyAppEnabledUsers = '120'; AnyAppActiveUsers = (55 * $scale).ToString()
        CopilotChatEnabledUsers = '120'; CopilotChatActiveUsers = (48 * $scale).ToString()
        TotalPromptsSubmitted = (900 * $scale).ToString(); AveragePromptsSubmitted = '7.5'
    }
}
Export-AppendCsv -Path (Join-Path $OutputPath 'copilot-user-count-summary.csv') -Rows @($summary) -Column (Get-CopilotColumn -Schema $schema -MapKey 'UserCountSummaryMap') -KeyColumn 'RunDate', 'ReportPeriod'

# Source 4
$trend = foreach ($day in 0..2) {
    $date = ([datetime]'2026-09-27').AddDays($day).ToString('yyyy-MM-dd')
    New-MapRow 'UserCountTrendMap' @{
        ReportRefreshDate = '2026-09-29'; ReportDate = $date; ReportPeriod = '7'
        MicrosoftTeamsEnabledUsers = '120'; MicrosoftTeamsActiveUsers = (20 + $day).ToString()
        OutlookEnabledUsers = '120'; OutlookActiveUsers = (25 + $day).ToString()
        AnyAppEnabledUsers = '120'; AnyAppActiveUsers = (40 + $day).ToString()
        CopilotChatEnabledUsers = '120'; CopilotChatActiveUsers = (35 + $day).ToString()
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
