#Requires -Version 7.0

<#
    BRO-481: every number on every page of these reports says what it is. Power BI
    Desktop reads formatting properties only in the PBIR expression form
    ({ "expr": { "Literal": { "Value": ... } } }); a bare JSON true or string is not
    rendered, which is how the cards in five of these reports lost their titles.

    Checks, per visual.json:
    - every projection has a nativeQueryRef;
    - every formatting property under title, categoryLabels, categoryAxis and
      valueAxis is an expression, not a bare value;
    - a card shows a title or a category label;
    - a chart or table shows a non-empty title, and a chart shows both axis titles.

    Read-only filesystem and JSON inspection.
#>

$script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
$script:Reports = @(
    'purview-ip/report/PurviewIPReport.Report'
    'identity-posture/report/IdentityPosture.Report'
    'guest-access/report/GuestAccess.Report'
    'teams-groups-lifecycle/report/TeamsGroupsLifecycle.Report'
    'copilot-studio-agents/report/CopilotStudioAgents.Report'
    'mailbox-exfiltration-risk/report/MailboxExfiltrationRisk.Report'
)
$script:Visuals = foreach ($report in $script:Reports) {
    $pages = Join-Path $script:RepoRoot "reports/$report/definition/pages"
    foreach ($file in (Get-ChildItem -LiteralPath $pages -Recurse -File -Filter 'visual.json')) {
        $type = (Get-Content -LiteralPath $file.FullName -Raw | ConvertFrom-Json -AsHashtable).visual.visualType
        @{ Report = $report.Split('/')[0]; Name = $file.Directory.Name; Path = $file.FullName; Type = $type }
    }
}

BeforeAll {
    function Get-LiteralValue {
        param($Property)
        if ($Property -is [System.Collections.IDictionary] -and $Property.Contains('expr')) {
            return $Property.expr.Literal.Value
        }
        return $null
    }

    function Get-ObjectProperty {
        param([System.Collections.IDictionary]$Objects, [string]$Object, [string]$Property)
        if (-not $Objects -or -not $Objects.Contains($Object)) { return $null }
        Get-LiteralValue $Objects[$Object][0].properties[$Property]
    }
}

Describe 'Labels on <Report> / <Name>' -ForEach $script:Visuals {
    BeforeAll {
        $script:Visual = (Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json -AsHashtable).visual
        $script:Objects = $script:Visual.objects
        $script:Container = $script:Visual.visualContainerObjects
    }

    It 'gives every projection a nativeQueryRef' {
        $missing = foreach ($role in $script:Visual.query.queryState.Keys) {
            foreach ($projection in $script:Visual.query.queryState[$role].projections) {
                if (-not $projection.nativeQueryRef) { "$role/$($projection.queryRef)" }
            }
        }
        ($missing -join ', ') | Should -BeNullOrEmpty
    }

    It 'writes its label properties as PBIR expressions' {
        $bare = foreach ($pair in @(@($script:Container, 'title'), @($script:Objects, 'categoryLabels'),
                @($script:Objects, 'categoryAxis'), @($script:Objects, 'valueAxis'))) {
            $holder, $object = $pair
            if (-not $holder -or -not $holder.Contains($object)) { continue }
            foreach ($entry in $holder[$object]) {
                foreach ($key in $entry.properties.Keys) {
                    if (-not ($entry.properties[$key] -is [System.Collections.IDictionary])) { "$object.$key" }
                }
            }
        }
        ($bare -join ', ') | Should -BeNullOrEmpty
    }

    It 'shows a title or a category label on a card' -Skip:($Type -ne 'card') {
        $title = Get-ObjectProperty $script:Container 'title' 'show'
        $label = Get-ObjectProperty $script:Objects 'categoryLabels' 'show'
        ($title -eq 'true' -or $label -eq 'true') | Should -BeTrue
    }

    It 'shows a non-empty title on a chart or table' -Skip:($Type -in 'card', 'slicer') {
        Get-ObjectProperty $script:Container 'title' 'show' | Should -Be 'true'
        Get-ObjectProperty $script:Container 'title' 'text' | Should -Match "^'.+'$"
    }

    It 'shows both axis titles on a chart' -Skip:($Type -notin 'clusteredBarChart', 'clusteredColumnChart', 'lineChart') {
        Get-ObjectProperty $script:Objects 'categoryAxis' 'showAxisTitle' | Should -Be 'true'
        Get-ObjectProperty $script:Objects 'valueAxis' 'showAxisTitle' | Should -Be 'true'
    }
}
