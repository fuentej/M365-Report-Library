#Requires -Version 7.0

<#
    BRO-481: every number on every page of these reports says what it is. Power BI
    Desktop reads formatting properties only in the PBIR expression form
    ({ "expr": { "Literal": { "Value": ... } } }); a bare JSON true or string is not
    rendered, which is how the cards in five of these reports lost their titles.

    A legacy card ("card") has no title by default (suppressDefaultTitle in its
    capabilities), and its category label is the field's display name resolved at
    render time. In Desktop that label did not appear on the Purview cards (BRO-461:
    "a card shows 25 and nothing says what it counts"), whose show flag was true and
    whose projections carried queryRef "_Measures.[DLP Matches]" and nativeQueryRef
    the same string. So the label on every card is the visual title, written as
    literal text in the visual: title.show true, title.text a non-empty string, and
    categoryLabels.show false so the text is not shown twice. A card that relies on
    its category label fails here.

    Checks, per visual.json:
    - every projection has a nativeQueryRef;
    - a Column or Measure projection is named the way Desktop writes it: queryRef
      "Entity.Property", nativeQueryRef "Property";
    - every formatting property under title, categoryLabels, categoryAxis and
      valueAxis is an expression, not a bare value;
    - a card shows a non-empty title as its label, with the category label off;
    - a chart or table shows a non-empty title, and a chart shows both axis titles;
    - a matrix keeps measures out of Rows and Columns (they take Grouping fields only);
    - each report has a .platform file, and each .pbip names its schema.

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

    It 'names each field the way Desktop writes it (queryRef Entity.Property, nativeQueryRef Property)' {
        $wrong = foreach ($role in $script:Visual.query.queryState.Keys) {
            foreach ($projection in $script:Visual.query.queryState[$role].projections) {
                $kind = @($projection.field.Keys)[0]
                if ($kind -notin 'Column', 'Measure') { continue }
                $entity = $projection.field[$kind].Expression.SourceRef.Entity
                $property = $projection.field[$kind].Property
                if ($projection.queryRef -cne "$entity.$property" -or $projection.nativeQueryRef -cne $property) {
                    "$role/$($projection.queryRef) ($($projection.nativeQueryRef))"
                }
            }
        }
        ($wrong -join ', ') | Should -BeNullOrEmpty
    }

    It 'labels a card with its title, not its category label' -Skip:($Type -ne 'card') {
        Get-ObjectProperty $script:Container 'title' 'show' | Should -Be 'true'
        Get-ObjectProperty $script:Container 'title' 'text' | Should -Match "^'.*\S.*'$"
        Get-ObjectProperty $script:Objects 'categoryLabels' 'show' | Should -Be 'false'
    }

    It 'shows a non-empty title on a chart or table' -Skip:($Type -in 'card', 'slicer') {
        Get-ObjectProperty $script:Container 'title' 'show' | Should -Be 'true'
        Get-ObjectProperty $script:Container 'title' 'text' | Should -Match "^'.+'$"
    }

    It 'keeps measures out of the Rows and Columns of a matrix' -Skip:($Type -ne 'pivotTable') {
        $misplaced = foreach ($role in 'Rows', 'Columns') {
            if (-not $script:Visual.query.queryState.Contains($role)) { continue }
            foreach ($projection in $script:Visual.query.queryState[$role].projections) {
                if ($projection.field.Contains('Measure')) { "$role/$($projection.queryRef)" }
            }
        }
        ($misplaced -join ', ') | Should -BeNullOrEmpty
    }

    It 'shows both axis titles on a chart' -Skip:($Type -notin 'clusteredBarChart', 'clusteredColumnChart', 'lineChart') {
        Get-ObjectProperty $script:Objects 'categoryAxis' 'showAxisTitle' | Should -Be 'true'
        Get-ObjectProperty $script:Objects 'valueAxis' 'showAxisTitle' | Should -Be 'true'
    }
}

$script:ReportFiles = foreach ($report in $script:Reports) {
    $reportRoot = Join-Path $script:RepoRoot "reports/$($report.Split('/')[0])/report"
    @{
        Report   = $report.Split('/')[0]
        Platform = Join-Path $script:RepoRoot "reports/$report/.platform"
        Pbip     = @(Get-ChildItem -LiteralPath $reportRoot -Filter '*.pbip' -File | ForEach-Object FullName)
    }
}

Describe 'Report item files for <Report>' -ForEach $script:ReportFiles {
    It 'has a .platform file with type Report' {
        Test-Path -LiteralPath $Platform | Should -BeTrue
        (Get-Content -LiteralPath $Platform -Raw | ConvertFrom-Json).metadata.type | Should -Be 'Report'
    }

    It 'names the pbipProperties schema in its .pbip' {
        $Pbip.Count | Should -Be 1
        (Get-Content -LiteralPath $Pbip[0] -Raw | ConvertFrom-Json).'$schema' | Should -Match 'pbipProperties/1\.0\.0/schema\.json$'
    }
}
