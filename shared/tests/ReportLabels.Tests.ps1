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
    the same string. So the cards in these reports carry the visual title as literal
    text. BRO-517 aligns this test with the BRO-481 rule as written: a card passes
    with either a shown category label or a non-empty title.

    Decision D-014 (BRO-517) makes the test accept the approved page designs as drawn:
    - a card whose measures all return text shows no number and needs no label. TMDL
      does not store a measure's data type (TOM Measure.DataType is read-only,
      computed by the engine), so a measure with no formatString counts as text;
    - a number card with no label of its own passes when a textbox with text names
      it from the same visual group (parentGroupName), or heads its strip: the
      textbox sits inside the same shape as the card, to the card's left, on the
      same rows;
    - a blank table column header passes when the visible column to its left has a
      readable header, so one header spans both columns.

    Checks, per visual.json:
    - every projection has a nativeQueryRef;
    - a Column or Measure projection is named the way Desktop writes it: queryRef
      "Entity.Property", nativeQueryRef "Property";
    - every formatting property under title, categoryLabels, categoryAxis and
      valueAxis is an expression, not a bare value;
    - a card (card, multiRowCard, kpi, cardVisual) shows a category label or a
      non-empty title (the BRO-481 rule), shows only text, or is named by a textbox
      (D-014);
    - a table or matrix gives every visible column a readable header, or shares the
      header of the column to its left (D-014);
    - any other data visual (a chart) shows a non-empty title, and a title on every
      axis it shows; textboxes, shapes, images, buttons and slicers show no number
      and need neither;
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
# Visuals that show no number need no label; tables are labelled by their column headers.
$script:DecorativeTypes = @('textbox', 'shape', 'basicShape', 'image', 'actionButton', 'slicer')
$script:CardTypes = @('card', 'multiRowCard', 'kpi', 'cardVisual')
$script:TableTypes = @('tableEx', 'pivotTable')
$script:NoTitleTypes = $script:DecorativeTypes + $script:CardTypes + $script:TableTypes
# "Entity.Measure" -> $true when the measure has a formatString (it returns a number).
function Get-MeasureFormats {
    param([string]$ReportFolder)
    $pbir = Get-Content -LiteralPath (Join-Path $ReportFolder 'definition.pbir') -Raw | ConvertFrom-Json
    $model = Join-Path $ReportFolder $pbir.datasetReference.byPath.path
    $formats = @{}
    foreach ($file in (Get-ChildItem -LiteralPath (Join-Path $model 'definition') -Recurse -File -Filter '*.tmdl')) {
        $table = $null; $measure = $null
        foreach ($line in (Get-Content -LiteralPath $file.FullName)) {
            if ($line -match "^table\s+('(?:[^']|'')+'|\S+)") {
                $table = $Matches[1].Trim("'").Replace("''", "'"); $measure = $null
            } elseif ($line -match "^\s+measure\s+('(?:[^']|'')+'|[^\s=]+)") {
                $measure = "$table." + $Matches[1].Trim("'").Replace("''", "'")
                $formats[$measure] = $false
            } elseif ($line -match '^\s+(column|partition|hierarchy|calculationGroup)\b') {
                $measure = $null
            } elseif ($measure -and $line -match '^\s+formatString(Definition)?\s*[:=]') {
                $formats[$measure] = $true
            }
        }
    }
    $formats
}

$script:Visuals = foreach ($report in $script:Reports) {
    $reportFolder = Join-Path $script:RepoRoot "reports/$report"
    $formats = Get-MeasureFormats $reportFolder
    $pages = Join-Path $reportFolder 'definition/pages'
    foreach ($file in (Get-ChildItem -LiteralPath $pages -Recurse -File -Filter 'visual.json')) {
        $type = (Get-Content -LiteralPath $file.FullName -Raw | ConvertFrom-Json -AsHashtable).visual.visualType
        @{
            Report = $report.Split('/')[0]; Name = $file.Directory.Name; Path = $file.FullName; Type = $type
            VisualsFolder = $file.Directory.Parent.FullName; MeasureFormats = $formats
        }
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

    function Get-TextboxText {
        param([System.Collections.IDictionary]$Visual)
        $runs = foreach ($paragraph in $Visual.objects.general[0].properties.paragraphs) {
            foreach ($run in $paragraph.textRuns) { $run.value }
        }
        ($runs -join ' ').Trim()
    }

    function Test-Inside {
        param($Inner, $Outer)
        $Inner.x -ge $Outer.x -and $Inner.y -ge $Outer.y -and
            $Inner.x + $Inner.width -le $Outer.x + $Outer.width -and
            $Inner.y + $Inner.height -le $Outer.y + $Outer.height
    }

    # D-014: a textbox with text in the card's visual group, or heading the card's
    # strip (inside the same shape, to the card's left, on the same rows).
    function Test-NamedByTextbox {
        param([System.Collections.IDictionary]$Container, [string]$VisualsFolder)
        $siblings = foreach ($file in (Get-ChildItem -LiteralPath $VisualsFolder -Recurse -File -Filter 'visual.json')) {
            Get-Content -LiteralPath $file.FullName -Raw | ConvertFrom-Json -AsHashtable
        }
        $card = $Container.position
        $texts = @($siblings | Where-Object { $_.visual.visualType -eq 'textbox' -and (Get-TextboxText $_.visual) })
        if ($Container.parentGroupName -and
            ($texts | Where-Object { $_.parentGroupName -eq $Container.parentGroupName })) { return $true }
        $strips = @($siblings | Where-Object { $_.visual.visualType -in 'shape', 'basicShape' -and (Test-Inside $card $_.position) })
        foreach ($text in $texts) {
            $box = $text.position
            $heads = $box.x + $box.width -le $card.x -and
                $box.y -lt $card.y + $card.height -and $card.y -lt $box.y + $box.height
            if ($heads -and ($strips | Where-Object { Test-Inside $box $_.position })) { return $true }
        }
        $false
    }
}

Describe 'Labels on <Report> / <Name>' -ForEach $script:Visuals {
    BeforeAll {
        $script:VisualContainer = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json -AsHashtable
        $script:Visual = $script:VisualContainer.visual
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

    It 'labels a card with a shown category label or a non-empty title' -Skip:($Type -notin $script:CardTypes) {
        $titled = (Get-ObjectProperty $script:Container 'title' 'show') -eq 'true' -and
            (Get-ObjectProperty $script:Container 'title' 'text') -match "^'.*\S.*'$"
        $labelObject = if ($Type -eq 'cardVisual') { 'label' } else { 'categoryLabels' }
        $labelled = (Get-ObjectProperty $script:Objects $labelObject 'show') -ne 'false'
        $projections = @(foreach ($role in $script:Visual.query.queryState.Keys) {
            $script:Visual.query.queryState[$role].projections
        })
        # D-014: every field is a measure with no formatString, so the card shows text.
        $textOnly = $projections.Count -gt 0 -and -not ($projections | Where-Object {
            -not $_.field.Contains('Measure') -or $MeasureFormats[$_.queryRef] -ne $false
        })
        $named = $titled -or $labelled -or $textOnly -or (Test-NamedByTextbox $script:VisualContainer $VisualsFolder)
        $named | Should -BeTrue -Because 'a card names its number with a category label, a title or a textbox that heads it'
    }

    It 'gives every table column a readable header' -Skip:($Type -notin $script:TableTypes) {
        $blank = foreach ($role in $script:Visual.query.queryState.Keys) {
            $leftReadable = $false
            foreach ($projection in $script:Visual.query.queryState[$role].projections) {
                if ($projection.hidden) { continue }
                $header = if ($projection.Contains('displayName')) { $projection.displayName } else { $projection.nativeQueryRef }
                $readable = $header -match '\S'
                # D-014: a blank header sits under the readable header of the column to its left.
                if (-not $readable -and -not $leftReadable) { "$role/$($projection.queryRef)" }
                $leftReadable = $readable
            }
        }
        ($blank -join ', ') | Should -BeNullOrEmpty
    }

    It 'shows a non-empty title on a chart' -Skip:($Type -in $script:NoTitleTypes) {
        Get-ObjectProperty $script:Container 'title' 'show' | Should -Be 'true'
        Get-ObjectProperty $script:Container 'title' 'text' | Should -Match "^'.*\S.*'$"
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

    It 'shows a title on every axis the chart shows' -Skip:($Type -in $script:NoTitleTypes) {
        $untitled = foreach ($axis in 'categoryAxis', 'valueAxis') {
            if (-not $script:Objects -or -not $script:Objects.Contains($axis)) { continue }
            if ((Get-ObjectProperty $script:Objects $axis 'show') -eq 'false') { continue }
            if ((Get-ObjectProperty $script:Objects $axis 'showAxisTitle') -ne 'true') { $axis }
        }
        ($untitled -join ', ') | Should -BeNullOrEmpty
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
