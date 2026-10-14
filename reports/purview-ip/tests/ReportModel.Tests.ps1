#Requires -Version 7.0

<#
    The Power BI report (reports/purview-ip/report/) is built from the same CSVs
    the collectors write. These tests hold the report to that:

    - every table, column and measure a visual's query references must exist in
      the TMDL semantic model (a typo'd field reference is silently blank in
      Power BI, not an error, so nothing else would catch it);
    - every semantic-model table sourced directly from a CSV must import exactly
      the columns that CSV has, in the same order -- the same guarantee
      SampleData.Tests.ps1 holds the collectors to, extended to the report;
    - the report has the approved page (BRO-502: Label downgrades), the report
      filter pane carries the anonymize toggle, and the downgrade table keys its
      rows on the hidden ActivityExplorerEvents[RecordIdentity] column.

    Calculated tables (Labels, RetentionLabels, SensitivityAtRest, RetentionAtRest,
    Calendar, Workload, User Display) and the disconnected Anonymize Toggle and
    Measures tables are not sourced from a CSV of their own, so they are exempt from
    the CSV-mapping check. $CsvBackedTables below is the complete, explicit list of
    what *is* checked.

    A model column may carry a different name from the CSV column it imports
    (Users[UserId] imports `Id`; ContentExplorerSnapshot[SnapshotDate] imports
    `RunDate`) so that the visuals and measures built on the name keep working. The
    check is on `sourceColumn`, which is the CSV header.
#>

BeforeAll {
    $script:Root = Join-Path $PSScriptRoot '../../..' | Resolve-Path | Select-Object -ExpandProperty Path
    $script:Samples = Join-Path $script:Root 'reports/purview-ip/samples'
    $script:ReportRoot = Join-Path $script:Root 'reports/purview-ip/report'
    $script:TablesFolder = Join-Path $script:ReportRoot 'PurviewIPReport.SemanticModel/definition/tables'
    $script:PagesFolder = Join-Path $script:ReportRoot 'PurviewIPReport.Report/definition/pages'

    Import-Module (Join-Path $script:Root 'shared/M365ReportLibrary.psm1') -Force
    . (Join-Path $PSScriptRoot '../../guest-access/tests/TmdlModel.ps1')

    $script:Model = Get-TmdlModel -TablesFolder $script:TablesFolder

    # Table name -> the sample CSV its `m` partition imports.
    $script:CsvBackedTables = @{
        Users                    = 'users.csv'
        Policies                 = 'policies.csv'
        ActivityExplorerEvents   = 'activity-explorer-events.csv'
        ContentExplorerSnapshot  = 'content-explorer-snapshot.csv'
        CopilotAccessedResources = 'copilot-accessed-resources.csv'
    }

    function Get-VisualFieldReference {
        <#
            .SYNOPSIS
                Every {Entity, Property, Kind} a visual.json's query pulls from the
                semantic model, found by walking the parsed JSON for `Column` /
                `Measure` field containers (see the semanticQuery schema: a field is
                `{ "Column": { "Expression": { "SourceRef": { "Entity": ... } },
                "Property": ... } }`, or the same shape under `"Measure"`).
        #>
        param([Parameter(Mandatory)]$Node)

        $results = [System.Collections.Generic.List[hashtable]]::new()

        function Walk($n) {
            if ($n -is [System.Collections.IDictionary]) {
                foreach ($kind in 'Column', 'Measure') {
                    if ($n.ContainsKey($kind)) {
                        $inner = $n[$kind]
                        $entity = $inner.Expression.SourceRef.Entity
                        $property = $inner.Property
                        if ($entity -and $property) {
                            $results.Add(@{ Kind = $kind; Entity = $entity; Property = $property })
                        }
                    }
                }
                foreach ($key in $n.Keys) { Walk $n[$key] }
            }
            elseif (($n -is [System.Collections.IEnumerable]) -and ($n -isnot [string])) {
                foreach ($item in $n) { Walk $item }
            }
        }

        Walk $Node
        return , $results.ToArray()
    }

    $script:VisualFiles = @(Get-ChildItem -LiteralPath $script:PagesFolder -Filter 'visual.json' -Recurse -File)
}

AfterAll {
    Remove-Module M365ReportLibrary -Force -ErrorAction SilentlyContinue
}

Describe 'The semantic model has tables to check' {
    It 'parses the five CSV-backed tables plus the derived and helper tables' {
        $script:Model.Keys.Count | Should -BeGreaterOrEqual 14
    }

    It 'finds visual.json files to check' {
        $script:VisualFiles.Count | Should -BeGreaterThan 0
    }
}

Describe 'Every visual field reference resolves against the TMDL model' {
    BeforeAll {
        $script:AllReferences = foreach ($file in $script:VisualFiles) {
            $json = Get-Content -LiteralPath $file.FullName -Raw | ConvertFrom-Json -AsHashtable
            foreach ($ref in (Get-VisualFieldReference -Node $json)) {
                [pscustomobject]@{
                    File     = $file.FullName.Substring($script:ReportRoot.Length + 1)
                    Kind     = $ref.Kind
                    Entity   = $ref.Entity
                    Property = $ref.Property
                }
            }
        }
    }

    It 'has field references to check' {
        # Guards the tests below against silently passing because nothing was found.
        # BRO-502 replaced the eight old pages with one approved page (about 30 references).
        $script:AllReferences.Count | Should -BeGreaterThan 25
    }

    It 'references only tables that exist in the model' {
        $missing = @($script:AllReferences | Where-Object { -not $script:Model.ContainsKey($_.Entity) } |
                ForEach-Object { '{0}: {1}' -f $_.File, $_.Entity } | Sort-Object -Unique)
        $missing -join '; ' | Should -BeNullOrEmpty
    }

    It 'references only columns that exist on their table' {
        $missing = @($script:AllReferences | Where-Object { $_.Kind -eq 'Column' } | ForEach-Object {
                $table = $script:Model[$_.Entity]
                if ($table -and ($table.Columns.Name -notcontains $_.Property)) {
                    '{0}: {1}.{2}' -f $_.File, $_.Entity, $_.Property
                }
            } | Sort-Object -Unique)
        $missing -join '; ' | Should -BeNullOrEmpty
    }

    It 'references only measures that exist on their table' {
        $missing = @($script:AllReferences | Where-Object { $_.Kind -eq 'Measure' } | ForEach-Object {
                $table = $script:Model[$_.Entity]
                if ($table -and ($table.Measures -notcontains $_.Property)) {
                    '{0}: {1}.{2}' -f $_.File, $_.Entity, $_.Property
                }
            } | Sort-Object -Unique)
        $missing -join '; ' | Should -BeNullOrEmpty
    }
}

Describe 'Every CSV-backed table matches its sample CSV' {
    It 'checks <Table>' -ForEach @(
        @{ Table = 'Users' }
        @{ Table = 'Policies' }
        @{ Table = 'ActivityExplorerEvents' }
        @{ Table = 'ContentExplorerSnapshot' }
        @{ Table = 'CopilotAccessedResources' }
    ) {
        $script:Model.ContainsKey($Table) | Should -BeTrue -Because "the model should have a $Table table"

        $csvPath = Join-Path $script:Samples $script:CsvBackedTables[$Table]
        $csvHeader = Get-CsvHeaderColumn -Path $csvPath
        $csvHeader | Should -Not -BeNullOrEmpty -Because "$csvPath should be readable"

        # Only the imported (sourceColumn-backed) columns, in file order -- a
        # calculated column layered on top of a CSV-backed table (e.g.
        # ActivityExplorerEvents' Retention Label Key) is not part of the CSV and
        # is deliberately excluded here. So is a column the table's Power Query
        # adds after reading the CSV (Users' manager chain and pseudonyms).
        $queryAdded = @{ Users = @(
            'Manager Path', 'Manager Path Length', 'Level 1 Id', 'Level 2 Id', 'Level 3 Id', 'Level 4 Id', 'Level 5 Id', 'Level 1 Name', 'Level 2 Name', 'Level 3 Name', 'Level 4 Name', 'Level 5 Name', 'Anonymous Id', 'Anonymous Display Name', 'Level 1 Anonymous Name', 'Level 2 Anonymous Name', 'Level 3 Anonymous Name', 'Level 4 Anonymous Name', 'Level 5 Anonymous Name'
        ) }
        $modelColumns = @($script:Model[$Table].Columns | Where-Object { -not $_.IsCalculated } |
                ForEach-Object { $_.SourceColumn } | Where-Object { $_ -notin $queryAdded[$Table] })

        ($modelColumns -join ',') | Should -Be ($csvHeader -join ',')
    }

    It 'reads <Table> from the file name the collector writes' -ForEach @(
        @{ Table = 'Users'; File = 'users.csv' }
        @{ Table = 'Policies'; File = 'policies.csv' }
        @{ Table = 'ActivityExplorerEvents'; File = 'activity-explorer-events.csv' }
        @{ Table = 'ContentExplorerSnapshot'; File = 'content-explorer-snapshot.csv' }
        @{ Table = 'CopilotAccessedResources'; File = 'copilot-accessed-resources.csv' }
    ) {
        $tmdl = Get-Content -LiteralPath $script:Model[$Table].Path -Raw
        $tmdl | Should -Match ([regex]::Escape('CsvFolder & "\' + $File + '"'))
        $columnCount = (Get-CsvHeaderColumn -Path (Join-Path $script:Samples $File)).Count
        $tmdl | Should -Match "Columns=$columnCount,"
    }

    It 'names every CSV-backed table exactly once' {
        (@($script:CsvBackedTables.Keys) | Sort-Object) -join ',' |
            Should -Be ((@('Users', 'Policies', 'ActivityExplorerEvents', 'ContentExplorerSnapshot', 'CopilotAccessedResources') | Sort-Object) -join ',')
    }

    It 'no longer imports the columns the port dropped' {
        $dropped = @('Environment', 'TenantId', 'CollectedAt')
        $stillThere = @(foreach ($table in $script:CsvBackedTables.Keys) {
                $script:Model[$table].Columns | Where-Object { $_.SourceColumn -in $dropped } |
                    ForEach-Object { '{0}.{1}' -f $table, $_.SourceColumn }
            })
        $stillThere -join '; ' | Should -BeNullOrEmpty
    }

    It 'keeps the latest <Table> snapshot so relationship keys stay unique' -ForEach @(
        @{ Table = 'Users' }
        @{ Table = 'Policies' }
    ) {
        # users.csv and policies.csv append one block of rows per RunDate. UserId,
        # UserPrincipalName, and the label ObjectId derived from Policies are the
        # one-side of relationships, which Power BI requires to be unique.
        $tmdl = Get-Content -LiteralPath $script:Model[$Table].Path -Raw
        $tmdl | Should -Match 'KeepLatestRunDate\('
        $expressions = Get-Content -LiteralPath (Join-Path $script:ReportRoot 'PurviewIPReport.SemanticModel/definition/expressions.tmdl') -Raw
        $expressions | Should -Match 'expression KeepLatestRunDate'
        $expressions | Should -Match 'List\.IsEmpty'
    }

    It 'maps content explorer workload codes onto the activity explorer names' {
        # Export-ContentExplorerData accepts EXO/ODB/SPO; the collector writes those
        # codes. Activity Explorer's Workload values for the same locations are
        # Exchange, OneDrive and SharePoint. The shared slicer filters both facts.
        $tmdl = Get-Content -LiteralPath $script:Model['ContentExplorerSnapshot'].Path -Raw
        $tmdl | Should -Match 'MapContentExplorerWorkload\('
        $expressions = Get-Content -LiteralPath (Join-Path $script:ReportRoot 'PurviewIPReport.SemanticModel/definition/expressions.tmdl') -Raw
        $expressions | Should -Match 'if _ = "EXO" then "Exchange"'
        $expressions | Should -Match 'else if _ = "ODB" then "OneDrive"'
        $expressions | Should -Match 'else if _ = "SPO" then "SharePoint"'
    }
}

Describe 'The report has the approved Label downgrades page and its filter' {
    # BRO-502 replaced the eight pages (Overview ... Org decomposition) with the approved
    # page 1 from docs/design/purview-ip. The design has no slicers: the anonymize toggle
    # is a report-level filter, as on the Teams and Groups lifecycle pages (BRO-499).
    It 'lists exactly the approved page in pages.json' {
        $pages = Get-Content -LiteralPath (Join-Path $script:PagesFolder 'pages.json') -Raw | ConvertFrom-Json -AsHashtable
        @($pages.pageOrder) -join ',' | Should -Be 'label-downgrades'
    }

    It 'names the page Label downgrades' {
        $page = Get-Content -LiteralPath (Join-Path $script:PagesFolder 'label-downgrades/page.json') -Raw | ConvertFrom-Json -AsHashtable
        $page.displayName | Should -Be 'Label downgrades'
    }

    It 'carries the anonymize toggle in the report filter pane' {
        $report = Get-Content -LiteralPath (Join-Path $script:ReportRoot 'PurviewIPReport.Report/definition/report.json') -Raw | ConvertFrom-Json -AsHashtable
        $fields = @($report.filterConfig.filters | ForEach-Object { '{0}.{1}' -f $_.field.Column.Expression.SourceRef.Entity, $_.field.Column.Property })
        $fields | Should -Contain 'Anonymize Toggle.Anonymize Toggle'
    }

    It 'shows people only through a measure that honours the anonymize toggle' {
        # Users[DisplayName] or ActivityExplorerEvents[User] on a visual would bypass the toggle.
        $raw = @($script:VisualFiles | ForEach-Object {
                $json = Get-Content -LiteralPath $_.FullName -Raw | ConvertFrom-Json -AsHashtable
                foreach ($ref in (Get-VisualFieldReference -Node $json)) {
                    if ($ref.Kind -eq 'Column' -and ('{0}.{1}' -f $ref.Entity, $ref.Property) -in 'Users.DisplayName', 'Users.UserPrincipalName', 'ActivityExplorerEvents.User') { $_.FullName }
                }
            })
        $raw -join '; ' | Should -BeNullOrEmpty
        $tmdl = Get-Content -LiteralPath $script:Model['_Measures'].Path -Raw
        $tmdl | Should -Match "measure 'Downgrade Person'[\s\S]*?\[Anonymize Toggle Value\] = 1"
    }

    It 'keys each downgrade table row on the hidden RecordIdentity column' {
        $table = Get-Content -LiteralPath (Join-Path $script:PagesFolder 'label-downgrades/visuals/downgrades-table/visual.json') -Raw | ConvertFrom-Json -AsHashtable
        $key = @($table.visual.query.queryState.Values.projections | Where-Object { $_.queryRef -eq 'ActivityExplorerEvents.RecordIdentity' })
        $key.Count | Should -Be 1
        $key[0].hidden | Should -BeTrue
    }
}
