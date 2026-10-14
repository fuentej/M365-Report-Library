#Requires -Version 7.0

<#
    The Power BI report (reports/copilot-studio-agents/report/) is built from the same CSVs
    the collectors write. These tests hold the report to that:

    - every table, column and measure a visual's query references must exist in the
      TMDL semantic model (a mistyped field reference is silently blank in Power BI,
      not an error, so nothing else would catch it);
    - every semantic-model table sourced directly from a CSV must have exactly the
      columns that CSV has, in the same order, for the populated sample and for the
      header-only file a tenant without the licence gets;
    - every relationship names columns that exist, and the inventory totals
      (capabilitiesCounts) are shown next to the returned counts.

    AgentsCurrent and EnvironmentsCurrent (latest-snapshot imports of their source CSVs) and the helper tables that are
    not sourced from any CSV (DateDim and the AnonymizeMode toggle) are exempt from the
    CSV-mapping check. $CsvBackedTables below is the complete list of what is checked.
#>

BeforeAll {
    $script:Root = Join-Path $PSScriptRoot '../../..' | Resolve-Path | Select-Object -ExpandProperty Path
    $script:Samples = Join-Path $script:Root 'reports/copilot-studio-agents/samples'
    $script:ReportRoot = Join-Path $script:Root 'reports/copilot-studio-agents/report'
    $script:DefinitionFolder = Join-Path $script:ReportRoot 'CopilotStudioAgents.SemanticModel/definition'
    $script:TablesFolder = Join-Path $script:DefinitionFolder 'tables'
    $script:PagesFolder = Join-Path $script:ReportRoot 'CopilotStudioAgents.Report/definition/pages'

    Import-Module (Join-Path $script:Root 'shared/M365ReportLibrary.psm1') -Force
    . (Join-Path $PSScriptRoot '../../guest-access/tests/TmdlModel.ps1')

    $script:Model = Get-TmdlModel -TablesFolder $script:TablesFolder

    # Table name -> the sample CSV its `m` partition imports.
    $script:CsvBackedTables = [ordered]@{
        Environments       = 'environments.csv'
        Agents             = 'agents.csv'
        AgentConnectors    = 'agent-connectors.csv'
        AgentComponents    = 'agent-components.csv'
        AgentModifications = 'agent-modifications.csv'
        AgentAuditEvents   = 'agent-audit-events.csv'
    }

    function Get-VisualFieldReference {
        <#
            .SYNOPSIS
                Every {Entity, Property, Kind} a visual.json's query pulls from the
                semantic model, found by walking the parsed JSON for `Column` /
                `Measure` field containers (semanticQuery schema: a field is
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
    It 'parses the six CSV-backed tables plus the calculated and helper tables' {
        $script:Model.Keys.Count | Should -Be 10
        foreach ($name in $script:CsvBackedTables.Keys + @('AgentsCurrent', 'EnvironmentsCurrent', 'DateDim', 'AnonymizeMode')) {
            $script:Model.ContainsKey($name) | Should -BeTrue -Because "the model should have a $name table"
        }
    }

    It 'finds visual.json files to check' {
        $script:VisualFiles.Count | Should -BeGreaterThan 50
    }

    It 'lists every table in model.tmdl' {
        $refs = @(Get-Content -LiteralPath (Join-Path $script:DefinitionFolder 'model.tmdl') |
                Select-String -Pattern '^ref table (.+)$' | ForEach-Object { ConvertFrom-TmdlName -Name $_.Matches[0].Groups[1].Value })
        ($refs | Sort-Object) -join ',' | Should -Be (($script:Model.Keys | Sort-Object) -join ',')
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
        # The one approved page (BRO-504) makes about 25 field references.
        $script:AllReferences.Count | Should -BeGreaterThan 20
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

    It 'matches queryRef to the Entity.Property it projects' {
        $bad = [System.Collections.Generic.List[string]]::new()
        foreach ($file in $script:VisualFiles) {
            $json = Get-Content -LiteralPath $file.FullName -Raw | ConvertFrom-Json -AsHashtable
            foreach ($role in $json.visual.query.queryState.Values) {
                foreach ($projection in $role.projections) {
                    $inner = if ($projection.field.ContainsKey('Column')) { $projection.field.Column } else { $projection.field.Measure }
                    $expected = '{0}.{1}' -f $inner.Expression.SourceRef.Entity, $inner.Property
                    if ($projection.queryRef -ne $expected) { $bad.Add("$($file.Name): $($projection.queryRef) != $expected") }
                }
            }
        }
        $bad -join '; ' | Should -BeNullOrEmpty
    }

    It 'projects no raw agent, environment or person column' {
        # Names and ids reach a visual only through the *_Display measures and the
        # pseudonym columns, which honour the Anonymize toggle.
        $raw = @($script:AllReferences | Where-Object {
                $_.Kind -eq 'Column' -and (
                    ($_.Property -in 'DisplayName', 'Name', 'BotSchemaName') -or
                    ($_.Property -in 'OwnerId', 'CreatedBy', 'ModifiedBy', 'UserId') -or
                    ($_.Property -eq 'AgentId' -and $_.Entity -eq 'Agents'))
            } | ForEach-Object { '{0}: {1}.{2}' -f $_.File, $_.Entity, $_.Property } | Sort-Object -Unique)
        $raw -join '; ' | Should -BeNullOrEmpty
    }
}

Describe 'Every CSV-backed table matches its sample CSV' {
    It 'checks <Table>' -ForEach @(
        @{ Table = 'Environments' }
        @{ Table = 'Agents' }
        @{ Table = 'AgentConnectors' }
        @{ Table = 'AgentComponents' }
        @{ Table = 'AgentModifications' }
        @{ Table = 'AgentAuditEvents' }
    ) {
        $script:Model.ContainsKey($Table) | Should -BeTrue -Because "the model should have a $Table table"

        $csvPath = Join-Path $script:Samples $script:CsvBackedTables[$Table]
        $csvHeader = Get-CsvHeaderColumn -Path $csvPath
        $csvHeader | Should -Not -BeNullOrEmpty -Because "$csvPath should be readable"

        # Only the imported (sourceColumn-backed) columns, in file order: a calculated
        # column layered on a CSV-backed table (PublishState, DaysSinceModified)
        # is not part of the CSV and is excluded, as it would be from an Import-Csv header.
        $modelColumns = @($script:Model[$Table].Columns | Where-Object { -not $_.IsCalculated } |
                ForEach-Object { $_.SourceColumn })

        ($modelColumns -join ',') | Should -Be ($csvHeader -join ',')
    }

    It 'names every CSV-backed table exactly once' {
        $script:CsvBackedTables.Keys.Count | Should -Be 6
        (@($script:CsvBackedTables.Values) | Sort-Object -Unique).Count | Should -Be 6
    }

    It 'reads <Table> from the CsvFolder parameter and the CSV named in the map' -ForEach @(
        @{ Table = 'Environments' }, @{ Table = 'Agents' }, @{ Table = 'AgentConnectors' },
        @{ Table = 'AgentComponents' }, @{ Table = 'AgentModifications' }, @{ Table = 'AgentAuditEvents' }
    ) {
        $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder "$Table.tmdl") -Raw
        $tmdl | Should -Match ([regex]::Escape('File.Contents(CsvFolder & "\' + $script:CsvBackedTables[$Table] + '")'))
    }

    It 'keeps the CsvFolder parameter as the only parameter query' {
        $expressions = Get-Content -LiteralPath (Join-Path $script:DefinitionFolder 'expressions.tmdl') -Raw
        ([regex]::Matches($expressions, '(?m)^expression ')).Count | Should -Be 1
        $expressions | Should -Match '(?m)^expression CsvFolder = '
    }
}

Describe 'The header-only connector files of GCC and GCC High still fit the model' {
    # samples/gcc and samples/gcchigh hold agent-connectors.csv with the header only,
    # because the connector inventory is not available in those clouds.
    It 'matches the model columns of AgentConnectors to samples/<Cloud>/agent-connectors.csv' -ForEach @(
        @{ Cloud = 'gcc' }, @{ Cloud = 'gcchigh' }
    ) {
        $path = Join-Path $script:Samples "$Cloud/agent-connectors.csv"
        $header = Get-CsvHeaderColumn -Path $path
        $header | Should -Not -BeNullOrEmpty
        @(Get-Content -LiteralPath $path).Count | Should -Be 1 -Because 'a header-only file means not collected'

        $modelColumns = @($script:Model['AgentConnectors'].Columns | Where-Object { -not $_.IsCalculated } | ForEach-Object { $_.SourceColumn })
        ($modelColumns -join ',') | Should -Be ($header -join ',')
    }

    It 'leaves a source with no rows blank, never zero, in the snapshot measures' {
        foreach ($name in 'Environments', 'Agents', 'AgentConnectors', 'AgentComponents', 'AgentModifications') {
            $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder "$name.tmdl") -Raw
            $tmdl | Should -Match 'IF\(ISBLANK\(Snap\), BLANK\(\)' -Because "$name measures must return blank when the file is header-only"
        }
        $audit = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'AgentAuditEvents.tmdl') -Raw
        $audit | Should -Match ([regex]::Escape('IF(COUNTROWS(ALL(AgentAuditEvents)) = 0, BLANK()'))
    }
}

Describe 'The inventory totals sit next to the returned counts' {
    It 'has a total (capabilitiesCounts) measure for each listed count' {
        $measures = $script:Model['Agents'].Measures
        $measures | Should -Contain 'Connectors Listed'
        $measures | Should -Contain 'Connectors Total (capabilitiesCounts)'
        $measures | Should -Contain 'Connector Operations Listed'
        $measures | Should -Contain 'Connector Operations Total (capabilitiesCounts)'
        $measures | Should -Contain 'Agents With Partial Connector Rows'
    }

    It 'counts connectors on the passport from the complete capabilitiesCounts total' {
        # BRO-504 removed the connectors page that showed listed and total counts side by side.
        # The passport shows the total, which the inventory does not cap at a random 200.
        $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'AgentsCurrent.tmdl') -Raw
        $tmdl | Should -Match ([regex]::Escape('SELECTEDVALUE(AgentsCurrent[CapabilitiesDistinctConnectors])'))
        $tmdl | Should -Not -Match 'AgentsCurrent\[ListedConnectorCount\]'
    }

    It 'flags an agent whose listed operations fall below the total' {
        $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'Agents.tmdl') -Raw
        $tmdl | Should -Match ([regex]::Escape('Agents[ListedConnectorOperationCount] < Agents[CapabilitiesDistinctConnectorOperations]'))
    }

    It 'flags an agent whose listed connectors fall below the capabilitiesCounts total' {
        # The inventory caps each resource type at a random 200. Connector rows are partial
        # when either the connector count or the operation count is short of its total.
        # https://learn.microsoft.com/microsoft-copilot-studio/admin-agent-inventory#known-limitations
        $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'Agents.tmdl') -Raw
        $tmdl | Should -Match ([regex]::Escape('Agents[ListedConnectorCount] < Agents[CapabilitiesDistinctConnectors]'))
        $tmdl | Should -Match 'NOT ISBLANK\(Agents\[ConnectorRowsPartial\]\)'
    }
}

Describe 'Empty values are Unknown, never zero' {
    It 'sums only the values that came back and is blank when none did' {
        $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'Agents.tmdl') -Raw
        $tmdl | Should -Match ([regex]::Escape('IF(ISBLANK(Snap) || Known = 0, BLANK()'))
    }

    It 'labels an empty authentication, channels or sharing flag Unknown' {
        $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'Agents.tmdl') -Raw
        $tmdl | Should -Match ([regex]::Escape('IF(ISBLANK(Agents[Authentication]), "Unknown"'))
        $tmdl | Should -Match ([regex]::Escape('IF(ISBLANK(Agents[Channels]), "Unknown"'))
        $tmdl | Should -Match ([regex]::Escape('IF(ISBLANK(Agents[ViewerEntireTenant]), "Unknown"'))
    }

    It 'derives days since modified as blank when ModifiedOn is empty' {
        $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'AgentModifications.tmdl') -Raw
        $tmdl | Should -Match ([regex]::Escape('IF(ISBLANK(AgentModifications[ModifiedOn]), BLANK(), DATEDIFF('))
        $script:Model['AgentModifications'].Measures | Should -Contain 'Unknown Last Modified'
    }

    It 'counts days since modified from the UTC date, not the clock time' {
        # DateTime.FromText with a literal Z keeps the UTC clock time, and RunDate is midnight.
        # DATEDIFF on the raw datetime is negative when the modification is later that same day.
        # https://learn.microsoft.com/powerquery-m/datetime-fromtext
        $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'AgentModifications.tmdl') -Raw
        $tmdl | Should -Match ([regex]::Escape('DATEDIFF(DATE(YEAR(AgentModifications[ModifiedOn]), MONTH(AgentModifications[ModifiedOn]), DAY(AgentModifications[ModifiedOn])), AgentModifications[RunDate], DAY)'))
        $tmdl | Should -Not -Match ([regex]::Escape('DATEDIFF(AgentModifications[ModifiedOn], AgentModifications[RunDate], DAY)'))
    }

    It 'never compares a nullable boolean to FALSE without excluding blanks' {
        # In DAX, BLANK() = FALSE() is true, so an empty flag would be counted as No.
        $offenders = [System.Collections.Generic.List[string]]::new()
        foreach ($file in Get-ChildItem -LiteralPath $script:TablesFolder -Filter '*.tmdl') {
            $lines = Get-Content -LiteralPath $file.FullName
            for ($i = 0; $i -lt $lines.Count; $i++) {
                if ($lines[$i] -match '(\w+\[\w+\]) = FALSE\(\)') {
                    $column = $Matches[1]
                    if ($lines[$i] -notmatch ('ISBLANK\(' + [regex]::Escape($column) + '\)')) {
                        $offenders.Add("$($file.Name):$($i + 1) $column")
                    }
                }
            }
        }
        $offenders -join '; ' | Should -BeNullOrEmpty
    }
}

Describe 'Relationships and the anonymize toggle' {
    BeforeAll {
        $text = Get-Content -LiteralPath (Join-Path $script:DefinitionFolder 'relationships.tmdl') -Raw
        $script:Relationships = @([regex]::Matches($text, '(?ms)^relationship (?<id>\S+)\s+fromColumn: (?<from>\S+)\s+toColumn: (?<to>\S+)') |
                ForEach-Object { @{ Id = $_.Groups['id'].Value; From = $_.Groups['from'].Value; To = $_.Groups['to'].Value } })
    }

    It 'declares relationships' {
        $script:Relationships.Count | Should -Be 13
    }

    It 'uses unique relationship ids' {
        @($script:Relationships.Id | Sort-Object -Unique).Count | Should -Be $script:Relationships.Count
    }

    It 'names a column that exists at both ends of <From> -> <To>' -ForEach @(
        $text = Get-Content -LiteralPath (Join-Path $PSScriptRoot '../report/CopilotStudioAgents.SemanticModel/definition/relationships.tmdl') -Raw
        [regex]::Matches($text, '(?ms)^relationship \S+\s+fromColumn: (?<from>\S+)\s+toColumn: (?<to>\S+)') |
            ForEach-Object { @{ From = $_.Groups['from'].Value; To = $_.Groups['to'].Value } }
    ) {
        foreach ($end in $From, $To) {
            $tableName, $columnName = $end -split '\.', 2
            $script:Model.ContainsKey($tableName) | Should -BeTrue -Because "$end names table $tableName"
            $script:Model[$tableName].Columns.Name | Should -Contain $columnName -Because "$end should be a column"
        }
    }

    It 'points the many side at the one side, as the TMDL overview does' {
        # fromColumn is the many side and toColumn the one side
        # (https://learn.microsoft.com/analysis-services/tmdl/tmdl-overview#relationship).
        $oneSide = 'DateDim.Date', 'AgentsCurrent.AgentId', 'EnvironmentsCurrent.EnvironmentId'
        $wrong = @($script:Relationships | Where-Object { $_.To -notin $oneSide -or $_.From -in $oneSide })
        ($wrong | ForEach-Object { "$($_.From) -> $($_.To)" }) -join '; ' | Should -BeNullOrEmpty
    }

    It 'offers Show names and Anonymize on the toggle table' {
        $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'AnonymizeMode.tmdl') -Raw
        $tmdl | Should -Match ([regex]::Escape('{{"Show names"}, {"Anonymize"}}'))
    }

    It 'switches the agent and environment display names on the toggle and uses Id-derived pseudonyms' {
        $agents = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'AgentsCurrent.tmdl') -Raw
        $agents | Should -Match ([regex]::Escape('SELECTEDVALUE(AnonymizeMode[Mode], "Show names")'))
        $agents | Should -Match ([regex]::Escape('IF(Mode = "Anonymize", Pseudo, RealName)'))
        $agents | Should -Match ([regex]::Escape('VAR IdText = AgentsCurrent[AgentId]'))
        $environments = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'EnvironmentsCurrent.tmdl') -Raw
        $environments | Should -Match ([regex]::Escape('IF(Mode = "Anonymize", Pseudo, RealName)'))
        $environments | Should -Match ([regex]::Escape('VAR IdText = EnvironmentsCurrent[EnvironmentId]'))
    }

    It 'switches owner, creator and modifier ids on the toggle' {
        foreach ($pair in @(@('Agents', 'Owner (Display)'), @('Agents', 'Creator (Display)'), @('AgentModifications', 'Modified By (Display)'))) {
            $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder "$($pair[0]).tmdl") -Raw
            $tmdl | Should -Match ([regex]::Escape("measure '$($pair[1])'"))
            $tmdl | Should -Match ([regex]::Escape('IF(Mode = "Anonymize"'))
        }
    }
}

Describe 'The page says what the data cannot show' {
    # BRO-504 replaced the seven pages these checks used to read with the Agent passports page.
    BeforeAll {
        $script:Passports = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'AgentsCurrent.tmdl') -Raw
        $script:PageText = (@(Get-ChildItem -LiteralPath $script:PagesFolder -Filter visual.json -Recurse |
                    ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw })) -join "`n"
    }

    It 'shows sharing as counts of people and groups, never names' {
        $script:Passports | Should -Match ([regex]::Escape('SELECTEDVALUE(AgentsCurrent[ViewerUserCount])'))
        $script:Passports | Should -Match ([regex]::Escape('SELECTEDVALUE(AgentsCurrent[ViewerGroupCount])'))
        $script:Passports | Should -Match '" people"'
    }

    It 'reads nothing from the connector file, which GCC and GCC High do not collect' {
        $script:PageText | Should -Not -Match '"Entity": "AgentConnectors"'
        $script:Passports | Should -Not -Match 'AgentConnectors\['
    }

    It 'shows an owner id, not a name, and leaves owner names to Premium' {
        # There is no directory file in this report, so the owner has no name to show.
        $script:Passports | Should -Match ([regex]::Escape('"owner id " & LEFT(OwnerKey, 4) & "..."'))
        $script:PageText | Should -Match 'Owner names, conversations, what each agent was asked'
        $script:PageText | Should -Match 'Premium edition at backlogkrew.com'
    }

    It 'says when the audit log was not collected instead of saying the records agree' {
        $script:Passports | Should -Match ([regex]::Escape('AuditRows = 0, "agent-audit-events.csv has no rows'))
        $script:Passports | Should -Match '"BotUpdateOperation-BotPublish"'
    }

    It 'shows agent and environment names through the anonymize toggle' {
        $script:Passports | Should -Match ([regex]::Escape('VAR RawName = [Agent Display Name]'))
        $script:Passports | Should -Match ([regex]::Escape('CALCULATE([Environment Display Name], TREATAS({EnvKey}, EnvironmentsCurrent[EnvironmentId]))'))
        $script:Passports | Should -Match ([regex]::Escape('IF(Mode = "Anonymize", EnvironmentsCurrent[Pseudonym], EnvironmentsCurrent[DisplayName])'))
        $script:Passports | Should -Match ([regex]::Escape('IF(Mode = "Anonymize", AgentsCurrent[Pseudonym], AgentsCurrent[DisplayName])'))
    }

    It 'shows an empty sign-in, channel list or editor count as Unknown' {
        $script:Passports | Should -Match ([regex]::Escape('ISBLANK(Auth), "Unknown"'))
        $script:Passports | Should -Match ([regex]::Escape('IF(ISBLANK(ChannelList), "Unknown"'))
        $script:Passports | Should -Match ([regex]::Escape('ISBLANK(EditorUsers) && ISBLANK(EditorGroups), "Unknown"'))
    }

    It 'labels an empty connector use or component category Unknown' {
        # A tabular connector is stored as one row with the operation columns empty.
        # https://learn.microsoft.com/microsoft-copilot-studio/admin-agent-inventory#connector-properties
        $connectors = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'AgentConnectors.tmdl') -Raw
        $connectors | Should -Match ([regex]::Escape('IF(ISBLANK(AgentConnectors[UsedAs]), "Unknown", AgentConnectors[UsedAs])'))
        $components = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'AgentComponents.tmdl') -Raw
        $components | Should -Match ([regex]::Escape('IF(ISBLANK(AgentComponents[Category]), "Unknown", AgentComponents[Category])'))
    }
}

Describe 'The report has the approved Agent passports page and its filter' {
    BeforeAll {
        $script:PageVisuals = Join-Path $script:PagesFolder 'agent-passports/visuals'
        $script:Passports = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'AgentsCurrent.tmdl') -Raw
    }

    It 'lists only the agent-passports page' {
        $pages = Get-Content -LiteralPath (Join-Path $script:PagesFolder 'pages.json') -Raw | ConvertFrom-Json
        @($pages.pageOrder) -join ',' | Should -Be 'agent-passports'
        (Get-Content -LiteralPath (Join-Path $script:PagesFolder 'agent-passports/page.json') -Raw | ConvertFrom-Json).displayName | Should -Be 'Agent passports'
    }

    It 'puts the anonymize toggle in the report filter pane' {
        $report = Get-Content -LiteralPath (Join-Path $script:ReportRoot 'CopilotStudioAgents.Report/definition/report.json') -Raw | ConvertFrom-Json -AsHashtable
        $filter = @($report.filterConfig.filters | Where-Object { $_.field.Column.Property -eq 'Mode' })
        $filter.Count | Should -Be 1
        $filter[0].field.Column.Expression.SourceRef.Entity | Should -Be 'AnonymizeMode'
        $filter[0].displayName | Should -Be 'Anonymize names'
    }

    It 'shows passport <N> in an image bound to [Passport <N>]' -ForEach (1..12 | ForEach-Object { @{ N = $_ } }) {
        $visual = Get-Content -LiteralPath (Join-Path $script:PageVisuals "passports-card-$N/visual.json") -Raw | ConvertFrom-Json -AsHashtable
        $visual.visual.visualType | Should -Be 'image'
        $visual.visual.objects.image[0].properties.sourceField.expr.Measure.Property | Should -Be "Passport $N"
        $script:Passports | Should -Match ([regex]::Escape("measure 'Passport $N' = CALCULATE([Passport Card], AgentsCurrent[Passport Order] = $N)"))
    }

    It 'counts the agents and environments past the slots, using the real slot counts' {
        $slots = @(Get-ChildItem -LiteralPath $script:PageVisuals -Directory -Filter 'passports-card-*').Count
        $slots | Should -Be 12
        $script:Passports | Should -Match ([regex]::Escape("measure 'Passport Overflow Note'"))
        $script:Passports | Should -Match ([regex]::Escape("VAR AgentSlots = $slots"))
        $script:Passports | Should -Match ([regex]::Escape('COUNTROWS(AgentsCurrent) - AgentSlots'))
        # Where They Live draws rows 1 to 6 only.
        $script:Passports | Should -Match ([regex]::Escape('FILTER(Ranked, [RowNo] <= 6)'))
        $script:Passports | Should -Match ([regex]::Escape('VAR EnvSlots = 6'))
        $script:Passports | Should -Match ([regex]::Escape('COUNTROWS(EnvironmentsCurrent) - EnvSlots'))
    }

    It 'shows that note on the page' {
        $visual = Get-Content -LiteralPath (Join-Path $script:PageVisuals 'passports-overflow-note/visual.json') -Raw | ConvertFrom-Json
        $visual.visual.visualType | Should -Be 'cardVisual'
        $visual.visual.query.queryState.Data.projections[0].field.Measure.Property | Should -Be 'Passport Overflow Note'
    }

    It 'marks every SVG measure as an image URL' {
        foreach ($name in @('Passport Card', 'Where They Live Image') + (1..12 | ForEach-Object { "Passport $_" })) {
            $script:Passports | Should -Match ("(?s)measure '" + [regex]::Escape($name) + "' =.*?dataCategory: ImageUrl")
        }
    }

    It 'orders the passports worst first and breaks ties by name and id' {
        $script:Passports | Should -Match ([regex]::Escape('AgentsCurrent[Passport Concern] * 1000000'))
        $script:Passports | Should -Match ([regex]::Escape('AgentsCurrent[AgentId] < ThisId'))
    }

    It 'uses no VAR name Desktop 2.158 rejects' {
        $rejected = 'Value', 'Status', 'Scope', 'Disabled', 'Hidden', 'FirstDate', 'LastDate', 'Lead', 'At', 'Rows', 'Level', 'Key', 'Top', 'Weight'
        $found = foreach ($file in Get-ChildItem -LiteralPath $script:TablesFolder -Filter '*.tmdl') {
            $text = Get-Content -LiteralPath $file.FullName -Raw
            foreach ($name in $rejected) {
                if ($text -cmatch "\bVAR $name\s*=") { "$($file.Name): VAR $name" }
            }
        }
        $found -join '; ' | Should -BeNullOrEmpty
    }

    It 'compares no column with a measure inside a CALCULATE filter' {
        # Desktop 2.158 rejects CALCULATE(..., Table[Column] = [Measure]); read the measure into a VAR first.
        $found = foreach ($file in Get-ChildItem -LiteralPath $script:TablesFolder -Filter '*.tmdl') {
            $text = Get-Content -LiteralPath $file.FullName -Raw
            foreach ($m in [regex]::Matches($text, '\w+\[\w+\] = \[[^\]]+\]')) { "$($file.Name): $($m.Value)" }
        }
        $found -join '; ' | Should -BeNullOrEmpty
    }
}
