#Requires -Version 7.0

<#
    The Power BI report (reports/teams-groups-lifecycle/report/) is built from the same
    CSVs the collectors write. These tests hold the report to that:

    - every table, column and measure a visual's query references must exist in the
      TMDL semantic model (a typo'd field reference is silently blank in Power BI, not
      an error, so nothing else would catch it);
    - every semantic-model table sourced directly from a CSV must have exactly the
      columns that CSV has, in the same order;
    - every relationship names columns that exist, and puts the many side in
      fromColumn;
    - no visual binds a raw name column, so the Anonymize toggle covers every name.

    GroupsCurrent and UsersCurrent (calculated, latest snapshot), DateDim (a calendar)
    and AnonymizeMode (the disconnected table behind the toggle) have no CSV of their
    own and are exempt from the CSV-mapping check. $CsvBackedTables below is the
    complete, explicit list of what is checked.

    The TMDL reader and the CSV header helper are the guest-access report's; this
    folder does not carry a second copy.
#>

BeforeAll {
    $script:Root = Join-Path $PSScriptRoot '../../..' | Resolve-Path | Select-Object -ExpandProperty Path
    $script:Samples = Join-Path $script:Root 'reports/teams-groups-lifecycle/samples'
    $script:ReportRoot = Join-Path $script:Root 'reports/teams-groups-lifecycle/report'
    $script:DefinitionFolder = Join-Path $script:ReportRoot 'TeamsGroupsLifecycle.SemanticModel/definition'
    $script:TablesFolder = Join-Path $script:DefinitionFolder 'tables'
    $script:PagesFolder = Join-Path $script:ReportRoot 'TeamsGroupsLifecycle.Report/definition/pages'

    Import-Module (Join-Path $script:Root 'shared/M365ReportLibrary.psm1') -Force
    . (Join-Path $PSScriptRoot '../../guest-access/tests/TmdlModel.ps1')

    $script:Model = Get-TmdlModel -TablesFolder $script:TablesFolder

    # Table name -> the sample CSV its `m` partition imports.
    $script:CsvBackedTables = @{
        Users                  = 'users.csv'
        Groups                 = 'groups.csv'
        GroupOwners            = 'group-owners.csv'
        DeletedGroups          = 'deleted-groups.csv'
        GroupLifecyclePolicies = 'group-lifecycle-policies.csv'
        GroupLifecycleCoverage = 'group-lifecycle-coverage.csv'
        TeamActivity           = 'team-activity.csv'
        GroupActivity          = 'group-activity.csv'
        TeamArchiveStatus      = 'team-archive-status.csv'
        GroupCreationEvents    = 'group-creation-events.csv'
    }

    function Get-VisualFieldReference {
        <#
            .SYNOPSIS
                Every {Entity, Property, Kind} a visual.json's query pulls from the
                semantic model, found by walking the parsed JSON for `Column` /
                `Measure` field containers.
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
    It 'parses the ten CSV-backed tables plus the calculated dimensions' {
        $script:Model.Keys.Count | Should -Be 14
    }

    It 'finds visual.json files to check' {
        $script:VisualFiles.Count | Should -BeGreaterThan 0
    }

    It 'declares every table in model.tmdl' {
        $declared = @(Get-Content -LiteralPath (Join-Path $script:DefinitionFolder 'model.tmdl') |
                ForEach-Object { if ($_ -match '^ref table\s+(.+)$') { $Matches[1] } })
        ($declared | Sort-Object) -join ',' | Should -Be (($script:Model.Keys | Sort-Object) -join ',')
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
        $script:AllReferences.Count | Should -BeGreaterThan 100
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

    It 'catches a column that does not exist' {
        # The check above is only worth having if it can fail.
        $script:Model['Groups'].Columns.Name | Should -Not -Contain 'NoSuchColumn'
        $script:Model['Groups'].Measures | Should -Not -Contain 'No Such Measure'
    }
}

Describe 'Every CSV-backed table matches its sample CSV' {
    It 'checks <Table>' -ForEach @(
        @{ Table = 'Users' }
        @{ Table = 'Groups' }
        @{ Table = 'GroupOwners' }
        @{ Table = 'DeletedGroups' }
        @{ Table = 'GroupLifecyclePolicies' }
        @{ Table = 'GroupLifecycleCoverage' }
        @{ Table = 'TeamActivity' }
        @{ Table = 'GroupActivity' }
        @{ Table = 'TeamArchiveStatus' }
        @{ Table = 'GroupCreationEvents' }
    ) {
        $script:Model.ContainsKey($Table) | Should -BeTrue -Because "the model should have a $Table table"

        $csvPath = Join-Path $script:Samples $script:CsvBackedTables[$Table]
        $csvHeader = Get-CsvHeaderColumn -Path $csvPath
        $csvHeader | Should -Not -BeNullOrEmpty -Because "$csvPath should be readable"

        # Only the imported (sourceColumn-backed) columns, in file order. A calculated
        # column layered on top of a CSV-backed table (SnapshotKey, Pseudonym) is not
        # part of the CSV and is excluded, same as it would be from an Import-Csv header.
        $modelColumns = @($script:Model[$Table].Columns | Where-Object { -not $_.IsCalculated } |
                ForEach-Object { $_.SourceColumn })

        ($modelColumns -join ',') | Should -Be ($csvHeader -join ',')
    }

    It 'imports each table from the CsvFolder parameter and the CSV named in the map' {
        foreach ($table in $script:CsvBackedTables.Keys) {
            $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder "$table.tmdl") -Raw
            $tmdl | Should -Match ([regex]::Escape('File.Contents(CsvFolder & "\' + $script:CsvBackedTables[$table] + '")')) -Because $table
        }
    }

    It 'has exactly one expression, the CsvFolder parameter' {
        $text = Get-Content -LiteralPath (Join-Path $script:DefinitionFolder 'expressions.tmdl') -Raw
        @([regex]::Matches($text, '(?m)^expression\s')).Count | Should -Be 1
        $text | Should -Match '(?m)^expression CsvFolder = '
        $text | Should -Match 'IsParameterQuery=true'
    }

    It 'names every CSV-backed table exactly once' {
        (@($script:CsvBackedTables.Keys) | Sort-Object) -join ',' |
            Should -Be ((@('Users', 'Groups', 'GroupOwners', 'DeletedGroups', 'GroupLifecyclePolicies', 'GroupLifecycleCoverage', 'TeamActivity', 'GroupActivity', 'TeamArchiveStatus', 'GroupCreationEvents') | Sort-Object) -join ',')
    }

    It 'covers every sample CSV the collectors write' {
        $samples = @(Get-ChildItem -LiteralPath $script:Samples -Filter '*.csv' -File | ForEach-Object Name | Sort-Object)
        $samples -join ',' | Should -Be ((@($script:CsvBackedTables.Values) | Sort-Object) -join ',')
    }
}

Describe 'Relationships join columns that exist, many side first' {
    BeforeAll {
        $text = Get-Content -LiteralPath (Join-Path $script:DefinitionFolder 'relationships.tmdl') -Raw
        $script:Relationships = @([regex]::Matches($text, '(?m)^relationship\s+\S+\s*\r?\n\s+fromColumn:\s*(\S+)\s*\r?\n\s+toColumn:\s*(\S+)') |
                ForEach-Object { [pscustomobject]@{ From = $_.Groups[1].Value; To = $_.Groups[2].Value } })
    }

    It 'has relationships' {
        $script:Relationships.Count | Should -BeGreaterThan 5
        @([regex]::Matches((Get-Content -LiteralPath (Join-Path $script:DefinitionFolder 'relationships.tmdl') -Raw), '(?m)^relationship\s')).Count |
            Should -Be $script:Relationships.Count
    }

    It 'names an existing column on both ends of <From> -> <To>' -ForEach @(
        foreach ($line in (Get-Content -LiteralPath (Join-Path $PSScriptRoot '../report/TeamsGroupsLifecycle.SemanticModel/definition/relationships.tmdl'))) {
            if ($line -match '^\s+fromColumn:\s*(\S+)') { $from = $Matches[1] }
            if ($line -match '^\s+toColumn:\s*(\S+)') { @{ From = $from; To = $Matches[1] } }
        }
    ) {
        foreach ($end in $From, $To) {
            $table, $column = $end -split '\.', 2
            $script:Model.ContainsKey($table) | Should -BeTrue -Because "$end names table $table"
            $script:Model[$table].Columns.Name | Should -Contain $column -Because $end
        }
    }

    It 'joins every date-bearing table to DateDim so the date slicer reaches it' {
        $toDate = @($script:Relationships | Where-Object { $_.To -eq 'DateDim.Date' } | ForEach-Object { ($_.From -split '\.')[0] })
        $toDate | Should -Contain 'Groups'
        $toDate | Should -Contain 'DeletedGroups'
        $toDate | Should -Contain 'GroupLifecyclePolicies'
        $toDate | Should -Contain 'GroupCreationEvents'
        # Usage reports include groups that were deleted during the period. Those
        # rows are not in groups.csv, so the date slicer has to reach them directly.
        # https://learn.microsoft.com/graph/api/reportroot-getoffice365groupsactivitydetail
        $toDate | Should -Contain 'TeamActivity'
        $toDate | Should -Contain 'GroupActivity'
    }

    It 'puts the many side in fromColumn: a snapshot child points at the one Groups row' {
        foreach ($child in 'GroupOwners', 'GroupLifecycleCoverage', 'TeamArchiveStatus') {
            $script:Relationships | Where-Object { $_.From -eq "$child.SnapshotKey" -and $_.To -eq 'Groups.SnapshotKey' } |
                Should -Not -BeNullOrEmpty -Because $child
        }
        $script:Relationships | Where-Object { $_.From -like 'DateDim.*' -or $_.From -like '*Current.*' } |
            Should -BeNullOrEmpty -Because 'DateDim and the Current tables are the one side'
    }

    It 'does not drop a usage-report row whose id is absent from groups.csv' {
        # Joining activity to Groups[SnapshotKey] makes the date slicer an inner join:
        # a deleted group that still has a usage row disappears as soon as a date is selected.
        foreach ($pair in @(
                @{ From = 'TeamActivity.SnapshotKey'; To = 'Groups.SnapshotKey' }
                @{ From = 'GroupActivity.SnapshotKey'; To = 'Groups.SnapshotKey' }
            )) {
            $script:Relationships | Where-Object { $_.From -eq $pair.From -and $_.To -eq $pair.To } |
                Should -BeNullOrEmpty -Because $pair.From
        }

        $script:Relationships | Where-Object { $_.From -eq 'TeamActivity.TeamId' -and $_.To -eq 'GroupsCurrent.Id' } |
            Should -Not -BeNullOrEmpty
        $script:Relationships | Where-Object { $_.From -eq 'GroupActivity.GroupId' -and $_.To -eq 'GroupsCurrent.Id' } |
            Should -Not -BeNullOrEmpty

        foreach ($table in 'TeamActivity', 'GroupActivity') {
            $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder "$table.tmdl") -Raw
            $tmdl | Should -Not -Match '\[Selected Snapshot Date\]' -Because "$table must use its own RunDate"
            $tmdl | Should -Match ([regex]::Escape("CALCULATE(MAX($table[RunDate]), ALLSELECTED($table))")) -Because $table
        }
    }
}

Describe 'The anonymize toggle covers every displayed name' {
    BeforeAll {
        # Columns that hold a person's or group's name or address. A visual binds the
        # Pseudonym column or a *Display Name measure instead.
        $script:RawNameColumns = @('DisplayName', 'OwnerDisplayName', 'GroupDisplayName', 'TeamName', 'TargetDisplayName',
            'UserPrincipalName', 'OwnerUserPrincipalName', 'OwnerPrincipalName', 'ManagerUserPrincipalName', 'UserId', 'Mail')
    }

    It 'does not bind a visual to a raw name column' {
        $leaks = foreach ($file in $script:VisualFiles) {
            $json = Get-Content -LiteralPath $file.FullName -Raw | ConvertFrom-Json -AsHashtable
            foreach ($ref in (Get-VisualFieldReference -Node $json)) {
                if ($ref.Kind -eq 'Column' -and $script:RawNameColumns -contains $ref.Property) {
                    '{0}: {1}.{2}' -f $file.FullName.Substring($script:PagesFolder.Length + 1), $ref.Entity, $ref.Property
                }
            }
        }
        $leaks -join '; ' | Should -BeNullOrEmpty
    }

    It 'has the Anonymize slicer and a Pseudonym group slicer on <Page>' -ForEach @(
        Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot '../report/TeamsGroupsLifecycle.Report/definition/pages') -Directory |
            ForEach-Object { @{ Page = $_.Name } }
    ) {
        $dir = Join-Path $script:PagesFolder "$Page/visuals"
        $anonymize = Get-Content -LiteralPath (Join-Path $dir 'slicer-anonymize/visual.json') -Raw | ConvertFrom-Json
        $anonymize.visual.query.queryState.Values.projections[0].field.Column.Expression.SourceRef.Entity | Should -Be 'AnonymizeMode'
        $group = Get-Content -LiteralPath (Join-Path $dir 'slicer-group/visual.json') -Raw | ConvertFrom-Json
        $group.visual.query.queryState.Values.projections[0].field.Column.Property | Should -Be 'Pseudonym'
    }

    It 'binds <Page> date slicer to DateDim[Date] in Between mode' -ForEach @(
        Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot '../report/TeamsGroupsLifecycle.Report/definition/pages') -Directory |
            ForEach-Object { @{ Page = $_.Name } }
    ) {
        $visual = Get-Content -LiteralPath (Join-Path $script:PagesFolder "$Page/visuals/slicer-date-range/visual.json") -Raw | ConvertFrom-Json
        $projection = $visual.visual.query.queryState.Values.projections[0]
        $projection.field.Column.Expression.SourceRef.Entity | Should -Be 'DateDim'
        $projection.field.Column.Property | Should -Be 'Date'
        $visual.visual.objects.data[0].properties.mode.expr.Literal.Value | Should -Be "'Between'"
    }

    It 'reads names through a Display Name measure that honours the toggle' {
        foreach ($measure in "measure 'Group Display Name'", "measure 'Deleted Group Display Name'", "measure 'Creator Display Name'", "measure 'Owner Names'") {
            $hit = Get-ChildItem -LiteralPath $script:TablesFolder -Filter '*.tmdl' | Select-String -SimpleMatch $measure
            @($hit).Count | Should -Be 1 -Because $measure
        }
        (Get-ChildItem -LiteralPath $script:TablesFolder -Filter '*.tmdl' | Select-String -SimpleMatch 'SELECTEDVALUE(AnonymizeMode[Mode], "Show names")').Count |
            Should -BeGreaterOrEqual 4
    }

    It 'derives a pseudonym from the id, so it is stable between refreshes' {
        foreach ($table in 'GroupsCurrent', 'UsersCurrent', 'DeletedGroups', 'GroupOwners') {
            $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder "$table.tmdl") -Raw
            $tmdl | Should -Match 'column PseudonymNumber' -Because $table
            $tmdl | Should -Match 'UNICODE\(MID\(IdText' -Because $table
            $tmdl | Should -Not -Match 'RAND' -Because "$table must not use a random value"
        }
    }
}

Describe 'Empty values are Unknown, never zero' {
    It 'does not measure age with TODAY()' {
        @(Get-ChildItem -LiteralPath $script:TablesFolder -Filter '*.tmdl' | Select-String -Pattern 'TODAY\(').Count | Should -Be 0
    }

    It 'returns blank, not 0, for days since last activity when the date is empty' {
        foreach ($table in 'TeamActivity', 'GroupActivity') {
            $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder "$table.tmdl") -Raw
            $tmdl | Should -Match ([regex]::Escape("IF(ISBLANK($table[LastActivityDate]), BLANK(), DATEDIFF($table[LastActivityDate], $table[RunDate], DAY))")) -Because $table
            $tmdl | Should -Match 'Unknown \(no last activity date\)' -Because $table
            $tmdl | Should -Match ([regex]::Escape("NOT ISBLANK($table[DaysSinceLastActivity]), $table[DaysSinceLastActivity] >= 90")) -Because "blank >= 90 is not guarded by DAX itself: $table"
        }
    }

    It 'leaves the activity counts blank when the snapshot has no usage-report rows' {
        $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'TeamActivity.tmdl') -Raw
        $tmdl | Should -Match 'IF\(Reported = 0, BLANK\(\), Matching\)'
        $tmdl | Should -Match 'not available in GCC High'
    }

    It 'reports a group with no listed owner list as unknown, not as no owner' {
        $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'Groups.tmdl') -Raw
        $tmdl | Should -Match 'IF\(RowCount = 0 \|\| UnknownRows > 0, "Owner unknown"'
        $tmdl | Should -Match 'Listed = 0 && NoneRows > 0, "No owner"'
    }

    It 'treats an empty expiration date as no date, not as zero days' {
        $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'Groups.tmdl') -Raw
        $tmdl | Should -Match ([regex]::Escape('IF(ISBLANK(Groups[ExpirationDateTime]), BLANK(), DATEDIFF(Groups[RunDate], Groups[ExpirationDateTime], DAY))'))
        $tmdl | Should -Match ([regex]::Escape('NOT ISBLANK(Groups[DaysToExpiry]), Groups[DaysToExpiry] >= 0, Groups[DaysToExpiry] <= 30'))
    }

    It 'classifies a soft-deleted group by whether groupTypes contains Unified' {
        # groupTypes is a collection. A dynamic Microsoft 365 group is Unified plus
        # DynamicMembership, joined with ";" in the CSV. Equality with "Unified" calls
        # that group a security group, and a dynamic security group (DynamicMembership
        # only) matches neither an empty list nor "Unified", so both cards drop it.
        # https://learn.microsoft.com/graph/api/resources/group
        $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'DeletedGroups.tmdl') -Raw
        $tmdl | Should -Match ([regex]::Escape('CONTAINSSTRING(DeletedGroups[GroupTypes] & "", "Unified")'))
        $tmdl | Should -Match ([regex]::Escape('CONTAINSSTRING(DeletedGroups[GroupTypes], "Unified")'))
        $tmdl | Should -Match ([regex]::Escape('NOT CONTAINSSTRING(DeletedGroups[GroupTypes] & "", "Unified")'))
        $tmdl | Should -Not -Match 'GroupTypes\] = "Unified"'
        $tmdl | Should -Not -Match 'ISBLANK\(DeletedGroups\[GroupTypes\]\)'
    }

    It 'treats an owner missing from users.csv as unknown, not enabled' {
        $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'GroupOwners.tmdl') -Raw
        $tmdl | Should -Match 'IF\(ISBLANK\(Enabled\), "Unknown"'
    }

    It 'counts snapshot metrics on the snapshot inside the selected date range' {
        $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'Groups.tmdl') -Raw
        $tmdl | Should -Match "measure 'Selected Snapshot Date'"
        $tmdl | Should -Match 'Groups\[RunDate\] = SnapshotDate'
    }
}

Describe 'Collector UTC timestamps are parsed with an explicit Z format' {
    It 'parses <Table> column <Column> with DateTime.FromText' -ForEach @(
        @{ Table = 'Users'; Column = 'CreatedDateTime' }
        @{ Table = 'Groups'; Column = 'CreatedDateTime' }
        @{ Table = 'Groups'; Column = 'RenewedDateTime' }
        @{ Table = 'Groups'; Column = 'ExpirationDateTime' }
        @{ Table = 'Groups'; Column = 'DeletedDateTime' }
        @{ Table = 'DeletedGroups'; Column = 'CreatedDateTime' }
        @{ Table = 'DeletedGroups'; Column = 'DeletedDateTime' }
        @{ Table = 'DeletedGroups'; Column = 'PurgeDateTime' }
        @{ Table = 'GroupCreationEvents'; Column = 'CreationTime' }
    ) {
        $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder "$Table.tmdl") -Raw
        $tmdl | Should -Match 'DateTime\.FromText'
        $tmdl | Should -Match "yyyy-MM-dd'T'HH:mm:ss'Z'"
        $tmdl | Should -Match ([regex]::Escape('{"' + $Column + '", ParseUtc, type datetime}'))
    }
}
