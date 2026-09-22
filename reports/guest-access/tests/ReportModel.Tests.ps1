#Requires -Version 7.0

<#
    The Power BI report (reports/guest-access/report/) is built from the same CSVs
    the collectors write. These tests hold the report to that:

    - every table, column and measure a visual's query references must exist in
      the TMDL semantic model (a typo'd field reference is silently blank in
      Power BI, not an error, so nothing else would catch it);
    - every semantic-model table sourced directly from a CSV must have exactly
      the columns that CSV has, in the same order -- the same guarantee
      SampleData.Tests.ps1 holds the collectors to, extended to the report.

    Calculated tables (GuestsCurrent, UsersCurrent, GuestMembershipsCurrent) and
    helper tables that are not sourced from any CSV (DateDim, a calendar; and
    AnonymizeMode, the disconnected table behind the anonymize toggle) are exempt
    from the CSV-mapping check -- there is no CSV for a calculated or disconnected
    table to match. $CsvBackedTables below is the complete, explicit list of what
    *is* checked.
#>

BeforeAll {
    $script:Root = Join-Path $PSScriptRoot '../../..' | Resolve-Path | Select-Object -ExpandProperty Path
    $script:Samples = Join-Path $script:Root 'reports/guest-access/samples'
    $script:ReportRoot = Join-Path $script:Root 'reports/guest-access/report'
    $script:TablesFolder = Join-Path $script:ReportRoot 'GuestAccess.SemanticModel/definition/tables'
    $script:PagesFolder = Join-Path $script:ReportRoot 'GuestAccess.Report/definition/pages'

    Import-Module (Join-Path $script:Root 'shared/M365ReportLibrary.psm1') -Force
    . (Join-Path $PSScriptRoot 'TmdlModel.ps1')

    $script:Model = Get-TmdlModel -TablesFolder $script:TablesFolder

    # Table name -> the sample CSV its `m` partition imports.
    $script:CsvBackedTables = @{
        Users            = 'users.csv'
        Guests           = 'guests.csv'
        GuestInvitations = 'guest-invitations.csv'
        GuestSignIns     = 'guest-signins.csv'
        SharingEvents    = 'sharing-events.csv'
        GuestMemberships = 'guest-memberships.csv'
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
    It 'parses at least the six CSV-backed tables plus the calculated dimensions' {
        $script:Model.Keys.Count | Should -BeGreaterOrEqual 9
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
        # Guards the test below against silently passing because nothing was found.
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
}

Describe 'Every CSV-backed table matches its sample CSV' {
    It 'checks <Table>' -ForEach @(
        @{ Table = 'Users' }
        @{ Table = 'Guests' }
        @{ Table = 'GuestInvitations' }
        @{ Table = 'GuestSignIns' }
        @{ Table = 'SharingEvents' }
        @{ Table = 'GuestMemberships' }
    ) {
        $script:Model.ContainsKey($Table) | Should -BeTrue -Because "the model should have a $Table table"

        $csvPath = Join-Path $script:Samples $script:CsvBackedTables[$Table]
        $csvHeader = Get-CsvHeaderColumn -Path $csvPath
        $csvHeader | Should -Not -BeNullOrEmpty -Because "$csvPath should be readable"

        # Only the imported (sourceColumn-backed) columns, in file order -- a
        # calculated column layered on top of a CSV-backed table (e.g.
        # GuestInvitations' EventCategory) is not part of the CSV and is
        # deliberately excluded here, same as it would be from an Import-Csv header.
        $modelColumns = @($script:Model[$Table].Columns | Where-Object { -not $_.IsCalculated } |
                ForEach-Object { $_.SourceColumn })

        ($modelColumns -join ',') | Should -Be ($csvHeader -join ',')
    }

    It 'names every CSV-backed table exactly once' {
        # Every key in $CsvBackedTables is the single source of truth for this
        # check; this just confirms the six names above match what it defines.
        (@($script:CsvBackedTables.Keys) | Sort-Object) -join ',' |
            Should -Be ((@('Users', 'Guests', 'GuestInvitations', 'GuestSignIns', 'SharingEvents', 'GuestMemberships') | Sort-Object) -join ',')
    }
}
