#Requires -Version 7.0

<#
    Two model-load errors Power BI Desktop raises that nothing else in the suite
    could see (BRO-345):

      Could not add Measure with the name 'Membership Count' because a Measure with
      the same name already exists in the 'Model' Model.

      The Column with the name of 'PseudonymNumber' already exists in the
      'GuestsCurrent' Table.

    Measure names are unique across the whole model, not per table. A calculated
    table inherits one column for each column its partition expression returns
    (a CalculatedTableColumn, sourceColumn = that name). FILTER(<Base>, ...) returns
    every column of <Base>. SELECTCOLUMNS returns only its aliases. The table may
    list an inherited column only as a bare declaration with `sourceColumn:`
    (relationships resolve against columns declared in the table's own file). A
    `column X = <DAX>` with the same name adds a second column called X.

    https://learn.microsoft.com/en-us/analysis-services/tabular-models/create-a-calculated-table-ssas-tabular

    Microsoft.AnalysisServices.NetCore.retail.amd64 19.84.1 (TmdlSerializer)
    deserializes the broken guest-access model without error, because the duplicate
    check happens when the engine loads the model, not when the folder is read. So
    these rules are written against the TMDL text instead. TMDL indent is either
    four spaces or a tab; both are one object level.

    https://learn.microsoft.com/analysis-services/tmdl/tmdl-overview
#>

BeforeDiscovery {
    $script:Models = Get-ChildItem -Path (Join-Path $PSScriptRoot '../../reports/*/report/*.SemanticModel/definition') -Directory |
        ForEach-Object { @{ Name = $_.FullName.Substring((Resolve-Path (Join-Path $PSScriptRoot '../../reports')).Path.Length + 1); Path = $_.FullName } }
}

BeforeAll {
    function ConvertFrom-TmdlObjectName {
        param([Parameter(Mandatory)][string]$Name)

        $trimmed = $Name.Trim()
        if ($trimmed.Length -ge 2 -and $trimmed[0] -eq "'" -and $trimmed[-1] -eq "'") {
            return $trimmed.Substring(1, $trimmed.Length - 2) -replace "''", "'"
        }
        return $trimmed
    }

    function Read-TmdlTableText {
        <#
            .SYNOPSIS
                Name, measures, columns, FILTER base, and SELECTCOLUMNS aliases.
                A column counts as calculated when its declaration carries `=`.
        #>
        param([Parameter(Mandatory)][AllowEmptyString()][string[]]$Line)

        # One leading tab and four leading spaces are the same object level.
        # Purview IP uses tabs; the other models use spaces.
        $Line = foreach ($raw in @($Line)) {
            $lead = [regex]::Match([string]$raw, '^[\t ]*').Value
            $width = ($lead -replace "`t", '    ').Length
            (' ' * $width) + ([string]$raw).Substring($lead.Length)
        }

        $name = $null
        $measures = [System.Collections.Generic.List[string]]::new()
        $columns = [System.Collections.Generic.List[object]]::new()
        $filterBase = $null
        $projected = $null
        for ($i = 0; $i -lt $Line.Count; $i++) {
            $text = $Line[$i]
            if ($text -match "^table\s+(?<n>'(?:[^']|'')+'|\S+)\s*$") {
                $name = ConvertFrom-TmdlObjectName $Matches['n']
            }
            elseif ($text -match "^    measure\s+(?<n>'(?:[^']|'')+'|[^\s=]+)\s*=") {
                $measures.Add((ConvertFrom-TmdlObjectName $Matches['n']))
            }
            elseif ($text -match "^    column\s+(?<n>'(?:[^']|'')+'|[^\s=]+)\s*(?<eq>=?)") {
                $columns.Add([pscustomobject]@{ Name = (ConvertFrom-TmdlObjectName $Matches['n']); Calculated = ($Matches['eq'] -eq '='); LineNumber = $i + 1 })
            }
            elseif ($text -match '^    partition\s+.*=\s*calculated\s*$') {
                $body = [System.Collections.Generic.List[string]]::new()
                for ($j = $i + 1; $j -lt $Line.Count -and $Line[$j] -match '^\s{8}|^\s*$'; $j++) {
                    $body.Add($Line[$j])
                }
                $hasSelect = $false
                foreach ($bodyLine in $body) {
                    if ($bodyLine -match 'SELECTCOLUMNS\s*\(') { $hasSelect = $true }
                }
                if ($hasSelect) {
                    # The expression returns the alias list, not every column of the inner FILTER.
                    $projected = [System.Collections.Generic.List[string]]::new()
                    foreach ($bodyLine in $body) {
                        if ($bodyLine -match '^\s+"(?<alias>[^"]+)"\s*,') {
                            $projected.Add($Matches['alias'])
                        }
                    }
                }
                else {
                    for ($j = 0; $j -lt $body.Count; $j++) {
                        if ($body[$j] -match "^\s+FILTER\(\s*(?<base>'(?:[^']|'')+'|\w+)\s*,") {
                            $filterBase = ConvertFrom-TmdlObjectName $Matches['base']
                            break
                        }
                        if ($body[$j] -match '^\s+FILTER\(\s*$') {
                            for ($k = $j + 1; $k -lt $body.Count -and $body[$k] -match '^\s*$'; $k++) { }
                            if ($k -lt $body.Count -and $body[$k] -match "^\s+(?<base>'(?:[^']|'')+'|\w+)\s*,") {
                                $filterBase = ConvertFrom-TmdlObjectName $Matches['base']
                            }
                            break
                        }
                    }
                }
            }
        }
        [pscustomobject]@{ Name = $name; Measures = $measures; Columns = $columns; FilterBase = $filterBase; Projected = $projected }
    }

    function Read-TmdlModelFolder {
        param([Parameter(Mandatory)][string]$Path)

        Get-ChildItem -LiteralPath (Join-Path $Path 'tables') -File -Filter '*.tmdl' |
            ForEach-Object { Read-TmdlTableText -Line (Get-Content -LiteralPath $_.FullName) }
    }

    function Get-DuplicateMeasure {
        param([Parameter(Mandatory)][object[]]$Table)

        $Table | ForEach-Object { $t = $_; $t.Measures | ForEach-Object { [pscustomobject]@{ Measure = $_; Table = $t.Name } } } |
            Group-Object Measure | Where-Object Count -gt 1 |
            ForEach-Object { "'$($_.Name)' in $(($_.Group.Table | Sort-Object) -join ', ')" }
    }

    function Get-RedeclaredInheritedColumn {
        param([Parameter(Mandatory)][object[]]$Table)

        foreach ($t in $Table) {
            $inherited = $null
            $from = $null
            if ($null -ne $t.Projected) {
                $inherited = @($t.Projected)
                $from = 'its calculated partition'
            }
            elseif ($t.FilterBase) {
                $base = @($Table | Where-Object Name -eq $t.FilterBase)
                if ($base.Count -ne 1) { continue }
                $inherited = @($base[0].Columns.Name)
                $from = $t.FilterBase
            }
            else { continue }

            foreach ($c in @($t.Columns) | Where-Object { $_.Calculated -and $inherited -contains $_.Name }) {
                "$($t.Name)[$($c.Name)] (line $($c.LineNumber)) repeats a column of $from"
            }
        }
    }
}

Describe 'TMDL model integrity' {
    It 'reads tables and columns from <Name>, so a tab-indented model is not a vacuous pass' -ForEach $script:Models {
        $tables = @(Read-TmdlModelFolder -Path $Path)
        $tables.Count | Should -BeGreaterThan 3
        $columnCount = 0
        foreach ($t in $tables) { $columnCount += @($t.Columns).Count }
        $columnCount | Should -BeGreaterThan 0
    }

    It '<Name> has no measure name used twice' -ForEach $script:Models {
        $dup = @(Get-DuplicateMeasure -Table @(Read-TmdlModelFolder -Path $Path))
        $dup -join '; ' | Should -BeNullOrEmpty
    }

    It '<Name> has no calculated table repeating a returned column as a calculated column' -ForEach $script:Models {
        $faults = @(Get-RedeclaredInheritedColumn -Table @(Read-TmdlModelFolder -Path $Path))
        $faults -join '; ' | Should -BeNullOrEmpty
    }

    It 'finds calculated tables in identity-posture, so the inheritance rule is exercised' {
        # BRO-347 rebuilt guest-access's *Current tables as import tables, so it has
        # none left to check. The rule still runs against the other reports' models.
        $path = (Resolve-Path (Join-Path $PSScriptRoot '../../reports/identity-posture/report/IdentityPosture.SemanticModel/definition')).Path
        $calculated = @(Read-TmdlModelFolder -Path $path | Where-Object { $_.FilterBase -or $null -ne $_.Projected })
        $calculated.Count | Should -BeGreaterThan 0
    }

    Context 'detection' {
        It 'flags a duplicate measure when the table is indented with a tab' {
            $a = Read-TmdlTableText -Line @('table A', "`tmeasure 'Membership Count' = COUNTROWS(A)")
            $b = Read-TmdlTableText -Line @('table B', "`tmeasure 'Membership Count' = COUNTROWS(B)")
            @($a.Measures).Count | Should -Be 1
            @(Get-DuplicateMeasure -Table @($a, $b)).Count | Should -Be 1
        }

        It 'flags a measure name defined on two tables' {
            $a = Read-TmdlTableText -Line @('table A', "    measure 'Membership Count' = COUNTROWS(A)")
            $b = Read-TmdlTableText -Line @('table B', "    measure 'Membership Count' = COUNTROWS(B)")
            @(Get-DuplicateMeasure -Table @($a, $b)).Count | Should -Be 1
        }

        It 'flags a calculated column that repeats a column of the FILTER base' {
            $base = Read-TmdlTableText -Line @('table Guests', '    column Id', '        dataType: string', '    column PseudonymNumber = ', '            RETURN 1')
            $cur = Read-TmdlTableText -Line @(
                'table GuestsCurrent'
                '    column Id'
                '        sourceColumn: Id'
                '    column PseudonymNumber = '
                '            RETURN 1'
                '    partition GuestsCurrent = calculated'
                '        mode: import'
                '        source ='
                '            FILTER(Guests, Guests[RunDate] = MAX(Guests[RunDate]))'
            )
            @(Get-RedeclaredInheritedColumn -Table @($base, $cur)).Count | Should -Be 1
        }

        It 'accepts the inherited column as a bare sourceColumn declaration' {
            $base = Read-TmdlTableText -Line @('table Guests', '    column PseudonymNumber = ', '            RETURN 1')
            $cur = Read-TmdlTableText -Line @(
                'table GuestsCurrent'
                '    column PseudonymNumber'
                '        sourceColumn: PseudonymNumber'
                '    partition GuestsCurrent = calculated'
                '        source ='
                '            FILTER(Guests, Guests[RunDate] = MAX(Guests[RunDate]))'
            )
            @(Get-RedeclaredInheritedColumn -Table @($base, $cur)).Count | Should -Be 0
        }

        It 'reads a FILTER base when the table name is on the next line' {
            $base = Read-TmdlTableText -Line @('table Users', '    column City')
            $cur = Read-TmdlTableText -Line @(
                'table UsersCurrent'
                '    column City = "x"'
                '    partition UsersCurrent = calculated'
                '        source ='
                '            FILTER('
                '                Users,'
                '                Users[RunDate] = MAX(Users[RunDate])'
                '            )'
            )
            $cur.FilterBase | Should -Be 'Users'
            @(Get-RedeclaredInheritedColumn -Table @($base, $cur)).Count | Should -Be 1
        }

        It 'flags a calculated column that repeats a SELECTCOLUMNS alias' {
            $base = Read-TmdlTableText -Line @('table Users', '    column Id')
            $cur = Read-TmdlTableText -Line @(
                'table UsersCurrent'
                '    column Pseudonym = "x"'
                '    partition UsersCurrent = calculated'
                '        source ='
                '            SELECTCOLUMNS('
                '                FILTER(Users, Users[RunDate] = MAX(Users[RunDate])),'
                '                "Id", Users[Id],'
                '                "Pseudonym", Users[DisplayName]'
                '            )'
            )
            @(Get-RedeclaredInheritedColumn -Table @($base, $cur)).Count | Should -Be 1
        }

        It 'does not treat a base column omitted from SELECTCOLUMNS as inherited' {
            $base = Read-TmdlTableText -Line @('table Users', '    column Id', '    column City')
            $cur = Read-TmdlTableText -Line @(
                'table UsersCurrent'
                '    column City = "x"'
                '    partition UsersCurrent = calculated'
                '        source ='
                '            SELECTCOLUMNS('
                '                FILTER(Users, Users[RunDate] = MAX(Users[RunDate])),'
                '                "Id", Users[Id]'
                '            )'
            )
            @(Get-RedeclaredInheritedColumn -Table @($base, $cur)).Count | Should -Be 0
        }

        It 'accepts a new calculated column that the base does not have' {
            $base = Read-TmdlTableText -Line @('table Users', '    column Id')
            $cur = Read-TmdlTableText -Line @(
                'table UsersCurrent'
                '    column Pseudonym = "x"'
                '    partition UsersCurrent = calculated'
                '        source ='
                '            FILTER(Users, Users[RunDate] = MAX(Users[RunDate]))'
            )
            @(Get-RedeclaredInheritedColumn -Table @($base, $cur)).Count | Should -Be 0
        }
    }
}
