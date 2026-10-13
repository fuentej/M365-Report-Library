#Requires -Version 7.0

<#
    Power BI Desktop 2.158 rejected the purview-ip project with
    PFE_TM_RELATIONSHIP_END_COLUMN_INVALID ("Relationship
    'a0518db9-efa3-41cd-b0b7-6b8705b8f1b5' uses an invalid column ID 242."). That
    relationship ends on Calendar[Date], a column of the CALENDAR(...) calculated
    table that was declared with the import-table form `sourceColumn: Date`.
    Desktop writes a column bound to a calculated table's expression output as a
    bracketed `sourceColumn: [Name]`, where Name is the column the DAX expression
    returns, plus `isNameInferred` when the model column keeps that name. This is
    the same failure and the same fix as guest-access (BRO-347,
    reports/guest-access/tests/CalculatedTableColumns.Tests.ps1).

    Rules, checked on the purview-ip model (TAB indentation):

    1. Every column of a calculated table that is bound to the expression output
       (it has a sourceColumn line; `column X = <expr>` columns do not) uses a
       bracketed sourceColumn, and the bracketed name is a column the expression
       returns: CALENDAR returns [Date], GENERATESERIES returns [Value],
       ROW("A", ...) returns [A], SELECTCOLUMNS returns its aliases and UNION the
       names of its first table.
    2. Such a column carries `isNameInferred` when its declared name equals the
       output name. 'Anonymize Toggle'[Anonymize Toggle] is the one bound column
       whose declared name differs (GENERATESERIES returns [Value]); its name is set
       by the model, not inferred, so it has no `isNameInferred` (the What-If
       parameter form).
    3. Every relationship end names a column declared in its table, and no
       relationship ends on a calculated-table column that breaks rules 1 or 2.

    https://learn.microsoft.com/en-us/analysis-services/tmdl/tmdl-overview
    https://learn.microsoft.com/en-us/dotnet/api/microsoft.analysisservices.tabular.calculatedtablecolumn.sourcecolumn
    https://learn.microsoft.com/en-us/dotnet/api/microsoft.analysisservices.tabular.calculatedtablecolumn.isnameinferred
    https://learn.microsoft.com/en-us/dax/calendar-function-dax
    https://learn.microsoft.com/en-us/dax/generateseries-function-dax
#>

BeforeAll {
    $script:Definition = Join-Path $PSScriptRoot '../report/PurviewIPReport.SemanticModel/definition' | Resolve-Path | Select-Object -ExpandProperty Path
    $script:TablesFolder = Join-Path $script:Definition 'tables'

    function ConvertFrom-TmdlQuotedName {
        param([Parameter(Mandatory)][string]$Name)
        $trimmed = $Name.Trim()
        if ($trimmed.Length -ge 2 -and $trimmed[0] -eq "'" -and $trimmed[-1] -eq "'") {
            return $trimmed.Substring(1, $trimmed.Length - 2) -replace "''", "'"
        }
        return $trimmed
    }

    function Read-CalculatedTable {
        <#
            .SYNOPSIS
                Table name, bound columns and expression output names of one
                tab-indented TMDL table file. IsCalculated is false for an import table.
        #>
        param([Parameter(Mandatory)][AllowEmptyString()][string[]]$Line)

        $name = $null
        $columns = [System.Collections.Generic.List[object]]::new()
        $isCalculated = $false
        $expression = [System.Collections.Generic.List[string]]::new()
        $current = $null
        $inPartition = $false
        foreach ($text in $Line) {
            if ($text -match "^table\s+(?<n>'(?:[^']|'')+'|\S+)\s*$") {
                $name = ConvertFrom-TmdlQuotedName $Matches['n']
                continue
            }
            if ($text -match "^\t(?!\t)\S") {
                $current = $null
                $inPartition = $false
                if ($text -match "^\tcolumn\s+(?<n>'(?:[^']|'')+'|[^\s=]+)\s*$") {
                    $current = [pscustomobject]@{ Name = (ConvertFrom-TmdlQuotedName $Matches['n']); SourceColumn = $null; IsNameInferred = $false }
                    $columns.Add($current)
                }
                elseif ($text -match "^\tcolumn\s+(?<n>'(?:[^']|'')+'|[^\s=]+)\s*=") {
                    # A calculated column: declared, but not bound to the expression output.
                    $columns.Add([pscustomobject]@{ Name = (ConvertFrom-TmdlQuotedName $Matches['n']); SourceColumn = $null; IsNameInferred = $false })
                }
                elseif ($text -match '^\tpartition\s+.+=\s*calculated\s*$') {
                    $isCalculated = $true
                    $inPartition = $true
                }
                continue
            }
            if ($current) {
                if ($text -match '^\t\tsourceColumn:\s*(?<s>.+?)\s*$') { $current.SourceColumn = $Matches['s'] }
                elseif ($text -match '^\t\tisNameInferred\s*$') { $current.IsNameInferred = $true }
            }
            elseif ($inPartition -and $text -match '^\t\t\t') {
                $expression.Add($text.Trim())
            }
        }
        $bound = @($columns | Where-Object { $null -ne $_.SourceColumn })
        [pscustomobject]@{
            Name         = $name
            IsCalculated = $isCalculated
            Expression   = ($expression -join "`n")
            Columns      = @($columns)
            BoundColumns = $bound
        }
    }

    function Get-ExpressionOutputName {
        <#
            .SYNOPSIS
                The column names a calculated table's DAX expression returns, for the
                functions this model uses.
        #>
        param([Parameter(Mandatory)][AllowEmptyString()][string]$Expression)

        $e = $Expression.TrimStart()
        if ($e -match '^CALENDAR(AUTO)?\s*\(') { return , @('Date') }
        if ($e -match '^GENERATESERIES\s*\(') { return , @('Value') }
        if ($e -match '^ROW\s*\(') {
            return , @([regex]::Matches($e, '(?:ROW\s*\(|,)\s*"(?<a>[^"]+)"\s*,') | ForEach-Object { $_.Groups['a'].Value })
        }
        if ($e -match '^(UNION\s*\(\s*)?SELECTCOLUMNS\s*\(') {
            # Only the first SELECTCOLUMNS names the columns of a UNION.
            $first = ($e -split 'SELECTCOLUMNS\s*\(')[1]
            return , @([regex]::Matches($first, '(?m)^\s*"(?<a>[^"]+)"\s*,') | ForEach-Object { $_.Groups['a'].Value })
        }
        if ($e -match '^DATATABLE\s*\(') {
            return , @([regex]::Matches(($e -split '\{')[0], '"(?<a>[^"]+)"\s*,\s*\w+') | ForEach-Object { $_.Groups['a'].Value })
        }
        return , @()
    }

    function Get-CalculatedTableColumnProblem {
        <#
            .SYNOPSIS
                The output-bound columns of a calculated table that do not use the form
                Power BI Desktop writes. Empty for an import table.
        #>
        param([Parameter(Mandatory)][AllowEmptyString()][string[]]$Line)

        $table = Read-CalculatedTable -Line $Line
        if (-not $table.IsCalculated) { return , @() }
        $outputs = Get-ExpressionOutputName -Expression $table.Expression
        $problems = [System.Collections.Generic.List[string]]::new()
        if ($outputs.Count -eq 0) { $problems.Add("$($table.Name): output columns of the expression are not known to this test") }
        foreach ($column in $table.BoundColumns) {
            $source = $column.SourceColumn
            if ($source -notmatch '^\[(?<o>[^\]]+)\]$') {
                $problems.Add("$($column.Name): sourceColumn '$source' is not bracketed")
                continue
            }
            $output = $Matches['o']
            if ($outputs.Count -gt 0 -and $outputs -cnotcontains $output) {
                $problems.Add("$($column.Name): sourceColumn '[$output]' is not a column the expression returns ($($outputs -join ', '))")
            }
            if ($column.Name -ceq $output -and -not $column.IsNameInferred) {
                $problems.Add("$($column.Name): missing isNameInferred")
            }
            if ($column.Name -cne $output -and $column.IsNameInferred) {
                $problems.Add("$($column.Name): isNameInferred on a column renamed from [$output]")
            }
        }
        return , $problems.ToArray()
    }

    function Get-RelationshipEnd {
        param([Parameter(Mandatory)][string]$Path)
        foreach ($line in Get-Content -LiteralPath $Path) {
            if ($line -match "^\s+(?:fromColumn|toColumn):\s*(?<table>'(?:[^']|'')+'|[^.\s]+)\.(?<column>'(?:[^']|'')+'|\S+)\s*$") {
                [pscustomobject]@{ Table = (ConvertFrom-TmdlQuotedName $Matches['table']); Column = (ConvertFrom-TmdlQuotedName $Matches['column']) }
            }
        }
    }

    function Get-RelationshipEndProblem {
        <#
            .SYNOPSIS
                Relationship ends that do not resolve to a declared column, or that are
                columns of a calculated table whose bound columns break the rule.
        #>
        param([Parameter(Mandatory)][string]$Definition)

        $declared = @{}
        $bad = @{}
        foreach ($file in Get-ChildItem -LiteralPath (Join-Path $Definition 'tables') -Filter '*.tmdl' -File) {
            $lines = Get-Content -LiteralPath $file.FullName
            $table = Read-CalculatedTable -Line $lines
            foreach ($c in $table.Columns) { $declared["$($table.Name).$($c.Name)"] = $true }
            foreach ($problem in (Get-CalculatedTableColumnProblem -Line $lines)) {
                $bad["$($table.Name).$(($problem -split ': ')[0])"] = $problem
            }
        }
        foreach ($end in Get-RelationshipEnd -Path (Join-Path $Definition 'relationships.tmdl')) {
            $key = "$($end.Table).$($end.Column)"
            if (-not $declared.ContainsKey($key)) { "$key (relationship end): not a declared column" }
            elseif ($bad.ContainsKey($key)) { "$key (relationship end): $($bad[$key])" }
        }
    }

    $script:Tables = @(
        Get-ChildItem -LiteralPath $script:TablesFolder -Filter '*.tmdl' -File |
            ForEach-Object { $t = Read-CalculatedTable -Line (Get-Content -LiteralPath $_.FullName); $t | Add-Member -NotePropertyName File -NotePropertyValue $_.Name -PassThru }
    )
    $script:CalculatedFiles = @($script:Tables | Where-Object IsCalculated | ForEach-Object File)
}

Describe 'Calculated table columns use the form Power BI Desktop writes' {
    It 'finds exactly the three calculated tables' {
        # User Display is a Power Query table; see the circular-dependency Describe below.
        ($script:CalculatedFiles | Sort-Object) -join ',' | Should -Be '_Measures.tmdl,AnonymizeToggle.tmdl,Calendar.tmdl'
    }

    It 'has no FILTER(...) calculated table that should be an import table' {
        # guest-access rebuilt its FILTER(Source) *Current tables as import tables
        # (BRO-347). purview-ip has none; keep it that way.
        @($script:Tables | Where-Object { $_.IsCalculated -and $_.Expression -match '^\s*FILTER\s*\(' }).File | Should -BeNullOrEmpty
    }

    It 'knows the output columns of <_>' -ForEach @('AnonymizeToggle.tmdl', 'Calendar.tmdl', '_Measures.tmdl') {
        $file = $_
        $table = $script:Tables | Where-Object File -eq $file
        $expected = @{
            'AnonymizeToggle.tmdl' = 'Value'
            'Calendar.tmdl'        = 'Date'
            '_Measures.tmdl'        = 'Measures Key'
        }[$file]
        (Get-ExpressionOutputName -Expression $table.Expression) -join ',' | Should -Be $expected
    }

    It '<_> binds every returned column with a bracketed sourceColumn and isNameInferred' -ForEach @('AnonymizeToggle.tmdl', 'Calendar.tmdl', '_Measures.tmdl') {
        $problems = Get-CalculatedTableColumnProblem -Line (Get-Content -LiteralPath (Join-Path $script:TablesFolder $_))
        $problems | Should -BeNullOrEmpty
    }

    It 'binds Calendar[Date] to [Date] and Anonymize Toggle to [Value]' {
        $calendar = $script:Tables | Where-Object Name -eq 'Calendar'
        ($calendar.BoundColumns | Where-Object Name -eq 'Date').SourceColumn | Should -Be '[Date]'
        $toggle = $script:Tables | Where-Object Name -eq 'Anonymize Toggle'
        ($toggle.BoundColumns | Where-Object Name -eq 'Anonymize Toggle').SourceColumn | Should -Be '[Value]'
    }

    It 'resolves every relationship end, with none on a broken calculated-table column' {
        @(Get-RelationshipEnd -Path (Join-Path $script:Definition 'relationships.tmdl')).Count | Should -BeGreaterThan 10
        Get-RelationshipEndProblem -Definition $script:Definition | Should -BeNullOrEmpty
    }

    Context 'on a deliberately broken copy' {
        BeforeAll {
            function New-BrokenCopy {
                param([string]$Table, [string]$Pattern, [string]$Replacement)
                $copy = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
                Copy-Item -LiteralPath $script:Definition -Destination $copy -Recurse
                $path = Join-Path $copy "tables/$Table.tmdl"
                $before = Get-Content -LiteralPath $path -Raw
                $text = $before -replace $Pattern, $Replacement
                if ($text -eq $before) { throw "The pattern '$Pattern' did not change $Table.tmdl" }
                Set-Content -LiteralPath $path -Value $text -NoNewline
                $copy
            }
        }

        It 'fails when a calculated-table column uses a bare sourceColumn' {
            $copy = New-BrokenCopy -Table 'Calendar' -Pattern 'sourceColumn: \[Date\]' -Replacement 'sourceColumn: Date'
            $problems = Get-CalculatedTableColumnProblem -Line (Get-Content -LiteralPath (Join-Path $copy 'tables/Calendar.tmdl'))
            $problems | Should -Contain "Date: sourceColumn 'Date' is not bracketed"
        }

        It 'fails when a calculated-table column lacks isNameInferred' {
            $copy = New-BrokenCopy -Table 'Calendar' -Pattern '(?m)^\t\tisNameInferred\r?\n' -Replacement ''
            $problems = Get-CalculatedTableColumnProblem -Line (Get-Content -LiteralPath (Join-Path $copy 'tables/Calendar.tmdl'))
            $problems | Should -Contain 'Date: missing isNameInferred'
        }

        It 'fails when the bracketed sourceColumn is the declared name, not the output name' {
            $copy = New-BrokenCopy -Table 'AnonymizeToggle' -Pattern 'sourceColumn: \[Value\]' -Replacement 'sourceColumn: [Anonymize Toggle]'
            $problems = Get-CalculatedTableColumnProblem -Line (Get-Content -LiteralPath (Join-Path $copy 'tables/AnonymizeToggle.tmdl'))
            $problems | Should -Contain "Anonymize Toggle: sourceColumn '[Anonymize Toggle]' is not a column the expression returns (Value)"
        }

        It 'fails when a relationship end is a column of a broken calculated table' {
            $copy = New-BrokenCopy -Table 'Calendar' -Pattern 'sourceColumn: \[Date\]' -Replacement 'sourceColumn: Date'
            $ends = @(Get-RelationshipEndProblem -Definition $copy)
            $ends | Should -Contain "Calendar.Date (relationship end): Date: sourceColumn 'Date' is not bracketed"
        }

        It 'fails when a relationship end is a quoted column of a broken calculated table' {
            $copy = New-BrokenCopy -Table 'AnonymizeToggle' -Pattern 'sourceColumn: \[Value\]' -Replacement 'sourceColumn: Value'
            $ends = @(Get-RelationshipEndProblem -Definition $copy)
            $ends | Should -Contain "Anonymize Toggle.Anonymize Toggle (relationship end): Anonymize Toggle: sourceColumn 'Value' is not bracketed"
        }

        It 'fails when a relationship end names a column the table does not declare' {
            $copy = New-BrokenCopy -Table 'Calendar' -Pattern '(?m)^\tcolumn Date\s*$' -Replacement "`tcolumn Day"
            $ends = @(Get-RelationshipEndProblem -Definition $copy)
            $ends | Should -Contain 'Calendar.Date (relationship end): not a declared column'
        }
    }
}

Describe 'Users and User Display have no DAX that can form a circular dependency' {
    # 'User Display' has a relationship to Users, which gives Users a blank row.
    # A DAX calculated column on Users that scans the table (PATH, RANKX,
    # LOOKUPVALUE), read by a DAX 'User Display', is a circular dependency, and
    # Power BI Desktop refuses the refresh. Both are built in Power Query.
    It 'Users has no calculated columns' {
        $users = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'Users.tmdl')
        @($users | Select-String -Pattern '^\t\ttype: calculated\s*$').Count | Should -Be 0
        @($users | Select-String -Pattern '^\tcolumn [^=]*=').Count | Should -Be 0
    }

    It 'User Display is a Power Query table built from Users' {
        $userDisplay = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'UserDisplay.tmdl') -Raw
        $userDisplay | Should -Match "(?m)^\tpartition 'User Display' = m\s*$"
        $userDisplay | Should -Match 'Table\.SelectColumns\(Users,'
    }

    It 'User Display reads only columns the Users query produces' {
        $users = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'Users.tmdl') -Raw
        $userDisplay = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'UserDisplay.tmdl') -Raw
        $produced = @([regex]::Matches($users, '(?m)^\t\tsourceColumn: (.+?)\s*$') | ForEach-Object { $_.Groups[1].Value })
        $read = @([regex]::Matches($userDisplay, 'Table\.SelectColumns\(Users, \{([^}]*)\}') | ForEach-Object {
                [regex]::Matches($_.Groups[1].Value, '"([^"]+)"') | ForEach-Object { $_.Groups[1].Value } })
        $read.Count | Should -Be 14
        @($read | Where-Object { $_ -notin $produced }) | Should -BeNullOrEmpty
    }
}
