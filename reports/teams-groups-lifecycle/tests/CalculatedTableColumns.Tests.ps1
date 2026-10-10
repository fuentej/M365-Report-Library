#Requires -Version 7.0

<#
    Power BI Desktop 2.158 rejected the teams-groups-lifecycle project with
    PFE_TM_RELATIONSHIP_END_COLUMN_INVALID ("Relationship ... uses an invalid column
    ID 74") (BRO-434), the same failure the guest-access report had (BRO-347). A
    relationship end that is a column of a calculated table has to be bound the way
    Desktop binds it: `isNameInferred` plus a bracketed `sourceColumn: [Name]`. The
    bare `sourceColumn: Name` form is what an import table uses.

    Two rules, both checked on the teams-groups-lifecycle model:

    1. Every column of a calculated table that is bound to the expression's output
       carries `isNameInferred` and a bracketed `sourceColumn`. Calculated columns
       (`column X = <expr>`) are not part of the output and are skipped.
    2. Every relationship end that is a column of a calculated table meets rule 1.
       GroupsCurrent and UsersCurrent are import tables loaded from the CSV, so the
       only relationship ends on calculated tables are DateDim[Date] and
       GroupKind[Label].

    https://learn.microsoft.com/en-us/dotnet/api/microsoft.analysisservices.tabular.calculatedtablecolumn.sourcecolumn
    https://learn.microsoft.com/en-us/dotnet/api/microsoft.analysisservices.tabular.calculatedtablecolumn.isnameinferred
    https://learn.microsoft.com/en-us/analysis-services/tmdl/tmdl-overview
#>

BeforeAll {
    $script:Definition = Join-Path $PSScriptRoot '../report/TeamsGroupsLifecycle.SemanticModel/definition' | Resolve-Path | Select-Object -ExpandProperty Path
    $script:TablesFolder = Join-Path $script:Definition 'tables'

    function Test-TmdlCalculatedTable {
        param([Parameter(Mandatory)][AllowEmptyString()][string[]]$Line)
        $text = $Line -join "`n"
        return $text -match '(?m)^    partition\s+.+=\s*calculated\s*$'
    }

    function Get-CalculatedTableColumnProblem {
        <#
            .SYNOPSIS
                Returns the output-bound columns of a calculated table that lack
                isNameInferred or a bracketed sourceColumn. Empty for an import table.
        #>
        param([Parameter(Mandatory)][AllowEmptyString()][string[]]$Line)

        if (-not (Test-TmdlCalculatedTable -Line $Line)) { return , @() }
        $text = $Line -join "`n"
        $problems = [System.Collections.Generic.List[string]]::new()
        # A bound column has a sourceColumn line; `column X = <expr>` has none.
        $blocks = [regex]::Matches($text, '(?m)^    column\s+(\S+)\s*$\n(?:(?!^    \S).*\n?)*')
        foreach ($block in $blocks) {
            $name = $block.Groups[1].Value
            $source = [regex]::Match($block.Value, '(?m)^        sourceColumn:\s*(\S+)\s*$')
            if (-not $source.Success) { $problems.Add("${name}: no sourceColumn"); continue }
            if ($source.Groups[1].Value -notmatch '^\[[^\]]+\]$') {
                $problems.Add("${name}: sourceColumn '$($source.Groups[1].Value)' is not bracketed")
            }
            if ($block.Value -notmatch '(?m)^        isNameInferred\s*$') { $problems.Add("${name}: missing isNameInferred") }
        }
        return , $problems.ToArray()
    }

    function Get-RelationshipEnd {
        param([Parameter(Mandatory)][string]$Path)
        foreach ($line in Get-Content -LiteralPath $Path) {
            if ($line -match '^\s+(?:fromColumn|toColumn):\s*(?<table>.+)\.(?<column>[^.]+?)\s*$') {
                [pscustomobject]@{ Table = $Matches['table'].Trim("'"); Column = $Matches['column'].Trim("'") }
            }
        }
    }

    function Get-RelationshipEndOnBadCalculatedTable {
        <#
            .SYNOPSIS
                Relationship ends that are columns of a calculated table whose
                columns do not meet the Desktop form.
        #>
        param([Parameter(Mandatory)][string]$Definition)

        $tables = Join-Path $Definition 'tables'
        $bad = @{}
        foreach ($file in Get-ChildItem -LiteralPath $tables -Filter '*.tmdl' -File) {
            $lines = Get-Content -LiteralPath $file.FullName
            if (-not (Test-TmdlCalculatedTable -Line $lines)) { continue }
            $problems = Get-CalculatedTableColumnProblem -Line $lines
            foreach ($problem in $problems) { $bad[$file.BaseName + '.' + ($problem -split ':')[0]] = $problem }
        }
        foreach ($end in Get-RelationshipEnd -Path (Join-Path $Definition 'relationships.tmdl')) {
            $key = "$($end.Table).$($end.Column)"
            if ($bad.ContainsKey($key)) { "$key (relationship end): $($bad[$key])" }
        }
    }

    $script:CalculatedFiles = @(
        Get-ChildItem -LiteralPath $script:TablesFolder -Filter '*.tmdl' -File |
            Where-Object { Test-TmdlCalculatedTable -Line (Get-Content -LiteralPath $_.FullName) }
    )
}

Describe 'Calculated table columns use the form Power BI Desktop writes' {
    It 'finds exactly the calculated dimension tables' {
        ($script:CalculatedFiles.Name | Sort-Object) -join ',' | Should -Be 'AnonymizeMode.tmdl,DateDim.tmdl,GroupKind.tmdl'
    }

    It 'has rebuilt the *Current tables as import tables' {
        $script:CalculatedFiles.Name | Should -Not -Contain 'GroupsCurrent.tmdl'
        $script:CalculatedFiles.Name | Should -Not -Contain 'UsersCurrent.tmdl'
        foreach ($table in 'GroupsCurrent', 'UsersCurrent') {
            Get-Content -LiteralPath (Join-Path $script:TablesFolder "$table.tmdl") -Raw |
                Should -Match "(?m)^    partition $table = m\s*\r?\n        mode: import" -Because $table
        }
    }

    It '<_> binds every returned column with isNameInferred and a bracketed sourceColumn' -ForEach @('DateDim.tmdl', 'GroupKind.tmdl', 'AnonymizeMode.tmdl') {
        $problems = Get-CalculatedTableColumnProblem -Line (Get-Content -LiteralPath (Join-Path $script:TablesFolder $_))
        $problems | Should -BeNullOrEmpty
    }

    It 'names the column each calculated table expression returns' -ForEach @(
        @{ File = 'DateDim.tmdl'; Column = 'Date' }
        @{ File = 'GroupKind.tmdl'; Column = 'Label' }
        @{ File = 'AnonymizeMode.tmdl'; Column = 'Mode' }
    ) {
        # CALENDAR returns [Date]; DATATABLE returns the column names it is given.
        Get-Content -LiteralPath (Join-Path $script:TablesFolder $File) -Raw |
            Should -Match ("(?m)^    column $Column\s*\r?\n(?:        .*\r?\n)*?        sourceColumn: \[$Column\]")
    }

    It 'has relationship ends on DateDim and GroupKind, so the rule is exercised' {
        $ends = @(Get-RelationshipEnd -Path (Join-Path $script:Definition 'relationships.tmdl') | ForEach-Object { "$($_.Table).$($_.Column)" })
        @($ends | Where-Object { $_ -eq 'DateDim.Date' }).Count | Should -Be 6
        @($ends | Where-Object { $_ -eq 'GroupKind.Label' }).Count | Should -Be 2
    }

    It 'has no relationship end on a calculated-table column that breaks the rule' {
        Get-RelationshipEndOnBadCalculatedTable -Definition $script:Definition | Should -BeNullOrEmpty
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
                if ($text -eq $before) { throw "Pattern '$Pattern' did not change $Table.tmdl" }
                Set-Content -LiteralPath $path -Value $text -NoNewline
                $copy
            }
        }

        It 'fails when a calculated-table column uses a bare sourceColumn' {
            $copy = New-BrokenCopy -Table 'DateDim' -Pattern 'sourceColumn: \[Date\]' -Replacement 'sourceColumn: Date'
            $problems = Get-CalculatedTableColumnProblem -Line (Get-Content -LiteralPath (Join-Path $copy 'tables/DateDim.tmdl'))
            $problems | Should -Contain "Date: sourceColumn 'Date' is not bracketed"
        }

        It 'fails when a calculated-table column lacks isNameInferred' {
            $copy = New-BrokenCopy -Table 'GroupKind' -Pattern '(?m)^        isNameInferred\r?\n' -Replacement ''
            $problems = Get-CalculatedTableColumnProblem -Line (Get-Content -LiteralPath (Join-Path $copy 'tables/GroupKind.tmdl'))
            $problems | Should -Contain 'Label: missing isNameInferred'
        }

        It 'fails when a relationship end is a column of a broken calculated table' {
            $copy = New-BrokenCopy -Table 'GroupKind' -Pattern 'sourceColumn: \[Label\]' -Replacement 'sourceColumn: Label'
            $ends = @(Get-RelationshipEndOnBadCalculatedTable -Definition $copy)
            $ends | Should -Contain "GroupKind.Label (relationship end): Label: sourceColumn 'Label' is not bracketed"
        }

        It 'fails when <Table> goes back to a calculated table whose columns are bare' -ForEach @(
            @{ Table = 'GroupsCurrent'; Base = 'Groups'; End = 'GroupsCurrent.Id' }
            @{ Table = 'UsersCurrent'; Base = 'Users'; End = 'UsersCurrent.Id' }
        ) {
            $copy = New-BrokenCopy -Table $Table -Pattern "(?s)    partition $Table = m.*?(\r?\n\r?\n    annotation)" -Replacement "    partition $Table = calculated`n        mode: import`n        source =`n            FILTER($Base, $Base[RunDate] = MAX($Base[RunDate]))`$1"
            $ends = @(Get-RelationshipEndOnBadCalculatedTable -Definition $copy)
            ($ends -join "`n") | Should -BeLike "*$End (relationship end)*"
        }
    }
}
