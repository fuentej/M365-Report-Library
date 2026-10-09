#Requires -Version 7.0

<#
    Power BI Desktop rejects the guest-access project with
    PFE_TM_RELATIONSHIP_END_COLUMN_INVALID ("Relationship ... uses an invalid column
    ID") when a relationship end is a column of a calculated table that is not bound
    the way Desktop binds it. Desktop writes every column that a calculated table's
    expression returns as `isNameInferred` plus a bracketed `sourceColumn: [Name]`.
    The bare `sourceColumn: Name` form is what an import table uses.

    Two rules, both checked on the guest-access model:

    1. Every column of a calculated table that is bound to the expression's output
       carries `isNameInferred` and a bracketed `sourceColumn`. Calculated columns
       (`column X = <expr>`) are not part of the output and are skipped.
    2. Every relationship end that is a column of a calculated table meets rule 1.
       The three *Current tables are import tables loaded from the CSV, so no
       relationship ends on a calculated table except DateDim.

    Examples of the Desktop form, from projects saved by Power BI Desktop:
    RuiRomano/pbip-demo, src/Model02.SemanticModel/definition/tables/LocalDateTable_cde5d4de-d289-40fd-b18f-100e83821e85.tmdl
    data-goblin/power-bi-agentic-development, plugins/reports/skills/review-report/usage-metrics-dataset/Usage Metrics Report.SemanticModel/definition/tables/Users.tmdl
#>

BeforeAll {
    $script:Definition = Join-Path $PSScriptRoot '../report/GuestAccess.SemanticModel/definition' | Resolve-Path | Select-Object -ExpandProperty Path
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
    It 'finds the calculated tables' {
        $script:CalculatedFiles.Name | Should -Contain 'DateDim.tmdl'
        $script:CalculatedFiles.Name | Should -Contain 'AnonymizeMode.tmdl'
    }

    It 'has rebuilt the *Current tables as import tables' {
        $script:CalculatedFiles.Name | Should -Not -Contain 'GuestsCurrent.tmdl'
        $script:CalculatedFiles.Name | Should -Not -Contain 'UsersCurrent.tmdl'
        $script:CalculatedFiles.Name | Should -Not -Contain 'GuestMembershipsCurrent.tmdl'
    }

    It '<_> binds every returned column with isNameInferred and a bracketed sourceColumn' -ForEach @('DateDim.tmdl', 'AnonymizeMode.tmdl') {
        $problems = Get-CalculatedTableColumnProblem -Line (Get-Content -LiteralPath (Join-Path $script:TablesFolder $_))
        $problems | Should -BeNullOrEmpty
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
                $text = (Get-Content -LiteralPath $path -Raw) -replace $Pattern, $Replacement
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
            $copy = New-BrokenCopy -Table 'DateDim' -Pattern '(?m)^        isNameInferred\r?\n' -Replacement ''
            $problems = Get-CalculatedTableColumnProblem -Line (Get-Content -LiteralPath (Join-Path $copy 'tables/DateDim.tmdl'))
            $problems | Should -Contain 'Date: missing isNameInferred'
        }

        It 'fails when a relationship end is a column of a broken calculated table' {
            $copy = New-BrokenCopy -Table 'DateDim' -Pattern 'sourceColumn: \[Date\]' -Replacement 'sourceColumn: Date'
            $ends = @(Get-RelationshipEndOnBadCalculatedTable -Definition $copy)
            $ends | Should -Contain "DateDim.Date (relationship end): Date: sourceColumn 'Date' is not bracketed"
        }

        It 'fails when a *Current table goes back to a calculated table whose columns are bare' {
            $copy = New-BrokenCopy -Table 'GuestsCurrent' -Pattern '(?s)    partition GuestsCurrent = m.*?(\r?\n\r?\n    annotation)' -Replacement "    partition GuestsCurrent = calculated`n        mode: import`n        source =`n            FILTER(Guests, Guests[RunDate] = MAX(Guests[RunDate]))`$1"
            $ends = @(Get-RelationshipEndOnBadCalculatedTable -Definition $copy)
            ($ends -join "`n") | Should -BeLike '*GuestsCurrent.Id (relationship end)*'
        }
    }
}
