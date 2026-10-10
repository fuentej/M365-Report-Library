#Requires -Version 7.0

<#
    Power BI Desktop 2.158 rejected the identity-posture project with
    PFE_TM_RELATIONSHIP_END_COLUMN_INVALID ("Relationship '1a16a17c-...' uses an
    invalid column ID 71", the Users.Id -> UsersCurrent.Id relationship, BRO-433).
    That is the error guest-access raised before BRO-347. A relationship end that
    is a column of a calculated table has to be bound the way Desktop binds it:
    `isNameInferred` plus a bracketed `sourceColumn: [Name]`, where Name is the
    column the expression returns. The bare `sourceColumn: Name` form is what an
    import table uses.

    Two rules, both checked on the identity-posture model:

    1. Every column of a calculated table that is bound to the expression's output
       carries `isNameInferred` and a bracketed `sourceColumn`. Calculated columns
       (`column X = <expr>`) are not part of the output and are skipped.
    2. Every relationship end that is a column of a calculated table meets rule 1.
       UsersCurrent is an import table loaded from users.csv, so the relationship
       ends left on calculated tables are DateDim.Date and
       RoleDefinitions.RoleDefinitionId.

    https://learn.microsoft.com/dotnet/api/microsoft.analysisservices.tabular.calculatedtablecolumn.sourcecolumn
    https://learn.microsoft.com/dotnet/api/microsoft.analysisservices.tabular.calculatedtablecolumn.isnameinferred
    https://learn.microsoft.com/en-us/analysis-services/tmdl/tmdl-overview
#>

BeforeAll {
    $script:Definition = Join-Path $PSScriptRoot '../report/IdentityPosture.SemanticModel/definition' | Resolve-Path | Select-Object -ExpandProperty Path
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
        ($script:CalculatedFiles.Name | Sort-Object) -join ',' | Should -Be 'AnonymizeMode.tmdl,DateDim.tmdl,RoleDefinitions.tmdl'
    }

    It 'has rebuilt UsersCurrent as an import table' {
        $script:CalculatedFiles.Name | Should -Not -Contain 'UsersCurrent.tmdl'
        $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'UsersCurrent.tmdl') -Raw
        $tmdl | Should -Match '(?m)^    partition UsersCurrent = m\s*$'
        $tmdl | Should -Match '(?m)^        mode: import\s*$'
    }

    It '<File> binds <Column> with isNameInferred and sourceColumn [<Column>]' -ForEach @(
        @{ File = 'DateDim.tmdl'; Column = 'Date' }
        @{ File = 'AnonymizeMode.tmdl'; Column = 'Mode' }
        @{ File = 'RoleDefinitions.tmdl'; Column = 'RoleDefinitionId' }
    ) {
        $lines = Get-Content -LiteralPath (Join-Path $script:TablesFolder $File)
        Get-CalculatedTableColumnProblem -Line $lines | Should -BeNullOrEmpty
        ($lines -join "`n") | Should -Match ("(?m)^    column $Column\s*\n(?:        .*\n)*?        sourceColumn: \[$Column\]\s*$")
    }

    It 'has no relationship end on a calculated-table column that breaks the rule' {
        Get-RelationshipEndOnBadCalculatedTable -Definition $script:Definition | Should -BeNullOrEmpty
    }

    It 'still has relationships that end on DateDim, RoleDefinitions and UsersCurrent, so the check is not vacuous' {
        $ends = @(Get-RelationshipEnd -Path (Join-Path $script:Definition 'relationships.tmdl') | ForEach-Object { "$($_.Table).$($_.Column)" })
        @($ends | Where-Object { $_ -eq 'DateDim.Date' }).Count | Should -Be 9
        @($ends | Where-Object { $_ -eq 'RoleDefinitions.RoleDefinitionId' }).Count | Should -Be 3
        @($ends | Where-Object { $_ -eq 'UsersCurrent.Id' }).Count | Should -Be 8
    }

    Context 'on a deliberately broken copy' {
        BeforeAll {
            function New-BrokenCopy {
                param([string]$Table, [string]$Pattern, [string]$Replacement)
                $copy = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
                Copy-Item -LiteralPath $script:Definition -Destination $copy -Recurse
                $path = Join-Path $copy "tables/$Table.tmdl"
                $text = Get-Content -LiteralPath $path -Raw
                $broken = $text -replace $Pattern, $Replacement
                if ($broken -eq $text) { throw "The pattern did not change $Table.tmdl" }
                Set-Content -LiteralPath $path -Value $broken -NoNewline
                $copy
            }
        }

        It 'fails when a calculated-table column uses a bare sourceColumn' {
            $copy = New-BrokenCopy -Table 'DateDim' -Pattern 'sourceColumn: \[Date\]' -Replacement 'sourceColumn: Date'
            $problems = Get-CalculatedTableColumnProblem -Line (Get-Content -LiteralPath (Join-Path $copy 'tables/DateDim.tmdl'))
            $problems | Should -Contain "Date: sourceColumn 'Date' is not bracketed"
        }

        It 'fails when a calculated-table column lacks isNameInferred' {
            $copy = New-BrokenCopy -Table 'RoleDefinitions' -Pattern '(?m)^        isNameInferred\r?\n' -Replacement ''
            $problems = Get-CalculatedTableColumnProblem -Line (Get-Content -LiteralPath (Join-Path $copy 'tables/RoleDefinitions.tmdl'))
            $problems | Should -Contain 'RoleDefinitionId: missing isNameInferred'
        }

        It 'fails when a relationship end is a column of a broken calculated table' {
            $copy = New-BrokenCopy -Table 'RoleDefinitions' -Pattern 'sourceColumn: \[RoleDefinitionId\]' -Replacement 'sourceColumn: RoleDefinitionId'
            $ends = @(Get-RelationshipEndOnBadCalculatedTable -Definition $copy)
            $ends | Should -Contain "RoleDefinitions.RoleDefinitionId (relationship end): RoleDefinitionId: sourceColumn 'RoleDefinitionId' is not bracketed"
        }

        It 'fails when UsersCurrent goes back to the FILTER calculated table with bare columns' {
            $copy = New-BrokenCopy -Table 'UsersCurrent' -Pattern '(?s)    partition UsersCurrent = m.*?(\r?\n\r?\n    annotation)' -Replacement "    partition UsersCurrent = calculated`n        mode: import`n        source =`n            FILTER(Users, Users[RunDate] = CALCULATE(MAX(Users[RunDate]), ALLEXCEPT(Users, Users[Id])))`$1"
            $ends = @(Get-RelationshipEndOnBadCalculatedTable -Definition $copy)
            ($ends -join "`n") | Should -BeLike '*UsersCurrent.Id (relationship end): Id: *'
        }
    }
}
