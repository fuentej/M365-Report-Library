#Requires -Version 7.0

<#
    Power BI Desktop 2.158 rejected identity-posture, teams-groups-lifecycle,
    copilot-studio-agents and mailbox-exfiltration-risk with "Relationship ... uses
    an invalid column ID" (BRO-373), the error guest-access had before BRO-347.
    Guest-access opened after its *Current tables became import tables and its
    remaining calculated tables used the form Desktop writes (BRO-367). These four
    reports now follow the same rules, and this file holds them to it:

    1. Every *Current table is an import (`m`) partition, not a calculated table,
       and keeps the latest snapshot with List.Max over non-null RunDate values.
    2. Every column a calculated table's expression returns carries
       `isNameInferred` and a bracketed `sourceColumn: [Name]`.
       https://learn.microsoft.com/en-us/dotnet/api/microsoft.analysisservices.tabular.calculatedtablecolumn.isnameinferred
       https://learn.microsoft.com/en-us/dotnet/api/microsoft.analysisservices.tabular.calculatedtablecolumn.sourcecolumn
    3. No relationship end is a column of a calculated table that breaks rule 2,
       and every relationship end is a column the table declares.
#>

BeforeDiscovery {
    $script:Reports = @(
        @{ Report = 'identity-posture'; Current = @('UsersCurrent'); Calculated = @('DateDim', 'AnonymizeMode', 'RoleDefinitions') }
        @{ Report = 'teams-groups-lifecycle'; Current = @('GroupsCurrent', 'UsersCurrent'); Calculated = @('DateDim', 'AnonymizeMode', 'GroupKind') }
        @{ Report = 'copilot-studio-agents'; Current = @('AgentsCurrent', 'EnvironmentsCurrent'); Calculated = @('DateDim', 'AnonymizeMode') }
        @{ Report = 'mailbox-exfiltration-risk'; Current = @('UsersCurrent'); Calculated = @('DateDim', 'AnonymizeMode', 'Apps') }
    )
}

BeforeAll {
    $script:Root = Join-Path $PSScriptRoot '../..' | Resolve-Path | Select-Object -ExpandProperty Path

    function Get-Definition {
        param([Parameter(Mandatory)][string]$Report)
        $model = Get-ChildItem -LiteralPath (Join-Path $script:Root "reports/$Report/report") -Directory -Filter '*.SemanticModel' |
            Select-Object -First 1
        Join-Path $model.FullName 'definition'
    }

    function Get-PartitionKind {
        param([Parameter(Mandatory)][string]$Path)
        $match = Select-String -LiteralPath $Path -Pattern '^    partition\s+.+=\s*(\S+)\s*$' | Select-Object -First 1
        if ($match) { $match.Matches[0].Groups[1].Value }
    }

    function Get-ColumnBlock {
        <#
            .SYNOPSIS
                Every `column` block in a table file: its name, whether it is a
                calculated column (`column X = <expr>`), and its text.
        #>
        param([Parameter(Mandatory)][string]$Path)
        $text = (Get-Content -LiteralPath $Path -Raw) -replace "`r`n", "`n"
        foreach ($block in [regex]::Matches($text, '(?m)^    column\s+(.+?)\s*$\n(?:(?!^    \S).*\n?)*')) {
            $head = $block.Groups[1].Value
            $isCalculated = $head -match '='
            $name = ($head -split '=')[0].Trim().Trim("'")
            [pscustomobject]@{ Name = $name; IsCalculated = $isCalculated; Text = $block.Value }
        }
    }

    function Get-CalculatedTableColumnProblem {
        <#
            .SYNOPSIS
                The returned columns of a calculated table that lack isNameInferred
                or a bracketed sourceColumn. Empty for an import table.
        #>
        param([Parameter(Mandatory)][string]$Path)
        if ((Get-PartitionKind -Path $Path) -ne 'calculated') { return , @() }
        $problems = foreach ($column in Get-ColumnBlock -Path $Path | Where-Object { -not $_.IsCalculated }) {
            $source = [regex]::Match($column.Text, '(?m)^        sourceColumn:\s*(\S+)\s*$')
            if (-not $source.Success) { "$($column.Name): no sourceColumn"; continue }
            if ($source.Groups[1].Value -notmatch '^\[[^\]]+\]$') { "$($column.Name): sourceColumn '$($source.Groups[1].Value)' is not bracketed" }
            if ($column.Text -notmatch '(?m)^        isNameInferred\s*$') { "$($column.Name): missing isNameInferred" }
        }
        return , @($problems)
    }

    function Get-RelationshipProblem {
        <#
            .SYNOPSIS
                Relationship ends that name a column the table does not declare, or a
                column of a calculated table that breaks the Desktop form.
        #>
        param([Parameter(Mandatory)][string]$Definition)
        $tables = Join-Path $Definition 'tables'
        foreach ($line in Get-Content -LiteralPath (Join-Path $Definition 'relationships.tmdl')) {
            if ($line -notmatch '^\s+(?:fromColumn|toColumn):\s*(?<table>.+)\.(?<column>[^.]+?)\s*$') { continue }
            $table = $Matches['table'].Trim("'"); $column = $Matches['column'].Trim("'")
            $path = Join-Path $tables "$table.tmdl"
            if (-not (Test-Path -LiteralPath $path)) { "${table}.${column}: no table file"; continue }
            if (-not (Get-ColumnBlock -Path $path | Where-Object Name -EQ $column)) { "${table}.${column}: column not declared"; continue }
            foreach ($problem in Get-CalculatedTableColumnProblem -Path $path) {
                if ($problem -like "${column}:*") { "${table}.${column} (relationship end): $problem" }
            }
        }
    }

    function Test-LatestRunDateIgnoresNull {
        param([Parameter(Mandatory)][string]$Path)
        (Get-Content -LiteralPath $Path -Raw) -match 'List\.Max\(List\.RemoveNulls\((?:Selected)?\[RunDate\]\)\)'
    }

    function New-BrokenCopy {
        param([string]$Definition, [string]$Table, [string]$Pattern, [string]$Replacement)
        $copy = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        Copy-Item -LiteralPath $Definition -Destination $copy -Recurse
        $path = Join-Path $copy "tables/$Table.tmdl"
        $text = (Get-Content -LiteralPath $path -Raw) -replace $Pattern, $Replacement
        Set-Content -LiteralPath $path -Value $text -NoNewline
        $copy
    }
}

Describe '<Report> opens the way guest-access does' -ForEach $script:Reports {
    BeforeAll {
        $script:Definition = Get-Definition -Report $Report
        $script:Tables = Join-Path $script:Definition 'tables'
    }

    It '<_> is an import table, not a calculated table' -ForEach $Current {
        Get-PartitionKind -Path (Join-Path $script:Tables "$_.tmdl") | Should -Be 'm'
    }

    It '<_> keeps the latest snapshot with List.Max of the non-null RunDate values' -ForEach $Current {
        Test-LatestRunDateIgnoresNull -Path (Join-Path $script:Tables "$_.tmdl") | Should -BeTrue
    }

    It 'has exactly the expected calculated tables' {
        $found = @(Get-ChildItem -LiteralPath $script:Tables -Filter '*.tmdl' -File |
                Where-Object { (Get-PartitionKind -Path $_.FullName) -eq 'calculated' } | ForEach-Object BaseName)
        ($found | Sort-Object) -join ',' | Should -Be (($Calculated | Sort-Object) -join ',')
    }

    It '<_> binds every returned column with isNameInferred and a bracketed sourceColumn' -ForEach $Calculated {
        $problems = Get-CalculatedTableColumnProblem -Path (Join-Path $script:Tables "$_.tmdl")
        $problems | Should -BeNullOrEmpty
    }

    It 'has no relationship end that is undeclared or on a calculated-table column that breaks the rule' {
        Get-RelationshipProblem -Definition $script:Definition | Should -BeNullOrEmpty
    }
}

Describe 'The checks fail on a deliberately broken copy' {
    BeforeAll {
        $script:Identity = Get-Definition -Report 'identity-posture'
        $script:Mailbox = Get-Definition -Report 'mailbox-exfiltration-risk'
    }

    It 'flags a calculated-table column with a bare sourceColumn at a relationship end' {
        $copy = New-BrokenCopy -Definition $script:Identity -Table 'RoleDefinitions' -Pattern 'sourceColumn: \[RoleDefinitionId\]' -Replacement 'sourceColumn: RoleDefinitionId'
        $problems = @(Get-RelationshipProblem -Definition $copy)
        ($problems -join "`n") | Should -BeLike "*RoleDefinitions.RoleDefinitionId (relationship end): RoleDefinitionId: sourceColumn 'RoleDefinitionId' is not bracketed*"
    }

    It 'flags a calculated-table column missing isNameInferred' {
        $copy = New-BrokenCopy -Definition $script:Mailbox -Table 'Apps' -Pattern '(?m)^        isNameInferred\r?\n' -Replacement ''
        $problems = Get-CalculatedTableColumnProblem -Path (Join-Path $copy 'tables/Apps.tmdl')
        $problems | Should -Contain 'AppId: missing isNameInferred'
    }

    It 'flags a *Current table turned back into a calculated table with bare columns' {
        $copy = New-BrokenCopy -Definition $script:Identity -Table 'UsersCurrent' -Pattern '(?s)    partition UsersCurrent = m.*?(\r?\n\r?\n    annotation)' -Replacement "    partition UsersCurrent = calculated`n        mode: import`n        source =`n            FILTER(Users, Users[RunDate] = MAX(Users[RunDate]))`$1"
        Get-PartitionKind -Path (Join-Path $copy 'tables/UsersCurrent.tmdl') | Should -Be 'calculated'
        ((Get-RelationshipProblem -Definition $copy) -join "`n") | Should -BeLike '*UsersCurrent.Id (relationship end)*'
    }

    It 'flags a relationship end the table does not declare' {
        $copy = New-BrokenCopy -Definition $script:Mailbox -Table 'UsersCurrent' -Pattern '(?m)^    column UserPrincipalName\s*$' -Replacement '    column UserPrincipalNameRenamed'
        Get-RelationshipProblem -Definition $copy | Should -Contain 'UsersCurrent.UserPrincipalName: column not declared'
    }

    It 'flags List.Max over every RunDate, including null' {
        $copy = New-BrokenCopy -Definition $script:Mailbox -Table 'UsersCurrent' -Pattern 'List\.Max\(List\.RemoveNulls\(Selected\[RunDate\]\)\)' -Replacement 'List.Max(Selected[RunDate])'
        Test-LatestRunDateIgnoresNull -Path (Join-Path $copy 'tables/UsersCurrent.tmdl') | Should -BeFalse
    }
}
