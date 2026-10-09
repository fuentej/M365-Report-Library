#Requires -Version 7.0

<#
    A calculated table's declared columns must be exactly the columns its
    SELECTCOLUMNS expression outputs, one to one by name. A declared column the
    expression does not produce, or an output column that is not declared (for
    example a calculated column inherited from the source table), can leave a
    relationship end pointing at a column the engine never created. Desktop
    reports that as PFE_TM_RELATIONSHIP_END_COLUMN_INVALID.

    Calculated columns added on top of the table (declared with `= <expr>`) are
    not part of the expression output and are ignored here.

    https://learn.microsoft.com/analysis-services/tabular-models/create-a-calculated-table-ssas-tabular
#>

BeforeAll {
    $script:TablesFolder = Join-Path $PSScriptRoot '../report/GuestAccess.SemanticModel/definition/tables' | Resolve-Path | Select-Object -ExpandProperty Path

    function Get-CalculatedTableColumnMismatch {
        <#
            .SYNOPSIS
                Compares declared sourceColumn columns with the names given to
                SELECTCOLUMNS. Returns a list of problems, empty when they match.
        #>
        param([Parameter(Mandatory)][AllowEmptyString()][string[]]$Line)

        $text = $Line -join "`n"
        if ($text -notmatch '(?m)^    partition\s+.+=\s*calculated\s*$') { return , @() }

        $declared = @(
            [regex]::Matches($text, '(?m)^    column\s+(\S+)\s*$\n(?:(?!^    \S).*\n?)*?^        sourceColumn:\s*(\S+)\s*$') |
                ForEach-Object { $_.Groups[2].Value }
        )

        $source = ($text -split '(?m)^        source =\s*$\n', 2)[1]
        if (-not $source -or $source -notmatch 'SELECTCOLUMNS\(') {
            return , @('the source expression does not use SELECTCOLUMNS, so its output columns are not explicit')
        }
        $produced = @([regex]::Matches($source, '(?m)^\s+"([^"]+)",\s') | ForEach-Object { $_.Groups[1].Value })

        $problems = [System.Collections.Generic.List[string]]::new()
        foreach ($name in $declared | Where-Object { $_ -notin $produced }) { $problems.Add("declared but not produced: $name") }
        foreach ($name in $produced | Where-Object { $_ -notin $declared }) { $problems.Add("produced but not declared: $name") }
        foreach ($name in $declared | Group-Object | Where-Object Count -gt 1) { $problems.Add("declared twice: $($name.Name)") }
        foreach ($name in $produced | Group-Object | Where-Object Count -gt 1) { $problems.Add("produced twice: $($name.Name)") }
        return , $problems.ToArray()
    }

    $script:CalculatedFiles = @(
        Get-ChildItem -LiteralPath $script:TablesFolder -Filter '*.tmdl' -File |
            Where-Object { (Get-Content -LiteralPath $_.FullName -Raw) -match '(?m)^    partition\s+.+=\s*calculated\s*$' }
    )
}

Describe 'Calculated table columns match their SELECTCOLUMNS output' {
    It 'finds the calculated tables' {
        $script:CalculatedFiles.Name | Should -Contain 'GuestsCurrent.tmdl'
        $script:CalculatedFiles.Name | Should -Contain 'UsersCurrent.tmdl'
        $script:CalculatedFiles.Name | Should -Contain 'GuestMembershipsCurrent.tmdl'
    }

    It '<_> declares exactly the columns its expression outputs' -ForEach @('GuestsCurrent.tmdl', 'UsersCurrent.tmdl', 'GuestMembershipsCurrent.tmdl') {
        $problems = Get-CalculatedTableColumnMismatch -Line (Get-Content -LiteralPath (Join-Path $script:TablesFolder $_))
        $problems | Should -BeNullOrEmpty
    }

    It 'fails on a copy with an inherited column missing from the declarations' {
        $lines = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'GuestsCurrent.tmdl')
        $broken = $lines -join "`n" -replace '(?m)^    column Mail\n(?:        .*\n)+', ''
        $problems = Get-CalculatedTableColumnMismatch -Line ($broken -split "`n")
        $problems | Should -Contain 'produced but not declared: Mail'
    }

    It 'fails on a copy that declares a column the expression does not output' {
        $lines = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'GuestsCurrent.tmdl')
        $broken = $lines -join "`n" -replace '(?m)^    partition', "    column DaysSinceCreated`n        dataType: int64`n        sourceColumn: DaysSinceCreated`n`n    partition"
        $problems = Get-CalculatedTableColumnMismatch -Line ($broken -split "`n")
        $problems | Should -Contain 'declared but not produced: DaysSinceCreated'
    }

    It 'fails on a copy that goes back to a bare FILTER' {
        $lines = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'GuestsCurrent.tmdl')
        $broken = $lines -join "`n" -replace '(?s)(        source =\n).*?(\n\n    annotation)', "`$1            FILTER(Guests, Guests[RunDate] = MAX(Guests[RunDate]))`$2"
        $problems = Get-CalculatedTableColumnMismatch -Line ($broken -split "`n")
        $problems | Should -Not -BeNullOrEmpty
    }
}
