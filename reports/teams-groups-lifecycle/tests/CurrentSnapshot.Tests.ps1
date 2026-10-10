#Requires -Version 7.0

<#
    GroupsCurrent and UsersCurrent keep the latest snapshot with Power Query
    (BRO-434). They read the same CSV as Groups and Users, keep the CSV columns
    with Table.SelectColumns, and keep the rows whose RunDate is the latest one
    with Table.SelectRows. Blank CSV cells are replaced with null before that
    step. List.Max includes nulls unless they are removed first, and a null
    maximum matches no row, so the snapshot would come back empty.

    https://learn.microsoft.com/en-us/powerquery-m/list-max
    https://learn.microsoft.com/en-us/powerquery-m/list-removenulls
    https://learn.microsoft.com/en-us/powerquery-m/table-selectrows
    https://learn.microsoft.com/en-us/powerquery-m/table-selectcolumns
#>

BeforeAll {
    $script:TablesFolder = Join-Path $PSScriptRoot '../report/TeamsGroupsLifecycle.SemanticModel/definition/tables' |
        Resolve-Path | Select-Object -ExpandProperty Path

    function Test-LatestRunDateIgnoresNull {
        <#
            .SYNOPSIS
                True when the latest-snapshot step drops null RunDate values
                before List.Max and keeps only the rows on that date.
        #>
        param([Parameter(Mandatory)][string]$Path)

        $text = Get-Content -LiteralPath $Path -Raw
        return ($text -match 'LatestRunDate = List\.Max\(List\.RemoveNulls\(Selected\[RunDate\]\)\)') -and
            ($text -match 'Table\.SelectRows\(Selected, each \[RunDate\] = LatestRunDate\)') -and
            ($text -match '(?m)^\s+Latest\s*$')
    }

    function Get-PartitionBody {
        <#
            .SYNOPSIS
                The M of a table's partition, up to the line before `in`.
        #>
        param([Parameter(Mandatory)][string]$Path)
        $text = Get-Content -LiteralPath $Path -Raw
        $m = [regex]::Match($text, '(?s)source =\s*\r?\n\s*let\s*\r?\n(?<body>.*?)\r?\n\s*in\s*\r?\n')
        if (-not $m.Success) { return $null }
        return $m.Groups['body'].Value
    }
}

Describe 'Current snapshot tables ignore a blank RunDate' {
    It '<_> takes List.Max of the non-null RunDate values' -ForEach @(
        'GroupsCurrent.tmdl'
        'UsersCurrent.tmdl'
    ) {
        Test-LatestRunDateIgnoresNull -Path (Join-Path $script:TablesFolder $_) | Should -BeTrue
    }

    It '<Table> reads its CSV the same way <Base> does' -ForEach @(
        @{ Table = 'GroupsCurrent'; Base = 'Groups' }
        @{ Table = 'UsersCurrent'; Base = 'Users' }
    ) {
        # The Current table copies the source partition's steps through
        # #"Changed Type", then adds the select and latest-snapshot steps.
        $base = Get-PartitionBody -Path (Join-Path $script:TablesFolder "$Base.tmdl")
        $current = Get-PartitionBody -Path (Join-Path $script:TablesFolder "$Table.tmdl")
        $base | Should -Not -BeNullOrEmpty
        $current | Should -Not -BeNullOrEmpty
        $current.StartsWith($base.TrimEnd() + ',') | Should -BeTrue -Because "$Table should start with the $Base partition's M"
    }

    It 'fails when GroupsCurrent goes back to List.Max of every RunDate, including null' {
        $text = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'GroupsCurrent.tmdl') -Raw
        $broken = $text -replace 'List\.Max\(List\.RemoveNulls\(Selected\[RunDate\]\)\)', 'List.Max(Selected[RunDate])'
        $broken | Should -Not -Be $text
        $path = Join-Path $TestDrive 'GroupsCurrent.tmdl'
        Set-Content -LiteralPath $path -Value $broken -NoNewline
        Test-LatestRunDateIgnoresNull -Path $path | Should -BeFalse
    }

    It 'fails when UsersCurrent keeps every snapshot instead of the latest' {
        $text = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'UsersCurrent.tmdl') -Raw
        $broken = $text -replace '(?m)^(\s+)Latest = Table\.SelectRows\(Selected, each \[RunDate\] = LatestRunDate\)', '$1Latest = Selected'
        $broken | Should -Not -Be $text
        $path = Join-Path $TestDrive 'UsersCurrent.tmdl'
        Set-Content -LiteralPath $path -Value $broken -NoNewline
        Test-LatestRunDateIgnoresNull -Path $path | Should -BeFalse
    }

    It 'fails when GroupsCurrent stops reading the same CSV as Groups' {
        $text = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'GroupsCurrent.tmdl') -Raw
        $broken = $text -replace [regex]::Escape('{"IsTeam", type logical}'), '{"IsTeam", type text}'
        $broken | Should -Not -Be $text
        $path = Join-Path $TestDrive 'GroupsCurrent-type.tmdl'
        Set-Content -LiteralPath $path -Value $broken -NoNewline
        $base = Get-PartitionBody -Path (Join-Path $script:TablesFolder 'Groups.tmdl')
        (Get-PartitionBody -Path $path).StartsWith($base.TrimEnd() + ',') | Should -BeFalse
    }
}
