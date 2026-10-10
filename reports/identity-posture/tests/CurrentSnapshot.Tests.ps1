#Requires -Version 7.0

<#
    UsersCurrent is an import table that reads users.csv (BRO-433). It keeps, for
    each user Id, the row from that user's own latest RunDate, which is what the
    calculated table it replaces did with
    CALCULATE(MAX(Users[RunDate]), ALLEXCEPT(Users, Users[Id])). It does not keep
    only the newest file date: a user missing from the last snapshot would lose
    every historical row through the one-side relationship.

    Blank CSV cells are replaced with null before this step. List.Max includes
    nulls unless they are removed first. Id is the one side of eight
    relationships, so the result holds one row per Id.

    https://learn.microsoft.com/en-us/powerquery-m/table-group
    https://learn.microsoft.com/en-us/powerquery-m/list-max
    https://learn.microsoft.com/en-us/powerquery-m/list-removenulls
    https://learn.microsoft.com/en-us/powerquery-m/table-nestedjoin
    https://learn.microsoft.com/en-us/powerquery-m/table-distinct
#>

BeforeAll {
    $script:TablesFolder = Join-Path $PSScriptRoot '../report/IdentityPosture.SemanticModel/definition/tables' |
        Resolve-Path | Select-Object -ExpandProperty Path
    $script:UsersCsv = Join-Path $PSScriptRoot '../samples/users.csv' | Resolve-Path | Select-Object -ExpandProperty Path

    function Get-CurrentSnapshotFault {
        <#
            .SYNOPSIS
                What is wrong with the per-Id latest-snapshot steps of a
                UsersCurrent.tmdl. Empty when the steps are all there.
        #>
        param([Parameter(Mandatory)][string]$Path)

        $text = Get-Content -LiteralPath $Path -Raw
        $faults = [System.Collections.Generic.List[string]]::new()
        if ($text -notmatch 'Table\.Group\(Selected,\s*\{"Id"\},\s*\{\{"LatestRunDate",\s*each List\.Max\(List\.RemoveNulls\(\[RunDate\]\)\)') {
            $faults.Add('no per-Id List.Max of the non-null RunDate values')
        }
        if ($text -notmatch 'Table\.NestedJoin\(Selected,\s*\{"Id",\s*"RunDate"\},\s*LatestPerId,\s*\{"Id",\s*"LatestRunDate"\},\s*"LatestPerId",\s*JoinKind\.Inner\)') {
            $faults.Add('does not inner-join each row to its Id''s latest RunDate')
        }
        if ($text -notmatch 'Table\.Distinct\([^\r\n]*\{"Id"\}\)') {
            $faults.Add('does not keep one row per Id')
        }
        if ($text -match 'List\.Max\([^\r\n]*Selected\[RunDate\]') {
            $faults.Add('takes the newest RunDate of the whole file, not per Id')
        }
        return , $faults.ToArray()
    }

    function Select-LatestPerId {
        <#
            .SYNOPSIS
                The same rule as the UsersCurrent query, on Import-Csv rows: per Id,
                the rows whose RunDate is that Id's latest non-empty RunDate, then
                the first of them.
        #>
        param([Parameter(Mandatory)][object[]]$Row)

        foreach ($group in $Row | Group-Object Id) {
            $dates = @($group.Group.RunDate | Where-Object { $_ })
            $latest = ($dates | Sort-Object -Descending | Select-Object -First 1)
            $group.Group | Where-Object { $_.RunDate -eq $latest } | Select-Object -First 1
        }
    }
}

Describe 'UsersCurrent keeps each user at that user''s latest snapshot' {
    It 'has every per-Id latest-snapshot step' {
        Get-CurrentSnapshotFault -Path (Join-Path $script:TablesFolder 'UsersCurrent.tmdl') | Should -BeNullOrEmpty
    }

    It 'reads users.csv, the file Users reads' {
        $current = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'UsersCurrent.tmdl') -Raw
        $users = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'Users.tmdl') -Raw
        $read = [regex]::Escape('File.Contents(CsvFolder & "\users.csv")')
        $current | Should -Match $read
        $users | Should -Match $read
    }

    It 'gives one row per Id on samples/users.csv' {
        $rows = @(Import-Csv -LiteralPath $script:UsersCsv)
        $ids = @($rows.Id | Sort-Object -Unique)
        $rows.Count | Should -BeGreaterThan $ids.Count -Because 'the sample holds more than one snapshot per user'
        $latest = @(Select-LatestPerId -Row $rows)
        $latest.Count | Should -Be $ids.Count
        @($latest.Id | Sort-Object -Unique).Count | Should -Be $latest.Count
        foreach ($row in $latest) {
            $row.RunDate | Should -Be (@($rows | Where-Object Id -eq $row.Id | ForEach-Object RunDate) | Sort-Object -Descending | Select-Object -First 1)
        }
    }

    It 'keeps a user who is missing from the newest snapshot, which a file-wide latest would drop' {
        $rows = @(
            [pscustomobject]@{ RunDate = '2026-09-01'; Id = 'a' }
            [pscustomobject]@{ RunDate = '2026-09-01'; Id = 'b' }
            [pscustomobject]@{ RunDate = '2026-10-01'; Id = 'a' }
            [pscustomobject]@{ RunDate = ''; Id = 'b' }
        )
        $latest = @(Select-LatestPerId -Row $rows | Sort-Object Id)
        ($latest | ForEach-Object { "$($_.Id)=$($_.RunDate)" }) -join ',' | Should -Be 'a=2026-10-01,b=2026-09-01'
    }

    Context 'on a deliberately broken copy' {
        BeforeAll {
            function New-BrokenCopy {
                param([string]$Pattern, [string]$Replacement)
                $text = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'UsersCurrent.tmdl') -Raw
                $broken = $text -replace $Pattern, $Replacement
                if ($broken -eq $text) { throw 'The pattern did not change UsersCurrent.tmdl' }
                $path = Join-Path $TestDrive ([guid]::NewGuid().ToString('N') + '.tmdl')
                Set-Content -LiteralPath $path -Value $broken -NoNewline
                $path
            }
        }

        It 'fails when the per-Id maximum includes null RunDate values' {
            $path = New-BrokenCopy -Pattern 'List\.Max\(List\.RemoveNulls\(\[RunDate\]\)\)' -Replacement 'List.Max([RunDate])'
            $faults = Get-CurrentSnapshotFault -Path $path
            $faults | Should -Contain 'no per-Id List.Max of the non-null RunDate values'
        }

        It 'fails when UsersCurrent keeps only the newest file date, as guest-access does' {
            $path = New-BrokenCopy -Pattern '(?s)LatestPerId = Table\.Group.*?OnePerId = [^\r\n]*' -Replacement "LatestRunDate = List.Max(List.RemoveNulls(Selected[RunDate])),`r`n                OnePerId = Table.SelectRows(Selected, each [RunDate] = LatestRunDate)"
            $faults = Get-CurrentSnapshotFault -Path $path
            $faults | Should -Contain 'takes the newest RunDate of the whole file, not per Id'
            $faults | Should -Contain 'does not keep one row per Id'
        }

        It 'fails when the join is not an inner join on Id and RunDate' {
            $path = New-BrokenCopy -Pattern 'JoinKind\.Inner' -Replacement 'JoinKind.LeftOuter'
            $faults = Get-CurrentSnapshotFault -Path $path
            $faults | Should -Contain 'does not inner-join each row to its Id''s latest RunDate'
        }

        It 'fails when the one-row-per-Id step is removed' {
            $path = New-BrokenCopy -Pattern 'Table\.Distinct\(Table\.Buffer\(Latest\), \{"Id"\}\)' -Replacement 'Latest'
            $faults = Get-CurrentSnapshotFault -Path $path
            $faults | Should -Contain 'does not keep one row per Id'
        }
    }
}
