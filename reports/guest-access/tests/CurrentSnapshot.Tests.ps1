#Requires -Version 7.0

<#
    GuestsCurrent, UsersCurrent and GuestMembershipsCurrent keep the latest
    snapshot with Power Query. Blank CSV cells are replaced with null before
    that step. List.Max includes nulls unless they are removed first, and a
    null maximum matches no row, so the snapshot comes back empty.

    https://learn.microsoft.com/en-us/powerquery-m/list-max
    https://learn.microsoft.com/en-us/powerquery-m/list-removenulls
#>

BeforeAll {
    $script:TablesFolder = Join-Path $PSScriptRoot '../report/GuestAccess.SemanticModel/definition/tables' |
        Resolve-Path | Select-Object -ExpandProperty Path

    function Test-LatestRunDateIgnoresNull {
        <#
            .SYNOPSIS
                True when the latest-snapshot step drops null RunDate values
                before List.Max.
        #>
        param([Parameter(Mandatory)][string]$Path)

        $text = Get-Content -LiteralPath $Path -Raw
        return $text -match 'List\.Max\(List\.RemoveNulls\(Selected\[RunDate\]\)\)'
    }
}

Describe 'Current snapshot tables ignore a blank RunDate' {
    It '<_> takes List.Max of the non-null RunDate values' -ForEach @(
        'GuestsCurrent.tmdl'
        'UsersCurrent.tmdl'
        'GuestMembershipsCurrent.tmdl'
    ) {
        Test-LatestRunDateIgnoresNull -Path (Join-Path $script:TablesFolder $_) | Should -BeTrue
    }

    It 'fails when GuestsCurrent goes back to List.Max of every RunDate, including null' {
        $text = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'GuestsCurrent.tmdl') -Raw
        $broken = $text -replace 'List\.Max\(List\.RemoveNulls\(Selected\[RunDate\]\)\)', 'List.Max(Selected[RunDate])'
        $path = Join-Path $TestDrive 'GuestsCurrent.tmdl'
        Set-Content -LiteralPath $path -Value $broken -NoNewline
        Test-LatestRunDateIgnoresNull -Path $path | Should -BeFalse
    }
}
