#Requires -Version 7.0

BeforeAll {
    $script:DocPath = Join-Path $PSScriptRoot '../unified-audit-log.md'
    $script:Doc = Get-Content -LiteralPath $script:DocPath -Raw
}

Describe 'unified audit log sources' {
    It 'pages Search-UnifiedAuditLog until the window is complete and keeps dates in UTC' {
        # ReturnLargeSet stops at 50,000 and can still have more records.
        # A date with no time is midnight UTC. Without HighCompleteness, results can be missing.
        $source1 = @($script:Doc -split '\r?\n' | Where-Object { $_ -match '^\| 1 \| Unified audit log through Exchange' })
        $source1 | Should -Match 'at most 100 records|Default 100 records'
        $source1 | Should -Match 'maximum 5,000'
        $source1 | Should -Match 'moreRecordsAvailable'
        $source1 | Should -Match 'while `moreRecordsAvailable` is still true'
        $source1 | Should -Match 'HighCompleteness'
        $source1 | Should -Match 'midnight UTC'
        $source1 | Should -Match 'must not be treated as complete'
        $source1 | Should -Match 'results can be missing'
    }
}
