#Requires -Version 7.0

BeforeAll {
    $script:DocPath = Join-Path $PSScriptRoot '../exchange-activity.md'
    $script:Doc = Get-Content -LiteralPath $script:DocPath -Raw
}

Describe 'exchange activity sources' {
    It 'lists every mailbox instead of stopping at the default 1000' {
        # Get-EXOMailbox returns 1000 mailboxes unless ResultSize is Unlimited,
        # and it omits inactive and soft-deleted mailboxes unless those switches are set.
        $source6 = @($script:Doc -split '\r?\n' | Where-Object { $_ -match '^\| 6 \| Mailbox list' })
        $source6 | Should -Match '-ResultSize Unlimited'
        $source6 | Should -Match 'defaults to 1000'
        $source6 | Should -Match '-IncludeInactiveMailbox'
        $source6 | Should -Match '-SoftDeletedMailbox'
        $source6 | Should -Match 'View recipient properties'
        $source6 | Should -Not -Match 'were not re-read'
    }

    It 'continues message trace from the last Received time' {
        # A full ResultSize round is not the end of the window. EndDate has to be
        # that row's Received time, and messages to more than 1000 recipients need
        # MessageTraceId or the trace is incomplete.
        $source8 = @($script:Doc -split '\r?\n' | Where-Object { $_ -match '^\| 8 \|' })
        $source8 | Should -Match 'last row''s `Received`'
        $source8 | Should -Match 'RecipientAddress'
        $source8 | Should -Match 'more than 1000 recipients'
        $source8 | Should -Match '-MessageTraceId'
        $source8 | Should -Match 'regional short date'
        $source8 | Should -Match 'not a count of distinct messages'
        $source8 | Should -Match 'does not name a retry'
        $source8 | Should -Not -Match 'from the last row'
    }
}
