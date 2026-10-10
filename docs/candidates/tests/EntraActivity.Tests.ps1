#Requires -Version 7.0

BeforeAll {
    $script:DocPath = Join-Path $PSScriptRoot '../entra-activity.md'
    $script:Doc = Get-Content -LiteralPath $script:DocPath -Raw
}

Describe 'entra activity sources' {
    It 'keeps every sign-in error code and the user on a failed sign-in' {
        # errorCode is Int32. Success is 0. The signInStatus example is 1024, not only 5-6 digits.
        # createdDateTime is UTC. Guests are not stored as #EXT#. 50058 can omit the username.
        $script:Doc | Should -Not -Match 'is a 5-6 digit integer'
        $script:Doc | Should -Match 'JSON example is `1024`'
        $script:Doc | Should -Match 'does not require 5-6 digits'
        $script:Doc | Should -Match '`additionalDetails`'
        $script:Doc | Should -Match 'always lowercase'
        $script:Doc | Should -Match '`#EXT#`'
        $script:Doc | Should -Match '`00000000-0000-0000`'
        $script:Doc | Should -Match '`createdDateTime` is UTC'
        $script:Doc | Should -Match '`unknownFutureValue`'
        $source3 = @($script:Doc -split '\r?\n' | Where-Object { $_ -match '^\| 3 \| Conditional Access result' })
        $source3 | Should -Match 'report-only four'
    }

    It 'pages sign-ins and directory audits and retries 429' {
        # Sign-ins cap a page at 1,000 and do not list $skip. Directory audits state no page size.
        # 429 on these resources is Retry-After, then a shorter window. Five calls per 10 seconds per tenant.
        $source1 = @($script:Doc -split '\r?\n' | Where-Object { $_ -match '^\| 1 \| Interactive' })
        $source4 = @($script:Doc -split '\r?\n' | Where-Object { $_ -match '^\| 4 \| Directory audit' })
        $source1 | Should -Match 'does not list `\$skip` or `\$select`'
        $source1 | Should -Match 'Do not send `\$top` above 1,000'
        $source4 | Should -Match 'does not state a page size'
        $source4 | Should -Match 'one page is not the window'
        $source4 | Should -Match '`activityDateTime` is always UTC'
        $script:Doc | Should -Match 'DirectoryPageTokenNotFoundException'
        $script:Doc | Should -Match 'five requests per 10 seconds per app per tenant'
        $script:Doc | Should -Match 'Wait the `Retry-After` seconds'
        $script:Doc | Should -Match 'starts at three days'
        $script:Doc | Should -Match 'not an empty log'
    }

    It 'keeps directory audit rows the four category examples leave out' {
        # Policy is the Conditional Access category. result includes timeout.
        # Application and App are both documented type strings. Example 2 id is not a GUID.
        $source4 = @($script:Doc -split '\r?\n' | Where-Object { $_ -match '^\| 4 \| Directory audit' })
        $script:Doc | Should -Not -Match 'did not research'
        $source4 | Should -Match 'not a closed set'
        $source4 | Should -Match 'category `Policy`'
        $source4 | Should -Match '`timeout`'
        $source4 | Should -Match 'not limited to'
        $source4 | Should -Match '`Application`'
        $source4 | Should -Match '`ServicePrincipal`'
        $source4 | Should -Match '`N/A`'
        $script:Doc | Should -Match 'Add Conditional Access policy'
        $script:Doc | Should -Match 'do not rewrite `Application` to `App`'
        $script:Doc | Should -Match 'SSGM_b662f17a-4e4d-4e1c-9248-cdec180024b2_MCDC4_88453290'
        $page5 = @($script:Doc -split '\r?\n' | Where-Object { $_ -match '^\| 5 \| Directory changes' })
        $page5 | Should -Match '`timeout` kept apart'
    }
}
