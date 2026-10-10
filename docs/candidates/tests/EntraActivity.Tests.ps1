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
        $source3 = @($script:Doc -split '\r?\n' | Where-Object { $_ -match '^\| 3 \|' })
        $source3 | Should -Match 'report-only four'
    }
}
