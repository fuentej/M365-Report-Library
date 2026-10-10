#Requires -Version 7.0

BeforeAll {
    $script:DocPath = Join-Path $PSScriptRoot '../license-utilization.md'
    $script:Doc = Get-Content -LiteralPath $script:DocPath -Raw
}

Describe 'license utilization sources' {
    It 'calls Copilot usage on default v2 and reads the CSV body' {
        # v2 is the default and the supported version. D30 is not a v2 period.
        # v1.0 returns the CSV in a 200 body. Global Reader is not a /copilot role.
        $script:Doc | Should -Match 'defaults to `v2`'
        $script:Doc | Should -Not -Match "getMicrosoft365CopilotUsageUserDetail\(period='D30'\)"
        $script:Doc | Should -Not -Match 'Global Reader \(beta page\)'
        $script:Doc | Should -Match '`200 OK`'
        $script:Doc | Should -Match 'not a redirect'
        $script:Doc | Should -Match 'Edge Last Activity Date'
        $script:Doc | Should -Match '`ALL` is the four periods'
    }
}
