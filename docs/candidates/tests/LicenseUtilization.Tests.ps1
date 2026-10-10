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

    It 'does not cap last activity at the selected report period' {
        # Last activity is the latest intentional use, not the aggregate window.
        $script:Doc | Should -Match 'regardless of the selected time period'
        $script:Doc | Should -Not -Match 'cannot show inactivity older than'
        $script:Doc | Should -Match '24 to 72 hours'
        $script:Doc | Should -Match 'perpetual license'
        $script:Doc | Should -Match 'within 30 days'
    }

    It 'records roles that cannot see usage-report user detail' {
        $source5a = @($script:Doc -split '\r?\n' | Where-Object { $_ -match '^\| 5a \|' })
        $source8 = @($script:Doc -split '\r?\n' | Where-Object { $_ -match '^\| 8 \| Fallback' })
        $source5a | Should -Match 'without visibility into detailed metrics'
        $source8 | Should -Match 'User Experience Success Manager'
    }

    It 'records the conceal-names setting and leaves GCC High report settings unverified' {
        $source6 = @($script:Doc -split '\r?\n' | Where-Object { $_ -match '^\| 6 \| Whether' })
        $source6 | Should -Match 'Conceal user, group, and site names'
        $source6 | Should -Not -Match '\[NotAvailable\]'
        $source6 | Should -Match 'UNVERIFIED'
        $script:Doc | Should -Not -Match 'menu path is not stated'
        $script:Doc | Should -Match 'Org settings'
        $script:Doc | Should -Match 'By default, reports hide'
    }

    It 'pages licensed users and the Microsoft 365 apps JSON report' {
        $source1 = @($script:Doc -split '\r?\n' | Where-Object { $_ -match '^\| 1 \| Licences' })
        $source2 = @($script:Doc -split '\r?\n' | Where-Object { $_ -match '^\| 2 \| Users' })
        $source1 | Should -Match 'only `\$select`'
        $source2 | Should -Match 'follow `@odata.nextLink`'
        $source2 | Should -Match '`\$skip` is not supported'
        $source2 | Should -Match 'capped at 500'
        $source2 | Should -Match 'DirectoryPageTokenNotFoundException'
        $script:Doc | Should -Match 'One JSON page is not the full set'
    }
}
