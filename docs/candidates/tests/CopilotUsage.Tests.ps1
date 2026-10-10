#Requires -Version 7.0

BeforeAll {
    $script:DocPath = Join-Path $PSScriptRoot '../copilot-usage.md'
    $script:Doc = Get-Content -LiteralPath $script:DocPath -Raw
}

Describe 'copilot usage sources' {
    It 'uses the usage CSV headers and keeps both lag figures' {
        # The v1 header is Microsoft Teams Copilot Last Activity Date, not a paraphrased Teams date.
        # The usage page says 48 hours and the reports overview says 72, and that overview's chat periods use 30 days.
        $source2 = @($script:Doc -split '\r?\n' | Where-Object { $_ -match '^\| 2 \| Copilot usage' })
        $source3 = @($script:Doc -split '\r?\n' | Where-Object { $_ -match '^\| 3 \| Copilot enabled' })
        $source4 = @($script:Doc -split '\r?\n' | Where-Object { $_ -match '^\| 4 \| Daily trend' })
        $source5 = @($script:Doc -split '\r?\n' | Where-Object { $_ -match '^\| 5 \| Fallback' })
        $source2 | Should -Match 'Microsoft Teams Copilot Last Activity Date'
        $source2 | Should -Match 'Prompts submitted for all apps'
        $source2 | Should -Match 'Active Usage Days for all apps'
        $source2 | Should -Match 'within 48 hours'
        $source2 | Should -Match 'within 72 hours'
        $source2 | Should -Match '@odata.nextLink'
        $source3 | Should -Match 'Microsoft Teams Enabled Users'
        $source3 | Should -Match 'application/octet-stream'
        $source3 | Should -Match 'over the selected timeframe'
        $source4 | Should -Match 'Prompts submitted'
        $source4 | Should -Not -Match 'and prompts submitted'
        $source5 | Should -Match '7, 30, 90 and 180'
        $source5 | Should -Match 'Usage Summary Reports Reader'
        $source5 | Should -Match 'User Experience Success Manager'
        $source5 | Should -Match 'N/A\^1'
        $source5 | Should -Match 'N/A\^2'
        $script:Doc | Should -Match 'Edit with Copilot in Word, Excel, PowerPoint and OneNote'
        $script:Doc | Should -Match 'Edge sidebar'
        $script:Doc | Should -Match 'Themes by Copilot'
        $script:Doc | Should -Match '11 December 2025'
        $script:Doc | Should -Match 'Edit with Excel and Edit with PowerPoint do not'
    }
}
