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

    It 'reads Copilot audit fields from CopilotEventData and the record-type enum' {
        # Record type 261 is CopilotInteraction. The schema nests the app host under CopilotEventData.
        # The Graph enum does not include copilotInteraction.
        $source6 = @($script:Doc -split '\r?\n' | Where-Object { $_ -match '^\| 6 \| Copilot interactions in the unified' })
        $source7 = @($script:Doc -split '\r?\n' | Where-Object { $_ -match '^\| 7 \| Copilot interactions through the Graph' })
        $source6 | Should -Match 'record type 261'
        $source6 | Should -Match 'CopilotEventData'
        $source6 | Should -Match 'ConnectedAIAppInteraction` as 328'
        $source6 | Should -Match 'TeamCopilotInteraction` as 334'
        $source6 | Should -Match 'Audit.General'
        $source6 | Should -Match 'does not return `TeamCopilotInteraction`'
        $source6 | Should -Not -Match 'is not stated on the pages read'
        $source7 | Should -Match 'does not include `copilotInteraction`'
        $source7 | Should -Match 'operationFilters'
        $source7 | Should -Match 'copilotSessionSharing'
        $source7 | Should -Match 'AuditLogsQuery-Entra.Read.All'
        $script:Doc | Should -Not -Match 'record type filter includes Copilot. Neither is stated'
    }

    It 'bounds the interaction export and does not treat a 429 as empty' {
        # createdDateTime needs both bounds. $top 100 is not the whole history. A 429 is throttling.
        $source8 = @($script:Doc -split '\r?\n' | Where-Object { $_ -match '^\| 8 \| Copilot prompts' })
        $source8 | Should -Match 'both a minimum and a maximum'
        $source8 | Should -Match 'not the full history'
        $source8 | Should -Match 'userPrompt'
        $source8 | Should -Match 'deleted users'
        $source8 | Should -Match '1,500 requests per second per app and 30 per app per tenant'
        $source8 | Should -Match 'a `429` is not an empty history'
        $source8 | Should -Match 'Outlook, PowerPoint, OneNote and Loop are not in this export'
        $source8 | Should -Not -Match 'Recommended `\$top` is 100\. Returns prompt'
    }
}
