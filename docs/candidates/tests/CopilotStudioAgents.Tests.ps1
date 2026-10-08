#Requires -Version 7.0

BeforeAll {
    $script:DocPath = Join-Path $PSScriptRoot '../copilot-studio-agents.md'
    $script:Doc = Get-Content -LiteralPath $script:DocPath -Raw
}

Describe 'Copilot Studio agents source list' {
    It 'keeps never-published draft agents in the inventory' {
        $script:Doc | Should -Not -Match 'unpublished drafts are not reflected'
        $script:Doc | Should -Match 'includes unpublished draft agents and published agents'
        $script:Doc | Should -Match 'lastPublishedAt'
        $script:Doc | Should -Match 'newer unpublished changes are omitted until that agent is published again'
    }

    It 'records that the 200-resource cap is a random subset' {
        $script:Doc | Should -Not -Match 'lists at most 200 resources'
        $script:Doc | Should -Match 'returns a random 200 of that type'
        $script:Doc | Should -Match 'capabilitiesCounts'
        $script:Doc | Should -Match 'complete count for each type'
    }

    It 'does not mark bot and botcomponent Available in GCC or GCC High' {
        $source4 = ($script:Doc -split "`n" | Where-Object { $_ -match '^\| 4 \|' })
        $source4 | Should -Not -BeNullOrEmpty
        $cells = $source4 -split '\|'
        # Columns: empty, #, Source, Endpoint, Role, License, Event, Retention, Commercial, GCC, GCC High, empty
        $cells[9].Trim() | Should -Match '^UNVERIFIED'
        $cells[10].Trim() | Should -Match '^UNVERIFIED'
        $cells[9] | Should -Not -Match '\[Available\]'
        $cells[10] | Should -Not -Match '\[Available\]'
        $script:Doc | Should -Not -Match 'for the Dataverse API\. That the `bot` tables'
    }

    It 'records the Search-UnifiedAuditLog page cap and the usage record shape' {
        $script:Doc | Should -Match 'returns at most 100 records'
        $script:Doc | Should -Match 'SessionCommand ReturnLargeSet'
        $script:Doc | Should -Match 'session cap 50,000'
        $script:Doc | Should -Match 'ResultSize` maximum 5,000'
        $script:Doc | Should -Match 'StartDate` and `EndDate` are UTC'
        $script:Doc | Should -Match 'RecordType CopilotInteraction'
        $script:Doc | Should -Match 'Copilot\.Studio\.'
        $script:Doc | Should -Match 'does not return the authoring operations'
        $script:Doc | Should -Match 'BotCreate'
    }

    It 'names Audit Reader as the least privileged audit role group' {
        $script:Doc | Should -Not -Match 'Purview Audit Reader'
        $script:Doc | Should -Match 'Audit Reader role group, which grants View-Only Audit Logs'
        $script:Doc | Should -Match 'Exchange admin center View-Only Audit Logs or Audit Logs role'
    }

    It 'does not treat Audit Premium as a one-year window for Copilot Studio' {
        $script:Doc | Should -Not -Match 'longer retention needs Audit \(Premium\)'
        $script:Doc | Should -Match 'default one-year policy covers only Exchange, SharePoint, OneDrive and Microsoft Entra'
        $script:Doc | Should -Match 'stay at 180 days unless a custom retention policy applies'
        $script:Doc | Should -Match 'CreationTime` on the Copilot Studio schema is UTC'
    }

    It 'leaves pay-as-you-go billing unverified when Learn disagrees' {
        $script:Doc | Should -Not -Match 'prerequisites\)\); Audit \(Standard\)'
        $script:Doc | Should -Match 'pay-as-you-go does not apply to them'
        $script:Doc | Should -Match 'Whether pay-as-you-go is required: UNVERIFIED'
    }

    It 'pages the inventory query until resultTruncated clears' {
        $script:Doc | Should -Match 'TableName` `PowerPlatformResources'
        $script:Doc | Should -Match 'Options\.SkipToken'
        $script:Doc | Should -Match 'resultTruncated'
        $script:Doc | Should -Match 'One response is not the full set'
        $script:Doc | Should -Match 'including `SkipToken` paging'
    }

    It 'names AI Reader as the least privileged inventory role' {
        $script:Doc | Should -Not -Match 'environment groups only'
        $script:Doc | Should -Match 'AI Reader, the least privileged role for this report'
        $script:Doc | Should -Match 'agentic apps, agent flows, environments and environment groups'
    }

    It 'lists every connector usedAs value, including knowledge connectors' {
        $script:Doc | Should -Not -Match 'exposes only the web search flag'
        $script:Doc | Should -Match 'usedAs` is `Tool`, `Topic Tool` or `Knowledge'
        $script:Doc | Should -Match 'usedAs` is `Knowledge'
    }
}
