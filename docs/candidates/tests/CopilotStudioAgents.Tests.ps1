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
}
