#Requires -Version 7.0

BeforeAll {
    $script:Doc = Get-Content -LiteralPath (Join-Path $PSScriptRoot '../../../docs/candidates/copilot-studio-agents.md') -Raw
    $script:Schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot '../collectors/CopilotStudioSchema.psd1')
}

Describe 'The collectors follow the contract doc' {
    It 'has a collector for each of the six source rows' {
        $script:Schema.SourceAvailability.Keys.Count | Should -Be 6
        foreach ($script in 'Get-PowerPlatformEnvironments', 'Get-CopilotStudioAgents', 'Get-AgentConnectors', 'Get-AgentComponents', 'Get-AgentModifications', 'Get-AgentAuditEvents') {
            Test-Path -LiteralPath (Join-Path $PSScriptRoot "../collectors/$script.ps1") | Should -BeTrue
        }
    }

    It 'copies the authoring audit labels listed in the doc' {
        foreach ($label in 'BotCreate', 'BotDelete', 'BotUpdateOperation-BotPublish', 'BotUpdateOperation-BotShare') {
            $script:Doc | Should -Match ([regex]::Escape($label))
            $script:Schema.AuditOperations | Should -Contain $label
        }
    }

    It 'reads connector availability from the doc: NotAvailable in GCC and GCC High only' {
        $source3 = @($script:Doc -split "`n" | Where-Object { $_ -match '^\| 3 \|' })[0]
        $cells = $source3 -split '\|'
        $cells[9].Trim() | Should -Match '^\[NotAvailable\]'
        $cells[10].Trim() | Should -Match '^\[NotAvailable\]'
        $script:Schema.SourceAvailability.AgentConnectors.GCC.Status | Should -Be 'NotAvailable'
        $script:Schema.SourceAvailability.AgentConnectors.GCCHigh.Status | Should -Be 'NotAvailable'
        $script:Schema.SourceAvailability.AgentConnectors.Commercial.Status | Should -Be 'Available'
    }
}
