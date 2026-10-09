#Requires -Version 7.0

BeforeAll {
    $script:Doc = Get-Content -LiteralPath (Join-Path $PSScriptRoot '../../../docs/candidates/mailbox-exfiltration-risk.md') -Raw
    $script:Schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot '../collectors/MailboxExfiltrationSchema.psd1')
    $script:Collectors = Join-Path $PSScriptRoot '../collectors'
}

Describe 'The collectors follow the contract doc' {
    It 'has a collector for every source row other than the shared users source' {
        foreach ($script in 'Get-AcceptedDomains', 'Get-MailboxForwarding', 'Get-SendOnBehalf', 'Get-InboxRules', 'Get-TransportRules',
            'Get-MailboxFullAccess', 'Get-SendAsPermissions', 'Get-DelegatedConsents', 'Get-AppRoleAssignments',
            'Get-MailboxChangeEvents', 'Get-MailAccessEvents', 'Get-AuditConfiguration') {
            Test-Path -LiteralPath (Join-Path $script:Collectors "$script.ps1") | Should -BeTrue -Because $script
        }
        # Source rows 1, 1b, 2, 3, 4a, 4b, 4c, 5a, 5b, 6a, 6b, 6c.
        $sources = ($script:Doc -split '## Sources')[1] -split '## Proposed report pages' | Select-Object -First 1
        @($sources -split "`n" | Where-Object { $_ -match '^\| (1|1b|2|3|4a|4b|4c|5a|5b|6a|6b|6c) \|' }).Count | Should -Be 12
        $script:Schema.SourceAvailability.Keys.Count | Should -Be 12
        Test-Path -LiteralPath (Join-Path $script:Collectors 'Get-Users.ps1') | Should -BeFalse
    }

    It 'copies the audit operation names the doc lists' {
        foreach ($operation in $script:Schema.MailboxChangeOperations + $script:Schema.MailAccessOperations) {
            $script:Doc | Should -Match ([regex]::Escape($operation))
        }
    }

    It 'marks Get-AcceptedDomain and the audit configuration checks UNVERIFIED in GCC and GCC High, as the doc does' {
        foreach ($source in 'AcceptedDomains', 'AuditConfiguration') {
            $script:Schema.SourceAvailability[$source].GCC.Status | Should -Be 'Unverified'
            $script:Schema.SourceAvailability[$source].GCCHigh.Status | Should -Be 'Unverified'
            $script:Schema.SourceAvailability[$source].Commercial.Status | Should -Be 'Available'
        }
        $script:Doc | Should -Match 'UNVERIFIED \(no page found that states GCC availability of `Get-OrganizationConfig` or `Get-AdminAuditLogConfig`'
    }
}
