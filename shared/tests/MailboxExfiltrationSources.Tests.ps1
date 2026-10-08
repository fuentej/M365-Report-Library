#Requires -Version 7.0

BeforeAll {
    $script:DocPath = Join-Path $PSScriptRoot '../../docs/candidates/mailbox-exfiltration-risk.md'
    $script:Doc = Get-Content -LiteralPath $script:DocPath -Raw
}

Describe 'Mailbox exfiltration source list' {
    It 'pages Search-UnifiedAuditLog past the 100-record default and stops at 50000' {
        # A bare Search-UnifiedAuditLog call returns at most 100 records.
        # ReturnLargeSet is unsorted and caps the session at 50,000.
        $script:Doc | Should -Match 'at most 100 records'
        $script:Doc | Should -Match 'ReturnLargeSet'
        $script:Doc | Should -Match '50,000'
        $script:Doc | Should -Match 'moreRecordsAvailable'
        $script:Doc | Should -Match 'HighCompleteness'
        $script:Doc | Should -Match 'must not be treated as the full window'
    }

    It 'names the mailbox audit operations for rules, forwarding, send, and permissions' {
        # UpdateInboxRules is the Outlook client operation. Send is not MailItemsAccessed.
        $script:Doc | Should -Match 'UpdateInboxRules'
        $script:Doc | Should -Match 'Add-MailboxPermission'
        $script:Doc | Should -Match 'Remove-MailboxPermission'
        $script:Doc | Should -Match 'operation `Send`'
        $script:Doc | Should -Match 'SendOnBehalf'
        $script:Doc | Should -Match 'RecordType ExchangeAdmin'
        $script:Doc | Should -Not -Match 'must hold the Audit \(Premium\) license for the record to be generated'
    }

    It 'drops self and inherited Full Access rows that are not delegates' {
        # FullAccess on NT AUTHORITY\SELF and inherited Deny entries are not delegates.
        $script:Doc | Should -Match 'NT AUTHORITY\\SELF'
        $script:Doc | Should -Match 'IsInherited -eq \$false'
        $script:Doc | Should -Match 'Deny -eq \$false'
        $script:Doc | Should -Match 'Identity` is mandatory'
    }

    It 'requests the Delivery property set and lifts the 1000-row Exchange cap' {
        # Get-EXOMailbox Minimum omits forwarding properties. ResultSize defaults to 1000.
        $script:Doc | Should -Match 'PropertySets Delivery'
        $script:Doc | Should -Match 'Get-EXOMailbox -ResultSize Unlimited'
        $script:Doc | Should -Match 'Get-InboxRule -Mailbox <id> -IncludeHidden -ResultSize Unlimited'
        $script:Doc | Should -Match 'Get-TransportRule -ResultSize Unlimited'
        $script:Doc | Should -Match 'Get-RecipientPermission -ResultSize Unlimited'
        $script:Doc | Should -Match 'ExcludeConditionActionDetails'
    }

    It 'requires the 10-year audit add-on and does not treat AuditEnabled as a per-mailbox switch' {
        $script:Doc | Should -Match '10-year audit log retention add-on'
        $script:Doc | Should -Match 'Audit \(Premium\) alone does not retain for 10 years'
        $script:Doc | Should -Match 'UnifiedAuditLogIngestionEnabled'
        $script:Doc | Should -Match 'DefaultAuditSet'
        $script:Doc | Should -Match 'always displays `AuditEnabled` as `True`'
        $script:Doc | Should -Not -Match '10 years need Audit \(Premium\)'
    }
}
