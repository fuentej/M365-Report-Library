#Requires -Version 7.0

BeforeAll {
    $script:DocPath = Join-Path $PSScriptRoot '../unified-audit-log.md'
    $script:Doc = Get-Content -LiteralPath $script:DocPath -Raw
}

Describe 'unified audit log sources' {
    It 'pages Search-UnifiedAuditLog until the window is complete and keeps dates in UTC' {
        # ReturnLargeSet stops at 50,000 and can still have more records.
        # A date with no time is midnight UTC. Without HighCompleteness, results can be missing.
        $source1 = @($script:Doc -split '\r?\n' | Where-Object { $_ -match '^\| 1 \| Unified audit log through Exchange' })
        $source1 | Should -Match 'at most 100 records|Default 100 records'
        $source1 | Should -Match 'maximum 5,000'
        $source1 | Should -Match 'moreRecordsAvailable'
        $source1 | Should -Match 'while `moreRecordsAvailable` is still true'
        $source1 | Should -Match 'HighCompleteness'
        $source1 | Should -Match 'midnight UTC'
        $source1 | Should -Match 'must not be treated as complete'
        $source1 | Should -Match 'results can be missing'
    }

    It 'treats the Management Activity API as a seven-day feed and lists every content type' {
        # Listing content older than 7 days fails. Exchange and General are separate content types.
        # startTime selects when the blob was published, not when the event happened.
        $source3 = @($script:Doc -split '\r?\n' | Where-Object { $_ -match '^\| 3 \| Unified audit log through the Office 365' })
        $source3 | Should -Match 'Audit\.Exchange'
        $source3 | Should -Match 'Audit\.General'
        $source3 | Should -Match 'DLP\.All'
        $source3 | Should -Match 'Read DLP sensitive data'
        $source3 | Should -Match 'PublisherIdentifier'
        $source3 | Should -Match 'contentCreated'
        $source3 | Should -Match 'no more than 7 days'
        $source3 | Should -Match 'AF20051'
        $source3 | Should -Match 'not the 180-day or one-year search'
        $script:Doc | Should -Not -Match 'How far back content stays listable was not found'
        $script:Doc | Should -Not -Match 'Content types seen on the pages read'
    }

    It 'keeps one-year retention to E5 and the audit add-on, and guests at 180 days' {
        # The one-year default does not cover non-E5 users or guests.
        # Get-UnifiedAuditLogRetentionPolicy omits the default policy.
        $script:Doc | Should -Match 'Microsoft Purview Suite'
        $script:Doc | Should -Match 'E5 eDiscovery and Audit add-on'
        $script:Doc | Should -Match 'guest users stay at 180 days'
        $script:Doc | Should -Match 'custom retention policy overrides the default and can be shorter'
        $script:Doc | Should -Match 'does not return the default'
        $script:Doc | Should -Match '7 Days, 30 Days, 3 Years, 5 Years, and 7 Years'
        $script:Doc | Should -Not -Match 'Entra ID, Exchange, OneDrive and SharePoint records are kept one year by default'
    }

    It 'treats MailItemsAccessed as Audit Standard for E3 and E5 and names Logon_type' {
        # The event is not E5-only. SensitivityLabel on that record is Premium.
        # Folder permission changes that are audited are UpdateFolderPermissions.
        $script:Doc | Should -Match 'Office 365 E3/E5 or Microsoft 365 E3/E5'
        $script:Doc | Should -Match 'Audit \(Standard\)'
        $script:Doc | Should -Match 'SensitivityLabel'
        $script:Doc | Should -Match 'Logon_type'
        $script:Doc | Should -Match 'MailboxUPN'
        $script:Doc | Should -Match 'UpdateFolderPermissions'
        $script:Doc | Should -Match 'not audited separately'
        $script:Doc | Should -Match 'Remove-MailboxPermission'
        $script:Doc | Should -Not -Match 'MailItemsAccessed needs E5'
        $script:Doc | Should -Not -Match 'shared mailbox page says E5'
        $script:Doc | Should -Not -Match 'field names were not confirmed'
    }
}
