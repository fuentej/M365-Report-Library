#Requires -Version 7.0

BeforeAll {
    $script:DocPath = Join-Path $PSScriptRoot '../../docs/candidates/oversharing.md'
    $script:Doc = Get-Content -LiteralPath $script:DocPath -Raw
}

Describe 'Oversharing source list' {
    It 'pages Search-UnifiedAuditLog until the window is complete and keeps dates in UTC' {
        # A bare call returns 100 records. ReturnLargeSet stops at 50,000 and can still
        # have moreRecordsAvailable. A date with no time is midnight UTC.
        $script:Doc | Should -Match 'at most 100 records'
        $script:Doc | Should -Match 'maximum is 5,000'
        $script:Doc | Should -Match 'moreRecordsAvailable'
        $script:Doc | Should -Match 'ReturnLargeSet'
        $script:Doc | Should -Match 'not the full window'
        $script:Doc | Should -Match 'HighCompleteness'
        $script:Doc | Should -Match 'midnight UTC'
        $script:Doc | Should -Match 'must not be treated as complete'
    }

    It 'does not treat a Security and Compliance false as auditing off' {
        # UnifiedAuditLogIngestionEnabled is always False in Security & Compliance PowerShell.
        $script:Doc | Should -Match 'always `False` in Security & Compliance PowerShell'
        $script:Doc | Should -Match 'Business Basic, Business Standard, and Business Premium'
    }

    It 'keeps one-year retention to E5 and the audit add-on, and guests at 180 days' {
        $script:Doc | Should -Match 'Microsoft Purview Suite'
        $script:Doc | Should -Match 'E5 eDiscovery and Audit add-on'
        $script:Doc | Should -Match 'guest users stay at 180 days'
        $script:Doc | Should -Match 'custom retention policy overrides the default and can be shorter'
        $script:Doc | Should -Not -Match 'One year for Exchange, SharePoint, OneDrive and Entra records of E5 users'
    }

    It 'includes withdrawn, blocked, and updated sharing invitations' {
        # Created and accepted invitations are not the whole invitation set.
        $script:Doc | Should -Match 'SharingInvitationBlocked'
        $script:Doc | Should -Match 'SharingInvitationUpdated'
        $script:Doc | Should -Match 'SharingInvitationRevoked'
        $script:Doc | Should -Match 'SharingInheritanceReset'
        $script:Doc | Should -Match 'specific-people link'
    }
}
