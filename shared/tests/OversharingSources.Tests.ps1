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

    It 'passes the mandatory site-permissions parameters and the E5 activity cap' {
        # CountOfUsersMoreThan and Name are mandatory on the site-permissions parameter set.
        # E5 without SAM gets no snapshots, and activity reports stop at 10,000 sites.
        $script:Doc | Should -Match 'CountOfUsersMoreThan'
        $script:Doc | Should -Match 'more than 1000 users'
        $script:Doc | Should -Match 'at most 10,000 sites'
        $script:Doc | Should -Match 'no snapshot reports'
        $script:Doc | Should -Match '16\.0\.25409'
        $row3a = ($script:Doc -split '\r?\n' | Where-Object { $_ -match '^\| 3a \|' })
        $row3a | Should -Match '-Name <name> -CountOfUsersMoreThan 0'
    }

    It 'records both data-collection entity spellings and the status that allows a report' {
        # The accepted-values list drops the underscore. The example keeps it.
        $script:Doc | Should -Match 'SharingLinksAnyone'
        $script:Doc | Should -Match 'Which string the cmdlet accepts is UNVERIFIED'
        $script:Doc | Should -Match 'NotInitiated'
        $script:Doc | Should -Match 'InProgress'
        $script:Doc | Should -Match 'Paused'
    }

    It 'includes OneDrive sites and names the sharing setting values' {
        # IncludePersonalSite defaults to false. Limit defaults to 200.
        # None is the widest link, not an absent link.
        $script:Doc | Should -Match 'IncludePersonalSite \$true'
        $script:Doc | Should -Match 'defaults to 200'
        $script:Doc | Should -Match 'defaults to `\$false`'
        $script:Doc | Should -Match 'ExternalUserAndGuestSharing'
        $script:Doc | Should -Match 'ExistingExternalUserSharingOnly'
        $script:Doc | Should -Match 'AnonymousAccess'
        $script:Doc | Should -Match 'not "no link"'
        $script:Doc | Should -Match 'do not call `Set-SPOTenant`'
        $script:Doc | Should -Match 'oneDrive.getAllSites'
    }

    It 'puts a Learn link on every Available or NotAvailable cell in the source table' {
        $sources = $script:Doc -split '## Sources', 2
        $sourceSection = ($sources[1] -split 'Consolidated notes:', 2)[0]
        $rows = @($sourceSection -split '\r?\n' | Where-Object { $_ -match '^\| \d' })
        $rows.Count | Should -BeGreaterThan 0
        foreach ($row in $rows) {
            $cells = $row -split '\|' | Select-Object -Skip 1
            foreach ($cell in $cells) {
                # Match the status token, not the word "available" inside an UNVERIFIED note.
                if ($cell -match '\[(?:Not)?Available\]') {
                    $cell | Should -Match 'https://learn\.microsoft\.com/'
                }
            }
        }
    }
}
