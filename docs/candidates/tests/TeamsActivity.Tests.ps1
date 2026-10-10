#Requires -Version 7.0

BeforeAll {
    $script:DocPath = Join-Path $PSScriptRoot '../teams-activity.md'
    $script:Doc = Get-Content -LiteralPath $script:DocPath -Raw
}

Describe 'Teams activity sources' {
    It 'pages call records and does not list them with the get-by-id cmdlet' {
        # Get-MgCommunicationCallRecord requires CallRecordId. The list defaults to 60
        # rows and type includes unknown. organizer stopped returning data on 2026-06-30.
        $source6 = @($script:Doc -split '\r?\n' | Where-Object { $_ -match '^\| 6 \| Calls' })
        $source6 | Should -Match 'requires `-CallRecordId`'
        $source6 | Should -Match 'follow `@odata.nextLink`'
        $source6 | Should -Match 'One page is not the full set'
        $source6 | Should -Match '`unknown`, `groupCall`, `peerToPeer`'
        $source6 | Should -Match '`videoBasedScreenSharing`'
        $source6 | Should -Match '`screenSharing`'
        $source6 | Should -Match '2026-06-30'
        $source6 | Should -Match 'organizer_v2'
        $source6 | Should -Match 'participants_v2'
        $source6 | Should -Match '130'
    }

    It 'follows export nextLink and records the evaluation-mode conflict' {
        # $top is a hint. model is ignored, and the export page still describes a cap.
        $source7 = @($script:Doc -split '\r?\n' | Where-Object { $_ -match '^\| 7 \| Chat' })
        $source7 | Should -Match 'not a guaranteed page size'
        $source7 | Should -Match 'at most 250'
        $source7 | Should -Match 'lastModifiedDateTime lt'
        $source7 | Should -Match 'model` is no longer required'
        $source7 | Should -Match 'evaluation mode'
        $source7 | Should -Match 'complete export'
        $source7 | Should -Match 'Purview DLP'
        $source7 | Should -Match 'follow `@odata.nextLink`'
    }

    It 'pages the audit search and keeps Teams retention at 180 days' {
        # A bare call returns 100 records. Teams is not in the one-year workload set.
        # MessageSent is not a complete chat count. The compliance cmdlet always reports auditing off.
        $source8 = @($script:Doc -split '\r?\n' | Where-Object { $_ -match '^\| 8 \| Teams events' })
        $source8 | Should -Match 'at most 100 records'
        $source8 | Should -Match 'maximum is 5,000'
        $source8 | Should -Match 'moreRecordsAvailable'
        $source8 | Should -Match 'not the full window'
        $source8 | Should -Match 'HighCompleteness'
        $source8 | Should -Match 'midnight UTC'
        $source8 | Should -Match 'always `False` in Security & Compliance PowerShell'
        $source8 | Should -Match 'Exchange admin center'
        $source8 | Should -Match 'Microsoft Purview Suite'
        $source8 | Should -Match 'E5 eDiscovery and Audit add-on'
        $source8 | Should -Match 'guest users stay at 180 days'
        $source8 | Should -Match 'can be shorter'
        $source8 | Should -Match 'not a complete chat count'
        $source8 | Should -Match 'footnote 12'
        $source8 | Should -Match 'manage-gcc.office.com'
        $source8 | Should -Match 'manage.office365.us'
        $source8 | Should -Match 'ActivityFeed.Read'
        $source8 | Should -Match 'Audit.General'
        $source8 | Should -Match 'at most 24 hours apart'
        $source8 | Should -Match 'at most 7 days ago'
        $source8 | Should -Match 'does not replace a 180-day'
        $source8 | Should -Not -Match 'was not checked'
    }

    It 'does not add overlapping Teams usage columns' {
        # Team Chat Message Count already includes posts and replies. Call Count is 1:1.
        # The team-activity header list spells Active users; the schema spells Active Users.
        $source1 = @($script:Doc -split '\r?\n' | Where-Object { $_ -match '^\| 1 \| Teams activity per user' })
        $source2 = @($script:Doc -split '\r?\n' | Where-Object { $_ -match '^\| 2 \| Teams activity per day' })
        $source4 = @($script:Doc -split '\r?\n' | Where-Object { $_ -match '^\| 4 \| Teams activity per team' })
        $source1 | Should -Match 'includes original posts and replies'
        $source1 | Should -Match '1:1 calls'
        $source1 | Should -Match 'being phased out'
        $source1 | Should -Match 'UTC date'
        $source1 | Should -Match '`Report Period`'
        $source2 | Should -Match '`Team Chat Messages`'
        $source2 | Should -Match '`Private Chat Messages`'
        $source4 | Should -CMatch '`Active users`'
        $source4 | Should -CMatch '`Active Users`'
        $source4 | Should -Match '00:00 through 23:59 UTC'
        $source4 | Should -Match 'do not add `Guests`'
        $source4 | Should -Match 'unique messages posted in a team chat'
        $source4 | Should -Match 'Cross-posted messages count only'
    }

    It 'puts a Learn link on every Available or NotAvailable cell in the source table' {
        $sources = $script:Doc -split '## Sources', 2
        $sourceSection = ($sources[1] -split 'Consolidated notes:', 2)[0]
        $rows = @($sourceSection -split '\r?\n' | Where-Object { $_ -match '^\| \d' })
        $rows.Count | Should -Be 9
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
