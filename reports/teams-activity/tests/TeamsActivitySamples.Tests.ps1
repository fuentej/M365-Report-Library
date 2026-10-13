#Requires -Version 7.0

BeforeAll {
    $script:Root = Join-Path $PSScriptRoot '../../..' | Resolve-Path | Select-Object -ExpandProperty Path
    $script:Collectors = Join-Path $script:Root 'reports/teams-activity/collectors'
    $script:Samples = Join-Path $script:Root 'reports/teams-activity/samples'

    Import-Module (Join-Path $script:Root 'shared/M365ReportLibrary.psm1') -Force
    $script:Schema = Import-PowerShellDataFile -LiteralPath (Join-Path $script:Collectors 'TeamsActivitySchema.psd1')

    # CSV name -> schema key. One row per source in docs/candidates/teams-activity.md that this report collects.
    # Source 4 is collected here and, from the same API, by reports/teams-groups-lifecycle; source 7 is not built; source 9 is manual.
    $script:Map = [ordered]@{
        'teams-user-activity-user-detail.csv' = 'TeamsUserActivityUserDetail'
        'teams-user-activity-counts.csv'      = 'TeamsUserActivityCounts'
        'teams-device-usage-user-detail.csv'  = 'TeamsDeviceUsageUserDetail'
        'teams-team-activity-detail.csv'      = 'TeamsTeamActivityDetail'
        'report-settings.csv'                 = 'ReportSettings'
        'call-records.csv'                    = 'CallRecords'
        'teams-audit-events.csv'              = 'TeamsAuditEvents'
    }
    $script:Unavailable = 'TeamsUserActivityUserDetail', 'TeamsUserActivityCounts', 'TeamsDeviceUsageUserDetail', 'TeamsTeamActivityDetail'
}

AfterAll {
    Remove-Module M365ReportLibrary -Force -ErrorAction SilentlyContinue
}

Describe 'The schema covers every collector source' {
    It 'has columns and availability for each source' {
        foreach ($key in $script:Map.Values) {
            $script:Schema.ContainsKey($key) | Should -BeTrue -Because $key
            $script:Schema.SourceAvailability.ContainsKey($key) | Should -BeTrue -Because $key
            foreach ($cloud in 'Commercial', 'GCC', 'GCCHigh') {
                $script:Schema.SourceAvailability[$key][$cloud].Status | Should -BeIn 'Available', 'NotAvailable', 'Unverified'
                $script:Schema.SourceAvailability[$key][$cloud].Reference | Should -Match '^https://learn\.microsoft\.com/'
            }
        }
    }

    It 'repeats no column within a CSV' {
        foreach ($key in $script:Map.Values) {
            @($script:Schema[$key] | Select-Object -Unique).Count | Should -Be $script:Schema[$key].Count -Because $key
        }
    }

    It 'matches the availability table in the contract' {
        $a = $script:Schema.SourceAvailability
        # Sources 1 to 3: the Graph usage reports are marked unsupported for US Government L4,
        # and GCC is UNVERIFIED because the API pages and the cloud table disagree.
        foreach ($key in $script:Unavailable) {
            $a[$key].Commercial.Status | Should -Be 'Available'
            $a[$key].GCC.Status | Should -Be 'Unverified'
            $a[$key].GCCHigh.Status | Should -Be 'NotAvailable'
        }
        $a.ReportSettings.Commercial.Status | Should -Be 'Available'
        $a.ReportSettings.GCC.Status | Should -Be 'Unverified'
        $a.ReportSettings.GCCHigh.Status | Should -Be 'Unverified'
        foreach ($cloud in 'Commercial', 'GCC', 'GCCHigh') {
            $a.CallRecords[$cloud].Status | Should -Be 'Available'
            $a.TeamsAuditEvents[$cloud].Status | Should -Be 'Available'
        }
    }

    It 'lists the five Teams audit operations from the contract and marks three UNVERIFIED' {
        $script:Schema.TeamsAuditOperations | Should -Be @('MeetingDetail', 'MeetingParticipantDetail', 'CallParticipantDetail', 'MessageSent', 'ChatCreated')
        $script:Schema.UnverifiedAuditOperations | Should -Be @('CallParticipantDetail', 'MessageSent', 'ChatCreated')
    }
}

Describe 'The sample CSVs match the collectors' {
    It 'has a sample for every collector with exactly the schema columns' {
        foreach ($name in $script:Map.Keys) {
            $path = Join-Path $script:Samples $name
            Test-Path -LiteralPath $path | Should -BeTrue -Because $name
            (Get-CsvHeaderColumn -Path $path) -join ',' | Should -Be ($script:Schema[$script:Map[$name]] -join ',') -Because $name
            @(Import-Csv -LiteralPath $path).Count | Should -BeGreaterThan 0 -Because $name
        }
    }

    It 'has a header-only sample for each source that is not available in GCC High' {
        foreach ($key in $script:Unavailable) {
            $name = ($script:Map.GetEnumerator() | Where-Object Value -EQ $key).Key
            $path = Join-Path $script:Samples "gcchigh/$name"
            Test-Path -LiteralPath $path | Should -BeTrue -Because $name
            (Get-CsvHeaderColumn -Path $path) -join ',' | Should -Be ($script:Schema[$key] -join ',')
            @(Import-Csv -LiteralPath $path).Count | Should -Be 0
        }
    }

    It 'uses no address outside example.com' {
        $text = foreach ($file in Get-ChildItem -LiteralPath $script:Samples -Filter '*.csv' -Recurse) { Get-Content -LiteralPath $file.FullName -Raw }
        $addresses = [regex]::Matches(($text -join "`n"), '[A-Za-z0-9._%+-]+@([A-Za-z0-9.-]+\.[A-Za-z]{2,})') | ForEach-Object { $_.Groups[1].Value.ToLowerInvariant() } | Sort-Object -Unique
        $outside = @($addresses | Where-Object { $_ -notmatch '(^|\.)example\.com$' })
        $outside -join '; ' | Should -BeNullOrEmpty
    }

    It 'covers a group call and a peer-to-peer call in call-records.csv, and a later record version' {
        $rows = Import-Csv -LiteralPath (Join-Path $script:Samples 'call-records.csv')
        ($rows.Type | Sort-Object -Unique) | Should -Be @('groupCall', 'peerToPeer')
        ($rows.Version | Sort-Object -Unique).Count | Should -BeGreaterThan 1
    }

    It 'keeps two rows for one session id when the endpoint or the times differ' {
        $rows = @(Import-Csv -LiteralPath (Join-Path $script:Samples 'call-records.csv') | Where-Object SessionId -EQ '6b2e0001-0000-0000-0000-000000000002')
        $rows.Count | Should -Be 2
        @($rows.CalleeUserId | Sort-Object -Unique).Count | Should -Be 2
    }

    It 'has at least 6 teams in teams-team-activity-detail.csv, one of them with no activity' {
        $rows = @(Import-Csv -LiteralPath (Join-Path $script:Samples 'teams-team-activity-detail.csv'))
        $latest = @($rows | Where-Object RunDate -EQ ($rows.RunDate | Sort-Object | Select-Object -Last 1))
        $latest.Count | Should -BeGreaterOrEqual 6
        $quiet = @($latest | Where-Object { -not $_.LastActivityDate })
        $quiet.Count | Should -BeGreaterOrEqual 1
        foreach ($team in $quiet) {
            foreach ($column in 'ActiveUsers', 'ActiveChannels', 'Reactions', 'MeetingsOrganized', 'PostMessages', 'ReplyMessages', 'ChannelMessages', 'Mentions') {
                $team.$column | Should -Be '0' -Because "$($team.TeamName) $column"
            }
        }
        @($latest | Where-Object { $_.LastActivityDate }).Count | Should -BeGreaterThan 1
    }

    It 'keeps teams-team-activity-detail.csv as the report states it: channel messages are posts plus replies, last activity does not follow the snapshot, and counts need activity inside the period' {
        $rows = @(Import-Csv -LiteralPath (Join-Path $script:Samples 'teams-team-activity-detail.csv'))
        $rows.Count | Should -BeGreaterThan 0
        foreach ($row in $rows) {
            $because = "$($row.TeamName) on $($row.RunDate)"
            [int]$row.ChannelMessages | Should -Be ([int]$row.PostMessages + [int]$row.ReplyMessages) -Because $because
            $refresh = [datetime]::ParseExact($row.ReportRefreshDate, 'yyyy-MM-dd', [cultureinfo]::InvariantCulture)
            if ($row.LastActivityDate) {
                $last = [datetime]::ParseExact($row.LastActivityDate, 'yyyy-MM-dd', [cultureinfo]::InvariantCulture)
                $last | Should -BeLessOrEqual $refresh -Because $because
            }
            $counted = [int]$row.ChannelMessages + [int]$row.MeetingsOrganized + [int]$row.Reactions + [int]$row.Mentions
            if ($counted -gt 0) {
                $row.LastActivityDate | Should -Not -BeNullOrEmpty -Because $because
                $last | Should -BeGreaterThan $refresh.AddDays(-[int]$row.ReportPeriod) -Because $because
            }
        }
    }

    It 'covers every audited operation in teams-audit-events.csv' {
        $ops = (Import-Csv -LiteralPath (Join-Path $script:Samples 'teams-audit-events.csv')).Operation | Sort-Object -Unique
        $ops | Should -Be (@($script:Schema.TeamsAuditOperations) | Sort-Object)
    }
}
