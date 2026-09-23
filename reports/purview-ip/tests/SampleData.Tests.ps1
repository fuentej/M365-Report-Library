#Requires -Version 7.0

<#
    The committed sample set is what a report built on this data is worked against, so
    its shape matters as much as the collectors'. These tests hold it to the shape a
    Purview information protection review needs: a full policy configuration, activity
    across every category, a growing content snapshot, and Copilot resources that
    include some blocked by policy.
#>

BeforeAll {
    $script:Root = Join-Path $PSScriptRoot '../../..' | Resolve-Path | Select-Object -ExpandProperty Path
    $script:Samples = Join-Path $script:Root 'reports/purview-ip/samples'

    Import-Module (Join-Path $script:Root 'shared/M365ReportLibrary.psm1') -Force

    $script:Schema = Import-PowerShellDataFile -LiteralPath (Join-Path $script:Root 'reports/purview-ip/collectors/PurviewIpSchema.psd1')

    $script:Users = @(Import-Csv -LiteralPath (Join-Path $script:Samples 'users.csv'))
    $script:Policies = @(Import-Csv -LiteralPath (Join-Path $script:Samples 'policies.csv'))
    $script:Events = @(Import-Csv -LiteralPath (Join-Path $script:Samples 'activity-explorer-events.csv'))
    $script:Content = @(Import-Csv -LiteralPath (Join-Path $script:Samples 'content-explorer-snapshot.csv'))
    $script:Copilot = @(Import-Csv -LiteralPath (Join-Path $script:Samples 'copilot-accessed-resources.csv'))

    $script:LatestUserRunDate = ($script:Users.RunDate | Sort-Object -Descending | Select-Object -First 1)
    $script:LatestUsers = @($script:Users | Where-Object RunDate -EQ $script:LatestUserRunDate)
    $script:LatestPolicyRunDate = ($script:Policies.RunDate | Sort-Object -Descending | Select-Object -First 1)
    $script:LatestPolicies = @($script:Policies | Where-Object RunDate -EQ $script:LatestPolicyRunDate)
}

AfterAll {
    Remove-Module M365ReportLibrary -Force -ErrorAction SilentlyContinue
}

Describe 'The sample data is fake' {
    It 'uses no address outside example.com' {
        $addresses = @(
            $script:Users.UserPrincipalName
            $script:Users.Mail
            $script:Events.User
            $script:Copilot.UserKey
        ) | Where-Object { $_ -and $_.Contains('@') }

        $outside = @($addresses | Where-Object { $_ -notmatch '@([a-z0-9-]+\.)*example\.com$' } | Sort-Object -Unique)
        $outside -join '; ' | Should -BeNullOrEmpty
    }

    It 'uses only example.com SharePoint site URLs' {
        $urls = @($script:Events.FullUrl; $script:Copilot.SiteUrl) | Where-Object { $_ }
        $outside = @($urls | Where-Object { $_ -notmatch '^https://example\.sharepoint\.com/' } | Sort-Object -Unique)
        $outside -join '; ' | Should -BeNullOrEmpty
    }
}

Describe 'The directory is deep and wide enough to report on' {
    It 'holds 60 members' {
        $script:LatestUsers.Count | Should -Be 60
    }

    It 'spreads them over at least five departments' {
        ($script:LatestUsers.Department | Sort-Object -Unique).Count | Should -BeGreaterOrEqual 5
    }

    It 'has a management chain three levels deep' {
        $byId = @{}
        foreach ($member in $script:LatestUsers) { $byId[$member.Id] = $member }

        $depths = foreach ($member in $script:LatestUsers) {
            $depth = 0
            $current = $member
            while ($current -and -not [string]::IsNullOrWhiteSpace($current.ManagerId) -and $depth -lt 20) {
                $depth++
                $current = $byId[$current.ManagerId]
            }
            $depth
        }

        ($depths | Measure-Object -Maximum).Maximum | Should -BeGreaterOrEqual 3
    }
}

Describe 'The policy configuration covers every object type' {
    It 'uses every ObjectType PurviewIpSchema.psd1 lists' {
        $found = @($script:LatestPolicies.ObjectType | Sort-Object -Unique)
        foreach ($objectType in $script:Schema.PolicyObjectTypes) {
            $found | Should -Contain $objectType
        }
    }

    It 'has 26 rows in the latest snapshot' {
        $script:LatestPolicies.Count | Should -Be 26
    }

    It 'gives at least one sensitivity label a parent' {
        $sublabels = @($script:LatestPolicies | Where-Object { $_.ObjectType -eq 'SensitivityLabel' -and $_.ParentId })
        $sublabels.Count | Should -BeGreaterOrEqual 1
        foreach ($sublabel in $sublabels) { $sublabel.ParentName | Should -Not -BeNullOrEmpty }
    }

    It 'marks at least one DLP policy as applying to Copilot' {
        @($script:LatestPolicies | Where-Object { $_.ObjectType -eq 'DlpPolicy' -and $_.AppliesToCopilot -eq 'True' }).Count |
            Should -BeGreaterOrEqual 1
    }

    It 'is keyed on RunDate plus ObjectType and ObjectId: re-running does not duplicate a row' {
        $keys = @($script:Policies | ForEach-Object { '{0}|{1}|{2}' -f $_.RunDate, $_.ObjectType, $_.ObjectId })
        ($keys | Sort-Object -Unique).Count | Should -Be $keys.Count
    }
}

Describe 'Activity Explorer events cover every category' {
    It 'uses only activity values PurviewIpSchema.psd1 recognises, or Other' {
        $known = @($script:Schema.ActivityCategories.Values | ForEach-Object { $_ })
        $unrecognised = @($script:Events | Where-Object { $known -notcontains $_.Activity -and $_.ActivityCategory -ne 'Other' })
        $unrecognised.Count | Should -Be 0
    }

    It 'includes a label downgrade' {
        @($script:Events | Where-Object IsLabelDowngrade -EQ 'True').Count | Should -BeGreaterThan 0
    }

    It 'includes a DLP match with a justification' {
        @($script:Events | Where-Object { $_.ActivityCategory -eq 'Dlp' -and $_.Justification }).Count | Should -BeGreaterThan 0
    }

    It 'covers at least six of the seven activity categories' {
        ($script:Events.ActivityCategory | Sort-Object -Unique).Count | Should -BeGreaterOrEqual 6
    }

    It 'has no duplicate RecordIdentity' {
        $ids = @($script:Events.RecordIdentity)
        ($ids | Sort-Object -Unique).Count | Should -Be $ids.Count
    }

    It 'dates no event after the moment the set is generated around' {
        # The generator's -EndDate is its "now". An event after it has not happened yet.
        $endOfWindow = [datetime]'2026-09-01T00:00:01Z'
        $strays = @($script:Events | Where-Object { ([datetime]$_.Happened) -ge $endOfWindow } | ForEach-Object { $_.RecordIdentity })
        $strays -join '; ' | Should -BeNullOrEmpty
    }
}

Describe 'The Content Explorer snapshot shows a trend' {
    It 'takes at least two snapshots' {
        ($script:Content.RunDate | Sort-Object -Unique).Count | Should -BeGreaterOrEqual 2
    }

    It 'grows a tracked count across snapshots' {
        $series = @($script:Content | Where-Object { $_.TagType -eq 'Sensitivity' -and $_.Workload -eq 'EXO' } | Sort-Object RunDate)
        $series.Count | Should -BeGreaterOrEqual 2
        [long]$series[-1].TotalCount | Should -BeGreaterThan ([long]$series[0].TotalCount)
    }

    It 'covers every workload the collector snapshots' {
        ($script:Content.Workload | Sort-Object -Unique) | Should -Be @('EXO', 'ODB', 'SPO', 'Teams')
    }
}

Describe 'Copilot accessed resources include a mix of allowed and blocked access' {
    It 'includes at least one blocked resource with a policy named' {
        $blocked = @($script:Copilot | Where-Object AccessBlocked -EQ 'True')
        $blocked.Count | Should -BeGreaterThan 0
        foreach ($row in $blocked) { $row.PolicyName | Should -Not -BeNullOrEmpty }
    }

    It 'includes at least one interaction that touched no resource' {
        @($script:Copilot | Where-Object { -not $_.ResourceId }).Count | Should -BeGreaterThan 0
    }

    It 'has no duplicate RecordId plus ResourceId combination' {
        $keys = @($script:Copilot | ForEach-Object { '{0}|{1}' -f $_.RecordId, $_.ResourceId })
        ($keys | Sort-Object -Unique).Count | Should -Be $keys.Count
    }
}
