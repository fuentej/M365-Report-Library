#Requires -Version 7.0

<#
    The committed sample set is what the Power BI report is built against, so its shape
    matters as much as the collectors'. These tests hold it to the shape the report needs:
    a directory deep enough to roll up by manager, guests in every state worth showing,
    and the sharing activity a guest-access review asks about.
#>

BeforeAll {
    $script:Root = Join-Path $PSScriptRoot '../../..' | Resolve-Path | Select-Object -ExpandProperty Path
    $script:Samples = Join-Path $script:Root 'reports/guest-access/samples'

    Import-Module (Join-Path $script:Root 'shared/M365ReportLibrary.psm1') -Force

    $script:Users = @(Import-Csv -LiteralPath (Join-Path $script:Samples 'users.csv'))
    $script:Guests = @(Import-Csv -LiteralPath (Join-Path $script:Samples 'guests.csv'))
    $script:Invitations = @(Import-Csv -LiteralPath (Join-Path $script:Samples 'guest-invitations.csv'))
    $script:SignIns = @(Import-Csv -LiteralPath (Join-Path $script:Samples 'guest-signins.csv'))
    $script:Sharing = @(Import-Csv -LiteralPath (Join-Path $script:Samples 'sharing-events.csv'))
    $script:Memberships = @(Import-Csv -LiteralPath (Join-Path $script:Samples 'guest-memberships.csv'))

    $script:LatestRunDate = ($script:Guests.RunDate | Sort-Object -Descending | Select-Object -First 1)
    $script:LatestGuests = @($script:Guests | Where-Object RunDate -EQ $script:LatestRunDate)
    $script:Members = @($script:Users | Where-Object { $_.RunDate -eq $script:LatestRunDate -and $_.UserType -eq 'Member' })
}

AfterAll {
    Remove-Module M365ReportLibrary -Force -ErrorAction SilentlyContinue
}

Describe 'The sample data is fake' {
    It 'uses no address outside example.com' {
        $addresses = @(
            $script:Users.UserPrincipalName
            $script:Users.Mail
            $script:Guests.Mail
            $script:Guests.UserPrincipalName
            $script:Invitations.InitiatedByUserPrincipalName
            $script:Invitations.TargetUserPrincipalName
            $script:SignIns.UserPrincipalName
        ) | Where-Object { $_ -and $_.Contains('@') }

        $outside = @($addresses | Where-Object { $_ -notmatch '@([a-z0-9-]+\.)*example\.(com|onmicrosoft\.com)$' } | Sort-Object -Unique)
        $outside -join '; ' | Should -BeNullOrEmpty
    }

    It 'uses only documentation IP addresses for sign-ins' {
        # 203.0.113.0/24 is TEST-NET-3, reserved for documentation by RFC 5737.
        @($script:SignIns.IpAddress | Where-Object { $_ -notlike '203.0.113.*' }).Count | Should -Be 0
    }
}

Describe 'The directory is deep and wide enough to report on' {
    It 'holds 200 members' {
        $script:Members.Count | Should -Be 200
    }

    It 'spreads them over at least five departments' {
        ($script:Members.Department | Sort-Object -Unique).Count | Should -BeGreaterOrEqual 5
    }

    It 'has a management chain three levels deep' {
        $byId = @{}
        foreach ($member in $script:Members) { $byId[$member.Id] = $member }

        $depths = foreach ($member in $script:Members) {
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

    It 'takes snapshots on six dates a month apart' {
        $dates = @($script:Guests.RunDate | Sort-Object -Unique | ForEach-Object { [datetime]$_ })
        $dates.Count | Should -Be 6

        for ($i = 1; $i -lt $dates.Count; $i++) {
            $gap = ($dates[$i] - $dates[$i - 1]).TotalDays
            $gap | Should -BeGreaterOrEqual 28
            $gap | Should -BeLessOrEqual 31
        }
    }
}

Describe 'The guests cover the states a guest-access review looks for' {
    It 'holds 60 guests from 12 partner domains' {
        $script:LatestGuests.Count | Should -Be 60
        ($script:LatestGuests.ExternalDomain | Sort-Object -Unique).Count | Should -Be 12
    }

    It 'includes guests who never accepted their invitation' {
        @($script:LatestGuests | Where-Object ExternalUserState -NE 'Accepted').Count | Should -BeGreaterThan 0
    }

    It 'includes guests with no sign-in for 90 days or more' {
        $cutoff = ([datetime]$script:LatestRunDate).AddDays(-90)
        $dormant = @($script:LatestGuests | Where-Object {
                -not [string]::IsNullOrWhiteSpace($_.LastSignInDateTime) -and ([datetime]$_.LastSignInDateTime) -lt $cutoff
            })
        $dormant.Count | Should -BeGreaterThan 0
    }

    It 'includes disabled guests' {
        @($script:LatestGuests | Where-Object AccountEnabled -EQ 'False').Count | Should -BeGreaterThan 0
    }

    It 'names every guest in users.csv as well' {
        $guestIdsInUsers = @($script:Users | Where-Object { $_.RunDate -eq $script:LatestRunDate -and $_.UserType -eq 'Guest' }).Id
        foreach ($guest in $script:LatestGuests) {
            $guestIdsInUsers | Should -Contain $guest.Id
        }
    }
}

Describe 'The events span the window the report charts' {
    It 'covers about 240 days of sharing activity' {
        $times = @($script:Sharing.CreationTime | ForEach-Object { [datetime]$_ })
        $span = (($times | Measure-Object -Maximum).Maximum - ($times | Measure-Object -Minimum).Minimum).TotalDays
        $span | Should -BeGreaterThan 200
    }

    It 'has invitations from at least twenty different members' {
        $senders = @($script:Invitations |
                Where-Object ActivityDisplayName -Like 'Invite*' |
                ForEach-Object { $_.InitiatedByUserPrincipalName } |
                Sort-Object -Unique)
        $senders.Count | Should -BeGreaterOrEqual 20
    }

    It 'records both invitations and redemptions' {
        $activities = @($script:Invitations.ActivityDisplayName | Sort-Object -Unique)
        $activities | Should -Contain 'Invite external user'
        $activities | Should -Contain 'Redeem external user invite'
    }

    It 'uses only activity names the Entra audit reference lists' {
        $known = (Import-PowerShellDataFile -LiteralPath (Join-Path $script:Root 'reports/guest-access/collectors/GuestAccessSchema.psd1')).InvitationActivities
        foreach ($activity in ($script:Invitations.ActivityDisplayName | Sort-Object -Unique)) {
            $known | Should -Contain $activity
        }
    }

    It 'includes <Operation> sharing events' -ForEach @(
        @{ Operation = 'AnonymousLinkCreated' }
        @{ Operation = 'AnonymousLinkUsed' }
        @{ Operation = 'AnonymousLinkRemoved' }
        @{ Operation = 'SecureLinkCreated' }
        @{ Operation = 'AddedToSecureLink' }
        @{ Operation = 'SecureLinkUsed' }
        @{ Operation = 'SharingSet' }
        @{ Operation = 'SharingRevoked' }
        @{ Operation = 'SharingInvitationCreated' }
        @{ Operation = 'SharingInvitationAccepted' }
        @{ Operation = 'SharingInvitationRevoked' }
    ) {
        @($script:Sharing | Where-Object Operation -EQ $Operation).Count | Should -BeGreaterThan 0
    }

    It 'dates no event after the moment the set is generated around' {
        # The generator's -EndDate is its "now". An event after it has not happened yet.
        $endOfWindow = ([datetime]$script:LatestRunDate).AddDays(1)

        $strays = @(
            $script:Invitations | ForEach-Object { [pscustomobject]@{ File = 'guest-invitations.csv'; When = [datetime]$_.ActivityDateTime; Id = $_.Id } }
            $script:SignIns | ForEach-Object { [pscustomobject]@{ File = 'guest-signins.csv'; When = [datetime]$_.CreatedDateTime; Id = $_.Id } }
            $script:Sharing | ForEach-Object { [pscustomobject]@{ File = 'sharing-events.csv'; When = [datetime]$_.CreationTime; Id = $_.Id } }
        ) | Where-Object { $_.When -ge $endOfWindow } | ForEach-Object { '{0}:{1}' -f $_.File, $_.Id }

        $strays -join '; ' | Should -BeNullOrEmpty
    }

    It 'records no sign-in before the guest redeemed their invitation' {
        $acceptedAt = @{}
        foreach ($guest in $script:LatestGuests) {
            if ($guest.ExternalUserState -eq 'Accepted' -and -not [string]::IsNullOrWhiteSpace($guest.ExternalUserStateChangeDateTime)) {
                $acceptedAt[$guest.Id] = [datetime]$guest.ExternalUserStateChangeDateTime
            }
        }

        $early = @($script:SignIns | Where-Object {
                $acceptedAt.ContainsKey($_.UserId) -and ([datetime]$_.CreatedDateTime) -lt $acceptedAt[$_.UserId]
            } | ForEach-Object { $_.Id })

        $early -join '; ' | Should -BeNullOrEmpty
    }

    It 'never shows a guest as accepted in a snapshot taken before they redeemed' {
        $wrong = @($script:Guests | Where-Object {
                $_.ExternalUserState -eq 'Accepted' -and
                -not [string]::IsNullOrWhiteSpace($_.ExternalUserStateChangeDateTime) -and
                ([datetime]$_.ExternalUserStateChangeDateTime) -gt ([datetime]$_.RunDate).AddDays(1)
            } | ForEach-Object { '{0}@{1}' -f $_.Id, $_.RunDate })

        $wrong -join '; ' | Should -BeNullOrEmpty
    }

    It 'shows at least one guest pending in an early snapshot and accepted in a later one' {
        # The per-snapshot state is what makes the redemption funnel visible over time.
        $byId = $script:Guests | Group-Object Id
        $changed = @($byId | Where-Object {
                ($_.Group.ExternalUserState | Sort-Object -Unique).Count -gt 1
            })

        $changed.Count | Should -BeGreaterThan 0
    }

    It 'attributes every sign-in to a guest' {
        $guestIds = @($script:Guests.Id | Sort-Object -Unique)
        $strays = @($script:SignIns.UserId | Sort-Object -Unique | Where-Object { $guestIds -notcontains $_ })
        $strays -join '; ' | Should -BeNullOrEmpty
    }
}

Describe 'Guests belong to both Teams and plain groups' {
    BeforeAll {
        $script:LatestMemberships = @($script:Memberships | Where-Object RunDate -EQ $script:LatestRunDate)
    }

    It 'puts guests in Teams' {
        @($script:LatestMemberships | Where-Object IsTeam -EQ 'True').Count | Should -BeGreaterThan 0
    }

    It 'puts guests in groups that are not Teams' {
        @($script:LatestMemberships | Where-Object IsTeam -EQ 'False').Count | Should -BeGreaterThan 0
    }

    It 'gives every guest at least one group' {
        $withGroups = @($script:LatestMemberships.GuestId | Sort-Object -Unique)
        $withGroups.Count | Should -Be $script:LatestGuests.Count
    }
}

Describe 'Every sample event file is keyed uniquely' {
    It 'has no duplicate Id in <File>' -ForEach @(
        @{ File = 'guest-invitations.csv' }
        @{ File = 'guest-signins.csv' }
        @{ File = 'sharing-events.csv' }
    ) {
        $ids = @((Import-Csv -LiteralPath (Join-Path $script:Samples $File)).Id)
        ($ids | Sort-Object -Unique).Count | Should -Be $ids.Count
    }
}
