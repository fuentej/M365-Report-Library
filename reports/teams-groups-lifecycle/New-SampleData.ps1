#Requires -Version 7.0

<#
    .SYNOPSIS
        Generates the fake Teams and Groups lifecycle data in ./samples.

    .DESCRIPTION
        The sample set lets the later Power BI report, and anyone reading this repository,
        work against realistic shapes without a tenant. Nothing here comes from a real
        directory: every address is under example.com, which RFC 2606 reserves for
        documentation.

        The generator is deterministic. The same -Seed and -EndDate always produce the
        same files, so a regenerated sample set shows up in a diff only when this script
        changes.

        Every file is written through Export-AppendCsv with the column list the collectors
        use, so a sample file cannot drift from its collector's output. ./samples/gcchigh
        holds the header-only files the two usage-report collectors write in GCC High,
        where Microsoft documents those APIs as not available.

    .PARAMETER EndDate
        The "now" the sample set is generated around. Fixed by default to keep the
        committed files stable.

    .EXAMPLE
        ./New-SampleData.ps1
#>
[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path $PSScriptRoot 'samples'),

    [datetime]$EndDate = [datetime]::new(2026, 9, 1, 0, 0, 0, [System.DateTimeKind]::Utc),

    [int]$Seed = 20260901
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot '../../shared/M365ReportLibrary.psm1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'collectors/TeamsGroupsSchema.psd1')

$EndDate = $EndDate.ToUniversalTime()
$random = [System.Random]::new($Seed)

# Three snapshots a month apart, the last one on EndDate.
$snapshotDates = 2..0 | ForEach-Object { $EndDate.AddMonths(-$_) }
$latest = $snapshotDates[-1]

$memberCount = 60
$groupCount = 48
$teamShare = 30
$deletedCount = 6
$policyLifetimeDays = 365
$tenantDomain = 'example.com'

$departments = @('Engineering', 'Sales', 'Marketing', 'Finance', 'Operations', 'Legal')
$topics = @(
    'Northwind Rollout', 'Contoso Migration', 'Fabrikam Pricing', 'Adventure Works Launch', 'Tailspin Support'
    'Litware Security', 'Proseware Hiring', 'Wingtip Events', 'Lucerne Budget', 'Margie Travel', 'Trey Research'
    'Woodgrove Audit', 'Alpine Ski House', 'Humongous Insurance', 'Coho Vineyard', 'Fourth Coffee Ops'
)
$suffixes = @('Project', 'Team', 'Working Group', 'Community', 'Planning', 'Review')
$givenNames = @('Avery', 'Blair', 'Casey', 'Devon', 'Emery', 'Finley', 'Gray', 'Harper', 'Indigo', 'Jordan', 'Kai', 'Logan', 'Marlowe', 'Nico', 'Oakley', 'Parker', 'Quinn', 'Reese', 'Sage', 'Tatum')
$familyNames = @('Abara', 'Bergstrom', 'Chaudhry', 'Dlamini', 'Ibarra', 'Jovanovic', 'Kowalski', 'Lindqvist', 'Moreau', 'Nakamura', 'Quintero', 'Silva')

function New-DeterministicGuid {
    <#
        .SYNOPSIS
            A GUID drawn from the seeded generator, so re-running produces the same ids.
    #>
    $bytes = [byte[]]::new(16)
    $random.NextBytes($bytes)
    return [guid]::new($bytes).ToString()
}

function Get-RandomItem {
    param([Parameter(Mandatory)][object[]]$Items)
    return $Items[$random.Next(0, $Items.Count)]
}

function Format-Stamp {
    param([datetime]$Value)
    return ConvertTo-CsvTimestamp $Value
}

if (Test-Path -LiteralPath $OutputPath) {
    Get-ChildItem -LiteralPath $OutputPath -Recurse -File -Filter '*.csv' | Remove-Item -Force
}
New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null

#region Users

$users = foreach ($i in 1..$memberCount) {
    $given = $givenNames[($i - 1) % $givenNames.Count]
    $family = $familyNames[($i * 5) % $familyNames.Count]
    $upn = ('{0}.{1}{2}@{3}' -f $given, $family, $(if ($i -gt $givenNames.Count) { $i } else { '' }), $tenantDomain).ToLowerInvariant()
    [pscustomobject]@{
        Id                = New-DeterministicGuid
        DisplayName       = "$given $family"
        UserPrincipalName = $upn
        # Every ninth account is disabled, so "owners who are disabled" has something to find.
        AccountEnabled    = (($i % 9) -ne 0)
        Department        = Get-RandomItem $departments
        Created           = $EndDate.AddDays(-$random.Next(200, 1200))
    }
}
$users = @($users)

foreach ($snapshot in $snapshotDates) {
    $rows = foreach ($user in $users) {
        [pscustomobject]@{
            RunDate                  = $snapshot.ToString('yyyy-MM-dd')
            Id                       = $user.Id
            DisplayName              = $user.DisplayName
            UserPrincipalName        = $user.UserPrincipalName
            Mail                     = $user.UserPrincipalName
            UserType                 = 'Member'
            AccountEnabled           = $user.AccountEnabled
            CreatedDateTime          = Format-Stamp $user.Created
            Department               = $user.Department
            JobTitle                 = 'Specialist'
            City                     = 'Seattle'
            Country                  = 'US'
            ManagerId                = ''
            ManagerUserPrincipalName = ''
        }
    }
    Export-AppendCsv -Path (Join-Path $OutputPath 'users.csv') -Rows @($rows) -Column (Get-EntraUserCsvColumn) -KeyColumn @('RunDate', 'Id')
}

#endregion

#region Groups

$groups = foreach ($i in 1..($groupCount + $deletedCount)) {
    $name = '{0} {1}' -f $topics[($i - 1) % $topics.Count], $suffixes[(($i - 1) % ($suffixes.Count * 2)) % $suffixes.Count]
    if ($i -gt $topics.Count) { $name = "$name $([math]::Floor(($i - 1) / $topics.Count) + 1)" }

    $isTeam = ($i -le $teamShare) -or (($i -gt $groupCount) -and ($i % 2 -eq 0))
    $created = $EndDate.AddDays(-$random.Next(20, 700))
    $renewed = if ($random.Next(0, 3) -eq 0) { $created } else { $EndDate.AddDays(-$random.Next(1, 330)) }
    if ($renewed -lt $created) { $renewed = $created }

    $ownerRoll = $random.Next(0, 10)
    $ownerTotal = if ($ownerRoll -lt 2) { 0 } elseif ($ownerRoll -lt 5) { 1 } else { $random.Next(2, 4) }

    [pscustomobject]@{
        Id          = New-DeterministicGuid
        DisplayName = $name
        Slug        = ($name -replace '[^A-Za-z0-9]', '').ToLowerInvariant()
        IsTeam      = $isTeam
        Created     = $created
        Renewed     = $renewed
        Visibility  = Get-RandomItem @('Private', 'Private', 'Public')
        OwnerTotal  = $ownerTotal
        # Two groups stand in for groups Graph cannot list owners for.
        Synced      = ($i -eq 7 -or $i -eq 19)
        IsDeleted   = ($i -gt $groupCount)
    }
}
$groups = @($groups)
$activeGroups = @($groups | Where-Object { -not $_.IsDeleted })
$deletedGroups = @($groups | Where-Object { $_.IsDeleted })

foreach ($snapshot in $snapshotDates) {
    $rows = foreach ($group in $activeGroups | Where-Object { $_.Created -le $snapshot }) {
        # Policy applies to the first 36 active groups only (a Selected scope).
        $covered = ([array]::IndexOf($activeGroups, $group) -lt 36)
        [pscustomobject]@{
            RunDate                     = $snapshot.ToString('yyyy-MM-dd')
            Id                          = $group.Id
            DisplayName                 = $group.DisplayName
            Mail                        = "$($group.Slug)@$tenantDomain"
            GroupTypes                  = 'Unified'
            SecurityEnabled             = $false
            MailEnabled                 = $true
            Visibility                  = $group.Visibility
            CreatedDateTime             = Format-Stamp $group.Created
            RenewedDateTime             = Format-Stamp $group.Renewed
            ExpirationDateTime          = if ($covered) { Format-Stamp $group.Renewed.AddDays($policyLifetimeDays) } else { '' }
            DeletedDateTime             = ''
            OnPremisesSyncEnabled       = if ($group.Synced) { $true } else { '' }
            ResourceProvisioningOptions = if ($group.IsTeam) { 'Team' } else { '' }
            IsTeam                      = $group.IsTeam
        }
    }
    Export-AppendCsv -Path (Join-Path $OutputPath 'groups.csv') -Rows @($rows) -Column $schema.Groups -KeyColumn @('RunDate', 'Id')
}

#endregion

#region Owners

foreach ($snapshot in $snapshotDates) {
    $rows = foreach ($group in $activeGroups | Where-Object { $_.Created -le $snapshot }) {
        if ($group.Synced) {
            [pscustomobject]@{
                RunDate = $snapshot.ToString('yyyy-MM-dd'); GroupId = $group.Id; OwnerListStatus = 'Unknown'
                OwnerId = ''; OwnerType = ''; OwnerDisplayName = ''; OwnerUserPrincipalName = ''
            }
        }
        elseif ($group.OwnerTotal -eq 0) {
            [pscustomobject]@{
                RunDate = $snapshot.ToString('yyyy-MM-dd'); GroupId = $group.Id; OwnerListStatus = 'None'
                OwnerId = ''; OwnerType = ''; OwnerDisplayName = ''; OwnerUserPrincipalName = ''
            }
        }
        else {
            $start = [array]::IndexOf($activeGroups, $group) % $users.Count
            foreach ($n in 0..($group.OwnerTotal - 1)) {
                $owner = $users[($start + $n * 7) % $users.Count]
                [pscustomobject]@{
                    RunDate = $snapshot.ToString('yyyy-MM-dd'); GroupId = $group.Id; OwnerListStatus = 'Listed'
                    OwnerId = $owner.Id; OwnerType = 'user'; OwnerDisplayName = $owner.DisplayName
                    OwnerUserPrincipalName = $owner.UserPrincipalName
                }
            }
        }
    }
    Export-AppendCsv -Path (Join-Path $OutputPath 'group-owners.csv') -Rows @($rows) -Column $schema.GroupOwners -KeyColumn @('RunDate', 'GroupId', 'OwnerId')
}

#endregion

#region Deleted groups

foreach ($snapshot in $snapshotDates | Select-Object -Last 2) {
    $rows = foreach ($group in $deletedGroups) {
        $deletedAt = $EndDate.AddDays(-(3 + 4 * [array]::IndexOf($deletedGroups, $group)))
        if ($deletedAt -gt $snapshot) { continue }
        [pscustomobject]@{
            RunDate                     = $snapshot.ToString('yyyy-MM-dd')
            Id                          = $group.Id
            DisplayName                 = $group.DisplayName
            GroupTypes                  = 'Unified'
            SecurityEnabled             = $false
            MailEnabled                 = $true
            CreatedDateTime             = Format-Stamp $group.Created
            DeletedDateTime             = Format-Stamp $deletedAt
            PurgeDateTime               = Format-Stamp $deletedAt.AddDays(30)
            ResourceProvisioningOptions = if ($group.IsTeam) { 'Team' } else { '' }
            IsTeam                      = $group.IsTeam
        }
    }

    # Soft-deleted security groups come back with securityEnabled false and an
    # empty groupTypes array. Microsoft 365 groups are groupTypes Unified.
    # https://learn.microsoft.com/graph/api/directory-deleteditems-list
    $securityDeletedAt = $EndDate.AddDays(-40)
    if ($securityDeletedAt -le $snapshot) {
        $rows = @($rows) + [pscustomobject]@{
            RunDate                     = $snapshot.ToString('yyyy-MM-dd')
            Id                          = 'c31799b8-0683-4d70-9e91-e032c89d3035'
            DisplayName                 = 'Contoso Role Assignable'
            GroupTypes                  = ''
            SecurityEnabled             = $false
            MailEnabled                 = $false
            CreatedDateTime             = Format-Stamp ($EndDate.AddDays(-400))
            DeletedDateTime             = Format-Stamp $securityDeletedAt
            PurgeDateTime               = Format-Stamp $securityDeletedAt.AddDays(30)
            ResourceProvisioningOptions = ''
            IsTeam                      = $false
        }
    }
    Export-AppendCsv -Path (Join-Path $OutputPath 'deleted-groups.csv') -Rows @($rows) -Column $schema.DeletedGroups -KeyColumn @('RunDate', 'Id')
}

#endregion

#region Lifecycle policy and coverage

$policyId = New-DeterministicGuid
foreach ($snapshot in $snapshotDates) {
    $policy = [pscustomobject]@{
        RunDate = $snapshot.ToString('yyyy-MM-dd'); Id = $policyId; GroupLifetimeInDays = $policyLifetimeDays
        ManagedGroupTypes = 'Selected'; AlternateNotificationEmails = "groups-admin@$tenantDomain"
    }
    Export-AppendCsv -Path (Join-Path $OutputPath 'group-lifecycle-policies.csv') -Rows @($policy) -Column $schema.GroupLifecyclePolicies -KeyColumn @('RunDate', 'Id')

    $rows = foreach ($group in $activeGroups | Where-Object { $_.Created -le $snapshot }) {
        $covered = ([array]::IndexOf($activeGroups, $group) -lt 36)
        [pscustomobject]@{
            RunDate        = $snapshot.ToString('yyyy-MM-dd')
            GroupId        = $group.Id
            PolicyId       = $policyId
            CoverageStatus = if ($covered) { 'Covered' } else { 'NotCovered' }
        }
    }
    Export-AppendCsv -Path (Join-Path $OutputPath 'group-lifecycle-coverage.csv') -Rows @($rows) -Column $schema.GroupLifecycleCoverage -KeyColumn @('RunDate', 'GroupId')
}

#endregion

#region Archive status

foreach ($snapshot in $snapshotDates) {
    $rows = foreach ($group in $activeGroups | Where-Object { $_.IsTeam -and $_.Created -le $snapshot }) {
        [pscustomobject]@{
            RunDate     = $snapshot.ToString('yyyy-MM-dd')
            TeamId      = $group.Id
            DisplayName = $group.DisplayName
            IsArchived  = ([array]::IndexOf($activeGroups, $group) % 6 -eq 5)
        }
    }
    Export-AppendCsv -Path (Join-Path $OutputPath 'team-archive-status.csv') -Rows @($rows) -Column $schema.TeamArchiveStatus -KeyColumn @('RunDate', 'TeamId')
}

#endregion

#region Usage reports

foreach ($snapshot in $snapshotDates) {
    $teamRows = [System.Collections.Generic.List[object]]::new()
    $groupRows = [System.Collections.Generic.List[object]]::new()

    foreach ($group in $activeGroups | Where-Object { $_.Created -le $snapshot }) {
        $daysIdle = $random.Next(0, 400)
        $last = $snapshot.AddDays(-$daysIdle)
        if ($last -lt $group.Created) { $last = $group.Created }
        $active = $daysIdle -lt 180
        $refresh = $snapshot.AddDays(-3).ToString('yyyy-MM-dd')

        if ($group.IsTeam) {
            $teamRows.Add([pscustomobject]@{
                    RunDate              = $snapshot.ToString('yyyy-MM-dd')
                    ReportRefreshDate    = $refresh
                    ReportPeriod         = 180
                    TeamId               = $group.Id
                    TeamName             = $group.DisplayName
                    TeamType             = 'Public'
                    LastActivityDate     = $last.ToString('yyyy-MM-dd')
                    ActiveUsers          = if ($active) { $random.Next(1, 40) } else { 0 }
                    ActiveChannels       = if ($active) { $random.Next(1, 8) } else { 0 }
                    Guests               = $random.Next(0, 5)
                    Reactions            = if ($active) { $random.Next(0, 300) } else { 0 }
                    MeetingsOrganized    = if ($active) { $random.Next(0, 40) } else { 0 }
                    PostMessages         = if ($active) { $random.Next(0, 500) } else { 0 }
                    ReplyMessages        = if ($active) { $random.Next(0, 800) } else { 0 }
                    ChannelMessages      = if ($active) { $random.Next(0, 1200) } else { 0 }
                    UrgentMessages       = if ($active) { $random.Next(0, 5) } else { 0 }
                    Mentions             = if ($active) { $random.Next(0, 90) } else { 0 }
                    ActiveSharedChannels = 0
                    ActiveExternalUsers  = $random.Next(0, 3)
                })
        }

        $groupRows.Add([pscustomobject]@{
                RunDate                        = $snapshot.ToString('yyyy-MM-dd')
                ReportRefreshDate              = $refresh
                ReportPeriod                   = 180
                GroupId                        = $group.Id
                GroupDisplayName               = $group.DisplayName
                IsDeleted                      = $false
                OwnerPrincipalName             = if ($group.OwnerTotal -gt 0) { $users[([array]::IndexOf($activeGroups, $group) % $users.Count)].UserPrincipalName } else { '' }
                LastActivityDate               = $last.ToString('yyyy-MM-dd')
                GroupType                      = $group.Visibility
                MemberCount                    = $random.Next(2, 60)
                ExternalMemberCount            = $random.Next(0, 5)
                ExchangeReceivedEmailCount     = if ($active) { $random.Next(0, 400) } else { 0 }
                SharePointActiveFileCount      = if ($active) { $random.Next(0, 200) } else { 0 }
                YammerPostedMessageCount       = 0
                YammerReadMessageCount         = 0
                YammerLikedMessageCount        = 0
                ExchangeMailboxTotalItemCount  = $random.Next(0, 5000)
                ExchangeMailboxStorageUsedByte = $random.Next(0, 90000000)
                SharePointTotalFileCount       = $random.Next(0, 3000)
                SharePointSiteStorageUsedByte  = $random.Next(0, 900000000)
            })
    }

    Export-AppendCsv -Path (Join-Path $OutputPath 'team-activity.csv') -Rows $teamRows.ToArray() -Column $schema.TeamActivity -KeyColumn @('RunDate', 'TeamId')
    Export-AppendCsv -Path (Join-Path $OutputPath 'group-activity.csv') -Rows $groupRows.ToArray() -Column $schema.GroupActivity -KeyColumn @('RunDate', 'GroupId')
}

#endregion

#region Creation events

# Audit (Standard) keeps 180 days, so only groups created inside that window have a record.
$auditStart = $EndDate.AddDays(-180)
$events = foreach ($group in $groups | Where-Object { $_.Created -ge $auditStart }) {
    $creator = $users[$random.Next(0, $users.Count)]
    $operation = if ($group.IsTeam) { 'TeamCreated' } else { 'AddGroup' }
    [pscustomobject]@{
        CreationTime      = Format-Stamp $group.Created
        Id                = New-DeterministicGuid
        Operation         = $operation
        UserId            = $creator.UserPrincipalName
        Workload          = if ($group.IsTeam) { 'MicrosoftTeams' } else { 'AzureActiveDirectory' }
        ObjectId          = $group.Id
        TargetDisplayName = if ($group.IsTeam) { $group.DisplayName } else { '' }
        TargetGroupId     = if ($group.IsTeam) { $group.Id } else { '' }
    }
}
Export-AppendCsv -Path (Join-Path $OutputPath 'group-creation-events.csv') -Rows @($events | Sort-Object CreationTime) -Column $schema.GroupCreationEvents -KeyColumn 'Id'

#endregion

#region GCC High: the two usage-report sources are not available

$gccHighPath = Join-Path $OutputPath 'gcchigh'
Export-AppendCsv -Path (Join-Path $gccHighPath 'team-activity.csv') -Column $schema.TeamActivity
Export-AppendCsv -Path (Join-Path $gccHighPath 'group-activity.csv') -Column $schema.GroupActivity

#endregion
