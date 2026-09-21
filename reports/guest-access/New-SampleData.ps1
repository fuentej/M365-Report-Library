#Requires -Version 7.0

<#
    .SYNOPSIS
        Generates the fake guest-access data in ./samples.

    .DESCRIPTION
        The sample set lets the Power BI report, and anyone reading this repository, work
        against realistic shapes without a tenant. Nothing here comes from a real
        directory: every address is under example.com or a partnerN.example.com subdomain,
        which RFC 2606 reserves for documentation.

        The generator is deterministic. The same -Seed and -EndDate always produce the
        same files, so a regenerated sample set shows up in a diff only when this script
        changes.

        Every file is written through Export-AppendCsv with the column list the collectors
        use, so a sample file cannot drift from its collector's output.

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

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'collectors/GuestAccessSchema.psd1')

$EndDate = $EndDate.ToUniversalTime()
$eventDays = 240
$eventStart = $EndDate.AddDays(-$eventDays)

# Six snapshots a month apart, the last one on EndDate.
$snapshotDates = 5..0 | ForEach-Object { $EndDate.AddMonths(-$_) }

$memberCount = 200
$partnerDomainCount = 12
$guestCount = 60
$tenantDomain = 'example.com'
$tenantOnMicrosoft = 'example.onmicrosoft.com'

$departments = @('Engineering', 'Sales', 'Marketing', 'Finance', 'Operations', 'Legal')
$cities = @(
    @{ City = 'Seattle'; Country = 'US' }
    @{ City = 'Austin'; Country = 'US' }
    @{ City = 'Dublin'; Country = 'IE' }
    @{ City = 'Toronto'; Country = 'CA' }
    @{ City = 'Sydney'; Country = 'AU' }
)

$random = [System.Random]::new($Seed)

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

function New-EventTime {
    <#
        .SYNOPSIS
            A timestamp inside the event window, biased to working hours.
    #>
    param([datetime]$From = $eventStart, [datetime]$To = $EndDate)

    $span = ($To - $From).TotalMinutes
    if ($span -le 1) { return $From }
    $moment = $From.AddMinutes($random.Next(0, [int]$span))
    return [datetime]::new($moment.Year, $moment.Month, $moment.Day,
        $random.Next(7, 20), $random.Next(0, 60), $random.Next(0, 60), [System.DateTimeKind]::Utc)
}

$firstNames = @(
    'Avery', 'Blair', 'Casey', 'Drew', 'Emery', 'Finley', 'Gray', 'Harper', 'Indigo', 'Jordan'
    'Kai', 'Logan', 'Marlowe', 'Noor', 'Oakley', 'Parker', 'Quinn', 'Reese', 'Sage', 'Tatum'
    'Umi', 'Vale', 'Wren', 'Xan', 'Yael', 'Zephyr'
)
$lastNames = @(
    'Abara', 'Bergstrom', 'Chaudhry', 'Delacroix', 'Eriksen', 'Fontaine', 'Gallagher', 'Haddad'
    'Ibarra', 'Jovanovic', 'Kowalski', 'Lindqvist', 'Moreau', 'Nakamura', 'Okonkwo', 'Petrov'
    'Quintero', 'Rasmussen', 'Silva', 'Takahashi', 'Ueda', 'Vasquez', 'Whitfield', 'Xiong'
    'Yilmaz', 'Zawadzki'
)

#region People

# The member hierarchy is three levels of manager under one chief executive:
# executive -> director (one per department) -> manager -> individual contributor.
$members = [System.Collections.Generic.List[object]]::new()
$usedAliases = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

function Get-UniqueAlias {
    param([string]$First, [string]$Last)
    $base = ('{0}.{1}' -f $First, $Last).ToLowerInvariant()
    $alias = $base
    $suffix = 2
    while (-not $usedAliases.Add($alias)) {
        $alias = '{0}{1}' -f $base, $suffix
        $suffix++
    }
    return $alias
}

for ($i = 0; $i -lt $memberCount; $i++) {
    $first = Get-RandomItem $firstNames
    $last = Get-RandomItem $lastNames
    $alias = Get-UniqueAlias -First $first -Last $last
    $place = Get-RandomItem $cities

    $members.Add([pscustomobject]@{
            Id                = New-DeterministicGuid
            DisplayName       = ('{0} {1}' -f $first, $last)
            UserPrincipalName = ('{0}@{1}' -f $alias, $tenantDomain)
            Mail              = ('{0}@{1}' -f $alias, $tenantDomain)
            Department        = $departments[$i % $departments.Count]
            JobTitle          = $null
            City              = $place.City
            Country           = $place.Country
            CreatedDateTime   = $eventStart.AddDays(-$random.Next(30, 1500))
            ManagerId         = $null
            ManagerUpn        = $null
            Level             = 3
        })
}

$executive = $members[0]
$executive.Level = 0
$executive.JobTitle = 'Chief Executive'
$executive.Department = 'Operations'

# One director per department.
$directors = @{}
$index = 1
foreach ($department in $departments) {
    $director = $members[$index]
    $director.Level = 1
    $director.Department = $department
    $director.JobTitle = "Director, $department"
    $director.ManagerId = $executive.Id
    $director.ManagerUpn = $executive.UserPrincipalName
    $directors[$department] = $director
    $index++
}

# Three managers per department, reporting to that department's director.
$managers = @{}
foreach ($department in $departments) { $managers[$department] = [System.Collections.Generic.List[object]]::new() }
for ($m = 0; $m -lt ($departments.Count * 3); $m++) {
    $manager = $members[$index]
    $department = $departments[$m % $departments.Count]
    $manager.Level = 2
    $manager.Department = $department
    $manager.JobTitle = "$department Manager"
    $manager.ManagerId = $directors[$department].Id
    $manager.ManagerUpn = $directors[$department].UserPrincipalName
    $managers[$department].Add($manager)
    $index++
}

# Everyone else reports to a manager in their own department.
for (; $index -lt $members.Count; $index++) {
    $person = $members[$index]
    $department = $person.Department
    $manager = Get-RandomItem $managers[$department].ToArray()
    $person.Level = 3
    $person.JobTitle = "$department Specialist"
    $person.ManagerId = $manager.Id
    $person.ManagerUpn = $manager.UserPrincipalName
}

$partnerDomains = 1..$partnerDomainCount | ForEach-Object { 'partner{0}.example.com' -f $_ }

# Members who send invitations: well over the twenty the report needs to show a spread.
$inviters = @($members | Where-Object { $_.Level -ge 1 } | Select-Object -First 26)

$guests = [System.Collections.Generic.List[object]]::new()
for ($g = 0; $g -lt $guestCount; $g++) {
    $first = Get-RandomItem $firstNames
    $last = Get-RandomItem $lastNames
    $domain = $partnerDomains[$g % $partnerDomainCount]
    $alias = ('{0}.{1}' -f $first, $last).ToLowerInvariant()
    # Two guests drawn from the same name pool can land on the same partner domain.
    $suffix = 2
    while (-not $usedAliases.Add(('{0}@{1}' -f $alias, $domain))) {
        $alias = ('{0}.{1}{2}' -f $first, $last, $suffix).ToLowerInvariant()
        $suffix++
    }
    $mail = '{0}@{1}' -f $alias, $domain

    # Ten guests never redeem their invitation; eight are later disabled; twelve of the
    # rest go quiet for more than ninety days.
    $neverAccepted = $g -lt 10
    $disabled = $g -ge 10 -and $g -lt 18
    $dormant = $g -ge 18 -and $g -lt 30

    # Spread invitations across the whole event window so snapshots grow over time. A
    # guest who has to have been quiet for ninety days has to have been invited long
    # enough ago for that to be true.
    $invited = if ($dormant) {
        New-EventTime -From $eventStart -To $EndDate.AddDays(-120)
    }
    else {
        New-EventTime -From $eventStart -To $EndDate.AddDays(-5)
    }

    $accepted = if ($neverAccepted) { $null } else { $invited.AddDays($random.Next(0, 6)).AddHours($random.Next(1, 20)) }

    $lastSignIn = $null
    if (-not $neverAccepted) {
        $lastSignIn = if ($dormant) {
            $EndDate.AddDays(-$random.Next(95, 200))
        }
        else {
            $EndDate.AddDays(-$random.Next(0, 25)).AddHours($random.Next(0, 23))
        }
        if ($lastSignIn -lt $accepted) { $lastSignIn = $accepted.AddDays(1) }
    }

    $guests.Add([pscustomobject]@{
            Id                = New-DeterministicGuid
            DisplayName       = ('{0} {1} ({2})' -f $first, $last, $domain)
            Mail              = $mail
            UserPrincipalName = ('{0}_{1}#EXT#@{2}' -f $alias, $domain, $tenantOnMicrosoft)
            ExternalDomain    = $domain
            CreatedDateTime   = $invited
            CreationType      = 'Invitation'
            ExternalUserState = if ($neverAccepted) { 'PendingAcceptance' } else { 'Accepted' }
            StateChanged      = if ($neverAccepted) { $invited } else { $accepted }
            AccountEnabled    = -not $disabled
            DisabledFrom      = if ($disabled) { $invited.AddDays($random.Next(30, 120)) } else { $null }
            LastSignIn        = $lastSignIn
            Inviter           = $inviters[$g % $inviters.Count]
            Dormant           = $dormant
        })
}

#endregion

#region Groups

$groupNames = @(
    'Project Northwind', 'Project Sunfish', 'Vendor Onboarding', 'Q3 Launch', 'Security Review'
    'Design Partners', 'Field Enablement', 'Contract Renewals', 'Data Migration', 'Support Escalations'
    'Partner Advisory', 'Release Readiness', 'Budget Planning', 'Legal Review', 'Analytics Guild'
    'Customer Research', 'Supplier Audit', 'Platform Upgrade', 'Brand Refresh', 'Training Program'
    'Incident Response', 'Localization', 'Procurement', 'Sustainability'
)

$groups = [System.Collections.Generic.List[object]]::new()
for ($i = 0; $i -lt $groupNames.Count; $i++) {
    $groups.Add([pscustomobject]@{
            Id          = New-DeterministicGuid
            DisplayName = $groupNames[$i]
            # Fourteen of the twenty-four are Teams; the rest are plain Microsoft 365 groups.
            IsTeam      = $i -lt 14
            Visibility  = if ($i % 3 -eq 0) { 'Public' } else { 'Private' }
        })
}

$membershipsByGuest = @{}
foreach ($guest in $guests) {
    $count = $random.Next(1, 4)
    $assigned = [System.Collections.Generic.List[object]]::new()
    # Everyone gets at least one Team and, for half the guests, a non-Team group too, so
    # the report can show both.
    $assigned.Add((Get-RandomItem @($groups | Where-Object IsTeam)))
    if ($count -gt 1) {
        $assigned.Add((Get-RandomItem @($groups | Where-Object { -not $_.IsTeam })))
    }
    if ($count -gt 2) {
        $extra = Get-RandomItem $groups.ToArray()
        if ($assigned.Id -notcontains $extra.Id) { $assigned.Add($extra) }
    }
    $membershipsByGuest[$guest.Id] = $assigned
}

#endregion

#region Write the files

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

# The generator rewrites the set from scratch; Export-AppendCsv would otherwise append
# to the files already committed.
foreach ($name in @('users.csv', 'guests.csv', 'guest-invitations.csv', 'guest-signins.csv', 'sharing-events.csv', 'guest-memberships.csv', 'run.log')) {
    $stale = Join-Path $OutputPath $name
    if (Test-Path -LiteralPath $stale) { Remove-Item -LiteralPath $stale -Force }
}

# users.csv and guests.csv: one snapshot block per snapshot date, holding the people who
# existed on that date.
$userRows = [System.Collections.Generic.List[object]]::new()
$guestRows = [System.Collections.Generic.List[object]]::new()
$membershipRows = [System.Collections.Generic.List[object]]::new()

foreach ($snapshot in $snapshotDates) {
    $runDate = $snapshot.ToString('yyyy-MM-dd')

    foreach ($member in $members) {
        $userRows.Add([pscustomobject]@{
                RunDate                  = $runDate
                Id                       = $member.Id
                DisplayName              = $member.DisplayName
                UserPrincipalName        = $member.UserPrincipalName
                Mail                     = $member.Mail
                UserType                 = 'Member'
                AccountEnabled           = $true
                CreatedDateTime          = ConvertTo-CsvTimestamp $member.CreatedDateTime
                Department               = $member.Department
                JobTitle                 = $member.JobTitle
                City                     = $member.City
                Country                  = $member.Country
                ManagerId                = $member.ManagerId
                ManagerUserPrincipalName = $member.ManagerUpn
            })
    }

    foreach ($guest in $guests) {
        if ($guest.CreatedDateTime -gt $snapshot) { continue }

        $enabled = $guest.AccountEnabled -or ($null -ne $guest.DisabledFrom -and $guest.DisabledFrom -gt $snapshot)
        $signInAsOf = if ($null -ne $guest.LastSignIn -and $guest.LastSignIn -le $snapshot) { $guest.LastSignIn } else { $null }

        $userRows.Add([pscustomobject]@{
                RunDate                  = $runDate
                Id                       = $guest.Id
                DisplayName              = $guest.DisplayName
                UserPrincipalName        = $guest.UserPrincipalName
                Mail                     = $guest.Mail
                UserType                 = 'Guest'
                AccountEnabled           = $enabled
                CreatedDateTime          = ConvertTo-CsvTimestamp $guest.CreatedDateTime
                Department               = ''
                JobTitle                 = ''
                City                     = ''
                Country                  = ''
                ManagerId                = ''
                ManagerUserPrincipalName = ''
            })

        $guestRows.Add([pscustomobject]@{
                RunDate                          = $runDate
                Id                               = $guest.Id
                DisplayName                      = $guest.DisplayName
                Mail                             = $guest.Mail
                UserPrincipalName                = $guest.UserPrincipalName
                ExternalDomain                   = $guest.ExternalDomain
                CreatedDateTime                  = ConvertTo-CsvTimestamp $guest.CreatedDateTime
                CreationType                     = $guest.CreationType
                ExternalUserState                = $guest.ExternalUserState
                ExternalUserStateChangeDateTime  = ConvertTo-CsvTimestamp $guest.StateChanged
                AccountEnabled                   = $enabled
                LastSignInDateTime               = ConvertTo-CsvTimestamp $signInAsOf
                LastNonInteractiveSignInDateTime = if ($null -eq $signInAsOf) { '' } else { ConvertTo-CsvTimestamp $signInAsOf.AddHours(-$random.Next(1, 72)) }
                LastSuccessfulSignInDateTime     = ConvertTo-CsvTimestamp $signInAsOf
            })

        foreach ($group in $membershipsByGuest[$guest.Id]) {
            $membershipRows.Add([pscustomobject]@{
                    RunDate          = $runDate
                    GuestId          = $guest.Id
                    GroupId          = $group.Id
                    GroupDisplayName = $group.DisplayName
                    IsTeam           = $group.IsTeam
                    Visibility       = $group.Visibility
                })
        }
    }
}

# guest-invitations.csv: an invitation for every guest, and a redemption for those who
# accepted.
$invitationRows = [System.Collections.Generic.List[object]]::new()
foreach ($guest in $guests) {
    $invitationRows.Add([pscustomobject]@{
            ActivityDateTime             = ConvertTo-CsvTimestamp $guest.CreatedDateTime
            Id                           = New-DeterministicGuid
            ActivityDisplayName          = 'Invite external user'
            Result                       = 'success'
            InitiatedByUserPrincipalName = $guest.Inviter.UserPrincipalName
            InitiatedByAppDisplayName    = 'Microsoft Entra admin center'
            TargetUserId                 = $guest.Id
            TargetUserPrincipalName      = $guest.UserPrincipalName
        })

    if ($guest.ExternalUserState -eq 'Accepted') {
        $invitationRows.Add([pscustomobject]@{
                ActivityDateTime             = ConvertTo-CsvTimestamp $guest.StateChanged
                Id                           = New-DeterministicGuid
                ActivityDisplayName          = 'Redeem external user invite'
                Result                       = 'success'
                InitiatedByUserPrincipalName = $guest.Mail
                InitiatedByAppDisplayName    = 'Microsoft Invitation Service'
                TargetUserId                 = $guest.Id
                TargetUserPrincipalName      = $guest.UserPrincipalName
            })
    }
    else {
        # An invitation that was re-sent and still not redeemed.
        $invitationRows.Add([pscustomobject]@{
                ActivityDateTime             = ConvertTo-CsvTimestamp $guest.CreatedDateTime.AddDays(14)
                Id                           = New-DeterministicGuid
                ActivityDisplayName          = 'Invite external user with reset invitation status'
                Result                       = 'success'
                InitiatedByUserPrincipalName = $guest.Inviter.UserPrincipalName
                InitiatedByAppDisplayName    = 'Microsoft Entra admin center'
                TargetUserId                 = $guest.Id
                TargetUserPrincipalName      = $guest.UserPrincipalName
            })
    }
}

# guest-signins.csv
$apps = @('Microsoft Teams', 'SharePoint Online Web Client Extensibility', 'Office 365 SharePoint Online', 'Microsoft Office', 'OneDrive SyncEngine')
$resources = @('Microsoft Teams Services', 'Office 365 SharePoint Online', 'Microsoft Graph', 'Windows Azure Active Directory')
$clientApps = @('Browser', 'Mobile Apps and Desktop clients')
$signInRows = [System.Collections.Generic.List[object]]::new()

foreach ($guest in $guests) {
    if ($guest.ExternalUserState -ne 'Accepted') { continue }

    $from = $guest.StateChanged
    $to = if ($guest.Dormant) { $guest.LastSignIn } else { $EndDate }
    if ($to -le $from) { continue }

    $signInCount = if ($guest.Dormant) { $random.Next(4, 15) } else { $random.Next(30, 110) }
    $place = Get-RandomItem $cities

    for ($s = 0; $s -lt $signInCount; $s++) {
        $when = New-EventTime -From $from -To $to
        # Roughly one sign-in in twelve fails, most often on a Conditional Access block.
        $failed = $random.Next(0, 12) -eq 0
        $signInRows.Add([pscustomobject]@{
                CreatedDateTime         = ConvertTo-CsvTimestamp $when
                Id                      = New-DeterministicGuid
                UserId                  = $guest.Id
                UserPrincipalName       = $guest.UserPrincipalName
                AppDisplayName          = Get-RandomItem $apps
                ResourceDisplayName     = Get-RandomItem $resources
                IpAddress               = ('203.0.113.{0}' -f $random.Next(1, 254))
                City                    = $place.City
                CountryOrRegion         = $place.Country
                ClientAppUsed           = Get-RandomItem $clientApps
                IsInteractive           = $random.Next(0, 4) -ne 0
                ErrorCode               = if ($failed) { '53003' } else { '0' }
                ConditionalAccessStatus = if ($failed) { 'failure' } else { Get-RandomItem @('success', 'notApplied') }
            })
    }
}

# sharing-events.csv
$sites = @(
    'https://example.sharepoint.com/sites/northwind'
    'https://example.sharepoint.com/sites/sunfish'
    'https://example.sharepoint.com/sites/vendors'
    'https://example.sharepoint.com/sites/launch'
    'https://example-my.sharepoint.com/personal/avery_abara_example_com'
)
$fileNames = @(
    'Statement of Work.docx', 'Pricing Model.xlsx', 'Architecture Review.pptx', 'Migration Plan.docx'
    'Test Results.xlsx', 'Security Questionnaire.docx', 'Rollout Schedule.xlsx', 'Meeting Notes.docx'
)
$sharingRows = [System.Collections.Generic.List[object]]::new()

function Add-SharingEvent {
    param(
        [Parameter(Mandatory)][string]$Operation,
        [Parameter(Mandatory)][datetime]$When,
        [Parameter(Mandatory)][string]$UserId,
        [AllowNull()][string]$TargetName,
        [AllowNull()][string]$TargetType
    )

    $site = Get-RandomItem $sites
    $file = Get-RandomItem $fileNames
    $workload = if ($site -like '*-my.sharepoint.com*') { 'OneDrive' } else { 'SharePoint' }

    $sharingRows.Add([pscustomobject]@{
            CreationTime          = ConvertTo-CsvTimestamp $When
            Id                    = New-DeterministicGuid
            Operation             = $Operation
            UserId                = $UserId
            Workload              = $workload
            SiteUrl               = ('{0}/' -f $site)
            ObjectId              = ('{0}/Shared Documents/{1}' -f $site, $file)
            SourceFileName        = $file
            TargetUserOrGroupName = $TargetName
            TargetUserOrGroupType = $TargetType
        })
}

foreach ($guest in $guests) {
    $sharer = Get-RandomItem $members.ToArray()
    $shareCount = $random.Next(6, 20)

    for ($e = 0; $e -lt $shareCount; $e++) {
        $when = New-EventTime -From $guest.CreatedDateTime -To $EndDate
        Add-SharingEvent -Operation 'SharingSet' -When $when -UserId $sharer.UserPrincipalName `
            -TargetName $guest.Mail -TargetType 'Guest'
    }

    # The invitation trail for a link sent to this guest.
    $invitedAt = New-EventTime -From $guest.CreatedDateTime -To $EndDate.AddDays(-10)
    Add-SharingEvent -Operation 'SharingInvitationCreated' -When $invitedAt -UserId $sharer.UserPrincipalName `
        -TargetName $guest.Mail -TargetType 'Guest'

    if ($guest.ExternalUserState -eq 'Accepted') {
        Add-SharingEvent -Operation 'SharingInvitationAccepted' -When $invitedAt.AddHours($random.Next(1, 72)) `
            -UserId $guest.Mail -TargetName $guest.Mail -TargetType 'Guest'

        # Secure (company-specific) links, which is how most guest sharing actually lands.
        Add-SharingEvent -Operation 'SecureLinkCreated' -When $invitedAt.AddHours(2) -UserId $sharer.UserPrincipalName `
            -TargetName $guest.Mail -TargetType 'Guest'
        Add-SharingEvent -Operation 'AddedToSecureLink' -When $invitedAt.AddHours(3) -UserId $sharer.UserPrincipalName `
            -TargetName $guest.Mail -TargetType 'Guest'
        foreach ($use in 1..($random.Next(1, 6))) {
            Add-SharingEvent -Operation 'SecureLinkUsed' -When (New-EventTime -From $invitedAt -To $EndDate) `
                -UserId $guest.Mail -TargetName $guest.Mail -TargetType 'Guest'
        }
    }
    else {
        Add-SharingEvent -Operation 'SharingInvitationRevoked' -When $invitedAt.AddDays($random.Next(20, 60)) `
            -UserId $sharer.UserPrincipalName -TargetName $guest.Mail -TargetType 'Guest'
    }

    # Access taken away again.
    if ($random.Next(0, 3) -eq 0) {
        Add-SharingEvent -Operation 'SharingRevoked' -When (New-EventTime -From $invitedAt -To $EndDate) `
            -UserId $sharer.UserPrincipalName -TargetName $guest.Mail -TargetType 'Guest'
    }
}

# Anonymous ("anyone") links, which have no named target.
for ($a = 0; $a -lt 120; $a++) {
    $sharer = Get-RandomItem $members.ToArray()
    $created = New-EventTime -From $eventStart -To $EndDate.AddDays(-3)
    Add-SharingEvent -Operation 'AnonymousLinkCreated' -When $created -UserId $sharer.UserPrincipalName `
        -TargetName '' -TargetType ''

    foreach ($use in 1..($random.Next(1, 9))) {
        Add-SharingEvent -Operation 'AnonymousLinkUsed' -When (New-EventTime -From $created -To $EndDate) `
            -UserId ('urn:spo:anon#{0}' -f $random.Next(100000, 999999)) -TargetName '' -TargetType ''
    }

    if ($random.Next(0, 3) -eq 0) {
        Add-SharingEvent -Operation 'AnonymousLinkRemoved' -When (New-EventTime -From $created -To $EndDate) `
            -UserId $sharer.UserPrincipalName -TargetName '' -TargetType ''
    }
}

# Events are written in the order a collector would see them: oldest first.
$sortedInvitations = @($invitationRows | Sort-Object ActivityDateTime, Id)
$sortedSignIns = @($signInRows | Sort-Object CreatedDateTime, Id)
$sortedSharing = @($sharingRows | Sort-Object CreationTime, Id)

Export-AppendCsv -Path (Join-Path $OutputPath 'users.csv') -Rows $userRows.ToArray() -Column (Get-EntraUserCsvColumn)
Export-AppendCsv -Path (Join-Path $OutputPath 'guests.csv') -Rows $guestRows.ToArray() -Column $schema.Guests
Export-AppendCsv -Path (Join-Path $OutputPath 'guest-memberships.csv') -Rows $membershipRows.ToArray() -Column $schema.GuestMemberships
Export-AppendCsv -Path (Join-Path $OutputPath 'guest-invitations.csv') -Rows $sortedInvitations -Column $schema.GuestInvitations -KeyColumn 'Id'
Export-AppendCsv -Path (Join-Path $OutputPath 'guest-signins.csv') -Rows $sortedSignIns -Column $schema.GuestSignIns -KeyColumn 'Id'
Export-AppendCsv -Path (Join-Path $OutputPath 'sharing-events.csv') -Rows $sortedSharing -Column $schema.SharingEvents -KeyColumn 'Id'

[pscustomobject]@{
    OutputPath     = $OutputPath
    Members        = $members.Count
    Guests         = $guests.Count
    PartnerDomains = $partnerDomains.Count
    Snapshots      = $snapshotDates.Count
    UserRows       = $userRows.Count
    GuestRows      = $guestRows.Count
    MembershipRows = $membershipRows.Count
    Invitations    = $sortedInvitations.Count
    SignIns        = $sortedSignIns.Count
    SharingEvents  = $sortedSharing.Count
}

#endregion
