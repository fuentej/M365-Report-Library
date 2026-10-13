#Requires -Version 7.0

<#
    .SYNOPSIS
        Regenerates samples/: fake CSVs with the exact columns each collector writes.

    .DESCRIPTION
        Every name, address and Id is invented; addresses use example.com and IP addresses
        are from the documentation range 203.0.113.0/24 (RFC 5737). The column order comes
        from collectors/EntraActivitySchema.psd1, so a sample cannot drift from its
        collector. No sign-in happens and no tenant is called.

        None of the five sources is marked NotAvailable in any cloud in the contract
        (docs/candidates/entra-activity.md), so there is no header-only sample.

    .EXAMPLE
        ./New-SampleData.ps1
#>
[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path $PSScriptRoot 'samples')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot '../../shared/M365ReportLibrary.psm1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'collectors/EntraActivitySchema.psd1')

if (Test-Path -LiteralPath $OutputPath) {
    Get-ChildItem -LiteralPath $OutputPath -Filter '*.csv' | Remove-Item -Force
}

function New-SignInSample {
    param([string]$Id, [string]$At, [string]$UserId, [string]$Upn, [string]$App, [string]$Resource,
        [string]$Ip, [string]$City, [string]$Country, [string]$Client, [string]$Os, [string]$Browser,
        [bool]$Interactive, [string]$EventTypes, [string]$ErrorCode, [string]$Reason, [string]$Details,
        [string]$CaStatus, [string]$Managed = 'True')

    [pscustomobject]@{
        CreatedDateTime = $At; Id = $Id; UserId = $UserId; UserPrincipalName = $Upn
        AppId = '00000002-0000-0ff1-ce00-000000000000'; AppDisplayName = $App; ResourceDisplayName = $Resource
        IpAddress = $Ip; City = $City; State = ''; CountryOrRegion = $Country; ClientAppUsed = $Client
        DeviceOperatingSystem = $Os; DeviceBrowser = $Browser; DeviceIsCompliant = $Managed; DeviceIsManaged = $Managed
        IsInteractive = $Interactive; SignInEventTypes = $EventTypes; ErrorCode = $ErrorCode
        FailureReason = $Reason; AdditionalDetails = $Details; ConditionalAccessStatus = $CaStatus
        RiskDetail = 'none'; RiskLevelAggregated = 'none'; RiskLevelDuringSignIn = 'none'; RiskState = 'none'
    }
}

# Volume and shape. One seeded random drives everything below, so the same seed gives the same
# files. The window is the 30 days 2026-09-01 to 2026-09-30 (UTC); the people work in US Eastern
# time, UTC-4 all month, so local business hours 08:00-17:59 are UTC 12:00-21:59.
$random = [System.Random]::new(20260930)
$windowStart = [datetime]::SpecifyKind([datetime]'2026-09-01', [DateTimeKind]::Utc)
$utcOffsetHours = 4

$firstNames = 'Avery', 'Casey', 'Devon', 'Emerson', 'Finley', 'Gray', 'Harper', 'Indigo', 'Jordan', 'Kai', 'Logan', 'Morgan', 'Nico', 'Oakley', 'Parker', 'Quinn', 'Reese', 'Sage', 'Tatum', 'Wren'
$lastNames = 'Abara', 'Chaudhry', 'Dube', 'Eriksen', 'Fontaine', 'Gupta', 'Hollis', 'Iyer', 'Jansen', 'Kowalski', 'Lindqvist', 'Moreau'
$cities = 'Tampa', 'Orlando', 'Miami'
$people = foreach ($index in 0..59) {
    $first = $firstNames[$index % $firstNames.Count]
    $last = $lastNames[$index % $lastNames.Count]
    [pscustomobject]@{
        UserId = if ($index -lt 3) { 'aaaaaaaa-0000-4000-8000-{0:D12}' -f ($index + 1) } else { 'aaaaaaaa-0000-4000-8000-{0:D12}' -f (1000 + $index) }
        Upn = ('{0}.{1}@example.com' -f $first, $last).ToLowerInvariant()
        Ip = "203.0.113.$(10 + $index)"
        City = $cities[$index % $cities.Count]
        Os = @('Windows 11', 'Windows 11', 'macOS 14', 'Windows 11', 'iOS 17')[$index % 5]
        Browser = @('Edge 128.0', 'Chrome 128.0', 'Safari 17.6', 'Edge 128.0', 'Teams')[$index % 5]
    }
}

$script:signInNumber = 100
function Format-Utc {
    param([datetime]$At)
    $At.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", [cultureinfo]::InvariantCulture)
}
function New-GeneratedSignIn {
    param($Person, [datetime]$At, [bool]$Interactive, [string]$App = 'Office 365 Exchange Online', [string]$Resource = 'Office 365 Exchange Online',
        [string]$Client = 'Browser', [string]$ErrorCode = '0', [string]$Ip = '', [string]$City = '', [string]$Os = '', [string]$Browser = '', [string]$Managed = 'True')

    $script:signInNumber++
    $failed = $ErrorCode -ne '0'
    $reason = if ($failed) { 'Invalid username or password or Invalid on-premise username or password.' } else { '' }
    $details = if ($failed) { 'Error validating credentials due to invalid username or password.' } else { '' }
    $eventType = if ($Interactive) { 'interactiveUser' } else { 'nonInteractiveUser' }
    New-SignInSample ('a1000000-0000-4000-8000-{0:D12}' -f $script:signInNumber) (Format-Utc $At) $Person.UserId $Person.Upn $App $Resource `
        $(if ($Ip) { $Ip } else { $Person.Ip }) $(if ($City) { $City } else { $Person.City }) 'US' $Client `
        $(if ($Os) { $Os } else { $Person.Os }) $(if ($Browser) { $Browser } else { $Person.Browser }) $Interactive $eventType $ErrorCode $reason $details 'notApplied' $Managed
}

# A weekday working sign-in time: most land 08:00-17:59 local on a weekday, a few in the
# early morning or evening, and a few on a weekend.
function Get-WorkingTime {
    $roll = $random.NextDouble()
    do {
        $date = $windowStart.AddDays($random.Next(0, 30))
        $weekend = $date.DayOfWeek -in 'Saturday', 'Sunday'
    } while ($weekend -ne ($roll -ge 0.96))
    $localHour = if ($roll -lt 0.90 -or $roll -ge 0.96) { $random.Next(8, 18) }
    elseif ($random.NextDouble() -lt 0.5) { $random.Next(6, 8) }
    else { $random.Next(18, 21) }
    $date.AddHours($localHour + $utcOffsetHours).AddMinutes($random.Next(0, 60)).AddSeconds($random.Next(0, 60))
}

$interactiveApps = foreach ($entry in @(
        @{ App = 'Office 365 Exchange Online'; Resource = 'Office 365 Exchange Online'; Client = 'Browser'; Weight = 4 }
        @{ App = 'Microsoft Teams'; Resource = 'Microsoft Teams'; Client = 'Mobile Apps and Desktop clients'; Weight = 4 }
        @{ App = 'SharePoint Online Web Client Extensibility'; Resource = 'Office 365 SharePoint Online'; Client = 'Browser'; Weight = 3 }
        @{ App = 'Azure Portal'; Resource = 'Windows Azure Service Management API'; Client = 'Browser'; Weight = 1 }
    )) {
    1..$entry.Weight | ForEach-Object { $entry }
}
$legacyClients = 'IMAP4', 'POP3', 'Authenticated SMTP', 'Exchange ActiveSync', 'Other clients'

# Hand-written rows: the Conditional Access and error-code cases the CA sample points at.
$interactiveRows = [System.Collections.Generic.List[object]]::new()
$interactiveRows.Add((New-SignInSample 'a1000000-0000-4000-8000-000000000001' '2026-09-10T14:05:00Z' 'aaaaaaaa-0000-4000-8000-000000000001' 'avery.abara@example.com' 'Office 365 Exchange Online' 'Office 365 Exchange Online' '203.0.113.10' 'Tampa' 'US' 'Browser' 'Windows 11' 'Edge 128.0' $true 'interactiveUser' '0' '' '' 'success'))
$interactiveRows.Add((New-SignInSample 'a1000000-0000-4000-8000-000000000002' '2026-09-10T15:30:00Z' 'aaaaaaaa-0000-4000-8000-000000000002' 'casey.chaudhry@example.com' 'Microsoft Teams' 'Microsoft Teams' '203.0.113.11' 'Orlando' 'US' 'Mobile Apps and Desktop clients' 'iOS 17' 'Teams' $true 'interactiveUser' '50126' 'Invalid username or password or Invalid on-premise username or password.' 'Error validating credentials due to invalid username or password.' 'notApplied'))
$interactiveRows.Add((New-SignInSample 'a1000000-0000-4000-8000-000000000003' '2026-09-11T16:15:00Z' '' '' 'Azure Portal' 'Windows Azure Service Management API' '203.0.113.12' '' 'US' 'Browser' 'macOS 14' 'Safari 17.6' $true 'interactiveUser' '50058' 'Session information is not sufficient for single-sign-on.' 'The user is not signed in; no user name was returned.' 'notApplied'))
$interactiveRows.Add((New-SignInSample 'a1000000-0000-4000-8000-000000000004' '2026-09-11T17:40:00Z' 'aaaaaaaa-0000-4000-8000-000000000003' 'devon.dube@example.com' 'SharePoint Online Web Client Extensibility' 'Office 365 SharePoint Online' '203.0.113.13' 'Miami' 'US' 'Browser' 'Windows 11' 'Chrome 128.0' $true 'interactiveUser' '53003' 'Access has been blocked by Conditional Access policies.' 'Access has been blocked by Conditional Access policies. The access policy does not allow token issuance.' 'failure'))

# Working-day sign-ins, a handful of failed passwords (50126, 2%).
foreach ($n in 1..2000) {
    $person = $people[$random.Next(0, $people.Count)]
    $app = $interactiveApps[$random.Next(0, $interactiveApps.Count)]
    $code = if ($random.NextDouble() -lt 0.02) { '50126' } else { '0' }
    $interactiveRows.Add((New-GeneratedSignIn -Person $person -At (Get-WorkingTime) -Interactive $true -App $app.App -Resource $app.Resource -Client $app.Client -ErrorCode $code))
}

# Odd hours: 12 successful sign-ins between 01:00 and 04:59 local, from four people.
foreach ($n in 1..12) {
    $person = $people[20 + ($n % 4)]
    $at = $windowStart.AddDays($random.Next(0, 30)).AddHours($random.Next(1, 5) + $utcOffsetHours).AddMinutes($random.Next(0, 60)).AddSeconds($random.Next(0, 60))
    $interactiveRows.Add((New-GeneratedSignIn -Person $person -At $at -Interactive $true))
}

# Legacy protocols: interactive, a few people on IMAP4, POP3, SMTP AUTH, EAS and unknown clients.
foreach ($n in 1..24) {
    $person = $people[30 + ($n % 6)]
    $client = $legacyClients[$n % $legacyClients.Count]
    $interactiveRows.Add((New-GeneratedSignIn -Person $person -At (Get-WorkingTime) -Interactive $true -Client $client -Os '' -Browser ''))
}

# One password spray: a single address tries one wrong password against 55 accounts in 18
# minutes, in the small hours of Thursday 2026-09-17 local time. Every attempt fails with 50126.
$sprayIp = '203.0.113.250'
$sprayStart = $windowStart.AddDays(16).AddHours(3 + $utcOffsetHours).AddMinutes(12)
foreach ($n in 0..54) {
    $interactiveRows.Add((New-GeneratedSignIn -Person $people[$n] -At $sprayStart.AddSeconds($n * 20 + $random.Next(0, 10)) -Interactive $true -Client 'Browser' `
            -ErrorCode '50126' -Ip $sprayIp -City 'Amsterdam' -Os 'Linux' -Browser 'Unknown' -Managed 'False'))
}
$interactive = @($interactiveRows | Sort-Object CreatedDateTime, Id)
Export-AppendCsv -Path (Join-Path $OutputPath 'signins-interactive.csv') -Rows $interactive -Column $schema.SignIns -KeyColumn 'Id'

$nonInteractiveRows = [System.Collections.Generic.List[object]]::new()
$nonInteractiveRows.Add((New-SignInSample 'b2000000-0000-4000-8000-000000000001' '2026-09-10T14:20:00Z' 'aaaaaaaa-0000-4000-8000-000000000001' 'avery.abara@example.com' 'Microsoft Office' 'Microsoft Graph' '203.0.113.10' 'Tampa' 'US' 'Mobile Apps and Desktop clients' 'Windows 11' '' $false 'nonInteractiveUser' '0' '' '' 'notApplied'))
$nonInteractiveRows.Add((New-SignInSample 'b2000000-0000-4000-8000-000000000002' '2026-09-11T18:00:00Z' 'aaaaaaaa-0000-4000-8000-000000000002' 'casey.chaudhry@example.com' 'Outlook Mobile' 'Office 365 Exchange Online' '203.0.113.11' 'Orlando' 'US' 'Mobile Apps and Desktop clients' 'iOS 17' '' $false 'nonInteractiveUser' '50173' 'The provided grant has expired due to it being revoked.' '' 'notApplied'))
$script:signInNumber = 100
foreach ($n in 1..360) {
    $person = $people[$random.Next(0, $people.Count)]
    $app = if ($random.Next(0, 2) -eq 0) { 'Microsoft Office' } else { 'Outlook Mobile' }
    $resource = if ($app -eq 'Microsoft Office') { 'Microsoft Graph' } else { 'Office 365 Exchange Online' }
    $row = New-GeneratedSignIn -Person $person -At (Get-WorkingTime) -Interactive $false -App $app -Resource $resource -Client 'Mobile Apps and Desktop clients' -Browser ''
    $row.Id = 'b2000000-0000-4000-8000-{0:D12}' -f $script:signInNumber
    $nonInteractiveRows.Add($row)
}
foreach ($n in 1..40) {
    $person = $people[30 + ($n % 6)]
    $row = New-GeneratedSignIn -Person $person -At (Get-WorkingTime) -Interactive $false -Client $legacyClients[$n % $legacyClients.Count] -Os '' -Browser ''
    $row.Id = 'b2000000-0000-4000-8000-{0:D12}' -f $script:signInNumber
    $nonInteractiveRows.Add($row)
}
$nonInteractive = @($nonInteractiveRows | Sort-Object CreatedDateTime, Id)
Export-AppendCsv -Path (Join-Path $OutputPath 'signins-noninteractive.csv') -Rows $nonInteractive -Column $schema.SignIns -KeyColumn 'Id'

function New-CaSample {
    param([string]$SignInId, [string]$At, [string]$UserId, [bool]$Interactive, [string]$Status,
        [string]$PolicyId, [string]$Name, [string]$Result, [string]$Grant, [string]$Session, [bool]$Readable)
    [pscustomobject]@{
        CreatedDateTime = $At; SignInId = $SignInId; UserId = $UserId; IsInteractive = $Interactive
        ConditionalAccessStatus = $Status; PolicyId = $PolicyId; PolicyDisplayName = $Name; PolicyResult = $Result
        EnforcedGrantControls = $Grant; EnforcedSessionControls = $Session; PolicyDetailReadable = $Readable
    }
}
$ca = @(
    New-CaSample 'a1000000-0000-4000-8000-000000000001' '2026-09-10T14:05:00Z' 'aaaaaaaa-0000-4000-8000-000000000001' $true 'success' 'cccccccc-0000-4000-8000-000000000001' 'Require MFA for all users' 'success' 'Mfa' '' $true
    New-CaSample 'a1000000-0000-4000-8000-000000000004' '2026-09-11T17:40:00Z' 'aaaaaaaa-0000-4000-8000-000000000003' $true 'failure' 'cccccccc-0000-4000-8000-000000000002' 'Block unmanaged devices' 'failure' 'Block' '' $true
    New-CaSample 'a1000000-0000-4000-8000-000000000004' '2026-09-11T17:40:00Z' 'aaaaaaaa-0000-4000-8000-000000000003' $true 'failure' 'cccccccc-0000-4000-8000-000000000003' 'Report-only: require compliant device' 'reportOnlySuccess' '' '' $true
    New-CaSample 'a1000000-0000-4000-8000-000000000002' '2026-09-10T15:30:00Z' 'aaaaaaaa-0000-4000-8000-000000000002' $true 'notApplied' '' '' '' '' '' $false
)
Export-AppendCsv -Path (Join-Path $OutputPath 'signin-conditional-access.csv') -Rows $ca -Column $schema.SignInConditionalAccess -KeyColumn 'SignInId', 'PolicyId', 'PolicyDisplayName'

function New-AuditSample {
    param([string]$Id, [string]$At, [string]$Activity, [string]$Category, [string]$Operation, [string]$Result,
        [string]$Reason, [string]$Service, [string]$ByUserId, [string]$ByUpn, [string]$ByAppId, [string]$ByApp,
        [string]$Types, [string]$Ids, [string]$Names, [string]$Upns, [string]$Modified)
    [pscustomobject]@{
        ActivityDateTime = $At; Id = $Id; ActivityDisplayName = $Activity; Category = $Category
        OperationType = $Operation; Result = $Result; ResultReason = $Reason; LoggedByService = $Service
        CorrelationId = 'dddddddd-0000-4000-8000-0000000000' + $Id.Substring($Id.Length - 2)
        InitiatedByUserId = $ByUserId; InitiatedByUserPrincipalName = $ByUpn; InitiatedByAppId = $ByAppId
        InitiatedByAppDisplayName = $ByApp; TargetResourceTypes = $Types; TargetResourceIds = $Ids
        TargetResourceDisplayNames = $Names; TargetResourceUserPrincipalNames = $Upns; ModifiedProperties = $Modified
    }
}
$audits = @(
    New-AuditSample 'Directory_eeeeeeee-0000-4000-8000-000000000001_01' '2026-09-10T09:00:00Z' 'Add user' 'UserManagement' 'Add' 'success' '' 'Core Directory' 'aaaaaaaa-0000-4000-8000-000000000009' 'admin.alvarez@example.com' '' '' 'User' 'aaaaaaaa-0000-4000-8000-000000000010' 'Riley Rosen' 'riley.rosen@example.com' '[{"displayName":"AccountEnabled","oldValue":"","newValue":"[true]"}]'
    New-AuditSample 'Directory_eeeeeeee-0000-4000-8000-000000000002_02' '2026-09-10T09:30:00Z' 'Update Conditional Access policy' 'Policy' 'Update' 'success' '' 'Conditional Access' 'aaaaaaaa-0000-4000-8000-000000000009' 'admin.alvarez@example.com' '' '' 'Policy' 'cccccccc-0000-4000-8000-000000000001' 'Require MFA for all users' '' '[{"displayName":"State","oldValue":"[\"enabledForReportingButNotEnforced\"]","newValue":"[\"enabled\"]"}]'
    New-AuditSample 'SSGM_b662f17a-4e4d-4e1c-9248-cdec180024b2_MCDC4_88453290' '2026-09-11T12:00:00Z' 'Add member to group' 'GroupManagement' 'Assign' 'failure' 'Insufficient privileges to complete the operation.' 'Core Directory' '' '' 'ffffffff-0000-4000-8000-000000000001' 'Provisioning Sample App' 'Group;User' 'aaaaaaaa-0000-4000-8000-000000000020;aaaaaaaa-0000-4000-8000-000000000010' 'Sales team;Riley Rosen' ';riley.rosen@example.com' ''
    New-AuditSample 'Directory_eeeeeeee-0000-4000-8000-000000000004_04' '2026-09-12T08:45:00Z' 'Add member to role' 'RoleManagement' 'Assign' 'success' '' 'Core Directory' 'aaaaaaaa-0000-4000-8000-000000000009' 'admin.alvarez@example.com' '' '' 'Role;User' 'bbbbbbbb-0000-4000-8000-000000000001;aaaaaaaa-0000-4000-8000-000000000010' 'Helpdesk Administrator;Riley Rosen' ';riley.rosen@example.com' ''
)
Export-AppendCsv -Path (Join-Path $OutputPath 'directory-audits.csv') -Rows $audits -Column $schema.DirectoryAudits -KeyColumn 'Id'

$retention = foreach ($environment in 'Commercial', 'GCC', 'GCCHigh') {
    foreach ($level in $schema.RetentionLevels) {
        [pscustomobject]@{
            RunDate = '2026-09-30'; Environment = $environment; LicenseLevel = $level.LicenseLevel
            SignInRetentionDays = $level.SignInRetentionDays; AuditRetentionDays = $level.AuditRetentionDays
            RiskySignInRetentionDays = $level.RiskySignInRetentionDays; RetentionStatus = 'UNVERIFIED'
            Reference = $schema.SourceAvailability.RetentionReference[$environment].Reference
        }
    }
}
Export-AppendCsv -Path (Join-Path $OutputPath 'retention-reference.csv') -Rows $retention -Column $schema.RetentionReference -KeyColumn 'RunDate', 'Environment', 'LicenseLevel'
