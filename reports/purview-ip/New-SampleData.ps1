#Requires -Version 7.0

<#
    .SYNOPSIS
        Generates the fake Purview information protection data in ./samples.

    .DESCRIPTION
        Lets the report, and anyone reading this repository, work against
        realistic shapes without a tenant. Nothing here comes from a real
        directory: every address is under example.com, and every label,
        policy and rule is an invented name.

        The generator is deterministic. The same -Seed and -EndDate always
        produce the same files, so a regenerated sample set shows up in a
        diff only when this script changes.

        Every file is written through Export-AppendCsv with the column list
        the collectors use, so a sample file cannot drift from its
        collector's output.

    .PARAMETER EndDate
        The "now" the sample set is generated around. Fixed by default to
        keep the committed files stable.

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
. (Join-Path $PSScriptRoot 'collectors/PurviewIpHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'collectors/PurviewIpSchema.psd1')

$EndDate = $EndDate.ToUniversalTime()
$tenantDomain = 'example.com'

$departments = @('Engineering', 'Sales', 'Marketing', 'Finance', 'Legal', 'Operations')
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
            A GUID drawn from the seeded generator, so re-running produces the
            same ids.
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
    param([datetime]$From, [datetime]$To = $EndDate)
    $span = ($To - $From).TotalMinutes
    if ($span -le 1) { return $From }
    return $From.AddMinutes($random.Next(0, [int]$span)).AddSeconds($random.Next(0, 60))
}

$firstNames = @(
    'Avery', 'Blair', 'Casey', 'Drew', 'Emery', 'Finley', 'Gray', 'Harper', 'Indigo', 'Jordan'
    'Kai', 'Logan', 'Marlowe', 'Noor', 'Oakley', 'Parker', 'Quinn', 'Reese', 'Sage', 'Tatum'
)
$lastNames = @(
    'Abara', 'Bergstrom', 'Chaudhry', 'Delacroix', 'Eriksen', 'Fontaine', 'Gallagher', 'Haddad'
    'Ibarra', 'Jovanovic', 'Kowalski', 'Lindqvist', 'Moreau', 'Nakamura', 'Okonkwo', 'Petrov'
    'Quintero', 'Rasmussen', 'Silva', 'Takahashi'
)

#region People

$memberCount = 60
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
            CreatedDateTime   = $EndDate.AddDays(-$random.Next(60, 1500))
            ManagerId         = $null
            ManagerUpn        = $null
            Level             = 3
        })
}

$executive = $members[0]
$executive.Level = 0
$executive.JobTitle = 'Chief Executive'
$executive.Department = 'Operations'

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

$managers = @{}
foreach ($department in $departments) { $managers[$department] = [System.Collections.Generic.List[object]]::new() }
for ($m = 0; $m -lt ($departments.Count * 2); $m++) {
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

for (; $index -lt $members.Count; $index++) {
    $person = $members[$index]
    $department = $person.Department
    $manager = Get-RandomItem $managers[$department].ToArray()
    $person.Level = 3
    $person.JobTitle = "$department Specialist"
    $person.ManagerId = $manager.Id
    $person.ManagerUpn = $manager.UserPrincipalName
}

#endregion

#region Policy configuration

# Sensitivity labels: 7, two of them sublabels of Confidential.
$labels = [ordered]@{
    Public                = [pscustomobject]@{ Id = New-DeterministicGuid; Name = 'Public'; Priority = 0; ParentId = ''; Encrypted = $false; Default = $false }
    General               = [pscustomobject]@{ Id = New-DeterministicGuid; Name = 'General'; Priority = 1; ParentId = ''; Encrypted = $false; Default = $true }
    Confidential          = [pscustomobject]@{ Id = New-DeterministicGuid; Name = 'Confidential'; Priority = 2; ParentId = ''; Encrypted = $false; Default = $false }
    ConfidentialInternal  = [pscustomobject]@{ Id = New-DeterministicGuid; Name = 'Confidential-Internal'; Priority = 3; ParentId = ''; Encrypted = $true; Default = $false }
    ConfidentialPartner   = [pscustomobject]@{ Id = New-DeterministicGuid; Name = 'Confidential-Partner'; Priority = 4; ParentId = ''; Encrypted = $true; Default = $false }
    HighlyConfidential    = [pscustomobject]@{ Id = New-DeterministicGuid; Name = 'Highly Confidential'; Priority = 5; ParentId = ''; Encrypted = $true; Default = $false }
    Restricted            = [pscustomobject]@{ Id = New-DeterministicGuid; Name = 'Restricted'; Priority = 6; ParentId = ''; Encrypted = $true; Default = $false }
}
$labels.ConfidentialInternal.ParentId = $labels.Confidential.Id
$labels.ConfidentialPartner.ParentId = $labels.Confidential.Id

$retentionLabels = [ordered]@{
    SevenYear = [pscustomobject]@{ Id = New-DeterministicGuid; Name = '7 Year Retention'; Action = 'Retain'; Duration = 2555; Type = 'ModifiedDate'; Record = $false }
    Permanent = [pscustomobject]@{ Id = New-DeterministicGuid; Name = 'Permanent Record'; Action = 'Retain'; Duration = -1; Type = 'CreationDate'; Record = $true }
    Delete90  = [pscustomobject]@{ Id = New-DeterministicGuid; Name = 'Delete after 90 days'; Action = 'Delete'; Duration = 90; Type = 'CreationDate'; Record = $false }
}

$dlpPolicies = [ordered]@{
    GlobalDlp  = [pscustomobject]@{ Id = New-DeterministicGuid; Name = 'Global DLP'; Copilot = $false }
    EndpointDlp = [pscustomobject]@{ Id = New-DeterministicGuid; Name = 'Endpoint DLP'; Copilot = $false }
    CopilotDlp = [pscustomobject]@{ Id = New-DeterministicGuid; Name = 'Copilot data protection'; Copilot = $true }
    TeamsDlp   = [pscustomobject]@{ Id = New-DeterministicGuid; Name = 'Teams DLP'; Copilot = $false }
}

$dlpRules = [ordered]@{
    CreditCardGlobal = [pscustomobject]@{ Id = New-DeterministicGuid; Name = 'Block credit card numbers'; Policy = $dlpPolicies.GlobalDlp; Sit = 'Credit Card Number' }
    SsnGlobal        = [pscustomobject]@{ Id = New-DeterministicGuid; Name = 'Notify on SSN'; Policy = $dlpPolicies.GlobalDlp; Sit = 'U.S. Social Security Number (SSN)' }
    RemovableMedia   = [pscustomobject]@{ Id = New-DeterministicGuid; Name = 'Block removable media copy'; Policy = $dlpPolicies.EndpointDlp; Sit = 'Credit Card Number' }
    CopilotSensitive = [pscustomobject]@{ Id = New-DeterministicGuid; Name = 'Block Copilot on Restricted label'; Policy = $dlpPolicies.CopilotDlp; Sit = '' }
    TeamsChatSsn     = [pscustomobject]@{ Id = New-DeterministicGuid; Name = 'Notify on SSN in chat'; Policy = $dlpPolicies.TeamsDlp; Sit = 'U.S. Social Security Number (SSN)' }
    PassportGlobal   = [pscustomobject]@{ Id = New-DeterministicGuid; Name = 'Block passport numbers'; Policy = $dlpPolicies.GlobalDlp; Sit = 'U.S. / U.K. Passport Number' }
}

$labelPolicies = [ordered]@{
    AllStaff       = [pscustomobject]@{ Id = New-DeterministicGuid; Name = 'All staff'; Labels = @($labels.Public, $labels.General, $labels.Confidential, $labels.ConfidentialInternal) }
    FinanceLegal   = [pscustomobject]@{ Id = New-DeterministicGuid; Name = 'Finance and Legal'; Labels = @($labels.Confidential, $labels.ConfidentialPartner, $labels.HighlyConfidential, $labels.Restricted) }
}

$autoLabelPolicies = [ordered]@{
    AutoSsn        = [pscustomobject]@{ Id = New-DeterministicGuid; Name = 'Auto-label SSNs'; Label = $labels.Confidential }
    AutoCreditCard = [pscustomobject]@{ Id = New-DeterministicGuid; Name = 'Auto-label credit cards'; Label = $labels.Restricted }
}

$retentionPolicies = [ordered]@{
    FinancialRetention = [pscustomobject]@{ Id = New-DeterministicGuid; Name = '7 year financial retention' }
    LegalHold          = [pscustomobject]@{ Id = New-DeterministicGuid; Name = 'Litigation hold - Legal' }
}

function New-PolicySnapshotRow {
    param([string]$RunDate, [string]$ObjectType, [hashtable]$Values)
    $row = [ordered]@{ RunDate = $RunDate; ObjectType = $ObjectType }
    foreach ($column in ($schema.Policies | Where-Object { $_ -ne 'RunDate' -and $_ -ne 'ObjectType' })) {
        $row[$column] = if ($Values.ContainsKey($column)) { $Values[$column] } else { '' }
    }
    [pscustomobject]$row
}

# Two snapshots a month apart: configuration rarely changes, so both carry
# the same policies, which is exactly what "keyed on RunDate plus ObjectId"
# is for - re-running does not duplicate a row, and a real change would only
# ever touch the newer snapshot.
$policySnapshotDates = @($EndDate.AddMonths(-1), $EndDate) | ForEach-Object { $_.ToString('yyyy-MM-dd') }
$policyRows = [System.Collections.Generic.List[object]]::new()

foreach ($runDate in $policySnapshotDates) {
    foreach ($label in $labels.Values) {
        $parentName = if ($label.ParentId) { ($labels.Values | Where-Object Id -EQ $label.ParentId).Name } else { '' }
        $policyRows.Add((New-PolicySnapshotRow -RunDate $runDate -ObjectType 'SensitivityLabel' -Values @{
                    ObjectId          = $label.Id
                    Name              = $label.Name
                    DisplayName       = $label.Name
                    ParentId          = $label.ParentId
                    ParentName        = $parentName
                    Priority          = $label.Priority
                    Enabled           = $true
                    Workload          = 'Exchange;SharePoint;OneDrive;Teams'
                    IsDefaultLabel    = $label.Default
                    EncryptionEnabled = $label.Encrypted
                    Comment           = "Sample sensitivity label $($label.Name)."
                    WhenCreatedUtc    = ConvertTo-CsvTimestamp $EndDate.AddMonths(-14)
                    WhenChangedUtc    = ConvertTo-CsvTimestamp $EndDate.AddMonths(-1)
                }))
    }

    foreach ($policy in $labelPolicies.Values) {
        $labelIds = @($policy.Labels | ForEach-Object { $_.Id })
        $labelNames = @($policy.Labels | ForEach-Object { $_.Name })
        $policyRows.Add((New-PolicySnapshotRow -RunDate $runDate -ObjectType 'LabelPolicy' -Values @{
                    ObjectId       = $policy.Id
                    Name           = $policy.Name
                    Mode           = 'Enable'
                    Enabled        = $true
                    Workload       = 'Exchange;SharePoint;OneDrive;Teams'
                    Locations      = 'All'
                    LabelIds       = ($labelIds -join ';')
                    LabelNames     = ($labelNames -join ';')
                    Comment        = "Sample label policy $($policy.Name)."
                    WhenCreatedUtc = ConvertTo-CsvTimestamp $EndDate.AddMonths(-13)
                    WhenChangedUtc = ConvertTo-CsvTimestamp $EndDate.AddMonths(-1)
                }))
    }

    foreach ($policy in $autoLabelPolicies.Values) {
        $policyRows.Add((New-PolicySnapshotRow -RunDate $runDate -ObjectType 'AutoLabelingPolicy' -Values @{
                    ObjectId       = $policy.Id
                    Name           = $policy.Name
                    Priority       = 0
                    Mode           = 'Enable'
                    Enabled        = $true
                    Locations      = 'SharePoint;OneDrive'
                    LabelIds       = $policy.Label.Id
                    LabelNames     = $policy.Label.Name
                    Comment        = "Sample auto-labeling policy $($policy.Name)."
                    WhenCreatedUtc = ConvertTo-CsvTimestamp $EndDate.AddMonths(-12)
                    WhenChangedUtc = ConvertTo-CsvTimestamp $EndDate.AddMonths(-2)
                }))
    }

    foreach ($policy in $dlpPolicies.Values) {
        $policyRows.Add((New-PolicySnapshotRow -RunDate $runDate -ObjectType 'DlpPolicy' -Values @{
                    ObjectId         = $policy.Id
                    Name             = $policy.Name
                    Priority         = 0
                    Mode             = 'Enable'
                    Enabled          = $true
                    Locations        = 'Exchange;SharePoint;OneDrive;Teams'
                    AppliesToCopilot = $policy.Copilot
                    Comment          = "Sample DLP policy $($policy.Name)."
                    WhenCreatedUtc   = ConvertTo-CsvTimestamp $EndDate.AddMonths(-11)
                    WhenChangedUtc   = ConvertTo-CsvTimestamp $EndDate.AddMonths(-1)
                }))
    }

    foreach ($rule in $dlpRules.Values) {
        $policyRows.Add((New-PolicySnapshotRow -RunDate $runDate -ObjectType 'DlpRule' -Values @{
                    ObjectId                  = $rule.Id
                    Name                      = $rule.Name
                    ParentId                  = $rule.Policy.Id
                    ParentName                = $rule.Policy.Name
                    Priority                  = 0
                    Enabled                   = $true
                    SensitiveInformationTypes = $rule.Sit
                    BlockAccess               = $true
                    BlockAccessScope          = 'All'
                    NotifyUser                = 'LastModifier'
                    AllowOverride             = 'True_With_Justification'
                    RequireJustification      = $true
                    GenerateIncidentReport    = 'SiteAdmin'
                    Comment                   = "Sample DLP rule $($rule.Name)."
                    WhenCreatedUtc            = ConvertTo-CsvTimestamp $EndDate.AddMonths(-11)
                    WhenChangedUtc            = ConvertTo-CsvTimestamp $EndDate.AddMonths(-1)
                }))
    }

    foreach ($policy in $retentionPolicies.Values) {
        $policyRows.Add((New-PolicySnapshotRow -RunDate $runDate -ObjectType 'RetentionPolicy' -Values @{
                    ObjectId       = $policy.Id
                    Name           = $policy.Name
                    Mode           = 'Enable'
                    Enabled        = $true
                    Workload       = 'Exchange;SharePoint;OneDrive'
                    Locations      = 'All'
                    Comment        = "Sample retention policy $($policy.Name)."
                    WhenCreatedUtc = ConvertTo-CsvTimestamp $EndDate.AddMonths(-10)
                    WhenChangedUtc = ConvertTo-CsvTimestamp $EndDate.AddMonths(-1)
                }))
    }

    foreach ($tag in $retentionLabels.Values) {
        $policyRows.Add((New-PolicySnapshotRow -RunDate $runDate -ObjectType 'RetentionLabel' -Values @{
                    ObjectId          = $tag.Id
                    Name              = $tag.Name
                    DisplayName       = $tag.Name
                    Priority          = 0
                    RetentionAction   = $tag.Action
                    RetentionDuration = $tag.Duration
                    RetentionType     = $tag.Type
                    IsRecordLabel     = $tag.Record
                    Comment           = "Sample retention label $($tag.Name)."
                    WhenCreatedUtc    = ConvertTo-CsvTimestamp $EndDate.AddMonths(-10)
                    WhenChangedUtc    = ConvertTo-CsvTimestamp $EndDate.AddMonths(-1)
                }))
    }
}

#endregion

#region Activity Explorer events (30-day retention window)

$eventStart = $EndDate.AddDays(-25)
$activityRows = [System.Collections.Generic.List[object]]::new()
$allLabels = @($labels.Values)
$allRetentionLabels = @($retentionLabels.Values)
$allDlpRules = @($dlpRules.Values)
$workloads = @('SharePoint', 'OneDrive', 'Exchange', 'Teams', 'Endpoint')
$fileNames = @('Statement of Work.docx', 'Pricing Model.xlsx', 'Architecture Review.pptx', 'Customer List.xlsx', 'Contract Draft.docx')

function Add-ActivityRow {
    param(
        [Parameter(Mandatory)][string]$RecordIdentity,
        [Parameter(Mandatory)][datetime]$Happened,
        [Parameter(Mandatory)][string]$Activity,
        [Parameter(Mandatory)][string]$Workload,
        [Parameter(Mandatory)][object]$User,
        [hashtable]$Extra = @{}
    )

    $labelEventType = if ($Extra.ContainsKey('LabelEventType')) { $Extra.LabelEventType } else { '' }
    $sitData = if ($Extra.ContainsKey('SensitiveInfoTypeData')) { $Extra.SensitiveInfoTypeData } else { $null }
    $sitSummary = Get-PurviewSensitiveInfoTypeSummary $sitData

    $row = [ordered]@{}
    foreach ($column in $schema.ActivityExplorerEvents) { $row[$column] = '' }

    $row.RecordIdentity = $RecordIdentity
    $row.Happened = ConvertTo-CsvTimestamp $Happened
    $row.EventDate = $Happened.ToString('yyyy-MM-dd')
    $row.Activity = $Activity
    $row.ActivityRaw = $Activity
    $row.ActivityCategory = Get-PurviewActivityCategory -Activity $Activity -Schema $schema
    $row.Workload = $Workload
    $row.User = $User.UserPrincipalName
    $row.UserType = 'Member'
    $row.LabelEventType = $labelEventType
    $row.IsLabelDowngrade = ($labelEventType -eq 'LabelDowngraded')
    $row.SensitiveInfoTypeName = $sitSummary.Name
    $row.SensitiveInfoTypeCount = $sitSummary.Count
    $row.SensitiveInfoTypeConfidence = $sitSummary.Confidence

    foreach ($key in $Extra.Keys) {
        if ($key -eq 'LabelEventType' -or $key -eq 'SensitiveInfoTypeData') { continue }
        $row[$key] = $Extra[$key]
    }

    $activityRows.Add([pscustomobject]$row)
}

$eventCounter = 0
function New-ActivityId {
    $script:eventCounter++
    'ae-{0:D5}' -f $script:eventCounter
}

# Labeling events: apply, change, remove, recommend, downgrade.
for ($i = 0; $i -lt 40; $i++) {
    $user = Get-RandomItem $members.ToArray()
    $label = Get-RandomItem $allLabels
    $happened = New-EventTime -From $eventStart
    $file = Get-RandomItem $fileNames
    $workload = Get-RandomItem @('SharePoint', 'OneDrive', 'Exchange')

    Add-ActivityRow -RecordIdentity (New-ActivityId) -Happened $happened -Activity 'LabelApplied' -Workload $workload -User $user -Extra @{
        LabelEventType   = 'LabelApplied'
        ItemName         = $file
        FilePath         = "/sites/team/$file"
        FullUrl          = "https://example.sharepoint.com/sites/team/$file"
        FileExtension    = [System.IO.Path]::GetExtension($file).TrimStart('.')
        Platform         = 'Web'
        SensitivityLabel = $label.Id
        HowApplied       = 'Manual'
        IsProtected      = ($label.Encrypted)
    }
}

for ($i = 0; $i -lt 10; $i++) {
    $user = Get-RandomItem $members.ToArray()
    $oldLabel = $labels.Restricted
    $newLabel = $labels.General
    $happened = New-EventTime -From $eventStart

    Add-ActivityRow -RecordIdentity (New-ActivityId) -Happened $happened -Activity 'LabelChanged' -Workload 'SharePoint' -User $user -Extra @{
        LabelEventType      = 'LabelDowngraded'
        ItemName            = Get-RandomItem $fileNames
        SensitivityLabel    = $newLabel.Id
        OldSensitivityLabel = $oldLabel.Id
        HowApplied          = 'Manual'
        Justification       = 'No longer contains restricted data'
    }
}

# Protection events.
for ($i = 0; $i -lt 15; $i++) {
    $user = Get-RandomItem $members.ToArray()
    Add-ActivityRow -RecordIdentity (New-ActivityId) -Happened (New-EventTime -From $eventStart) -Activity 'NewProtection' -Workload 'Exchange' -User $user -Extra @{
        ItemName        = Get-RandomItem $fileNames
        ProtectionType  = 'Rms'
        ProtectionOwner = $user.UserPrincipalName
        RMSEncrypted    = $true
    }
}

# DLP events, some with an override and justification.
for ($i = 0; $i -lt 25; $i++) {
    $user = Get-RandomItem $members.ToArray()
    $rule = Get-RandomItem $allDlpRules
    $overridden = ($i % 5 -eq 0)

    Add-ActivityRow -RecordIdentity (New-ActivityId) -Happened (New-EventTime -From $eventStart) -Activity 'DLPRuleMatch' -Workload (Get-RandomItem $workloads) -User $user -Extra @{
        ItemName                    = Get-RandomItem $fileNames
        PolicyId                    = $rule.Policy.Id
        PolicyName                  = $rule.Policy.Name
        RuleId                      = $rule.Id
        RuleName                    = $rule.Name
        RuleActions                 = 'BlockAccess;NotifyUser'
        EnforcementMode             = 'Enable'
        FalsePositive               = $false
        DlpPolicyMatchId            = New-DeterministicGuid
        SensitiveInfoTypeData       = @([pscustomobject]@{ SensitiveInfoTypeName = $rule.Sit; SensitiveInfoTypeCount = $random.Next(1, 5); SensitiveInfoTypeConfidence = $random.Next(65, 100) })
        Justification               = if ($overridden) { 'Approved by manager for external audit' } else { '' }
    }
}

# Retention classification events.
foreach ($tag in $allRetentionLabels) {
    for ($i = 0; $i -lt 5; $i++) {
        $user = Get-RandomItem $members.ToArray()
        Add-ActivityRow -RecordIdentity (New-ActivityId) -Happened (New-EventTime -From $eventStart) -Activity 'ClassificationAdded' -Workload (Get-RandomItem @('SharePoint', 'OneDrive')) -User $user -Extra @{
            ItemName       = Get-RandomItem $fileNames
            RetentionLabel = $tag.Name
        }
    }
}

# Copilot / AI interactions and a file discovery scan.
for ($i = 0; $i -lt 10; $i++) {
    $user = Get-RandomItem $members.ToArray()
    Add-ActivityRow -RecordIdentity (New-ActivityId) -Happened (New-EventTime -From $eventStart) -Activity 'CopilotInteraction' -Workload 'Teams' -User $user -Extra @{
        Application = 'Microsoft 365 Copilot'
    }
}
for ($i = 0; $i -lt 5; $i++) {
    $user = Get-RandomItem $members.ToArray()
    Add-ActivityRow -RecordIdentity (New-ActivityId) -Happened (New-EventTime -From $eventStart) -Activity 'FileDiscovered' -Workload 'SharePoint' -User $user -Extra @{
        ItemName = Get-RandomItem $fileNames
    }
}

# Endpoint events, to exercise the largest activity category.
for ($i = 0; $i -lt 10; $i++) {
    $user = Get-RandomItem $members.ToArray()
    Add-ActivityRow -RecordIdentity (New-ActivityId) -Happened (New-EventTime -From $eventStart) -Activity 'FileCopiedToRemovableMedia' -Workload 'Endpoint' -User $user -Extra @{
        ItemName   = Get-RandomItem $fileNames
        DeviceName = 'DESKTOP-{0}' -f $random.Next(10000, 99999)
        Platform   = 'Windows'
    }
}

$sortedActivity = @($activityRows | Sort-Object Happened, RecordIdentity)

#endregion

#region Content Explorer snapshot

$contentSnapshotDates = @($EndDate.AddMonths(-2), $EndDate.AddMonths(-1), $EndDate) | ForEach-Object { $_.ToString('yyyy-MM-dd') }
$contentTags = [System.Collections.Generic.List[object]]::new()
foreach ($label in $allLabels) { $contentTags.Add([pscustomobject]@{ TagType = 'Sensitivity'; TagName = $label.Name }) }
foreach ($tag in $allRetentionLabels) { $contentTags.Add([pscustomobject]@{ TagType = 'Retention'; TagName = $tag.Name }) }
foreach ($sit in @('Credit Card Number', 'U.S. Social Security Number (SSN)', 'U.S. / U.K. Passport Number')) {
    $contentTags.Add([pscustomobject]@{ TagType = 'SensitiveInformationType'; TagName = $sit })
}

$contentRows = [System.Collections.Generic.List[object]]::new()
$snapshotIndex = 0
foreach ($runDate in $contentSnapshotDates) {
    foreach ($tag in $contentTags) {
        foreach ($workload in @('EXO', 'ODB', 'SPO', 'Teams')) {
            # Counts grow snapshot over snapshot, so the report shows a trend.
            $base = 20 + ($contentTags.IndexOf($tag) * 7) + ($workload.Length * 3)
            $total = $base + ($snapshotIndex * $random.Next(5, 25))
            $contentRows.Add([pscustomobject]@{
                    RunDate    = $runDate
                    TagType    = $tag.TagType
                    TagName    = $tag.TagName
                    Workload   = $workload
                    TotalCount = $total
                })
        }
    }
    $snapshotIndex++
}

#endregion

#region Copilot accessed resources

$copilotRows = [System.Collections.Generic.List[object]]::new()
$agents = @(
    @{ Id = New-DeterministicGuid; Name = 'Sales Q&A Agent' }
    @{ Id = New-DeterministicGuid; Name = 'HR Policy Agent' }
    @{ Id = New-DeterministicGuid; Name = 'Microsoft 365 Copilot Chat' }
)
$resourceTypes = @('File', 'ListItem', 'Message')
$copilotWorkloads = @('Teams', 'SharePointOnline', 'Word', 'Outlook')

for ($i = 0; $i -lt 40; $i++) {
    $user = Get-RandomItem $members.ToArray()
    $agent = Get-RandomItem $agents
    $happened = New-EventTime -From $eventStart
    $recordId = 'copilot-{0:D5}' -f $i
    $threadId = New-DeterministicGuid

    $resourceCount = $random.Next(0, 3)
    if ($resourceCount -eq 0) {
        $copilotRows.Add([pscustomobject]@{
                RecordId     = $recordId
                CreationTime = ConvertTo-CsvTimestamp $happened
                EventDate    = $happened.ToString('yyyy-MM-dd')
                Operation    = 'CopilotInteraction'
                RecordType   = 'CopilotInteraction'
                Workload     = Get-RandomItem $copilotWorkloads
                UserId       = $user.Id
                UserKey      = $user.UserPrincipalName
                UserType     = 'Regular'
                AppHost      = 'Copilot'
                AppIdentity  = 'Microsoft 365 Copilot'
                AgentId      = $agent.Id
                AgentName    = $agent.Name
                ThreadId     = $threadId
            })
        continue
    }

    for ($r = 0; $r -lt $resourceCount; $r++) {
        $label = Get-RandomItem $allLabels
        $blocked = ($label.Name -eq 'Restricted' -and $random.Next(0, 2) -eq 0)
        $policy = if ($blocked) { $dlpPolicies.CopilotDlp } else { $null }
        $rule = if ($blocked) { $dlpRules.CopilotSensitive } else { $null }

        $copilotRows.Add([pscustomobject]@{
                RecordId           = $recordId
                CreationTime       = ConvertTo-CsvTimestamp $happened
                EventDate          = $happened.ToString('yyyy-MM-dd')
                Operation          = 'CopilotInteraction'
                RecordType         = 'CopilotInteraction'
                Workload           = Get-RandomItem $copilotWorkloads
                UserId             = $user.Id
                UserKey            = $user.UserPrincipalName
                UserType           = 'Regular'
                AppHost            = 'Copilot'
                AppIdentity        = 'Microsoft 365 Copilot'
                AgentId            = $agent.Id
                AgentName          = $agent.Name
                ThreadId           = $threadId
                ResourceId         = New-DeterministicGuid
                ResourceName       = Get-RandomItem $fileNames
                ResourceType       = Get-RandomItem $resourceTypes
                ResourceAction     = 'Read'
                SiteUrl            = 'https://example.sharepoint.com/sites/team'
                ListItemUniqueId   = New-DeterministicGuid
                SensitivityLabelId = $label.Id
                Status             = if ($blocked) { 'failure' } else { 'success' }
                AccessBlocked      = [bool]$blocked
                XpiaDetected       = $false
                PolicyId           = if ($policy) { $policy.Id } else { '' }
                PolicyName         = if ($policy) { $policy.Name } else { '' }
                PolicyRules        = if ($rule) { $rule.Name } else { '' }
            })
    }
}

$sortedCopilot = @($copilotRows | Sort-Object CreationTime, RecordId)

#endregion

#region Write the files

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

# The generator rewrites the set from scratch; Export-AppendCsv would
# otherwise append to the files already committed.
foreach ($name in @('users.csv', 'policies.csv', 'activity-explorer-events.csv', 'content-explorer-snapshot.csv', 'copilot-accessed-resources.csv', 'run.log')) {
    $stale = Join-Path $OutputPath $name
    if (Test-Path -LiteralPath $stale) { Remove-Item -LiteralPath $stale -Force }
}

$userRows = [System.Collections.Generic.List[object]]::new()
foreach ($runDate in $policySnapshotDates) {
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
}

Export-AppendCsv -Path (Join-Path $OutputPath 'users.csv') -Rows $userRows.ToArray() -Column (Get-EntraUserCsvColumn)
Export-AppendCsv -Path (Join-Path $OutputPath 'policies.csv') -Rows $policyRows.ToArray() -Column $schema.Policies
Export-AppendCsv -Path (Join-Path $OutputPath 'activity-explorer-events.csv') -Rows $sortedActivity -Column $schema.ActivityExplorerEvents -KeyColumn 'RecordIdentity'
Export-AppendCsv -Path (Join-Path $OutputPath 'content-explorer-snapshot.csv') -Rows $contentRows.ToArray() -Column $schema.ContentExplorerSnapshot
Export-AppendCsv -Path (Join-Path $OutputPath 'copilot-accessed-resources.csv') -Rows $sortedCopilot -Column $schema.CopilotAccessedResources -KeyColumn @('RecordId', 'ResourceId')

[pscustomobject]@{
    OutputPath        = $OutputPath
    Members           = $members.Count
    PolicyRows        = $policyRows.Count
    ActivityEvents    = $sortedActivity.Count
    ContentSnapshots  = $contentRows.Count
    CopilotRows       = $sortedCopilot.Count
}

#endregion
