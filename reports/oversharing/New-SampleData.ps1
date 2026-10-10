#Requires -Version 7.0

<#
    .SYNOPSIS
        Generates the fake oversharing data in ./samples.

    .DESCRIPTION
        The sample set lets the later Power BI report, and anyone reading this repository,
        work against realistic shapes without a tenant. Nothing here comes from a real
        directory: every address is under example.com, example.net or example.org, which
        RFC 2606 reserves for documentation, and every site is under
        example.sharepoint.com or example-my.sharepoint.com.

        The generator is deterministic. The same -Seed and -EndDate always produce the
        same files.

        Every file is written through Export-AppendCsv with the column list the collectors
        use, so a sample file cannot drift from its collector's output. Snapshot files hold
        three snapshots a month apart. The sets cover every branch a page reads: Anyone,
        organization and Specific-people links, inherited and direct permissions, sites open
        to Everyone except external users, and each SharingCapability and
        DefaultSharingLinkType value. The two audit-event files are inside the last 90 days.

        The ReportRow of the sharing link, EEEU activity and labeled-file CSVs is
        illustrative: Learn does not list the columns of those exports.

        The contract doc marks no source NotAvailable in any cloud, so there are no
        header-only gcc or gcchigh samples.

    .PARAMETER EndDate
        The "now" the sample set is generated around. Fixed by default to keep the
        committed files stable.

    .EXAMPLE
        ./New-SampleData.ps1
#>
[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path $PSScriptRoot 'samples'),

    [datetime]$EndDate = [datetime]::new(2026, 10, 1, 0, 0, 0, [System.DateTimeKind]::Utc),

    [int]$Seed = 20261001
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot '../../shared/M365ReportLibrary.psm1')

. (Join-Path $PSScriptRoot 'collectors/OversharingHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'collectors/OversharingSchema.psd1')

$EndDate = $EndDate.ToUniversalTime()
$random = [System.Random]::new($Seed)

# Three snapshots a month apart, the last one on EndDate.
$snapshotDates = 2..0 | ForEach-Object { $EndDate.AddMonths(-$_) }

$givenNames = @('Avery', 'Blair', 'Casey', 'Devon', 'Emery', 'Finley', 'Gray', 'Harper')
$familyNames = @('Abara', 'Bergstrom', 'Chaudhry', 'Dlamini', 'Ibarra', 'Jovanovic', 'Kowalski', 'Lindqvist')
$people = foreach ($i in 0..7) {
    [pscustomobject]@{
        Name = '{0} {1}' -f $givenNames[$i], $familyNames[$i]
        Upn  = ('{0}.{1}@example.com' -f $givenNames[$i], $familyNames[$i]).ToLowerInvariant()
    }
}
$people = @($people)

function New-DeterministicGuid {
    $bytes = [byte[]]::new(16)
    $random.NextBytes($bytes)
    return [guid]::new($bytes).ToString()
}

function Format-Stamp {
    param([AllowNull()][object]$Value)
    return ConvertTo-CsvTimestamp $Value
}

if (Test-Path -LiteralPath $OutputPath) {
    Get-ChildItem -LiteralPath $OutputPath -Recurse -File -Filter '*.csv' | Remove-Item -Force
}
New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null

# Twelve sites: the root, eight SharePoint sites and three OneDrive sites.
$siteDefinitions = @(
    @{ Name = 'Root Site'; Path = ''; Template = 'SITEPAGEPUBLISHING#0'; Personal = $false }
    @{ Name = 'All Company'; Path = '/sites/allcompany'; Template = 'SITEPAGEPUBLISHING#0'; Personal = $false }
    @{ Name = 'Finance'; Path = '/sites/finance'; Template = 'GROUP#0'; Personal = $false }
    @{ Name = 'Human Resources'; Path = '/sites/hr'; Template = 'GROUP#0'; Personal = $false }
    @{ Name = 'Legal'; Path = '/sites/legal'; Template = 'STS#3'; Personal = $false }
    @{ Name = 'Marketing'; Path = '/sites/marketing'; Template = 'GROUP#0'; Personal = $false }
    @{ Name = 'Engineering'; Path = '/sites/engineering'; Template = 'GROUP#0'; Personal = $false }
    @{ Name = 'Partner Portal'; Path = '/sites/partners'; Template = 'STS#3'; Personal = $false }
    @{ Name = 'Avery Abara'; Path = '/personal/avery_abara_example_com'; Template = 'SPSPERS#10'; Personal = $true }
    @{ Name = 'Blair Bergstrom'; Path = '/personal/blair_bergstrom_example_com'; Template = 'SPSPERS#10'; Personal = $true }
    @{ Name = 'Casey Chaudhry'; Path = '/personal/casey_chaudhry_example_com'; Template = 'SPSPERS#10'; Personal = $true }
    @{ Name = 'Devon Dlamini'; Path = '/personal/devon_dlamini_example_com'; Template = 'SPSPERS#10'; Personal = $true }
)
$sites = foreach ($definition in $siteDefinitions) {
    $hostName = if ($definition.Personal) { 'example-my.sharepoint.com' } else { 'example.sharepoint.com' }
    [pscustomobject]@{
        Id       = '{0},{1},{2}' -f $hostName, (New-DeterministicGuid), (New-DeterministicGuid)
        Name     = $definition.Name
        HostName = $hostName
        WebUrl   = "https://$hostName$($definition.Path)"
        Template = $definition.Template
        Personal = $definition.Personal
        Guid     = New-DeterministicGuid
    }
}
$sites = @($sites)

$labels = @(
    @{ Guid = '3f1c6a52-8b0e-4d7a-9c2e-5a1b7d9e0f11'; Name = 'Confidential' }
    @{ Guid = '8d4e2b90-1a6c-4f35-b7d8-2c9e0a1f3b22'; Name = 'Highly Confidential' }
)

foreach ($snapshot in $snapshotDates) {
    $runDate = $snapshot.ToString('yyyy-MM-dd')
    $snapshotIndex = [array]::IndexOf($snapshotDates, $snapshot)
    $reportDate = Format-Stamp $snapshot.AddHours(-30)

    # Source 1.
    $siteRows = foreach ($site in $sites) {
        [pscustomobject]@{
            RunDate = $runDate; SiteId = $site.Id; Name = $site.Name; WebUrl = $site.WebUrl
            IsPersonalSite = $site.Personal.ToString(); HostName = $site.HostName; DataLocationCode = ''
        }
    }
    Export-AppendCsv -Path (Join-Path $OutputPath 'sites.csv') -Rows @($siteRows) -Column $schema.Sites -KeyColumn @('RunDate', 'SiteId')

    # Source 2. A few folders and files in the first four sites' document libraries.
    $permissionRows = foreach ($siteIndex in 1..4) {
        $site = $sites[$siteIndex]
        $driveId = 'b!drive{0}' -f $siteIndex
        foreach ($itemIndex in 1..3) {
            $itemId = '01ITEM{0}{1}' -f $siteIndex, $itemIndex
            $item = @{ Name = @('Budget.xlsx', 'Handbook.docx', 'Roadmap.pptx')[$itemIndex - 1]; Url = "$($site.WebUrl)/Shared Documents/$itemId" }
            $kind = ($siteIndex + $itemIndex + $snapshotIndex) % 4
            $link = switch ($kind) {
                0 { @{ Scope = 'anonymous'; Type = 'view' } }
                1 { @{ Scope = 'organization'; Type = 'edit' } }
                2 { @{ Scope = 'users'; Type = 'edit' } }
                default { $null }
            }
            if ($link) {
                [pscustomobject]@{
                    RunDate = $runDate; SiteId = $site.Id; DriveId = $driveId; ItemId = $itemId; ItemName = $item.Name; ItemWebUrl = $item.Url
                    PermissionId = "link-$siteIndex-$itemIndex"; Roles = $(if ($link.Type -eq 'edit') { 'write' } else { 'read' })
                    LinkScope = $link.Scope; LinkType = $link.Type; LinkPreventsDownload = 'False'
                    HasPassword = $(if ($link.Scope -eq 'anonymous' -and $itemIndex -eq 2) { 'True' } else { 'False' })
                    ExpirationDateTime = $(if ($link.Scope -eq 'anonymous') { Format-Stamp $snapshot.AddDays(14) } else { '' })
                    IsInherited = 'False'; InheritedFromItemId = ''
                    GrantedTo = $(if ($link.Scope -eq 'users') { $people[$itemIndex].Name } else { '' })
                }
            }
            $direct = $people[($siteIndex + $itemIndex) % $people.Count]
            [pscustomobject]@{
                RunDate = $runDate; SiteId = $site.Id; DriveId = $driveId; ItemId = $itemId; ItemName = $item.Name; ItemWebUrl = $item.Url
                PermissionId = "grant-$siteIndex-$itemIndex"; Roles = 'write'; LinkScope = ''; LinkType = ''
                LinkPreventsDownload = ''; HasPassword = ''; ExpirationDateTime = ''
                IsInherited = ($itemIndex -eq 3).ToString(); InheritedFromItemId = $(if ($itemIndex -eq 3) { '01FOLDER' + $siteIndex } else { '' })
                GrantedTo = $direct.Name
            }
        }
    }
    Export-AppendCsv -Path (Join-Path $OutputPath 'item-permissions.csv') -Rows @($permissionRows) -Column $schema.ItemPermissions -KeyColumn @('RunDate', 'DriveId', 'ItemId', 'PermissionId')

    # Source 3a. Report ids change per snapshot and workload, as a new report does.
    foreach ($workload in 'SharePoint', 'OneDriveForBusiness') {
        $reportId = [guid]::new($snapshotIndex + 1, 1, $(if ($workload -eq 'SharePoint') { 1 } else { 2 }), [byte[]](1, 2, 3, 4, 5, 6, 7, 8)).ToString()
        $breadthRows = foreach ($site in $sites | Where-Object { $_.Personal -eq ($workload -eq 'OneDriveForBusiness') }) {
            $index = [array]::IndexOf($sites, $site)
            $users = if ($site.Name -eq 'All Company') { 4200 } elseif ($site.Personal) { 1 + ($index % 3) } else { 15 + ($index * 37 + $snapshotIndex * 11) % 900 }
            $eeeu = if ($site.Name -in 'All Company', 'Partner Portal') { 3 + $snapshotIndex } else { $index % 3 }
            [pscustomobject]@{
                RunDate = $runDate; Workload = $workload; ReportId = $reportId; ReportDate = $reportDate
                SiteId = $site.Id; SiteName = $site.Name; SiteUrl = $site.WebUrl; SiteTemplate = $(if ($site.Template -like 'GROUP*') { 'Team site' } elseif ($site.Personal) { 'Other sites' } else { 'Communication site' })
                PrimaryAdmin = $people[$index % $people.Count].Name; PrimaryAdminEmail = $people[$index % $people.Count].Upn
                ExternalSharing = $(if ($site.Name -eq 'Partner Portal') { 'Yes' } else { 'No' }); SitePrivacy = $(if ($site.Template -like 'GROUP*') { @('Private', 'Public')[$index % 2] } else { '' })
                SiteSensitivity = $(if ($site.Name -in 'Legal', 'Finance') { 'Confidential' } else { '' })
                UsersWithAccess = [string]$users; GuestUserPermissions = [string]($index % 4); ExternalParticipantPermissions = '0'
                EntraGroupPermissions = [string]($index % 5); FileCount = [string](200 + $index * 313); ItemsWithUniquePermissions = [string]($index * 7 % 40)
                PeopleInYourOrgLinks = [string]($index * 3 % 25); AnyoneLinks = [string]($index * 5 % 9 + $snapshotIndex); EeeuPermissions = [string]$eeeu
                EveryonePermissions = $(if ($site.Name -eq 'Partner Portal') { '1' } else { '0' })
            }
        }
        Export-AppendCsv -Path (Join-Path $OutputPath 'site-permission-breadth.csv') -Rows @($breadthRows) -Column $schema.SitePermissionBreadth -KeyColumn @('ReportId', 'SiteId')
    }

    # Source 3b. Direct and indirect EEEU and Everyone access.
    $exposureRows = foreach ($entity in 'EveryoneExceptExternalUsers', 'Everyone') {
        $reportId = [guid]::new($snapshotIndex + 1, 2, $(if ($entity -eq 'Everyone') { 2 } else { 1 }), [byte[]](8, 7, 6, 5, 4, 3, 2, 1)).ToString()
        $recipient = if ($entity -eq 'Everyone') { 'Everyone' } else { 'Everyone except external users' }
        foreach ($siteIndex in 1, 2, 7) {
            if ($entity -eq 'Everyone' -and $siteIndex -ne 7) { continue }
            $site = $sites[$siteIndex]
            foreach ($itemIndex in 0..2) {
                $indirect = ($itemIndex -eq 0)
                [pscustomobject]@{
                    RunDate = $runDate; ReportEntity = $entity; ReportId = $reportId; ReportDate = $reportDate
                    SiteId = $site.Guid; WebId = $site.Guid; ListId = [guid]::new($siteIndex, 3, 3, [byte[]](1, 1, 1, 1, 1, 1, 1, 1)).ToString()
                    ScopeId = [guid]::new($siteIndex, 4, $itemIndex, [byte[]](2, 2, 2, 2, 2, 2, 2, 2)).ToString()
                    UniqueId = [guid]::new($siteIndex, 5, $itemIndex, [byte[]](3, 3, 3, 3, 3, 3, 3, 3)).ToString()
                    ListItemId = [string]($itemIndex + 1); ItemType = @('Web', 'Folder', 'File')[$itemIndex]
                    ItemUrl = $(if ($itemIndex -eq 0) { ([uri]$site.WebUrl).AbsolutePath } else { '{0}/Shared Documents/Item{1}' -f ([uri]$site.WebUrl).AbsolutePath, $itemIndex })
                    RoleDefinition = @('Read', 'Edit', 'Read')[$itemIndex]; LinkId = ''; LinkScope = ''; Recipient = $recipient
                    ParentObjectId = $(if ($indirect) { [guid]::new($siteIndex, 6, 6, [byte[]](4, 4, 4, 4, 4, 4, 4, 4)).ToString() } else { '' })
                    ParentGroupName = $(if ($indirect) { "$($site.Name) visitors" } else { '' }); ParentGroupEmail = ''
                    ParentGroupType = $(if ($indirect) { 'SharePoint group' } else { '' }); TotalUserCount = [string](10 + $siteIndex * 13 + $itemIndex)
                }
            }
        }
    }
    Export-AppendCsv -Path (Join-Path $OutputPath 'everyone-item-exposure.csv') -Rows @($exposureRows) `
        -Column $schema.EveryoneItemExposure -KeyColumn @('ReportId', 'ReportEntity', 'SiteId', 'WebId', 'ListId', 'ScopeId', 'UniqueId', 'LinkId', 'ParentObjectId', 'RoleDefinition', 'Recipient')

    # Sources 3c and 3d. Rolling 28-day reports.
    $windowStart = Format-Stamp $snapshot.AddDays(-28)
    $windowEnd = Format-Stamp $snapshot
    $linkRows = foreach ($entity in 'SharingLinks_Anyone', 'SharingLinks_PeopleInYourOrg', 'SharingLinks_Guests') {
        $reportId = [guid]::new($snapshotIndex + 1, 7, [int]$entity.Length, [byte[]](1, 3, 5, 7, 9, 11, 13, 15)).ToString()
        foreach ($siteIndex in 1..4) {
            $site = $sites[$siteIndex]
            $row = [ordered]@{ 'Site ID' = $site.Guid; 'Site URL' = $site.WebUrl; 'Site Name' = $site.Name; 'Links created' = [string](2 + ($siteIndex * 5 + $snapshotIndex) % 17) }
            [pscustomobject]@{
                RunDate = $runDate; ReportEntity = $entity; Workload = 'SharePoint'; ReportId = $reportId
                ReportStartTime = $windowStart; ReportEndTime = $windowEnd; SiteId = $site.Guid; SiteUrl = $site.WebUrl
                ReportRow = ($row | ConvertTo-Json -Compress)
            }
        }
    }
    Export-AppendCsv -Path (Join-Path $OutputPath 'sharing-link-activity.csv') -Rows @($linkRows) -Column $schema.SharingLinkActivity -KeyColumn @('ReportId', 'ReportEntity', 'Workload', 'ReportRow')

    $eeeuRows = foreach ($entity in 'EveryoneExceptExternalUsersAtSite', 'EveryoneExceptExternalUsersForItems') {
        $reportId = [guid]::new($snapshotIndex + 1, 8, [int]$entity.Length, [byte[]](2, 4, 6, 8, 10, 12, 14, 16)).ToString()
        foreach ($siteIndex in 1, 2, 7) {
            $site = $sites[$siteIndex]
            $row = [ordered]@{ 'Site ID' = $site.Guid; 'Site URL' = $site.WebUrl; 'Shared with EEEU' = [string](1 + ($siteIndex + $snapshotIndex) % 6) }
            [pscustomobject]@{
                RunDate = $runDate; ReportEntity = $entity; Workload = 'SharePoint'; ReportId = $reportId
                ReportStartTime = $windowStart; ReportEndTime = $windowEnd; SiteId = $site.Guid; SiteUrl = $site.WebUrl
                ReportRow = ($row | ConvertTo-Json -Compress)
            }
        }
    }
    Export-AppendCsv -Path (Join-Path $OutputPath 'eeeu-activity.csv') -Rows @($eeeuRows) -Column $schema.EeeuActivity -KeyColumn @('ReportId', 'ReportEntity', 'Workload', 'ReportRow')

    # Source 3e. One report per label.
    $labelRows = foreach ($label in $labels) {
        $reportId = [guid]::new($snapshotIndex + 1, 9, [int]$label.Name.Length, [byte[]](9, 8, 7, 6, 5, 4, 3, 2)).ToString()
        foreach ($siteIndex in 2, 4) {
            $site = $sites[$siteIndex]
            $row = [ordered]@{ 'Site ID' = $site.Guid; 'Site URL' = $site.WebUrl; 'Labeled files' = [string](5 + $siteIndex * 4) }
            [pscustomobject]@{
                RunDate = $runDate; LabelGuid = $label.Guid; LabelName = $label.Name; Workload = 'SharePoint'; ReportId = $reportId
                ReportCreatedDateTime = $reportDate; SiteId = $site.Guid; SiteUrl = $site.WebUrl; ReportRow = ($row | ConvertTo-Json -Compress)
            }
        }
    }
    Export-AppendCsv -Path (Join-Path $OutputPath 'labeled-file-sites.csv') -Rows @($labelRows) -Column $schema.LabeledFileSites -KeyColumn @('ReportId', 'LabelGuid', 'ReportRow')

    # Source 4. The tenant row, then every site. Every SharingCapability and
    # DefaultSharingLinkType value appears.
    $capabilities = @('ExternalUserAndGuestSharing', 'Disabled', 'ExternalUserSharingOnly', 'ExistingExternalUserSharingOnly')
    $linkTypes = @('None', 'Direct', 'Internal', 'AnonymousAccess')
    $settingRows = @([pscustomobject]@{
            RunDate = $runDate; Scope = 'Tenant'; Url = ''; Title = ''; Template = ''; SharingCapability = 'ExternalUserAndGuestSharing'
            DefaultSharingLinkType = 'Direct'; DisableCompanyWideSharingLinks = 'False'; SensitivityLabel = ''
        })
    $settingRows += foreach ($index in 0..($sites.Count - 1)) {
        $site = $sites[$index]
        [pscustomobject]@{
            RunDate = $runDate; Scope = 'Site'; Url = $site.WebUrl; Title = $site.Name; Template = $site.Template
            SharingCapability = $capabilities[$index % 4]; DefaultSharingLinkType = $linkTypes[($index + $snapshotIndex) % 4]
            DisableCompanyWideSharingLinks = ($index % 5 -eq 0).ToString()
            SensitivityLabel = $(if ($site.Name -in 'Legal', 'Finance') { 'Confidential' } else { '' })
        }
    }
    Export-AppendCsv -Path (Join-Path $OutputPath 'site-sharing-settings.csv') -Rows $settingRows -Column $schema.SiteSharingSettings -KeyColumn @('RunDate', 'Scope', 'Url')

    # Source 5c.
    Export-AppendCsv -Path (Join-Path $OutputPath 'audit-log-status.csv') -Rows @([pscustomobject]@{
            RunDate = $runDate; UnifiedAuditLogIngestionEnabled = 'True'
        }) -Column $schema.AuditLogStatus -KeyColumn 'RunDate'
}

# Audit events, inside the last 90 days. RecordType is the display name that
# Search-UnifiedAuditLog -Formatted returns.
function New-SampleEvent {
    param([string]$Operation, [int]$Index, [string]$Target = '', [string]$TargetType = '')

    $user = $people[$Index % $people.Count]
    $site = $sites[1 + $Index % 7]
    $file = @('Budget.xlsx', 'Handbook.docx', 'Roadmap.pptx', 'Contract.pdf')[$Index % 4]
    $relative = 'Shared Documents'
    [pscustomobject]@{
        CreationTime          = Format-Stamp $EndDate.AddMinutes(-$random.Next(60, 60 * 24 * 90))
        Id                    = New-DeterministicGuid
        RecordType            = 'SharePointSharingOperation'
        Operation             = $Operation
        UserId                = $user.Upn
        Workload              = 'SharePoint'
        ObjectId              = '{0}/{1}/{2}' -f $site.WebUrl, $relative, $file
        ItemType              = 'File'
        SiteUrl               = $site.WebUrl
        SourceRelativeUrl     = $relative
        SourceFileName        = $file
        TargetUserOrGroupName = $Target
        TargetUserOrGroupType = $TargetType
        ClientIP              = "203.0.113.$(10 + $Index)"
    }
}

$anonymousEvents = foreach ($index in 0..19) {
    New-SampleEvent -Operation $schema.AnonymousLinkOperations[$index % 4] -Index $index
}
$anonymousEvents = @($anonymousEvents | Sort-Object CreationTime)
Export-AppendCsv -Path (Join-Path $OutputPath 'anonymous-link-events.csv') -Rows $anonymousEvents -Column $schema.AuditEvents -KeyColumn 'Id'

$sharingEvents = foreach ($index in 0..39) {
    $operation = $schema.SharingOperations[$index % $schema.SharingOperations.Count]
    $targetIsGuest = $operation -in 'SharingInvitationCreated', 'SharingInvitationAccepted', 'AddedToSecureLink'
    $target = if ($operation -in 'SharingSet', 'AddedToSecureLink', 'SharingInvitationCreated', 'SharingInvitationAccepted', 'AddedToGroup') {
        if ($targetIsGuest) { "guest$index#EXT#@example.net" } else { $people[($index + 3) % $people.Count].Upn }
    }
    else { '' }
    New-SampleEvent -Operation $operation -Index $index -Target $target -TargetType $(if (-not $target) { '' } elseif ($targetIsGuest) { 'Guest' } else { 'Member' })
}
$sharingEvents = @($sharingEvents | Sort-Object CreationTime)
Export-AppendCsv -Path (Join-Path $OutputPath 'sharing-events.csv') -Rows $sharingEvents -Column $schema.AuditEvents -KeyColumn 'Id'
