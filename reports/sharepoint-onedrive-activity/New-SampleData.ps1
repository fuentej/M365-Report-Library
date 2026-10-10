#Requires -Version 7.0

<#
    .SYNOPSIS
        Regenerates samples/: fake CSVs with exactly the columns the collectors write.

    .DESCRIPTION
        Every name is invented and every address is on example.com. Columns come from
        collectors/SharePointOneDriveSchema.psd1, so a sample cannot drift from its collector.
        samples/gcchigh/ holds the header-only files a GCC High run leaves for the six Graph usage
        reports, which are not available in that cloud.

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

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'collectors/SharePointOneDriveSchema.psd1')

function Write-Sample {
    param([string]$Folder, [string]$Name, [string[]]$Column, [object[]]$Row)
    $path = Join-Path $Folder $Name
    if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
    Export-AppendCsv -Path $path -Rows $Row -Column $Column
}

if (-not (Test-Path -LiteralPath $OutputPath)) { New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null }
$gccHigh = Join-Path $OutputPath 'gcchigh'
if (-not (Test-Path -LiteralPath $gccHigh)) { New-Item -Path $gccHigh -ItemType Directory -Force | Out-Null }

$gb = 1GB
$people = @(
    @{ Name = 'Avery Abara'; Upn = 'avery.abara@example.com'; Used = 12 * $gb; Files = 4200; Active = 310; Last = '2026-10-07' }
    @{ Name = 'Blake Bishop'; Upn = 'blake.bishop@example.com'; Used = 46 * $gb; Files = 15800; Active = 40; Last = '2026-10-02' }
    @{ Name = 'Casey Cho'; Upn = 'casey.cho@example.com'; Used = 3 * $gb; Files = 900; Active = 0; Last = '' }
    @{ Name = 'Devon Diaz'; Upn = 'devon.diaz@example.com'; Used = 20 * $gb; Files = 7300; Active = 120; Last = '2026-10-06' }
)
$sites = @(
    @{ Id = 'contoso.sharepoint.example.com,3b1c5d2e-1111-4a2b-9c3d-000000000001,4f6a7b8c-1111-4d5e-8f90-000000000001'; Url = 'https://contoso.sharepoint.example.com/sites/finance'; Title = 'Finance'; Owner = 'Avery Abara'; OwnerUpn = 'avery.abara@example.com'; Template = 'GROUP#0'; Used = 85 * $gb; Files = 52000; Active = 1900; Pages = 410; Visited = 120; Last = '2026-10-08' }
    @{ Id = 'contoso.sharepoint.example.com,3b1c5d2e-2222-4a2b-9c3d-000000000002,4f6a7b8c-2222-4d5e-8f90-000000000002'; Url = 'https://contoso.sharepoint.example.com/sites/projects'; Title = 'Projects'; Owner = 'Blake Bishop'; OwnerUpn = 'blake.bishop@example.com'; Template = 'STS#3'; Used = 12 * $gb; Files = 9100; Active = 85; Pages = 30; Visited = 14; Last = '2026-09-18' }
    @{ Id = 'contoso.sharepoint.example.com,3b1c5d2e-3333-4a2b-9c3d-000000000003,4f6a7b8c-3333-4d5e-8f90-000000000003'; Url = 'https://contoso.sharepoint.example.com/sites/archive'; Title = 'Archive'; Owner = 'Casey Cho'; OwnerUpn = 'casey.cho@example.com'; Template = 'STS#3'; Used = 140 * $gb; Files = 210000; Active = 0; Pages = 0; Visited = 0; Last = '' }
)
$runDates = '2026-10-01', '2026-10-08'
$refresh = '2026-10-06'

# Source 1
$rows = foreach ($run in $runDates) {
    foreach ($s in $sites) {
        $growth = if ($run -eq '2026-10-08') { 1 * $gb } else { 0 }
        [pscustomobject]@{
            RunDate = $run; ReportRefreshDate = $refresh; SiteId = $s.Id; SiteUrl = $s.Url; OwnerDisplayName = $s.Owner
            IsDeleted = 'False'; LastActivityDate = $s.Last; FileCount = $s.Files; ActiveFileCount = $s.Active
            PageViewCount = $s.Pages; VisitedPageCount = $s.Visited; StorageUsedByte = $s.Used + $growth
            StorageAllocatedByte = 1024 * $gb; RootWebTemplate = $s.Template; OwnerPrincipalName = $s.OwnerUpn; ReportPeriod = 30
        }
    }
}
Write-Sample $OutputPath 'sharepoint-site-usage-detail.csv' $schema.SharePointSiteUsageDetail $rows

# Source 2
$rows = foreach ($run in $runDates) {
    foreach ($p in $people) {
        $deleted = if ($p.Name -eq 'Casey Cho') { 'True' } else { 'False' }
        [pscustomobject]@{
            RunDate = $run; ReportRefreshDate = $refresh; SiteUrl = "https://contoso-my.sharepoint.example.com/personal/$($p.Upn.Split('@')[0].Replace('.', '_'))_example_com"
            OwnerDisplayName = $p.Name; IsDeleted = $deleted; LastActivityDate = $p.Last; FileCount = $p.Files; ActiveFileCount = $p.Active
            StorageUsedByte = $p.Used; StorageAllocatedByte = 1024 * $gb; OwnerPrincipalName = $p.Upn; ReportPeriod = 30
        }
    }
}
Write-Sample $OutputPath 'onedrive-usage-account-detail.csv' $schema.OneDriveUsageAccountDetail $rows

# Sources 3 and 4
foreach ($pair in @(@('SharePointSiteUsageStorage', 'sharepoint-site-usage-storage.csv', 'All', 230), @('OneDriveUsageStorage', 'onedrive-usage-storage.csv', 'All', 80))) {
    $key, $name, $type, $baseGb = $pair
    $rows = foreach ($day in 1..6) {
        [pscustomobject]@{
            RunDate = '2026-10-08'; ReportRefreshDate = $refresh; SiteType = $type
            StorageUsedByte = ($baseGb + $day) * $gb; ReportDate = ('2026-10-{0:00}' -f $day); ReportPeriod = 30
        }
    }
    Write-Sample $OutputPath $name $schema[$key] $rows
}

# Sources 5 and 6. A period report and a one-day report in the same file.
$activity = @(
    @{ Upn = 'avery.abara@example.com'; View = 540; Sync = 120; Int = 14; Ext = 3; Pages = 40; Last = '2026-10-07'; Deleted = ''; DeletedOn = '' }
    @{ Upn = 'blake.bishop@example.com'; View = 12; Sync = 0; Int = 0; Ext = 0; Pages = 2; Last = '2026-10-02'; Deleted = 'False'; DeletedOn = '' }
    @{ Upn = 'casey.cho@example.com'; View = 0; Sync = 0; Int = 0; Ext = 0; Pages = 0; Last = ''; Deleted = 'True'; DeletedOn = '2026-09-20' }
)
foreach ($shape in @(@('SharePointActivityUserDetail', 'sharepoint-activity-user-detail.csv', $true), @('OneDriveActivityUserDetail', 'onedrive-activity-user-detail.csv', $false))) {
    $key, $name, $withPages = $shape
    $rows = foreach ($variant in @(@{ Query = ''; Period = '7'; Scale = 1 }, @{ Query = '2026-10-05'; Period = ''; Scale = 0 })) {
        foreach ($a in $activity) {
            $row = [ordered]@{
                RunDate = '2026-10-08'; QueryDate = $variant.Query; ReportRefreshDate = $refresh; UserPrincipalName = $a.Upn
                IsDeleted = $(if ($a.Deleted) { $a.Deleted } else { 'False' }); DeletedDate = $a.DeletedOn
                LastActivityDate = $a.Last
                ViewedOrEditedFileCount = $(if ($variant.Scale) { $a.View } else { [math]::Floor($a.View / 7) })
                SyncedFileCount = $(if ($variant.Scale) { $a.Sync } else { [math]::Floor($a.Sync / 7) })
                SharedInternallyFileCount = $(if ($variant.Scale) { $a.Int } else { [math]::Floor($a.Int / 7) })
                SharedExternallyFileCount = $(if ($variant.Scale) { $a.Ext } else { [math]::Floor($a.Ext / 7) })
            }
            if ($withPages) { $row['VisitedPageCount'] = $(if ($variant.Scale) { $a.Pages } else { [math]::Floor($a.Pages / 7) }) }
            $row['AssignedProducts'] = 'MICROSOFT 365 E3'
            $row['ReportPeriod'] = $variant.Period
            [pscustomobject]$row
        }
    }
    Write-Sample $OutputPath $name $schema[$key] $rows
}

# Source 7
$rows = foreach ($run in $runDates) { [pscustomobject]@{ RunDate = $run; DisplayConcealedNames = $(if ($run -eq '2026-10-01') { 'True' } else { 'False' }) } }
Write-Sample $OutputPath 'report-settings.csv' $schema.ReportSettings $rows

# Source 9, route (b)
$rows = foreach ($run in $runDates) {
    [pscustomobject]@{ RunDate = $run; StorageQuota = 5242880; StorageQuotaAllocated = 1048576; ResourceQuota = 300; ResourceQuotaAllocated = 0; OneDriveStorageQuota = 1048576 }
}
Write-Sample $OutputPath 'tenant-storage.csv' $schema.TenantStorage $rows

# Source 10. The storage columns are empty in the second row: the cmdlet may not return them without -Detailed.
$rows = foreach ($s in $sites) {
    [pscustomobject]@{ RunDate = '2026-10-08'; Url = $s.Url; Title = $s.Title; Template = $s.Template; StorageUsageCurrent = [math]::Floor($s.Used / 1MB); ResourceUsageCurrent = 0; WebsCount = 1 }
}
$rows += [pscustomobject]@{ RunDate = '2026-10-08'; Url = 'https://contoso-my.sharepoint.example.com/personal/avery_abara_example_com'; Title = 'Avery Abara'; Template = 'SPSPERS#10'; StorageUsageCurrent = ''; ResourceUsageCurrent = ''; WebsCount = '' }
Write-Sample $OutputPath 'spo-sites.csv' $schema.SpoSites $rows

# Source 11
$rows = [System.Collections.Generic.List[object]]::new()
foreach ($s in $sites) {
    $state = if ($s.Used -gt 100 * $gb) { 'nearing' } else { 'normal' }
    $rows.Add([pscustomobject]@{
            RunDate = '2026-10-08'; SiteId = $s.Id; SiteWebUrl = $s.Url; IsPersonalSite = 'False'
            DriveId = 'b!' + ($s.Title.ToLowerInvariant() + '-documents'); DriveName = 'Documents'; DriveType = 'documentLibrary'
            DriveWebUrl = $s.Url + '/Shared%20Documents'; QuotaState = $state; QuotaUsed = $s.Used; QuotaTotal = 1024 * $gb
            QuotaRemaining = 1024 * $gb - $s.Used; QuotaDeleted = 1 * $gb; LastModifiedDateTime = '2026-10-07T14:03:11Z'
        })
}
$rows.Add([pscustomobject]@{
        RunDate = '2026-10-08'; SiteId = 'contoso-my.sharepoint.example.com,5d2e6f7a-4444-4b3c-8d9e-000000000004,6a7b8c9d-4444-4e5f-9a0b-000000000004'
        SiteWebUrl = 'https://contoso-my.sharepoint.example.com/personal/avery_abara_example_com'; IsPersonalSite = 'True'
        DriveId = 'b!avery-onedrive'; DriveName = 'OneDrive'; DriveType = 'business'; DriveWebUrl = 'https://contoso-my.sharepoint.example.com/personal/avery_abara_example_com/Documents'
        QuotaState = 'normal'; QuotaUsed = 12 * $gb; QuotaTotal = 1024 * $gb; QuotaRemaining = 1012 * $gb; QuotaDeleted = 0; LastModifiedDateTime = '2026-10-07T09:30:00Z'
    })
Write-Sample $OutputPath 'drive-quota.csv' $schema.DriveQuota $rows.ToArray()

# Source 12. Whole UTC days, counted by the library.
$rows = foreach ($day in '2026-10-05', '2026-10-06', '2026-10-07') {
    foreach ($e in @(@('SharePoint', 'avery.abara@example.com', 'FileAccessed', 120), @('SharePoint', 'avery.abara@example.com', 'FileModified', 35), @('OneDrive', 'devon.diaz@example.com', 'FileSyncUploadedFull', 18), @('SharePoint', 'blake.bishop@example.com', 'PageViewed', 6), @('SharePoint', 'blake.bishop@example.com', 'SharingSet', 1))) {
        [pscustomobject]@{ Date = $day; Workload = $e[0]; UserId = $e[1]; Operation = $e[2]; EventCount = $e[3] }
    }
}
Write-Sample $OutputPath 'file-events.csv' $schema.FileEvents $rows

# Source 13. The first interval carries only access and edit; the third is based on incomplete data.
$s = $sites[0]
$rows = @(
    [pscustomobject]@{ RunDate = '2026-10-08'; SiteId = $s.Id; SiteWebUrl = $s.Url; IntervalStart = '2026-10-05T00:00:00Z'; IntervalEnd = '2026-10-06T00:00:00Z'
        AccessActionCount = 52; AccessActorCount = 11; CreateActionCount = ''; CreateActorCount = ''; EditActionCount = 9; EditActorCount = 4
        DeleteActionCount = ''; DeleteActorCount = ''; MoveActionCount = ''; MoveActorCount = ''; IncompleteData = 'False'; MissingDataBeforeDateTime = ''; WasThrottled = '' }
    [pscustomobject]@{ RunDate = '2026-10-08'; SiteId = $s.Id; SiteWebUrl = $s.Url; IntervalStart = '2026-10-06T00:00:00Z'; IntervalEnd = '2026-10-07T00:00:00Z'
        AccessActionCount = 61; AccessActorCount = 13; CreateActionCount = 3; CreateActorCount = 2; EditActionCount = 14; EditActorCount = 5
        DeleteActionCount = 1; DeleteActorCount = 1; MoveActionCount = 2; MoveActorCount = 1; IncompleteData = 'False'; MissingDataBeforeDateTime = ''; WasThrottled = '' }
    [pscustomobject]@{ RunDate = '2026-10-08'; SiteId = $s.Id; SiteWebUrl = $s.Url; IntervalStart = '2026-10-07T00:00:00Z'; IntervalEnd = '2026-10-08T00:00:00Z'
        AccessActionCount = 0; AccessActorCount = 0; CreateActionCount = ''; CreateActorCount = ''; EditActionCount = ''; EditActorCount = ''
        DeleteActionCount = ''; DeleteActorCount = ''; MoveActionCount = ''; MoveActorCount = ''; IncompleteData = 'True'; MissingDataBeforeDateTime = '2026-10-07T06:00:00Z'; WasThrottled = 'False' }
)
Write-Sample $OutputPath 'site-activity.csv' $schema.SiteActivity $rows

# The six Graph usage reports are not available in GCC High: header only.
$notAvailable = @{
    'sharepoint-site-usage-detail.csv'     = 'SharePointSiteUsageDetail'
    'onedrive-usage-account-detail.csv'    = 'OneDriveUsageAccountDetail'
    'sharepoint-site-usage-storage.csv'    = 'SharePointSiteUsageStorage'
    'onedrive-usage-storage.csv'           = 'OneDriveUsageStorage'
    'sharepoint-activity-user-detail.csv'  = 'SharePointActivityUserDetail'
    'onedrive-activity-user-detail.csv'    = 'OneDriveActivityUserDetail'
}
foreach ($name in $notAvailable.Keys) {
    Write-Sample $gccHigh $name $schema[$notAvailable[$name]] @()
}
