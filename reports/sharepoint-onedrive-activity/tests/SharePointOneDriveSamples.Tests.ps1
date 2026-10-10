#Requires -Version 7.0

BeforeAll {
    $script:Root = Join-Path $PSScriptRoot '../../..' | Resolve-Path | Select-Object -ExpandProperty Path
    $script:Collectors = Join-Path $script:Root 'reports/sharepoint-onedrive-activity/collectors'
    $script:Samples = Join-Path $script:Root 'reports/sharepoint-onedrive-activity/samples'

    Import-Module (Join-Path $script:Root 'shared/M365ReportLibrary.psm1') -Force
    $script:Schema = Import-PowerShellDataFile -LiteralPath (Join-Path $script:Collectors 'SharePointOneDriveSchema.psd1')

    # CSV name -> schema key. One row per source in docs/candidates/sharepoint-onedrive-activity.md that is a collector.
    $script:Map = [ordered]@{
        'sharepoint-site-usage-detail.csv'    = 'SharePointSiteUsageDetail'
        'onedrive-usage-account-detail.csv'   = 'OneDriveUsageAccountDetail'
        'sharepoint-site-usage-storage.csv'   = 'SharePointSiteUsageStorage'
        'onedrive-usage-storage.csv'          = 'OneDriveUsageStorage'
        'sharepoint-activity-user-detail.csv' = 'SharePointActivityUserDetail'
        'onedrive-activity-user-detail.csv'   = 'OneDriveActivityUserDetail'
        'report-settings.csv'                 = 'ReportSettings'
        'tenant-storage.csv'                  = 'TenantStorage'
        'spo-sites.csv'                       = 'SpoSites'
        'drive-quota.csv'                     = 'DriveQuota'
        'file-events.csv'                     = 'FileEvents'
        'site-activity.csv'                   = 'SiteActivity'
    }
    $script:Unavailable = 'SharePointSiteUsageDetail', 'OneDriveUsageAccountDetail', 'SharePointSiteUsageStorage', 'OneDriveUsageStorage', 'SharePointActivityUserDetail', 'OneDriveActivityUserDetail'
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
        # Sources 1 to 6: Graph usage reports are marked unsupported for US Government L4.
        foreach ($key in $script:Unavailable) {
            $a[$key].Commercial.Status | Should -Be 'Available'
            $a[$key].GCC.Status | Should -Be 'Available'
            $a[$key].GCCHigh.Status | Should -Be 'NotAvailable'
        }
        $a.ReportSettings.Commercial.Status | Should -Be 'Available'
        $a.ReportSettings.GCC.Status | Should -Be 'Available'
        $a.ReportSettings.GCCHigh.Status | Should -Be 'Unverified'
        foreach ($key in 'TenantStorage', 'SpoSites') {
            foreach ($cloud in 'Commercial', 'GCC', 'GCCHigh') { $a[$key][$cloud].Status | Should -Be 'Unverified' -Because "$key $cloud" }
        }
        foreach ($key in 'DriveQuota', 'FileEvents') {
            foreach ($cloud in 'Commercial', 'GCC', 'GCCHigh') { $a[$key][$cloud].Status | Should -Be 'Available' -Because "$key $cloud" }
        }
        $a.SiteActivity.Commercial.Status | Should -Be 'Available'
        $a.SiteActivity.GCC.Status | Should -Be 'Available'
        $a.SiteActivity.GCCHigh.Status | Should -Be 'Unverified'
    }

    It 'counts only operations the contract lists' {
        $script:Schema.FileOperations | Should -Be @('FileAccessed', 'FileModified', 'FileDownloaded', 'FileUploaded', 'FileSyncDownloadedFull', 'FileSyncUploadedFull', 'PageViewed', 'SharingSet', 'AnonymousLinkCreated', 'SecureLinkCreated')
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

    It 'covers a concealed and an unconcealed run in report-settings.csv' {
        (Import-Csv -LiteralPath (Join-Path $script:Samples 'report-settings.csv')).DisplayConcealedNames | Sort-Object -Unique | Should -Be @('False', 'True')
    }

    It 'shows an interval with incomplete data and an absent action in site-activity.csv' {
        $rows = @(Import-Csv -LiteralPath (Join-Path $script:Samples 'site-activity.csv'))
        @($rows | Where-Object IncompleteData -EQ 'True').Count | Should -BeGreaterThan 0
        @($rows | Where-Object { $_.CreateActionCount -eq '' }).Count | Should -BeGreaterThan 0
    }
}
