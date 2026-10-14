#Requires -Version 7.0

BeforeAll {
    $script:Root = Join-Path $PSScriptRoot '../../..' | Resolve-Path | Select-Object -ExpandProperty Path
    $script:Collectors = Join-Path $script:Root 'reports/exchange-activity/collectors'
    $script:Samples = Join-Path $script:Root 'reports/exchange-activity/samples'

    Import-Module (Join-Path $script:Root 'shared/M365ReportLibrary.psm1') -Force
    $script:Schema = Import-PowerShellDataFile -LiteralPath (Join-Path $script:Collectors 'ExchangeActivitySchema.psd1')

    # CSV name -> schema key. One row per source in docs/candidates/exchange-activity.md that is a collector.
    $script:Map = [ordered]@{
        'mailbox-usage-detail.csv'          = 'MailboxUsageDetail'
        'mailbox-usage-storage.csv'         = 'MailboxUsageStorage'
        'email-activity-user-detail.csv'    = 'EmailActivityUserDetail'
        'email-app-usage-user-detail.csv'   = 'EmailAppUsageUserDetail'
        'report-settings.csv'               = 'ReportSettings'
        'mailboxes.csv'                     = 'Mailboxes'
        'mailbox-statistics.csv'            = 'MailboxStatistics'
        'message-trace.csv'                 = 'MessageTrace'
        'mobile-devices.csv'                = 'MobileDevices'
        'graph-message-trace.csv'           = 'GraphMessageTrace'
    }
    $script:Unavailable = 'MailboxUsageDetail', 'MailboxUsageStorage', 'EmailActivityUserDetail', 'EmailAppUsageUserDetail'
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
        # Sources 1 to 4: Graph usage reports are marked unsupported for US Government L4.
        foreach ($key in $script:Unavailable) {
            $a[$key].Commercial.Status | Should -Be 'Available'
            $a[$key].GCC.Status | Should -Be 'Available'
            $a[$key].GCCHigh.Status | Should -Be 'NotAvailable'
        }
        $a.ReportSettings.GCCHigh.Status | Should -Be 'Unverified'
        foreach ($key in 'Mailboxes', 'MailboxStatistics', 'MessageTrace', 'MobileDevices') {
            $a[$key].Commercial.Status | Should -Be 'Available'
            $a[$key].GCC.Status | Should -Be 'Unverified'
            $a[$key].GCCHigh.Status | Should -Be 'Unverified'
        }
        foreach ($cloud in 'Commercial', 'GCC', 'GCCHigh') { $a.GraphMessageTrace[$cloud].Status | Should -Be 'Unverified' }
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

    It 'covers every quota status in mailbox-usage-detail.csv' {
        $statuses = (Import-Csv -LiteralPath (Join-Path $script:Samples 'mailbox-usage-detail.csv')).QuotaStatus | Sort-Object -Unique
        $statuses | Should -Be @('CantSend', 'CantSendReceive', 'Good', 'Warning')
    }
}

Describe 'The sample data behaves like a real tenant' {
    BeforeAll {
        $script:Detail = @(Import-Csv -LiteralPath (Join-Path $script:Samples 'mailbox-usage-detail.csv'))
        $script:Storage = @(Import-Csv -LiteralPath (Join-Path $script:Samples 'mailbox-usage-storage.csv'))
        $script:Activity = @(Import-Csv -LiteralPath (Join-Path $script:Samples 'email-activity-user-detail.csv'))
    }

    It 'grows mailboxes at different rates between the two runs' {
        $growth = foreach ($upn in ($script:Detail.UserPrincipalName | Sort-Object -Unique)) {
            $rows = @($script:Detail | Where-Object UserPrincipalName -EQ $upn | Sort-Object RunDate)
            [int64]$rows[-1].StorageUsedByte - [int64]$rows[0].StorageUsedByte
        }
        @($growth | Sort-Object -Unique).Count | Should -BeGreaterThan 3
        @($growth | Where-Object { $_ -gt 0 }).Count | Should -BeGreaterThan 0
    }

    It 'does not repeat the tenant storage trend from one week to the next' {
        $windows = $script:Storage | Group-Object RunDate
        $windows.Count | Should -BeGreaterOrEqual 2
        $shapes = foreach ($window in $windows) { ($window.Group | ForEach-Object { $_.StorageUsedByte }) -join ',' }
        @($shapes | Sort-Object -Unique).Count | Should -Be $windows.Count
        $values = @($script:Storage | Sort-Object ReportDate | ForEach-Object { [int64]$_.StorageUsedByte })
        $values | Should -Be @($values | Sort-Object)
        @($values | Sort-Object -Unique).Count | Should -Be $values.Count
    }

    It 'shows no mail received for a mailbox at or over its receive limit' {
        $latest = ($script:Detail.RunDate | Sort-Object -Descending | Select-Object -First 1)
        $over = @($script:Detail | Where-Object { $_.RunDate -eq $latest -and [decimal]$_.StorageUsedByte -ge [decimal]$_.ProhibitSendReceiveQuotaByte })
        $over.Count | Should -BeGreaterThan 0
        foreach ($mailbox in $over) {
            $row = $script:Activity | Where-Object UserPrincipalName -EQ $mailbox.UserPrincipalName
            [int]$row.ReceiveCount | Should -Be 0 -Because $mailbox.UserPrincipalName
            $delivered = @(Import-Csv -LiteralPath (Join-Path $script:Samples 'message-trace.csv') | Where-Object { $_.RecipientAddress -eq $mailbox.UserPrincipalName -and $_.Status -eq 'Delivered' })
            $delivered.Count | Should -Be 0
        }
        @($script:Activity | Where-Object { [int]$_.ReceiveCount -gt 0 }).Count | Should -BeGreaterThan 0
    }
}
