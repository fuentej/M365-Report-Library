#Requires -Version 7.0

BeforeAll {
    $script:Root = Join-Path $PSScriptRoot '../../..' | Resolve-Path | Select-Object -ExpandProperty Path
    $script:Collectors = Join-Path $script:Root 'reports/unified-audit-log/collectors'
    $script:Samples = Join-Path $script:Root 'reports/unified-audit-log/samples'

    Import-Module (Join-Path $script:Root 'shared/M365ReportLibrary.psm1') -Force
    $script:Schema = Import-PowerShellDataFile -LiteralPath (Join-Path $script:Collectors 'UnifiedAuditLogSchema.psd1')

    # CSV name -> schema key. One row per source in docs/candidates/unified-audit-log.md.
    $script:Map = [ordered]@{
        'audit-search-cmdlet.csv'      = 'AuditSearchCmdlet'
        'audit-graph-records.csv'      = 'AuditGraphRecords'
        'audit-activity-feed.csv'      = 'AuditActivityFeed'
        'audit-ingestion.csv'          = 'AuditIngestion'
        'audit-retention-policies.csv' = 'AuditRetentionPolicies'
    }
}

AfterAll {
    Remove-Module M365ReportLibrary -Force -ErrorAction SilentlyContinue
}

Describe 'The schema covers every source in the contract' {
    It 'has columns and availability for each of the five sources' {
        $script:Map.Count | Should -Be 5
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
        # Source 1, 4, 5: UNVERIFIED in GCC and GCC High.
        foreach ($key in 'AuditSearchCmdlet', 'AuditIngestion', 'AuditRetentionPolicies') {
            $a[$key].Commercial.Status | Should -Be 'Available' -Because $key
            $a[$key].GCC.Status | Should -Be 'Unverified' -Because $key
            $a[$key].GCCHigh.Status | Should -Be 'Unverified' -Because $key
        }
        # Source 2: NotAvailable in GCC High (US Government L4 is marked unsupported on every page read).
        $a.AuditGraphRecords.Commercial.Status | Should -Be 'Available'
        $a.AuditGraphRecords.GCC.Status | Should -Be 'Available'
        $a.AuditGraphRecords.GCCHigh.Status | Should -Be 'NotAvailable'
        # Source 3: a root URL is documented for all three clouds.
        foreach ($cloud in 'Commercial', 'GCC', 'GCCHigh') { $a.AuditActivityFeed[$cloud].Status | Should -Be 'Available' }
    }

    It 'uses the documented Management Activity API root for each cloud' {
        $script:Schema.ActivityFeedRoot.Commercial | Should -Be 'https://manage.office.com'
        $script:Schema.ActivityFeedRoot.GCC | Should -Be 'https://manage-gcc.office.com'
        $script:Schema.ActivityFeedRoot.GCCHigh | Should -Be 'https://manage.office365.us'
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

    It 'has a header-only sample for the source that is not available in GCC High' {
        $path = Join-Path $script:Samples 'gcchigh/audit-graph-records.csv'
        Test-Path -LiteralPath $path | Should -BeTrue
        (Get-CsvHeaderColumn -Path $path) -join ',' | Should -Be ($script:Schema.AuditGraphRecords -join ',')
        @(Import-Csv -LiteralPath $path).Count | Should -Be 0
    }

    It 'has a header-only sample for exactly the sources the contract marks NotAvailable' {
        $notAvailable = foreach ($key in $script:Map.Values) {
            if ($script:Schema.SourceAvailability[$key].GCCHigh.Status -eq 'NotAvailable') { $key }
        }
        @($notAvailable) | Should -Be @('AuditGraphRecords')
        @(Get-ChildItem -LiteralPath (Join-Path $script:Samples 'gcchigh') -Filter '*.csv').Count | Should -Be 1
    }

    It 'holds AuditData that is JSON and agrees with the row it sits in' {
        foreach ($name in 'audit-search-cmdlet.csv', 'audit-graph-records.csv', 'audit-activity-feed.csv') {
            foreach ($row in Import-Csv -LiteralPath (Join-Path $script:Samples $name)) {
                $data = $row.AuditData | ConvertFrom-Json
                $data.Id | Should -Be $row.RecordId -Because $name
                $data.Operation | Should -Be $row.Operation -Because $name
            }
        }
    }

    It 'uses no address outside example.com' {
        $text = foreach ($file in Get-ChildItem -LiteralPath $script:Samples -Filter '*.csv' -Recurse) { Get-Content -LiteralPath $file.FullName -Raw }
        $hosts = [regex]::Matches(($text -join "`n"), '(?:[A-Za-z0-9._%+-]+@|https?://)([A-Za-z0-9.-]+\.[A-Za-z]{2,})') | ForEach-Object { $_.Groups[1].Value.ToLowerInvariant() } | Sort-Object -Unique
        $outside = @($hosts | Where-Object { $_ -notmatch '(^|\.)example\.com$' })
        $outside -join '; ' | Should -BeNullOrEmpty
    }

    It 'has the retention durations the cmdlet names' {
        $durations = (Import-Csv -LiteralPath (Join-Path $script:Samples 'audit-retention-policies.csv')).RetentionDuration | Sort-Object -Unique
        $durations | ForEach-Object { $_ | Should -BeIn 'ThreeMonths', 'SixMonths', 'NineMonths', 'TwelveMonths', 'TenYears' }
    }
}
