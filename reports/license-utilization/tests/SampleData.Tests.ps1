#Requires -Version 7.0

BeforeDiscovery {
    $script:SchemaForDiscovery = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot '../collectors/LicenseUtilizationSchema.psd1')
    $script:UsageCases = foreach ($key in $script:SchemaForDiscovery.UsageReports.Keys) { @{ Key = $key } }
}

BeforeAll {
    $script:Samples = Join-Path $PSScriptRoot '../samples'
    $script:Schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot '../collectors/LicenseUtilizationSchema.psd1')
    . (Join-Path $PSScriptRoot '../collectors/LicenseUtilizationHelpers.ps1')
    Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

    $script:StateFiles = @{
        'subscribed-skus.csv'      = 'SubscribedSkus'
        'sku-service-plans.csv'    = 'SkuServicePlans'
        'user-licenses.csv'        = 'UserLicenses'
        'license-details.csv'      = 'LicenseDetails'
        'user-signin-activity.csv' = 'UserSignInActivity'
        'report-settings.csv'      = 'ReportSettings'
        'access-reviews.csv'       = 'AccessReviewDefinitions'
        'attack-simulations.csv'   = 'AttackSimulations'
        'teams-pstn-calls.csv'     = 'TeamsPstnCalls'
    }
    $script:UsageFiles = @{
        ActiveUserUsage = 'usage-active-users.csv'; EmailActivityUsage = 'usage-email-activity.csv'
        TeamsActivityUsage = 'usage-teams-activity.csv'; SharePointActivityUsage = 'usage-sharepoint-activity.csv'
        OneDriveActivityUsage = 'usage-onedrive-activity.csv'; M365AppUsage = 'usage-m365-apps.csv'
        CopilotUsage = 'usage-copilot.csv'
    }
}

Describe 'The sample CSVs match the schema' {
    It '<File> has the columns the schema lists' -ForEach @(
        foreach ($file in 'subscribed-skus.csv', 'sku-service-plans.csv', 'user-licenses.csv', 'license-details.csv', 'user-signin-activity.csv', 'report-settings.csv', 'access-reviews.csv', 'attack-simulations.csv', 'teams-pstn-calls.csv') { @{ File = $file } }
    ) {
        $key = $script:StateFiles[$File]
        (Get-CsvHeaderColumn -Path (Join-Path $script:Samples $File)) -join ',' | Should -Be ($script:Schema[$key] -join ',')
    }

    It '<Key> has the columns derived from its usage-report headers, with rows' -ForEach $script:UsageCases {
        $path = Join-Path $script:Samples $script:UsageFiles[$Key]
        (Get-CsvHeaderColumn -Path $path) -join ',' | Should -Be ((Get-UsageColumn -Header $script:Schema.UsageReports[$Key]) -join ',')
        @(Import-Csv -LiteralPath $path).Count | Should -BeGreaterThan 0
    }

    It 'holds a header-only <Key> file for GCC High, where the doc marks it NotAvailable' -ForEach $script:UsageCases {
        $path = Join-Path $script:Samples "gcchigh/$($script:UsageFiles[$Key])"
        (Get-CsvHeaderColumn -Path $path) -join ',' | Should -Be ((Get-UsageColumn -Header $script:Schema.UsageReports[$Key]) -join ',')
        @(Import-Csv -LiteralPath $path).Count | Should -Be 0
    }

    It 'holds a header-only attack-simulations.csv for GCC High, where the doc marks it NotAvailable' {
        $path = Join-Path $script:Samples 'gcchigh/attack-simulations.csv'
        (Get-CsvHeaderColumn -Path $path) -join ',' | Should -Be ($script:Schema.AttackSimulations -join ',')
        @(Import-Csv -LiteralPath $path).Count | Should -Be 0
    }

    It 'has rows in each License value sample' -ForEach @(@{ File = 'access-reviews.csv' }, @{ File = 'attack-simulations.csv' }, @{ File = 'teams-pstn-calls.csv' }) {
        @(Import-Csv -LiteralPath (Join-Path $script:Samples $File)).Count | Should -BeGreaterThan 0
    }

    It 'uses only example.com addresses' {
        $text = Get-ChildItem -LiteralPath $script:Samples -Recurse -Filter '*.csv' | Get-Content -Raw
        $addresses = [regex]::Matches(($text -join "`n"), '[A-Za-z0-9._-]+@([A-Za-z0-9.-]+)') | ForEach-Object { $_.Groups[1].Value }
        $addresses | Should -Not -BeNullOrEmpty
        ($addresses | Where-Object { $_ -ne 'example.com' }) | Should -BeNullOrEmpty
    }
}
