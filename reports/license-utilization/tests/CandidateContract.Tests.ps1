#Requires -Version 7.0

BeforeAll {
    $script:Doc = Get-Content -LiteralPath (Join-Path $PSScriptRoot '../../../docs/candidates/license-utilization.md') -Raw
    $script:Schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot '../collectors/LicenseUtilizationSchema.psd1')
}

Describe 'The collectors follow the contract doc' {
    It 'has a collector script for each tenant source row (1 to 6 and 5a to 5g)' {
        foreach ($script in 'Get-SubscribedSkus', 'Get-UserLicenses', 'Get-LicenseDetails', 'Get-UserSignInActivity', 'Get-ReportSettings',
            'Get-ActiveUserUsage', 'Get-EmailActivityUsage', 'Get-TeamsActivityUsage', 'Get-SharePointActivityUsage',
            'Get-OneDriveActivityUsage', 'Get-M365AppUsage', 'Get-CopilotUsage') {
            Test-Path -LiteralPath (Join-Path $PSScriptRoot "../collectors/$script.ps1") | Should -BeTrue
        }
    }

    It 'has an availability entry for every source with all three clouds' {
        $script:Schema.SourceAvailability.Keys.Count | Should -Be 12
        foreach ($entry in $script:Schema.SourceAvailability.GetEnumerator()) {
            foreach ($cloud in 'Commercial', 'GCC', 'GCCHigh') {
                $entry.Value[$cloud].Status | Should -BeIn @('Available', 'NotAvailable', 'Unverified')
            }
        }
    }

    It 'marks the usage reports NotAvailable in GCC High, as the doc does' {
        foreach ($key in $script:Schema.UsageReports.Keys) {
            $script:Schema.SourceAvailability[$key].GCCHigh.Status | Should -Be 'NotAvailable'
            $script:Schema.SourceAvailability[$key].Commercial.Status | Should -Be 'Available'
        }
    }

    It 'marks signInActivity and report settings UNVERIFIED in GCC High, and the doc says so' {
        $script:Schema.SourceAvailability.UserSignInActivity.GCCHigh.Status | Should -Be 'Unverified'
        $script:Schema.SourceAvailability.ReportSettings.GCCHigh.Status | Should -Be 'Unverified'
        $row4 = @($script:Doc -split "`n" | Where-Object { $_ -match '^\| 4 \|' })[0]
        $row6 = @($script:Doc -split "`n" | Where-Object { $_ -match '^\| 6 \|' })[0]
        ($row4 -split '\|')[-2].Trim() | Should -Match '^UNVERIFIED'
        ($row6 -split '\|')[-2].Trim() | Should -Match '^UNVERIFIED'
    }

    It 'names the Copilot signed-in roles from the /copilot page, which does not list Global Reader' {
        $readme = Get-Content -LiteralPath (Join-Path $PSScriptRoot '../README.md') -Raw
        $readme | Should -Match 'AI Administrator'
        $readme | Should -Match 'does not list Global Reader'
    }

    It 'is on the CI Pester path, so a failure in this report fails the workflow' {
        # tests.yml lists paths explicitly. A merge that drops this folder leaves the
        # collectors untested while the workflow stays green.
        $workflow = Get-Content -LiteralPath (Join-Path $PSScriptRoot '../../../.github/workflows/tests.yml') -Raw
        $workflow | Should -Match '\./reports/license-utilization/tests'
    }

    It 'leaves out the two doc rows that are not tenant calls (7 and 8)' {
        $script:Doc | Should -Match '(?m)^\| 7 \| Product names'
        $script:Doc | Should -Match '(?m)^\| 8 \| Fallback for GCC High'
    }
}
