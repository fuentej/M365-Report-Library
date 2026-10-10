#Requires -Version 7.0

BeforeAll {
    $script:DocPath = Join-Path $PSScriptRoot '../sharepoint-onedrive-activity.md'
    $script:Doc = Get-Content -LiteralPath $script:DocPath -Raw
}

Describe 'SharePoint and OneDrive activity sources' {
    It 'counts every SharePoint site last-activity event group' {
        # A shortened event list marks a site inactive after a delete or a download.
        # Activity-report last activity is for the selected range, and Is Deleted is a removed licence.
        $row1 = @($script:Doc -split '\r?\n' | Where-Object { $_ -match '^\| 1 \| SharePoint storage' })
        $row5 = @($script:Doc -split '\r?\n' | Where-Object { $_ -match '^\| 5 \| SharePoint activity' })
        $row6 = @($script:Doc -split '\r?\n' | Where-Object { $_ -match '^\| 6 \| OneDrive activity' })
        $row1 | Should -Match 'ActiveFiles'
        $row1 | Should -Match 'FileDeleted'
        $row1 | Should -Not -Match 'anonymous-link or company-link creation'
        $row5 | Should -Match 'for the selected date range'
        $row5 | Should -Match 'license was removed'
        $row6 | Should -Match 'for the selected date range'
        $flat = $script:Doc -replace '\s+', ' '
        $flat | Should -Match 'regardless of the selected time period'
        $flat | Should -Match 'rows 5 and 6 must not be read as lifetime last activity'
    }
}
