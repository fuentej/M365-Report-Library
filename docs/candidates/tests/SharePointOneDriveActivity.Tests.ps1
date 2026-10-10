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

    It 'keeps the deleted-user windows and the storage units apart' {
        # The overview removes a deleted user in 30 days. The OneDrive usage page keeps them for 180.
        # Graph date is 30 days and its storage columns are bytes. The admin center day view is 28 days and megabytes.
        $row1 = @($script:Doc -split '\r?\n' | Where-Object { $_ -match '^\| 1 \| SharePoint storage' })
        $row2 = @($script:Doc -split '\r?\n' | Where-Object { $_ -match '^\| 2 \| OneDrive storage' })
        $question2 = @($script:Doc -split '\r?\n' | Where-Object { $_ -match '^\| 2 \| Files each user' })
        $flat = $script:Doc -replace '\s+', ' '
        $row1 | Should -Not -Match 'so plan for 28'
        $row2 | Should -Match 'temporarily empty'
        $row2 | Should -Match 'no Site Id'
        $flat | Should -Match 'do not drop OneDrive rows at 30 days'
        $flat | Should -Match 'perpetual license'
        $flat | Should -Match 'Storage used \(MB\)'
        $flat | Should -Match 'Do not add the two units'
        $question2 | Should -Match 'past 30 days'
        $question2 | Should -Not -Match 'last 28 days'
    }

    It 'follows the getAllSites nextLink onto OneDrive and does not treat a drive timestamp as last activity' {
        # The sample nextLink changes path. A rebuilt /sites/getAllSites drops OneDrive.
        # Drive lastModifiedDateTime is not the usage Last Activity Date. An incomplete activity interval is not zero.
        $row11 = @($script:Doc -split '\r?\n' | Where-Object { $_ -match '^\| 11 \| Site list' })
        $question3 = @($script:Doc -split '\r?\n' | Where-Object { $_ -match '^\| 3 \| Sites and OneDrive' })
        $flat = $script:Doc -replace '\s+', ' '
        $row11 | Should -Match 'oneDrive\.getAllSites'
        $row11 | Should -Match 'do not rebuild'
        $row11 | Should -Match 'lastModifiedDateTime'
        $row11 | Should -Match 'system facet'
        $flat | Should -Match 'getActivitiesByInterval'
        $flat | Should -Match 'incompleteData'
        $flat | Should -Match 'not yet available in all national deployments'
        $question3 | Should -Not -Match 'no page read shows a last-activity property'
    }
}
