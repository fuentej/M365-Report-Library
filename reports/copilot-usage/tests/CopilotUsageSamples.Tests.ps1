#Requires -Version 7.0

BeforeAll {
    $script:Root = Join-Path $PSScriptRoot '../../..' | Resolve-Path | Select-Object -ExpandProperty Path
    $script:Samples = Join-Path $script:Root 'reports/copilot-usage/samples'
    $script:Summary = @(Import-Csv -LiteralPath (Join-Path $script:Samples 'copilot-user-count-summary.csv'))
}

Describe 'The sample data behaves like a real tenant' {
    It 'lists at least 40 distinct people in the per-person usage file' {
        $people = @(Import-Csv -LiteralPath (Join-Path $script:Samples 'copilot-usage-user-detail.csv'))
        $people.Count | Should -BeGreaterOrEqual 40
        @($people.UserPrincipalName | Sort-Object -Unique).Count | Should -Be $people.Count
    }

    It 'has a different weekly-to-monthly active ratio for each Copilot app' {
        $weekly = $script:Summary | Where-Object ReportPeriod -EQ '7'
        $monthly = $script:Summary | Where-Object ReportPeriod -EQ '28'
        $ratios = foreach ($app in 'MicrosoftTeams', 'Word', 'PowerPoint', 'Outlook', 'Excel', 'OneNote', 'Loop', 'CopilotChat') {
            $month = [double]$monthly."$($app)ActiveUsers"
            $month | Should -BeGreaterThan 0
            [math]::Round(([double]$weekly."$($app)ActiveUsers") / $month, 2)
        }
        @($ratios | Sort-Object -Unique).Count | Should -Be $ratios.Count
        ($ratios | Measure-Object -Maximum).Maximum - ($ratios | Measure-Object -Minimum).Minimum | Should -BeGreaterThan 0.3
    }

    It 'counts the summary from the per-person file' {
        $people = @(Import-Csv -LiteralPath (Join-Path $script:Samples 'copilot-usage-user-detail.csv'))
        $refresh = [datetime]'2026-09-29'
        foreach ($row in $script:Summary) {
            $cutoff = $refresh.AddDays(1 - [int]$row.ReportPeriod)
            $active = @($people | Where-Object { $_.WordCopilotLastActivityDate -and [datetime]$_.WordCopilotLastActivityDate -ge $cutoff }).Count
            [int]$row.WordActiveUsers | Should -Be $active
            [int]$row.WordEnabledUsers | Should -Be $people.Count
        }
    }
}
