#Requires -Version 7.0

BeforeAll {
    $script:ThemePath = Join-Path $PSScriptRoot '../theme/BacklogKrewHouseStyle.json'
    $script:ThemeText = Get-Content -LiteralPath $script:ThemePath -Raw
    $script:Theme = $script:ThemeText | ConvertFrom-Json -AsHashtable
}

Describe 'Backlog Krew house-style theme' {
    It 'parses as JSON' {
        { $script:ThemeText | ConvertFrom-Json -ErrorAction Stop } | Should -Not -Throw
        $script:Theme | Should -Not -BeNullOrEmpty
    }

    It 'sets the theme name' {
        $script:Theme['name'] | Should -Be 'Backlog Krew house style'
    }

    It 'points $schema at a published report theme schema' {
        $script:Theme['$schema'] | Should -Match '^https://raw\.githubusercontent\.com/microsoft/powerbi-desktop-samples/main/Report%20Theme%20JSON%20Schema/reportThemeSchema-\d+\.\d+\.json$'
    }

    It 'has at least 6 data colours, all #RRGGBB' {
        $colors = @($script:Theme['dataColors'])
        $colors.Count | Should -BeGreaterOrEqual 6
        foreach ($color in $colors) {
            $color | Should -Match '^#[0-9A-Fa-f]{6}$'
        }
    }

    It 'puts teal first and navy second, as on the approved pages' {
        $script:Theme['dataColors'][0] | Should -Be '#138A8A'
        $script:Theme['dataColors'][1] | Should -Be '#16324A'
    }

    It 'turns the card title on by default' {
        # BRO-481: the legacy card shows no title unless it is turned on, and
        # categoryLabels did not render reliably in Desktop.
        $script:Theme['visualStyles']['card']['*']['title'][0]['show'] | Should -BeTrue
    }

    It 'sets the page background' {
        $background = $script:Theme['visualStyles']['page']['*']['background'][0]
        $background['color']['solid']['color'] | Should -Be '#EEF1F4'
    }

    It 'tells a PBIR report to register a GUID-suffixed copy whose name matches its file name' {
        $readme = Get-Content -LiteralPath (Join-Path $PSScriptRoot '../theme/README.md') -Raw
        $readme | Should -Match 'BacklogKrewHouseStyle-<guid>\.json'
        $readme | Should -Match 'customTheme\.name'
    }

    It 'keeps every colour in the file as #RRGGBB' {
        $colors = [regex]::Matches($script:ThemeText, '"#[^"]*"') | ForEach-Object { $_.Value.Trim('"') }
        $colors | Should -Not -BeNullOrEmpty
        foreach ($color in $colors) {
            $color | Should -Match '^#[0-9A-Fa-f]{6}$'
        }
    }
}
