#Requires -Version 7.0

BeforeAll {
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
}

Describe 'README em dashes' {
    It 'leaves no U+2014 in a file whose name starts with README' {
        # BRO-427 removes the em dash and does not swap it for a hyphen.
        $hits = foreach ($file in (Get-ChildItem -Path $script:RepoRoot -Recurse -File -Filter 'README*')) {
            $text = Get-Content -LiteralPath $file.FullName -Raw
            if ($text.Contains([char]0x2014)) {
                $file.FullName.Substring($script:RepoRoot.Length).TrimStart('\', '/')
            }
        }

        ($hits -join ', ') | Should -BeNullOrEmpty
    }

    It 'joins the Pester Gallery note with a colon' {
        $readme = Get-Content -LiteralPath (Join-Path $script:RepoRoot 'README.md') -Raw
        $readme | Should -Match 'PowerShell Gallery as usual: that is unchanged'
        $readme | Should -Not -Match 'PowerShell Gallery as usual;'
    }

    It 'joins the preview-features note with a comma' {
        $readme = Get-Content -LiteralPath (Join-Path $script:RepoRoot 'reports/guest-access/README.md') -Raw
        $readme | Should -Match 'under Preview features, both are required to read this project'
        $readme | Should -Not -Match 'under Preview features;'
    }

    It 'joins the label GUID note with a colon' {
        $readme = Get-Content -LiteralPath (Join-Path $script:RepoRoot 'reports/purview-ip/README.md') -Raw
        $readme | Should -Match 'hold label GUIDs: join them to'
        $readme | Should -Not -Match 'hold label GUIDs;'
    }

    It 'ends the NotAvailable note with a period' {
        $readme = Get-Content -LiteralPath (Join-Path $script:RepoRoot 'reports/teams-groups-lifecycle/README.md') -Raw
        $readme | Should -Match 'skips the source\. It writes a header-only CSV and logs why'
        $readme | Should -Not -Match 'skips the source;'
    }
}
