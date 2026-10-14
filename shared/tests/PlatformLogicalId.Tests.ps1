#Requires -Version 7.0

BeforeAll {
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
    $script:Items = foreach ($file in (Get-ChildItem -Path (Join-Path $script:RepoRoot 'reports') -Recurse -File -Force -Filter '.platform')) {
        $relative = $file.FullName.Substring($script:RepoRoot.Length).TrimStart('\', '/')
        $parsed = $null
        try { $parsed = Get-Content -LiteralPath $file.FullName -Raw | ConvertFrom-Json -ErrorAction Stop } catch { }
        [pscustomobject]@{ Path = $relative; Json = $parsed; LogicalId = $parsed.config.logicalId }
    }
    $script:GuidShape = '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'
}

Describe 'Fabric .platform logicalId' {
    It 'finds at least one .platform file under reports/' {
        @($script:Items).Count | Should -BeGreaterThan 0
    }

    It 'parses every .platform file as JSON' {
        $bad = @($script:Items | Where-Object { $null -eq $_.Json } | ForEach-Object Path)
        ($bad -join ', ') | Should -BeNullOrEmpty
    }

    It 'gives every .platform file a GUID-shaped config.logicalId' {
        $bad = @($script:Items | Where-Object { $_.LogicalId -isnot [string] -or $_.LogicalId -notmatch $script:GuidShape } | ForEach-Object Path)
        ($bad -join ', ') | Should -BeNullOrEmpty
    }

    It 'never repeats a logicalId across .platform files' {
        $dupes = @($script:Items | Group-Object { "$($_.LogicalId)".ToLowerInvariant() } | Where-Object Count -gt 1 |
                ForEach-Object { "$($_.Name): $(($_.Group.Path) -join ' | ')" })
        ($dupes -join '; ') | Should -BeNullOrEmpty
    }
}
