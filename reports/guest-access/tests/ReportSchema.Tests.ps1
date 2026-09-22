#Requires -Version 7.0

<#
    Every PBIR JSON file the guest-access report ships (report.json, version.json,
    pages.json, each page.json, each visual.json, definition.pbir, definition.pbism,
    and the two .platform files) declares a `$schema` property naming the exact
    Power BI / Fabric schema it is written against. This test validates each file
    against the schema it names.

    The schemas are vendored under tests/schemas/, mirroring their path in
    https://github.com/microsoft/json-schemas, so this runs fully offline -- no
    network access, no external validator package -- which is what keeps it safe
    to run in CI. SchemaValidator.ps1 documents exactly what "validates" checks.
#>

# This data has to be gathered here, at the top level of the file, rather than in a
# BeforeAll: Pester builds the "validates <RelativePath>" test below with -ForEach,
# and -ForEach is evaluated during Discovery (as the file is dot-sourced), before
# any BeforeAll runs. Everything here is read-only filesystem/JSON inspection, so
# doing it at Discovery time is safe. It does not call anything from
# SchemaValidator.ps1, so it does not need that file dot-sourced yet.
$script:Root = Join-Path $PSScriptRoot '../../..' | Resolve-Path | Select-Object -ExpandProperty Path
$script:ReportRoot = Join-Path $script:Root 'reports/guest-access/report'
$script:SchemaRoot = (Resolve-Path (Join-Path $PSScriptRoot 'schemas')).ProviderPath

# -Force: a `.platform` file is a dotfile and Get-ChildItem hides those by default
# even though only the *name*, not a separate extension, starts with the dot.
$script:AllFiles = @(Get-ChildItem -LiteralPath $script:ReportRoot -Recurse -File -Force |
        Where-Object { $_.Extension -eq '.json' -or $_.Name -eq '.platform' })

$script:SchemaBearingFiles = @($script:AllFiles | Where-Object {
        $data = Get-Content -LiteralPath $_.FullName -Raw | ConvertFrom-Json -AsHashtable
        $data.ContainsKey('$schema')
    })

# Pester only *executes* top-level script code (everything above, outside a block)
# during Discovery, to build the test tree; none of those script-scoped variables
# survive into the Run phase, where the It/BeforeAll bodies below actually execute.
# -ForEach values are the one exception -- Pester closes over them explicitly. So
# every one of those variables, and the SchemaValidator.ps1 dot-source, has to be
# redone here to exist when Run-phase code below reads them.
BeforeAll {
    . (Join-Path $PSScriptRoot 'SchemaValidator.ps1')

    $script:Root = Join-Path $PSScriptRoot '../../..' | Resolve-Path | Select-Object -ExpandProperty Path
    $script:ReportRoot = Join-Path $script:Root 'reports/guest-access/report'
    $script:SchemaRoot = (Resolve-Path (Join-Path $PSScriptRoot 'schemas')).ProviderPath

    $script:AllFiles = @(Get-ChildItem -LiteralPath $script:ReportRoot -Recurse -File -Force |
            Where-Object { $_.Extension -eq '.json' -or $_.Name -eq '.platform' })

    $script:SchemaBearingFiles = @($script:AllFiles | Where-Object {
            $data = Get-Content -LiteralPath $_.FullName -Raw | ConvertFrom-Json -AsHashtable
            $data.ContainsKey('$schema')
        })
}

Describe 'Every PBIR/platform JSON file declares and validates against its schema' {
    It 'finds report files to check' {
        $script:AllFiles.Count | Should -BeGreaterThan 30
    }

    It 'declares a `$schema` on every JSON/`.platform` file this report ships' {
        # If a file is missing `$schema` entirely, the per-file test below would
        # never run for it -- so this closes that gap explicitly.
        $missing = @($script:AllFiles | Where-Object { $script:SchemaBearingFiles.FullName -notcontains $_.FullName } |
                ForEach-Object { $_.FullName.Substring($script:ReportRoot.Length + 1) })
        $missing -join '; ' | Should -BeNullOrEmpty
    }

    It 'has a vendored schema for every `$schema` URL referenced' {
        $urls = @($script:SchemaBearingFiles | ForEach-Object {
                (Get-Content -LiteralPath $_.FullName -Raw | ConvertFrom-Json -AsHashtable)['$schema']
            } | Sort-Object -Unique)
        $urls.Count | Should -BeGreaterThan 0

        $prefix = 'https://developer.microsoft.com/json-schemas/'
        $missing = @($urls | Where-Object {
                -not $_.StartsWith($prefix) -or -not (Test-Path -LiteralPath (Join-Path $script:SchemaRoot $_.Substring($prefix.Length)))
            })
        $missing -join '; ' | Should -BeNullOrEmpty
    }

    It 'validates <RelativePath> against its declared schema' -ForEach @(
        $script:SchemaBearingFiles | ForEach-Object {
            @{ RelativePath = $_.FullName.Substring($script:ReportRoot.Length + 1); FullPath = $_.FullName }
        }
    ) {
        $errors = Test-PbirSchema -Path $FullPath -SchemaRoot $script:SchemaRoot
        $errors -join '; ' | Should -BeNullOrEmpty
    }
}

Describe 'The schema validator actually catches structural problems' {
    <#
        A validator that never fails is not proving anything about the files above
        -- these cases confirm Test-JsonSchema flags a missing required property,
        a disallowed extra one, and a wrong type, using the report's own real
        visualContainer schema and one of its own real visual.json files as the
        base case.
    #>
    BeforeAll {
        $script:SampleVisualPath = Join-Path $script:ReportRoot `
            'GuestAccess.Report/definition/pages/overview/visuals/card-guest-count/visual.json'
        $script:SampleVisual = Get-Content -LiteralPath $script:SampleVisualPath -Raw | ConvertFrom-Json -AsHashtable
    }

    It 'passes the unmodified file' {
        (Test-PbirSchema -Path $script:SampleVisualPath -SchemaRoot $script:SchemaRoot).Count | Should -Be 0
    }

    It 'flags a missing required property' {
        $broken = [hashtable]$script:SampleVisual.Clone()
        $broken.Remove('name')
        $tmp = New-TemporaryFile
        try {
            $broken | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $tmp
            $errors = Test-PbirSchema -Path $tmp.FullName -SchemaRoot $script:SchemaRoot
            $errors.Count | Should -BeGreaterThan 0
            $errors -join '; ' | Should -Match "missing required property 'name'"
        }
        finally {
            Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
        }
    }

    It 'flags a disallowed additional property' {
        $broken = [hashtable]$script:SampleVisual.Clone()
        $broken['notARealProperty'] = 'nope'
        $tmp = New-TemporaryFile
        try {
            $broken | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $tmp
            $errors = Test-PbirSchema -Path $tmp.FullName -SchemaRoot $script:SchemaRoot
            $errors -join '; ' | Should -Match 'not allowed'
        }
        finally {
            Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
        }
    }

    It 'flags the wrong type for position.width' {
        $broken = [hashtable]$script:SampleVisual.Clone()
        $broken['position'] = [hashtable]$broken['position'].Clone()
        $broken['position']['width'] = 'not-a-number'
        $tmp = New-TemporaryFile
        try {
            $broken | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $tmp
            $errors = Test-PbirSchema -Path $tmp.FullName -SchemaRoot $script:SchemaRoot
            $errors -join '; ' | Should -Match 'expected type'
        }
        finally {
            Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
        }
    }
}
