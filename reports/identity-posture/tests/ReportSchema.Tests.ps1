#Requires -Version 7.0

<#
    Every PBIR JSON file the identity-posture report ships (report.json, version.json,
    pages.json, each page.json, each visual.json, definition.pbir, definition.pbism and
    the two .platform files) declares a `$schema` property naming the exact Power BI /
    Fabric schema it is written against. This test validates each file against the
    schema it names.

    The validator (SchemaValidator.ps1) and the vendored schemas (tests/schemas/, which
    mirror their path in https://github.com/microsoft/json-schemas) are the guest-access
    report's: both reports are written against the same schema versions, and one copy
    keeps them from drifting apart. Everything runs offline.

    Unlike the guest-access test, this one also covers definition.pbir and
    definition.pbism, which are JSON files without a `.json` extension.
#>

# Gathered at the top level because Pester evaluates -ForEach while it dot-sources the
# file (Discovery), before any BeforeAll runs. Read-only filesystem and JSON inspection.
$script:Root = Join-Path $PSScriptRoot '../../..' | Resolve-Path | Select-Object -ExpandProperty Path
$script:ReportRoot = Join-Path $script:Root 'reports/identity-posture/report'
$script:SchemaRoot = (Resolve-Path (Join-Path $PSScriptRoot '../../guest-access/tests/schemas')).ProviderPath

# -Force: a `.platform` file is a dotfile and Get-ChildItem hides those by default.
# StaticResources/ holds registered resources, not PBIR: the house-style theme (BRO-501)
# names the Power BI report theme schema, which is not vendored here, and the SVG icons
# have no schema at all. `powerbi-report-author validate` checks the theme.
$script:AllFiles = @(Get-ChildItem -LiteralPath $script:ReportRoot -Recurse -File -Force |
        Where-Object { ($_.Extension -in '.json', '.pbir', '.pbism' -or $_.Name -eq '.platform') -and $_.FullName -notmatch '[\\/]StaticResources[\\/]' })

BeforeAll {
    . (Join-Path $PSScriptRoot '../../guest-access/tests/SchemaValidator.ps1')

    $script:Root = Join-Path $PSScriptRoot '../../..' | Resolve-Path | Select-Object -ExpandProperty Path
    $script:ReportRoot = Join-Path $script:Root 'reports/identity-posture/report'
    $script:SchemaRoot = (Resolve-Path (Join-Path $PSScriptRoot '../../guest-access/tests/schemas')).ProviderPath

    $script:AllFiles = @(Get-ChildItem -LiteralPath $script:ReportRoot -Recurse -File -Force |
            Where-Object { ($_.Extension -in '.json', '.pbir', '.pbism' -or $_.Name -eq '.platform') -and $_.FullName -notmatch '[\\/]StaticResources[\\/]' })

    $script:SchemaBearingFiles = @($script:AllFiles | Where-Object {
            $data = Get-Content -LiteralPath $_.FullName -Raw | ConvertFrom-Json -AsHashtable
            $data.ContainsKey('$schema')
        })
}

Describe 'Every PBIR/platform JSON file declares and validates against its schema' {
    It 'finds report files to check' {
        $script:AllFiles.Count | Should -BeGreaterThan 30
    }

    It 'ships the project, the model definition and the report definition files' {
        $names = $script:AllFiles.Name
        $names | Should -Contain 'definition.pbir'
        $names | Should -Contain 'definition.pbism'
        $names | Should -Contain 'report.json'
        $names | Should -Contain 'pages.json'
        @($names | Where-Object { $_ -eq '.platform' }).Count | Should -Be 2
    }

    It 'declares a `$schema` on every JSON/`.platform` file this report ships' {
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
        $script:AllFiles | ForEach-Object {
            @{ RelativePath = $_.FullName.Substring($script:ReportRoot.Length + 1); FullPath = $_.FullName }
        }
    ) {
        $errors = Test-PbirSchema -Path $FullPath -SchemaRoot $script:SchemaRoot
        $errors -join '; ' | Should -BeNullOrEmpty
    }
}

Describe 'The report is laid out as the one approved page' {
    # BRO-501 replaced the seven pages the contract proposed (Overview ... Risky users), each
    # with its own date, entity and anonymize slicers, with the approved Accounts to fix page.
    # The design has no slicers; the anonymize toggle is a report-level filter (ReportModel.Tests.ps1).
    BeforeAll {
        $script:PagesFolder = Join-Path $script:ReportRoot 'IdentityPosture.Report/definition/pages'
        $script:PageOrder = (Get-Content -LiteralPath (Join-Path $script:PagesFolder 'pages.json') -Raw | ConvertFrom-Json -AsHashtable).pageOrder
    }

    It 'lists the approved page only' {
        $script:PageOrder -join ',' | Should -Be 'accounts-to-fix'
    }

    It 'has a page.json for every listed page and no unlisted page folder' {
        $folders = @(Get-ChildItem -LiteralPath $script:PagesFolder -Directory | ForEach-Object Name | Sort-Object)
        $folders -join ',' | Should -Be (($script:PageOrder | Sort-Object) -join ',')
        foreach ($page in $script:PageOrder) {
            (Join-Path $script:PagesFolder "$page/page.json") | Should -Exist
        }
    }

    It 'registers the house-style theme and every image the page uses' {
        $report = Get-Content -LiteralPath (Join-Path $script:ReportRoot 'IdentityPosture.Report/definition/report.json') -Raw | ConvertFrom-Json -AsHashtable
        $theme = $report.themeCollection.customTheme.name
        $theme | Should -Match '^BacklogKrewHouseStyle-[0-9a-f]{8,12}\.json$'
        $items = @($report.resourcePackages | Where-Object name -EQ 'RegisteredResources' | ForEach-Object { $_.items })
        $resources = Join-Path $script:ReportRoot 'IdentityPosture.Report/StaticResources/RegisteredResources'
        foreach ($item in $items) { Join-Path $resources $item.path | Should -Exist }
        ($items | Where-Object type -EQ 'CustomTheme').name | Should -Be $theme
        (Get-Content -LiteralPath (Join-Path $resources $theme) -Raw | ConvertFrom-Json).name | Should -Be $theme
    }
}

Describe 'The schema validator flags a broken identity-posture visual' {
    It 'rejects a visual that lost its required name' {
        $path = Join-Path $script:ReportRoot 'IdentityPosture.Report/definition/pages/accounts-to-fix/visuals/fix-tile-1/visual.json'
        (Test-PbirSchema -Path $path -SchemaRoot $script:SchemaRoot).Count | Should -Be 0

        $broken = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json -AsHashtable
        $broken.Remove('name')
        $tmp = New-TemporaryFile
        try {
            $broken | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $tmp
            $errors = Test-PbirSchema -Path $tmp.FullName -SchemaRoot $script:SchemaRoot
            $errors -join '; ' | Should -Match "missing required property 'name'"
        }
        finally {
            Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
        }
    }
}
