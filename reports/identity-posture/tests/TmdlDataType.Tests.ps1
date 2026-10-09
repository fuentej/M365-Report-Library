#Requires -Version 7.0

<#
    Power BI Desktop refuses to open a project whose TMDL holds a dataType the parser
    does not know ("Failed to convert the value 'logical' to the expected type
    DataType"). `logical` is a Power Query type name, not a TMDL dataType; the TMDL
    value is `boolean`. Nothing else in the suite parses TMDL the way Desktop does, so
    this reads every `dataType:` line and holds it to the valid set.

    The valid set is the one BRO-324 names: string, int64, double, decimal, dateTime,
    boolean, binary, variant. TMDL reads values case-insensitively
    (https://learn.microsoft.com/analysis-services/tmdl/tmdl-overview#casing).
    A `type logical` inside a partition's M expression is valid M and is not a
    dataType line, so it is not checked here.
#>

BeforeDiscovery {
    $script:DefinitionFolder = Join-Path $PSScriptRoot '../report/IdentityPosture.SemanticModel/definition' | Resolve-Path | Select-Object -ExpandProperty Path
    $script:TmdlFiles = Get-ChildItem -LiteralPath $script:DefinitionFolder -Recurse -File -Filter '*.tmdl' |
        ForEach-Object { @{ Name = $_.FullName.Substring($script:DefinitionFolder.Length + 1); Path = $_.FullName } }
}

BeforeAll {
    $script:DefinitionFolder = Join-Path $PSScriptRoot '../report/IdentityPosture.SemanticModel/definition' | Resolve-Path | Select-Object -ExpandProperty Path
    $script:ValidDataTypes = @('string', 'int64', 'double', 'decimal', 'dateTime', 'boolean', 'binary', 'variant')

    function Get-TmdlDataType {
        <#
            .SYNOPSIS
                Every `dataType:` property in a TMDL text, with its line number.
        #>
        param([Parameter(Mandatory)][AllowEmptyString()][string[]]$Line)

        for ($i = 0; $i -lt $Line.Count; $i++) {
            if ($Line[$i] -match '^\s*datatype\s*:\s*(?<value>.*?)\s*$') {
                [pscustomobject]@{ LineNumber = $i + 1; Value = $Matches['value'] }
            }
        }
    }

    function Get-InvalidTmdlDataType {
        param([Parameter(Mandatory)][AllowEmptyString()][string[]]$Line)

        Get-TmdlDataType -Line $Line | Where-Object { $script:ValidDataTypes -inotcontains $_.Value }
    }
}

Describe 'TMDL dataType values' {
    It 'finds dataType lines in the semantic model, so the scan is not vacuous' {
        $files = Get-ChildItem -LiteralPath $script:DefinitionFolder -Recurse -File -Filter '*.tmdl'
        $count = ($files | ForEach-Object { @(Get-TmdlDataType -Line (Get-Content -LiteralPath $_.FullName)).Count } | Measure-Object -Sum).Sum
        $count | Should -BeGreaterThan 100
    }

    It '<Name> holds only valid dataType values' -ForEach $script:TmdlFiles {
        $invalid = Get-InvalidTmdlDataType -Line (Get-Content -LiteralPath $Path) |
            ForEach-Object { "line $($_.LineNumber): dataType: $($_.Value)" }

        $invalid -join '; ' | Should -BeNullOrEmpty
    }

    It 'rejects the Power Query name logical, which blocked opening the project' {
        $text = @('table T', '    column IsTeam', '        dataType: logical', '        sourceColumn: IsTeam')

        $invalid = @(Get-InvalidTmdlDataType -Line $text)

        $invalid.Count | Should -Be 1
        $invalid[0].Value | Should -Be 'logical'
        $invalid[0].LineNumber | Should -Be 3
    }

    It 'accepts every value in the valid set, in any case' {
        foreach ($value in $script:ValidDataTypes + @('DateTime', 'BOOLEAN')) {
            @(Get-InvalidTmdlDataType -Line @("        dataType: $value")).Count | Should -Be 0 -Because $value
        }
    }

    It 'does not flag a Power Query type logical inside a partition expression' {
        $text = @(
            '    partition T = m'
            '        source ='
            '            let'
            '                Typed = Table.TransformColumnTypes(Source, {{"IsTeam", type logical}})'
            '            in'
            '                Typed'
        )

        @(Get-InvalidTmdlDataType -Line $text).Count | Should -Be 0
    }

    It 'rejects other unknown values' {
        foreach ($value in 'bool', 'text', 'number', 'date', 'integer') {
            @(Get-InvalidTmdlDataType -Line @("        dataType: $value")).Count | Should -Be 1 -Because $value
        }
    }

    It 'leaves no dataType: logical anywhere under the report folder' {
        $hits = Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot '../report') -Recurse -File |
            Select-String -Pattern '^\s*dataType:\s*logical\s*$' -CaseSensitive:$false
        @($hits).Count | Should -Be 0
    }
}
