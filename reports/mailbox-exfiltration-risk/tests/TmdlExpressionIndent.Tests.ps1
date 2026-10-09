#Requires -Version 7.0

<#
    A multiline TMDL expression has to sit one indent level deeper than the
    object's properties. When it does not, the parser keeps reading property
    lines as expression text, so dataType, formatString and lineageTag never
    become properties and the DAX is no longer the expression that was written.

    https://learn.microsoft.com/en-us/analysis-services/tmdl/tmdl-overview#expressions
#>

BeforeAll {
    $script:DefinitionFolder = Join-Path $PSScriptRoot '../report/MailboxExfiltrationRisk.SemanticModel/definition' | Resolve-Path | Select-Object -ExpandProperty Path

    function Get-TmdlLeadingSpaceCount {
        <#
            .SYNOPSIS
                Counts leading spaces. This model indents with four spaces.
        #>
        param([Parameter(Mandatory)][AllowEmptyString()][string]$Text)

        $count = 0
        foreach ($char in $Text.ToCharArray()) {
            if ($char -ne ' ') { break }
            $count++
        }
        return $count
    }

    function Get-TmdlExpressionIndentFault {
        <#
            .SYNOPSIS
                Lines where a column or measure expression is not indented
                deeper than a property that follows it.
        #>
        param([Parameter(Mandatory)][AllowEmptyString()][string[]]$Line)

        $faults = [System.Collections.Generic.List[string]]::new()
        $indentUnit = 4
        for ($i = 0; $i -lt $Line.Count; $i++) {
            if ($Line[$i] -notmatch '^\s*(column|measure)\s+.*=\s*$') { continue }

            $declarationIndent = Get-TmdlLeadingSpaceCount -Text $Line[$i]
            $propertyIndent = $declarationIndent + $indentUnit
            $expressionLines = [System.Collections.Generic.List[object]]::new()

            for ($j = $i + 1; $j -lt $Line.Count; $j++) {
                if ([string]::IsNullOrWhiteSpace($Line[$j])) { continue }

                $indent = Get-TmdlLeadingSpaceCount -Text $Line[$j]
                if ($indent -le $declarationIndent) { break }

                $trimmed = $Line[$j].Trim()
                $isProperty = $indent -eq $propertyIndent -and $trimmed -match '^[A-Za-z][A-Za-z0-9]*\s*:'
                if ($isProperty) {
                    foreach ($expressionLine in $expressionLines) {
                        if ($expressionLine.Indent -le $indent) {
                            $faults.Add("line $($expressionLine.LineNumber) is not deeper than '$trimmed' on line $($j + 1)")
                        }
                    }
                    break
                }

                $expressionLines.Add([pscustomobject]@{
                        Indent     = $indent
                        LineNumber = $j + 1
                    })
            }
        }

        return $faults
    }
}

Describe 'TMDL multiline expression indent' {
    It 'flags an expression that shares its indent with the following property' {
        $text = @(
            '    measure Sales = '
            '        VAR x = 1'
            '        RETURN x'
            '        formatString: #,0'
        )

        $faults = @(Get-TmdlExpressionIndentFault -Line $text)

        $faults.Count | Should -Be 2
        $faults[0] | Should -Match 'line 2'
        $faults[1] | Should -Match 'line 3'
    }

    It 'accepts an expression indented one level deeper than its properties' {
        $text = @(
            '    column Days = '
            '            VAR d = 1'
            '            RETURN d'
            '        dataType: int64'
            '        formatString: #,0'
        )

        @(Get-TmdlExpressionIndentFault -Line $text).Count | Should -Be 0
    }

    It 'leaves a partition M expression alone' {
        $text = @(
            '    partition T = m'
            '        mode: import'
            '        source ='
            '            let'
            '                Source = Csv.Document(CsvFolder)'
            '                #"Changed Type" = Table.TransformColumnTypes(Source, {{"IsTeam", type logical}})'
            '            in'
            '                #"Changed Type"'
        )

        @(Get-TmdlExpressionIndentFault -Line $text).Count | Should -Be 0
    }

    It 'keeps every multiline column and measure expression deeper than its properties' {
        $files = Get-ChildItem -LiteralPath $script:DefinitionFolder -Recurse -File -Filter '*.tmdl'
        $hits = foreach ($file in $files) {
            $relative = $file.FullName.Substring($script:DefinitionFolder.Length + 1)
            foreach ($fault in @(Get-TmdlExpressionIndentFault -Line (Get-Content -LiteralPath $file.FullName))) {
                "${relative}: $fault"
            }
        }

        $hits -join '; ' | Should -BeNullOrEmpty
    }
}
