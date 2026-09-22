#Requires -Version 7.0

<#
    A small, purpose-built TMDL reader for the guest-access report's semantic model.
    It is not a general TMDL parser -- it only understands the shape this repo's
    generated tables use (one table per file, tab-stop indentation of four spaces,
    `column X` / `column X = <expr>` / `measure X = <expr>` declarations) -- but that
    is exactly what ReportModel.Tests.ps1 needs to check visual field references and
    CSV-column mappings against.
#>

function ConvertFrom-TmdlName {
    <#
        .SYNOPSIS
            Strips the single-quote wrapping TMDL uses for names with spaces/special
            characters, undoing the doubled-quote escape.
    #>
    param([Parameter(Mandatory)][string]$Name)

    $trimmed = $Name.Trim()
    if ($trimmed.Length -ge 2 -and $trimmed[0] -eq "'" -and $trimmed[-1] -eq "'") {
        return $trimmed.Substring(1, $trimmed.Length - 2) -replace "''", "'"
    }
    return $trimmed
}

function Get-TmdlTable {
    <#
        .SYNOPSIS
            Parses a single <table>.tmdl file into its name, columns and measures.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $lines = Get-Content -LiteralPath $Path
    $tableName = $null
    $columns = [System.Collections.Generic.List[hashtable]]::new()
    $measures = [System.Collections.Generic.List[string]]::new()

    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]

        if (-not $tableName -and $line -match '^table\s+(.+)$') {
            $tableName = ConvertFrom-TmdlName -Name $Matches[1]
            continue
        }

        if ($line -match '^    column\s+(.+?)(\s*=\s*(.*))?$') {
            $rawName = $Matches[1]
            $isCalculated = $Matches[2] -ne $null -and $Matches[2] -ne ''
            $name = ConvertFrom-TmdlName -Name $rawName

            $sourceColumn = $null
            for ($j = $i + 1; $j -lt $lines.Count; $j++) {
                if ($lines[$j] -match '^    (column|measure|partition)\s') { break }
                if ($lines[$j] -match '^\s+sourceColumn:\s*(.+)$') {
                    $sourceColumn = ConvertFrom-TmdlName -Name $Matches[1]
                    break
                }
            }

            $columns.Add(@{
                    Name         = $name
                    IsCalculated = [bool]$isCalculated -or (-not $sourceColumn)
                    SourceColumn = $sourceColumn
                })
            continue
        }

        if ($line -match '^    measure\s+(.+?)\s*=') {
            $measures.Add((ConvertFrom-TmdlName -Name $Matches[1]))
            continue
        }
    }

    if (-not $tableName) {
        throw "No 'table <name>' declaration found in $Path"
    }

    return @{
        Name     = $tableName
        Path     = $Path
        Columns  = $columns.ToArray()
        Measures = $measures.ToArray()
    }
}

function Get-TmdlModel {
    <#
        .SYNOPSIS
            Every table in the semantic model's definition/tables folder, keyed by
            table name.
    #>
    param([Parameter(Mandatory)][string]$TablesFolder)

    $model = @{}
    foreach ($file in Get-ChildItem -LiteralPath $TablesFolder -Filter '*.tmdl' -File) {
        $table = Get-TmdlTable -Path $file.FullName
        $model[$table.Name] = $table
    }
    return $model
}
