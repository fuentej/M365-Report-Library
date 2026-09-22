#Requires -Version 7.0

<#
    A small, offline, structural JSON Schema (draft-07) validator.

    It exists so ReportSchema.Tests.ps1 can check every PBIR file this report ships
    against the exact schema its own `$schema` property names, without CI needing
    network access to fetch schemas or install a validator package. The schemas
    themselves are vendored under tests/schemas/, mirroring their path in
    https://github.com/microsoft/json-schemas so a `$ref` between two vendored files
    resolves the same way it would against the published schema tree.

    This is deliberately not a complete draft-07 implementation. It checks type,
    required, additionalProperties (false or a schema), items, $ref, anyOf,
    oneOf (exactly one match), const, enum, and pattern. It does not check
    format, if/then/else, or unevaluatedProperties.
#>

$script:SchemaDocumentCache = @{}

function Get-SchemaDocument {
    <#
        .SYNOPSIS
            Loads and caches a vendored schema file by its resolved path.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $full = (Resolve-Path -LiteralPath $Path).ProviderPath
    if (-not $script:SchemaDocumentCache.ContainsKey($full)) {
        $script:SchemaDocumentCache[$full] = Get-Content -LiteralPath $full -Raw | ConvertFrom-Json -AsHashtable
    }
    return $script:SchemaDocumentCache[$full]
}

function Resolve-JsonPointer {
    param($Root, [string]$Pointer)

    if ([string]::IsNullOrEmpty($Pointer) -or $Pointer -eq '#') { return $Root }
    $trimmed = $Pointer.TrimStart('#').TrimStart('/')
    if ([string]::IsNullOrEmpty($trimmed)) { return $Root }

    $current = $Root
    foreach ($segment in ($trimmed -split '/')) {
        $decoded = $segment -replace '~1', '/' -replace '~0', '~'
        if ($current -is [System.Collections.IDictionary] -and $current.ContainsKey($decoded)) {
            $current = $current[$decoded]
        }
        elseif (($current -is [System.Collections.IEnumerable]) -and ($current -isnot [string])) {
            $current = @($current)[[int]$decoded]
        }
        else {
            throw "Cannot resolve JSON pointer segment '$decoded' of '$Pointer'"
        }
    }
    return $current
}

function Resolve-SchemaRef {
    <#
        .SYNOPSIS
            Resolves a `$ref` (local "#/definitions/X" or cross-file
            "../other/1.0.0/schema.json#/definitions/Y") relative to the file the
            referencing schema itself came from.
    #>
    param([Parameter(Mandatory)][string]$Ref, [Parameter(Mandatory)][string]$CurrentFile)

    $hashIndex = $Ref.IndexOf('#')
    if ($hashIndex -lt 0) { $filePart = $Ref; $pointerPart = '' }
    else { $filePart = $Ref.Substring(0, $hashIndex); $pointerPart = $Ref.Substring($hashIndex) }

    if ([string]::IsNullOrEmpty($filePart)) {
        $targetFile = $CurrentFile
    }
    else {
        $dir = Split-Path -Parent $CurrentFile
        $targetFile = (Resolve-Path -LiteralPath (Join-Path $dir $filePart)).ProviderPath
    }

    $doc = Get-SchemaDocument -Path $targetFile
    $sub = Resolve-JsonPointer -Root $doc -Pointer $pointerPart
    return @{ Schema = $sub; File = $targetFile }
}

function Test-JsonInstanceType {
    param($Value, [string]$Type)

    switch ($Type) {
        'object' { return $Value -is [System.Collections.IDictionary] }
        'array' { return ($Value -is [System.Collections.IEnumerable]) -and ($Value -isnot [string]) -and ($Value -isnot [System.Collections.IDictionary]) }
        'string' { return $Value -is [string] }
        'number' { return $Value -is [double] -or $Value -is [int] -or $Value -is [long] -or $Value -is [decimal] }
        'integer' { return $Value -is [int] -or $Value -is [long] -or (($Value -is [double]) -and ([math]::Floor($Value) -eq $Value)) }
        'boolean' { return $Value -is [bool] }
        'null' { return $null -eq $Value }
        default { return $true }
    }
}

function Test-JsonScalarEqual {
    param($A, $B)
    if ($null -eq $A -or $null -eq $B) { return $A -eq $B }
    return "$A" -ceq "$B"
}

function Test-JsonSchema {
    <#
        .SYNOPSIS
            Validates $Instance against $Schema, returning an array of error
            strings (empty means valid). $SchemaFile is the absolute path of the
            file $Schema was read from, needed to resolve any relative `$ref`.
    #>
    param(
        $Instance,
        [Parameter(Mandatory)][hashtable]$Schema,
        [Parameter(Mandatory)][string]$SchemaFile,
        [string]$InstancePath = '$',
        [int]$Depth = 0
    )

    # NOTE: every exit point returns with the unary comma (`,$errors`). Without it,
    # PowerShell enumerates the List[string] onto the pipeline -- an *empty* list
    # then produces zero pipeline objects, and a caller doing `$x = Test-JsonSchema
    # ...` gets $null back instead of an empty, usable collection.
    $errors = [System.Collections.Generic.List[string]]::new()
    if ($Depth -gt 75) {
        $errors.Add("$InstancePath : schema nesting exceeded the safety limit -- possible `$ref cycle")
        return , $errors
    }

    if ($Schema.ContainsKey('$ref')) {
        $resolved = Resolve-SchemaRef -Ref $Schema['$ref'] -CurrentFile $SchemaFile
        return , (Test-JsonSchema -Instance $Instance -Schema $resolved.Schema -SchemaFile $resolved.File `
                -InstancePath $InstancePath -Depth ($Depth + 1))
    }

    if ($Schema.ContainsKey('const') -and -not (Test-JsonScalarEqual $Instance $Schema['const'])) {
        $errors.Add("$InstancePath : expected const '$($Schema['const'])', got '$Instance'")
    }

    if ($Schema.ContainsKey('enum')) {
        $hit = @($Schema['enum'] | Where-Object { Test-JsonScalarEqual $_ $Instance })
        if ($hit.Count -eq 0) {
            $errors.Add("$InstancePath : '$Instance' is not one of the allowed enum values")
        }
    }

    if ($Schema.ContainsKey('pattern') -and $Instance -is [string] -and $Instance -notmatch $Schema['pattern']) {
        $errors.Add("$InstancePath : '$Instance' does not match pattern '$($Schema['pattern'])'")
    }

    if ($Schema.ContainsKey('type')) {
        $types = @($Schema['type'])
        if (-not ($types | Where-Object { Test-JsonInstanceType -Value $Instance -Type $_ })) {
            $errors.Add("$InstancePath : expected type [$($types -join '|')]")
        }
    }

    foreach ($combinatorKey in 'anyOf', 'oneOf') {
        if (-not $Schema.ContainsKey($combinatorKey)) { continue }
        # oneOf is exactly one match. anyOf is at least one. Stopping at the first
        # success makes oneOf accept an instance that also matches a later branch.
        $passCount = 0
        $firstAltErrors = $null
        foreach ($alt in $Schema[$combinatorKey]) {
            $altErrors = Test-JsonSchema -Instance $Instance -Schema $alt -SchemaFile $SchemaFile `
                -InstancePath $InstancePath -Depth ($Depth + 1)
            if ($altErrors.Count -eq 0) {
                $passCount++
                if ($combinatorKey -eq 'anyOf') { break }
            }
            elseif ($null -eq $firstAltErrors) { $firstAltErrors = $altErrors }
        }
        $ok = if ($combinatorKey -eq 'oneOf') { $passCount -eq 1 } else { $passCount -ge 1 }
        if (-not $ok) {
            if ($combinatorKey -eq 'oneOf' -and $passCount -gt 1) {
                $errors.Add("$InstancePath : matched $passCount oneOf alternatives; exactly one is required")
            }
            else {
                $detail = if ($firstAltErrors) { $firstAltErrors -join '; ' } else { '(no alternatives)' }
                $errors.Add("$InstancePath : matched no $combinatorKey alternative (closest: $detail)")
            }
        }
    }

    if ($Instance -is [System.Collections.IDictionary]) {
        if ($Schema.ContainsKey('required')) {
            foreach ($required in $Schema['required']) {
                if (-not $Instance.ContainsKey($required)) {
                    $errors.Add("$InstancePath : missing required property '$required'")
                }
            }
        }

        $declared = @()
        if ($Schema.ContainsKey('properties')) { $declared = @($Schema['properties'].Keys) }

        $additional = $true
        if ($Schema.ContainsKey('additionalProperties')) { $additional = $Schema['additionalProperties'] }

        foreach ($key in $Instance.Keys) {
            if ($declared -contains $key) {
                $subErrors = Test-JsonSchema -Instance $Instance[$key] -Schema $Schema['properties'][$key] `
                    -SchemaFile $SchemaFile -InstancePath "$InstancePath.$key" -Depth ($Depth + 1)
                $errors.AddRange($subErrors)
            }
            elseif ($additional -eq $false) {
                $errors.Add("$InstancePath.$key : property not allowed (additionalProperties: false)")
            }
            elseif ($additional -is [System.Collections.IDictionary]) {
                $subErrors = Test-JsonSchema -Instance $Instance[$key] -Schema $additional `
                    -SchemaFile $SchemaFile -InstancePath "$InstancePath.$key" -Depth ($Depth + 1)
                $errors.AddRange($subErrors)
            }
        }
    }
    elseif (($Instance -is [System.Collections.IEnumerable]) -and ($Instance -isnot [string])) {
        if ($Schema.ContainsKey('items')) {
            $i = 0
            foreach ($item in $Instance) {
                $subErrors = Test-JsonSchema -Instance $item -Schema $Schema['items'] -SchemaFile $SchemaFile `
                    -InstancePath "$InstancePath[$i]" -Depth ($Depth + 1)
                $errors.AddRange($subErrors)
                $i++
            }
        }
    }

    return , $errors
}

function Test-PbirSchema {
    <#
        .SYNOPSIS
            Reads a PBIR JSON file, resolves the schema its own `$schema` property
            names against the vendored copy under $SchemaRoot, and validates the
            file against it. Returns an array of error strings (empty is valid).
    #>
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$SchemaRoot)

    $data = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json -AsHashtable
    if (-not $data.ContainsKey('$schema')) {
        return , @("$Path : has no `$schema property")
    }

    $schemaUrl = $data['$schema']
    $prefix = 'https://developer.microsoft.com/json-schemas/'
    if (-not $schemaUrl.StartsWith($prefix)) {
        return , @("$Path : `$schema '$schemaUrl' is not a developer.microsoft.com/json-schemas URL")
    }

    $relativePath = $schemaUrl.Substring($prefix.Length)
    $schemaFile = Join-Path $SchemaRoot $relativePath
    if (-not (Test-Path -LiteralPath $schemaFile)) {
        return , @("$Path : no vendored schema for '$schemaUrl' (expected at $schemaFile)")
    }
    $schemaFile = (Resolve-Path -LiteralPath $schemaFile).ProviderPath

    $schemaDoc = Get-SchemaDocument -Path $schemaFile
    return , (Test-JsonSchema -Instance $data -Schema $schemaDoc -SchemaFile $schemaFile -InstancePath '$')
}
