#Requires -Version 7.0

# Dot-sourced by the teams-groups-lifecycle collectors.

function Get-TeamsGroupsSourceAvailability {
    <#
        .SYNOPSIS
            Says whether a source is available in a cloud, and why.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][ValidateSet('Commercial', 'GCC', 'GCCHigh')][string]$Environment,
        [Parameter(Mandatory)][hashtable]$Schema
    )

    $availability = $Schema.SourceAvailability
    if (-not $availability.ContainsKey($Source)) {
        throw "Unknown source '$Source'. Known sources: $(($availability.Keys | Sort-Object) -join ', ')."
    }

    $entry = $availability[$Source][$Environment]

    $reason = switch ($entry.Status) {
        'Available' { "$Source is documented as available in $Environment." }
        'NotAvailable' { "$Source is documented as unavailable in $Environment." }
        default { "Availability of $Source in $Environment is UNVERIFIED; the collector will attempt it anyway." }
    }

    [pscustomobject]@{
        Source      = $Source
        Environment = $Environment
        Status      = $entry.Status
        # Only NotAvailable skips. An unverified source is attempted so a guess
        # never drops data the tenant would have returned.
        ShouldSkip  = ($entry.Status -eq 'NotAvailable')
        Reason      = $reason
        Reference   = $entry.Reference
    }
}

function Join-ListValue {
    <#
        .SYNOPSIS
            Flattens a list into one CSV cell, separated by semicolons.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Position = 0)][AllowNull()]$Value)

    if ($null -eq $Value) { return '' }
    if ($Value -is [string]) { return $Value }
    return (@($Value) | Where-Object { $null -ne $_ } | ForEach-Object { [string]$_ }) -join ';'
}

function Test-ListContains {
    [CmdletBinding()]
    [OutputType([bool])]
    param([Parameter(Mandatory)][AllowNull()]$Value, [Parameter(Mandatory)][string]$Item)

    return [bool](@($Value) | Where-Object { $_ -eq $Item })
}

function Get-ReportColumnValue {
    <#
        .SYNOPSIS
            Reads a usage-report CSV column by any of its accepted header names.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Row,
        [Parameter(Mandatory)][string[]]$Header
    )

    foreach ($name in $Header) {
        $property = $Row.PSObject.Properties[$name]
        if ($property -and $null -ne $property.Value) { return [string]$property.Value }
    }
    return ''
}

function Get-GraphErrorIsNotFound {
    <#
        .SYNOPSIS
            True when a Graph error is a 404 / ResourceNotFound.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param([Parameter(Mandatory)][string]$Message)

    return ($Message -match 'NotFound|Not Found|404|does not exist')
}
