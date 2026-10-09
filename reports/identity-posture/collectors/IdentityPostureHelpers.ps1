#Requires -Version 7.0

# Dot-sourced by the identity-posture collectors.

function Get-IdentityPostureSourceAvailability {
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

function ConvertTo-SignInTimestamp {
    <#
        .SYNOPSIS
            A sign-in time as UTC text, or an empty string when Graph holds no value.

        .DESCRIPTION
            List users returns 0001-01-01T00:00:00Z for lastNonInteractiveSignInDateTime
            and "" for lastSuccessfulSignInDateTime when there is no value
            (https://learn.microsoft.com/graph/api/user-list#example-11-get-users-including-their-last-sign-in-time).
            Neither is a sign-in, so both become an empty cell instead of a date.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Position = 0)][AllowNull()][object]$Value)

    if ($null -eq $Value) { return '' }
    if ($Value -is [string] -and [string]::IsNullOrWhiteSpace($Value)) { return '' }

    $text = ConvertTo-CsvTimestamp $Value
    if ([string]::IsNullOrEmpty($text) -or $text.StartsWith('0001-01-01')) { return '' }
    return $text
}

function Test-LicenseError {
    <#
        .SYNOPSIS
            True when an error message says the tenant lacks the licence a source needs.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Message)

    return ($Message -match '(?i)licen[cs]e|NonPremium|premium|RequiresPremium|AadPremium')
}

function Get-IdentityPostureScope {
    <#
        .SYNOPSIS
            The shared sign-in's read scopes plus the ones a source adds.
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param([Parameter(Mandatory)][hashtable]$Schema, [Parameter(Mandatory)][string]$Source)

    $extra = @($Schema.ExtraScopes[$Source])
    return , [string[]]@((Get-DefaultGraphScope) + $extra | Where-Object { $_ } | Select-Object -Unique)
}

function Invoke-IdentityPostureSnapshot {
    <#
        .SYNOPSIS
            Runs one state collector: check availability, sign in, read, map, append.

        .DESCRIPTION
            Every state source follows the same steps, so they live here and each collector
            script only says what it reads and how a row is shaped.

              * A source the schema marks NotAvailable in this cloud writes the header only.
              * An Unverified source is attempted and logs a warning.
              * A missing licence is a logged skip: header only, no error and no exception.
              * Any other failure is logged as an error, leaves a header-only file when
                nothing was collected, and throws so the run does not report success.

        .PARAMETER Source
            The schema key, which names the column list, the availability entry and the
            extra Graph scopes.

        .PARAMETER Fetch
            Reads the source. Called with no arguments; returns the objects, every page.

        .PARAMETER Map
            Shapes one object into a row. Called with the object and the RunDate.

        .PARAMETER License
            What the source needs, for the log line written when it is refused.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][string]$CsvName,
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)][scriptblock]$Fetch,
        [Parameter(Mandatory)][scriptblock]$Map,
        [Parameter(Mandatory)][string[]]$KeyColumn,
        [Parameter(Mandatory)][string]$License,
        [string]$Description = $Source,

        [ValidateSet('Commercial', 'GCC', 'GCCHigh')]
        [string]$Environment = 'Commercial',

        [string]$AppId,
        [string]$CertificateThumbprint,
        [string]$TenantId,
        [string]$Organization,
        [string]$SchemaPath = (Join-Path $PSScriptRoot 'IdentityPostureSchema.psd1'),
        [switch]$SkipConnect
    )

    $schema = Import-PowerShellDataFile -LiteralPath $SchemaPath
    $columns = [string[]]$schema[$Source]
    $csvPath = Join-Path $OutputPath $CsvName
    $log = [System.IO.Path]::GetFileNameWithoutExtension($CsvName)
    $runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

    if (-not (Test-Path -LiteralPath $OutputPath)) {
        New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
    }

    $availability = Get-IdentityPostureSourceAvailability -Source $Source -Environment $Environment -Schema $schema
    if ($availability.ShouldSkip) {
        Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $log -Message (
            "Skipping $CsvName. $($availability.Reason) $($availability.Reference)")
        Export-AppendCsv -Path $csvPath -Column $columns
        return
    }
    if ($availability.Status -eq 'Unverified') {
        Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $log -Message $availability.Reason
    }

    if (-not $SkipConnect) {
        Connect-M365Service -Service Graph -Environment $Environment -AppId $AppId `
            -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization `
            -Scopes (Get-IdentityPostureScope -Schema $schema -Source $Source)
    }

    $items = $null
    try {
        $items = @(& $Fetch)
    }
    catch {
        $message = $_.Exception.Message
        if (Test-LicenseError -Message $message) {
            Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $log -Message (
                "Skipping ${CsvName}: ${Description} is not licensed in this tenant ($message). It needs $License. Writing the header only.")
        }
        else {
            Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $log -Message (
                "$Description is unavailable to this sign-in ($message). It needs $License. Writing the header only.")
            Export-AppendCsv -Path $csvPath -Column $columns
            throw "$Description is unavailable ($message)."
        }
        Export-AppendCsv -Path $csvPath -Column $columns
        return
    }

    $rows = foreach ($item in $items) {
        if ($null -ne $item) { & $Map $item $runDate }
    }

    $result = Export-AppendCsv -Path $csvPath -Rows @($rows) -Column $columns -KeyColumn $KeyColumn -PassThru
    Write-CollectorLog -OutputPath $OutputPath -Source $log -Message (
        '{0}: {1} rows written, {2} skipped.' -f $CsvName, $result.Written, $result.Skipped)
}
