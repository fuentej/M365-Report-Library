#Requires -Version 7.0

# Dot-sourced by the Entra activity collectors. The shared module has no event-window
# collector that takes a source list, so the window loop lives here; it is the same
# loop reports/identity-posture uses for its sign-in collector.

function Get-EntraActivitySourceAvailability {
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

function Get-EntraActivityScope {
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

function Get-PropertyValue {
    <#
        .SYNOPSIS
            A property of a Graph SDK object, from the object or its AdditionalProperties,
            as a string; empty when absent.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowNull()]$Object, [Parameter(Mandatory)][string]$Name)

    if ($null -eq $Object) { return '' }
    $value = Get-GraphAdditionalProperty -Object $Object -Name $Name
    if ($null -eq $value) { return '' }
    return [string]$value
}

function Test-ConditionalAccessReadable {
    <#
        .SYNOPSIS
            True when the current Graph session holds a Conditional Access read permission.

        .DESCRIPTION
            appliedConditionalAccessPolicies is omitted from a sign-in, without error,
            unless the caller can read Conditional Access data
            (https://learn.microsoft.com/graph/api/signin-list#permissions). An empty list
            therefore does not prove that no policy applied, so each row records this.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param([Parameter(Mandatory)][hashtable]$Schema)

    try {
        $context = Get-MgContext
    }
    catch {
        return $false
    }
    if ($null -eq $context) { return $false }
    $held = @($context.Scopes)
    foreach ($permission in $Schema.ConditionalAccessReadPermissions) {
        if ($held -contains $permission) { return $true }
    }
    return $false
}

function ConvertTo-EntraActivityUtc {
    <#
        .SYNOPSIS
            A sign-in or directory-audit timestamp as UTC.

        .DESCRIPTION
            createdDateTime and activityDateTime are UTC
            (https://learn.microsoft.com/graph/api/resources/signin#properties,
            https://learn.microsoft.com/graph/api/resources/directoryaudit#properties).
            An Unspecified DateTime is that UTC clock time. ToUniversalTime would treat
            it as the machine's local zone and move the query window.
    #>
    [CmdletBinding()]
    [OutputType([datetime])]
    param([Parameter(Mandatory)][datetime]$Value)

    if ($Value.Kind -eq [DateTimeKind]::Unspecified) {
        return [datetime]::SpecifyKind($Value, [DateTimeKind]::Utc)
    }
    return $Value.ToUniversalTime()
}

function Invoke-EntraActivityEventCollector {
    <#
        .SYNOPSIS
            Runs one event collector: check availability, sign in, resume from the
            watermark, read each window, map, append.

        .DESCRIPTION
            Event sources append from the last exported timestamp. The range is walked in
            windows so no single request asks for an unbounded span, and a window is
            committed only when it was read completely: a failure part-way leaves the
            earlier windows in the file and stops, so the next run resumes at the watermark.

              * A source the schema marks NotAvailable in this cloud writes the header only.
              * An Unverified source is attempted and logs a warning.
              * A missing licence is a logged skip: header only, no exception.
              * Any other failure is logged as an error and throws.

        .PARAMETER Fetch
            Reads one window. Called with the window's Start and End as UTC text; returns
            the objects, every page.

        .PARAMETER Map
            Shapes one object into zero or more rows. Called with the object.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][string]$CsvName,
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)][scriptblock]$Fetch,
        [Parameter(Mandatory)][scriptblock]$Map,
        [Parameter(Mandatory)][string[]]$KeyColumn,
        [Parameter(Mandatory)][string]$WatermarkColumn,
        [Parameter(Mandatory)][string]$License,
        [string]$Description = $Source,
        # The schema key holding the column list, when it is not the source's own name.
        [string]$ColumnKey = $Source,
        [scriptblock]$BeforeFetch,

        [ValidateSet('Commercial', 'GCC', 'GCCHigh')]
        [string]$Environment = 'Commercial',

        [string]$AppId,
        [string]$CertificateThumbprint,
        [string]$TenantId,
        [string]$Organization,
        [string]$SchemaPath = (Join-Path $PSScriptRoot 'EntraActivitySchema.psd1'),

        [Nullable[datetime]]$StartDate,
        [Nullable[datetime]]$EndDate,
        [int]$LookbackDays = 30,
        [int]$WindowHours = 24,

        [switch]$SkipConnect
    )

    $schema = Import-PowerShellDataFile -LiteralPath $SchemaPath
    $columns = [string[]]$schema[$ColumnKey]
    $csvPath = Join-Path $OutputPath $CsvName
    $log = [System.IO.Path]::GetFileNameWithoutExtension($CsvName)

    if (-not (Test-Path -LiteralPath $OutputPath)) {
        New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
    }

    $availability = Get-EntraActivitySourceAvailability -Source $Source -Environment $Environment -Schema $schema
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
            -Scopes (Get-EntraActivityScope -Schema $schema -Source $Source)
    }

    $context = if ($BeforeFetch) { & $BeforeFetch $schema } else { $null }

    $explicitRange = ($null -ne $StartDate) -or ($null -ne $EndDate)
    $watermark = Get-CsvWatermark -Path $csvPath -Column $WatermarkColumn

    $start = if ($null -ne $StartDate) { ConvertTo-EntraActivityUtc ([datetime]$StartDate) }
    elseif ($null -ne $watermark) { ConvertTo-EntraActivityUtc ([datetime]$watermark) }
    else { [datetime]::UtcNow.AddDays(-$LookbackDays) }

    $end = if ($null -ne $EndDate) { ConvertTo-EntraActivityUtc ([datetime]$EndDate) } else { [datetime]::UtcNow }

    if ($end -le $start) {
        if ($explicitRange) {
            # An inverted or empty range asked for explicitly is a mistake, and collecting
            # nothing while reporting success would hide it.
            throw ('The requested range is empty: the end ({0}) is not later than the start ({1}).' -f
                (ConvertTo-CsvTimestamp $end), (ConvertTo-CsvTimestamp $start))
        }

        Write-CollectorLog -OutputPath $OutputPath -Source $log -Message (
            'Nothing new to collect: the watermark ({0}) is already at or after the end of the range.' -f (ConvertTo-CsvTimestamp $start))
        Export-AppendCsv -Path $csvPath -Column $columns
        return
    }

    Write-CollectorLog -OutputPath $OutputPath -Source $log -Message (
        'Querying {0} from {1} to {2} in {3}-hour windows (watermark: {4}).' -f
        $Description, (ConvertTo-CsvTimestamp $start), (ConvertTo-CsvTimestamp $end), $WindowHours,
        $(if ($null -eq $watermark) { 'none' } else { ConvertTo-CsvTimestamp $watermark }))

    $rows = [System.Collections.Generic.List[object]]::new()
    $failure = $null

    foreach ($window in Split-DateRange -Start $start -End $end -WindowMinutes ($WindowHours * 60)) {
        try {
            $items = @(& $Fetch (ConvertTo-CsvTimestamp $window.Start) (ConvertTo-CsvTimestamp $window.End))
        }
        catch {
            $failure = $_.Exception.Message
            break
        }

        foreach ($item in $items) {
            if ($null -eq $item) { continue }
            foreach ($row in @(& $Map $item $context)) { $rows.Add($row) }
        }
    }

    $result = Export-AppendCsv -Path $csvPath -Rows $rows.ToArray() -Column $columns -KeyColumn $KeyColumn -PassThru
    Write-CollectorLog -OutputPath $OutputPath -Source $log -Message (
        '{0}: {1} rows written, {2} skipped as already collected.' -f $CsvName, $result.Written, $result.Skipped)

    if ($null -ne $failure) {
        if ($result.Written -eq 0 -and (Test-LicenseError -Message $failure)) {
            Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $log -Message (
                "Skipping ${CsvName}: ${Description} is not licensed in this tenant ($failure). It needs $License. Writing the header only.")
            return
        }

        $keptNote = if ($result.Written -gt 0) { ' The windows read before the failure were kept.' } else { '' }
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $log -Message (
            "$Description is unavailable to this sign-in ($failure). It needs $License.$keptNote")
        if ($result.Written -gt 0) {
            throw "Reading $Description stopped part-way ($failure). The complete windows were kept; re-run to resume."
        }
        throw "Reading $Description failed ($failure)."
    }
}

function ConvertTo-SignInRow {
    <#
        .SYNOPSIS
            Shapes a signIn object into a SignIns row.

        .DESCRIPTION
            Property names follow https://learn.microsoft.com/graph/api/resources/signin#properties.
            Nested objects (location, deviceDetail, status) are read through
            Get-GraphAdditionalProperty because the SDK may place them in AdditionalProperties.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$SignIn)

    $location = $SignIn.PSObject.Properties['Location']
    $device = $SignIn.PSObject.Properties['DeviceDetail']
    $status = $SignIn.PSObject.Properties['Status']
    $locationValue = if ($location) { $location.Value } else { $null }
    $deviceValue = if ($device) { $device.Value } else { $null }
    $statusValue = if ($status) { $status.Value } else { $null }

    $eventTypes = Get-GraphAdditionalProperty -Object $SignIn -Name 'signInEventTypes'

    [pscustomobject]@{
        CreatedDateTime         = ConvertTo-CsvTimestamp $SignIn.CreatedDateTime
        Id                      = $SignIn.Id
        UserId                  = $SignIn.UserId
        UserPrincipalName       = $SignIn.UserPrincipalName
        AppId                   = $SignIn.AppId
        AppDisplayName          = $SignIn.AppDisplayName
        ResourceDisplayName     = $SignIn.ResourceDisplayName
        IpAddress               = $SignIn.IPAddress
        City                    = Get-PropertyValue $locationValue 'city'
        State                   = Get-PropertyValue $locationValue 'state'
        CountryOrRegion         = Get-PropertyValue $locationValue 'countryOrRegion'
        ClientAppUsed           = $SignIn.ClientAppUsed
        DeviceOperatingSystem   = Get-PropertyValue $deviceValue 'operatingSystem'
        DeviceBrowser           = Get-PropertyValue $deviceValue 'browser'
        DeviceIsCompliant       = Get-PropertyValue $deviceValue 'isCompliant'
        DeviceIsManaged         = Get-PropertyValue $deviceValue 'isManaged'
        IsInteractive           = $SignIn.IsInteractive
        SignInEventTypes        = Join-ListValue $eventTypes
        ErrorCode               = Get-PropertyValue $statusValue 'errorCode'
        FailureReason           = Get-PropertyValue $statusValue 'failureReason'
        AdditionalDetails       = Get-PropertyValue $statusValue 'additionalDetails'
        ConditionalAccessStatus = [string]$SignIn.ConditionalAccessStatus
        RiskDetail              = [string]$SignIn.RiskDetail
        RiskLevelAggregated     = [string]$SignIn.RiskLevelAggregated
        RiskLevelDuringSignIn   = [string]$SignIn.RiskLevelDuringSignIn
        RiskState               = [string]$SignIn.RiskState
    }
}
