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
    $value = Get-EntraField -Object $Object -Name $Name
    if ($null -eq $value) { return '' }
    return [string]$value
}

function Get-EntraField {
    <#
        .SYNOPSIS
            One Graph property from a SDK object or from the hashtable Invoke-MgGraphRequest returns.
    #>
    [CmdletBinding()]
    param([AllowNull()]$Object, [Parameter(Mandatory)][string]$Name)

    if ($null -eq $Object) { return $null }
    $value = Get-GraphAdditionalProperty -Object $Object -Name $Name
    if ($null -ne $value) { return $value }
    if ($Object -isnot [System.Collections.IDictionary]) { return $null }
    foreach ($key in @($Object.Keys)) {
        if ([string]::Equals([string]$key, $Name, [StringComparison]::OrdinalIgnoreCase)) {
            return $Object[$key]
        }
    }
    return $null
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

function Test-EntraGraphReadUri {
    <#
        .SYNOPSIS
            True for a relative Graph URL or an absolute URL on a Microsoft Graph host.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param([Parameter(Mandatory)][string]$Uri)

    # Relative URLs use the signed-in host: graph.microsoft.com or graph.microsoft.us.
    # https://learn.microsoft.com/graph/deployments
    if ($Uri -notmatch '^[a-z][a-z0-9+.-]*://') { return $true }

    $parsed = $null
    if (-not [uri]::TryCreate($Uri, [UriKind]::Absolute, [ref]$parsed)) { return $false }
    if ($parsed.Scheme -ne 'https') { return $false }
    return $parsed.Host -eq 'graph.microsoft.com' -or $parsed.Host -eq 'graph.microsoft.us'
}

function Get-EntraGraphJsonValue {
    <#
        .SYNOPSIS
            One key from a Graph page. The name is not a dotted path, so @odata.nextLink works.
    #>
    [CmdletBinding()]
    param([AllowNull()]$Object, [Parameter(Mandatory)][string]$Name)

    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) {
        foreach ($key in @($Object.Keys)) {
            if ([string]::Equals([string]$key, $Name, [StringComparison]::OrdinalIgnoreCase)) {
                return $Object[$key]
            }
        }
        return $null
    }
    $property = $Object.PSObject.Properties[$Name]
    if ($property) { return $property.Value }
    return $null
}

function Get-EntraGraphHttpStatus {
    <#
        .SYNOPSIS
            The HTTP status on a failed Graph read, or $null.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$ErrorRecord)

    $candidates = @($ErrorRecord.Exception, $ErrorRecord.Exception.InnerException)
    foreach ($exception in $candidates) {
        if ($null -eq $exception) { continue }
        $response = $exception.PSObject.Properties['Response']
        if (-not $response -or $null -eq $response.Value) { continue }
        $status = $response.Value.PSObject.Properties['StatusCode']
        if ($status -and $null -ne $status.Value) { return [int]$status.Value }
    }
    if ($ErrorRecord.Exception.Message -match '\b(401|403|404|429|503)\b') {
        return [int]$Matches[1]
    }
    return $null
}

function Get-EntraGraphRetryDelaySeconds {
    <#
        .SYNOPSIS
            Seconds to wait before repeating the same throttled request.

        .DESCRIPTION
            Wait the Retry-After seconds. Do not retry immediately.
            https://learn.microsoft.com/graph/throttling
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$ErrorRecord, [Parameter(Mandatory)][int]$Attempt)

    $delay = [int][math]::Min(60, [math]::Pow(2, $Attempt - 1))
    $candidates = @($ErrorRecord.Exception, $ErrorRecord.Exception.InnerException)
    foreach ($exception in $candidates) {
        if ($null -eq $exception) { continue }
        $response = $exception.PSObject.Properties['Response']
        if (-not $response -or $null -eq $response.Value) { continue }
        $headers = $response.Value.PSObject.Properties['Headers']
        if (-not $headers -or $null -eq $headers.Value) { continue }

        $retryAfter = $headers.Value.PSObject.Properties['RetryAfter']
        if ($retryAfter -and $retryAfter.Value) {
            $delta = $retryAfter.Value.PSObject.Properties['Delta']
            if ($delta -and $null -ne $delta.Value) {
                $delay = [int][math]::Ceiling($delta.Value.TotalSeconds)
                break
            }
        }

        $named = $null
        if ($headers.Value -is [System.Collections.IDictionary] -and $headers.Value.Contains('Retry-After')) {
            $named = [string]$headers.Value['Retry-After']
        }
        if (-not [string]::IsNullOrWhiteSpace($named)) {
            $seconds = 0
            if ([int]::TryParse($named, [ref]$seconds) -and $seconds -gt 0) { $delay = $seconds }
        }
    }

    if ($ErrorRecord.Exception.Message -match 'Retry-After:\s*(\d+)') {
        $stated = [int]$Matches[1]
        if ($stated -gt $delay) { $delay = $stated }
    }

    if ($delay -lt 1) { return 1 }
    return $delay
}

function Test-EntraActivityThrottle {
    <#
        .SYNOPSIS
            True when a failed read is HTTP 429 or 503, which is throttling, not an empty log.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param([Parameter(Mandatory)]$ErrorRecord)

    $status = Get-EntraGraphHttpStatus -ErrorRecord $ErrorRecord
    return ($status -eq 429 -or $status -eq 503)
}

function Invoke-EntraGraphGet {
    <#
        .SYNOPSIS
            One Graph GET. HTTP 429 and 503 retry the same URL. A page token error does not.

        .DESCRIPTION
            Prefer: include-unknown-enum-members is sent on this request. Callers that follow
            @odata.nextLink call this again so the header is on every page. The SDK page
            iterator does not forward extra headers
            (https://learn.microsoft.com/graph/sdks/paging).
            A 429 waits Retry-After and retries this URL. It does not follow a nextLink
            taken from the error. DirectoryPageTokenNotFoundException is not retried with
            a different link (https://learn.microsoft.com/graph/paging).
            401 and 403 are not retried.
            https://learn.microsoft.com/graph/throttling
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [string]$OutputPath,
        [string]$LogSource,
        [int]$MaxAttempts = 4
    )

    if (-not (Test-EntraGraphReadUri -Uri $Uri)) {
        throw "Refusing to request a page outside Microsoft Graph: $Uri"
    }

    $headers = @{ Prefer = 'include-unknown-enum-members' }
    for ($attempt = 1; ; $attempt++) {
        try {
            return Invoke-MgGraphRequest -Method GET -Uri $Uri -Headers $headers -OutputType Hashtable -ErrorAction Stop
        }
        catch {
            if ($_.Exception.Message -match 'DirectoryPageTokenNotFoundException') { throw }
            $status = Get-EntraGraphHttpStatus -ErrorRecord $_
            if (($status -ne 429 -and $status -ne 503) -or $attempt -ge $MaxAttempts) { throw }
            $delay = Get-EntraGraphRetryDelaySeconds -ErrorRecord $_ -Attempt $attempt
            if ($OutputPath -and $LogSource) {
                Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $LogSource -Message (
                    'HTTP {0}. Waiting {1} seconds before attempt {2} of the same request.' -f $status, $delay, ($attempt + 1))
            }
            Start-Sleep -Seconds $delay
        }
    }
}

function New-EntraActivityListUri {
    <#
        .SYNOPSIS
            The first-page URL for a sign-in or directory-audit window.

        .DESCRIPTION
            Sign-in optional parameters are `$top`, `$skiptoken` and `$filter`. The page
            size maximum and default is 1,000, so this URL does not send `$top`. It does
            not send `$skip` or `$select`. Directory audits also do not list `$skip`.
            https://learn.microsoft.com/graph/api/signin-list
            https://learn.microsoft.com/graph/api/directoryaudit-list
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][ValidateSet('v1.0', 'beta')][string]$Version,
        [Parameter(Mandatory)][string]$RelativePath,
        [Parameter(Mandatory)][string]$Filter
    )

    if ($RelativePath -notin @('auditLogs/signIns', 'auditLogs/directoryAudits')) {
        throw "Unsupported Entra activity path '$RelativePath'."
    }
    if ([string]::IsNullOrWhiteSpace($Filter)) {
        throw 'A time-range filter is required so the list request cannot time out.'
    }

    $encoded = [uri]::EscapeDataString($Filter)
    return "/$Version/$RelativePath`?`$filter=$encoded"
}

function Get-EntraActivityPagedValues {
    <#
        .SYNOPSIS
            Every object in a sign-in or directory-audit window, following @odata.nextLink.

        .DESCRIPTION
            The nextLink URL is requested as returned. A repeated nextLink throws. A host
            other than graph.microsoft.com or graph.microsoft.us throws.
            https://learn.microsoft.com/graph/paging
            https://learn.microsoft.com/graph/deployments
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Version,
        [Parameter(Mandatory)][string]$RelativePath,
        [Parameter(Mandatory)][string]$Filter,
        [string]$OutputPath,
        [string]$LogSource
    )

    $next = New-EntraActivityListUri -Version $Version -RelativePath $RelativePath -Filter $Filter
    $visited = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
    while (-not [string]::IsNullOrWhiteSpace($next)) {
        if (-not (Test-EntraGraphReadUri -Uri $next)) {
            throw "Refusing to request a page outside Microsoft Graph: $next"
        }
        if (-not $visited.Add($next)) {
            throw "Graph returned a nextLink that was already requested: $next"
        }
        $page = Invoke-EntraGraphGet -Uri $next -OutputPath $OutputPath -LogSource $LogSource
        foreach ($item in @(Get-EntraGraphJsonValue -Object $page -Name 'value')) {
            if ($null -ne $item) { , $item }
        }
        $next = [string](Get-EntraGraphJsonValue -Object $page -Name '@odata.nextLink')
    }
}

function Get-EntraActivityWindowItems {
    <#
        .SYNOPSIS
            Reads one time window, and halves it when throttling persists.

        .DESCRIPTION
            Identity and access reports are limited to five requests per 10 seconds per
            app per tenant. If a 429 continues, shorten the timespan. The reports page
            starts at three days; these collectors already use windows of at most 24 hours
            and halve a throttled window down to one hour.
            https://learn.microsoft.com/graph/throttling-limits#identity-and-access-reports-service-limits
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][scriptblock]$Fetch,
        [Parameter(Mandatory)][datetime]$Start,
        [Parameter(Mandatory)][datetime]$End,
        [string]$OutputPath,
        [string]$LogSource
    )

    try {
        return @(& $Fetch (ConvertTo-CsvTimestamp $Start) (ConvertTo-CsvTimestamp $End))
    }
    catch {
        $span = ($End - $Start).TotalMinutes
        if (-not (Test-EntraActivityThrottle -ErrorRecord $_) -or $span -le 60) { throw }
        $mid = $Start.AddMinutes($span / 2)
        if ($OutputPath -and $LogSource) {
            Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $LogSource -Message (
                'HTTP 429 persisted for {0} to {1}. Shortening the window and reading each half.' -f
                (ConvertTo-CsvTimestamp $Start), (ConvertTo-CsvTimestamp $End))
        }
        $left = @(Get-EntraActivityWindowItems -Fetch $Fetch -Start $Start -End $mid -OutputPath $OutputPath -LogSource $LogSource)
        $right = @(Get-EntraActivityWindowItems -Fetch $Fetch -Start $mid -End $End -OutputPath $OutputPath -LogSource $LogSource)
        return @($left + $right)
    }
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
            $items = @(Get-EntraActivityWindowItems -Fetch $Fetch -Start $window.Start -End $window.End -OutputPath $OutputPath -LogSource $log)
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
            Invoke-MgGraphRequest returns a hashtable of the JSON, so nested objects are
            read by their Graph names. errorCode is kept as returned, including 0.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$SignIn)

    $locationValue = Get-EntraField -Object $SignIn -Name 'location'
    $deviceValue = Get-EntraField -Object $SignIn -Name 'deviceDetail'
    $statusValue = Get-EntraField -Object $SignIn -Name 'status'
    $eventTypes = Get-EntraField -Object $SignIn -Name 'signInEventTypes'

    [pscustomobject]@{
        CreatedDateTime         = ConvertTo-CsvTimestamp (Get-EntraField -Object $SignIn -Name 'createdDateTime')
        Id                      = Get-PropertyValue $SignIn 'id'
        UserId                  = Get-PropertyValue $SignIn 'userId'
        UserPrincipalName       = Get-PropertyValue $SignIn 'userPrincipalName'
        AppId                   = Get-PropertyValue $SignIn 'appId'
        AppDisplayName          = Get-PropertyValue $SignIn 'appDisplayName'
        ResourceDisplayName     = Get-PropertyValue $SignIn 'resourceDisplayName'
        IpAddress               = Get-PropertyValue $SignIn 'ipAddress'
        City                    = Get-PropertyValue $locationValue 'city'
        State                   = Get-PropertyValue $locationValue 'state'
        CountryOrRegion         = Get-PropertyValue $locationValue 'countryOrRegion'
        ClientAppUsed           = Get-PropertyValue $SignIn 'clientAppUsed'
        DeviceOperatingSystem   = Get-PropertyValue $deviceValue 'operatingSystem'
        DeviceBrowser           = Get-PropertyValue $deviceValue 'browser'
        DeviceIsCompliant       = Get-PropertyValue $deviceValue 'isCompliant'
        DeviceIsManaged         = Get-PropertyValue $deviceValue 'isManaged'
        IsInteractive           = Get-PropertyValue $SignIn 'isInteractive'
        SignInEventTypes        = Join-ListValue $eventTypes
        ErrorCode               = Get-PropertyValue $statusValue 'errorCode'
        FailureReason           = Get-PropertyValue $statusValue 'failureReason'
        AdditionalDetails       = Get-PropertyValue $statusValue 'additionalDetails'
        ConditionalAccessStatus = Get-PropertyValue $SignIn 'conditionalAccessStatus'
        RiskDetail              = Get-PropertyValue $SignIn 'riskDetail'
        RiskLevelAggregated     = Get-PropertyValue $SignIn 'riskLevelAggregated'
        RiskLevelDuringSignIn   = Get-PropertyValue $SignIn 'riskLevelDuringSignIn'
        RiskState               = Get-PropertyValue $SignIn 'riskState'
    }
}
