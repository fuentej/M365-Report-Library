#Requires -Version 7.0

# Dot-sourced by the unified audit log collectors. Report-specific: the shared connection and CSV
# functions come from shared/M365ReportLibrary.psm1.
#
# Two functions in this file send a POST. Both are read paths that start a read, and neither changes
# tenant data (decision D-007, docs/candidates/unified-audit-log.md "Read-only caveat"):
#   New-AuditGraphQuery            POST /security/auditLog/queries
#   Start-AuditActivitySubscription  POST .../activity/feed/subscriptions/start
# tests/ReadOnly.Tests.ps1 pins both to exactly those paths.

function Get-AuditSourceAvailability {
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

function Test-AuditSourceSkipped {
    <#
        .SYNOPSIS
            Applies the availability rule for a collector: a NotAvailable source writes its
            header only and logs why; an UNVERIFIED one logs a warning and is attempted.

        .OUTPUTS
            $true when the collector must stop.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][string]$LogSource,
        [Parameter(Mandatory)][string]$CsvPath,
        [Parameter(Mandatory)][string[]]$Column,
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)][string]$Environment,
        [Parameter(Mandatory)][hashtable]$Schema
    )

    $availability = Get-AuditSourceAvailability -Source $Source -Environment $Environment -Schema $Schema
    if ($availability.ShouldSkip) {
        Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $LogSource -Message (
            "Skipping $(Split-Path -Leaf $CsvPath). $($availability.Reason) $($availability.Reference)")
        Export-AppendCsv -Path $CsvPath -Column $Column
        return $true
    }
    if ($availability.Status -eq 'Unverified') {
        Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $LogSource -Message (
            "$($availability.Reason) $($availability.Reference)")
    }
    return $false
}

#region Values

function Get-ObjectValue {
    <#
        .SYNOPSIS
            A property of an object or key of a dictionary, or $null when it is not there.
            Audit records differ by workload, so a missing property is normal.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowNull()]$Object,
        [Parameter(Mandatory)][string]$Name
    )

    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) {
        foreach ($key in $Object.Keys) {
            if ([string]$key -ieq $Name) { return $Object[$key] }
        }
        return $null
    }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    return $property.Value
}

function Get-ObjectText {
    <#
        .SYNOPSIS
            A property as text, or an empty string. Several values are joined with a semicolon.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][AllowNull()]$Object,
        [Parameter(Mandatory)][string]$Name
    )

    $value = Get-ObjectValue -Object $Object -Name $Name
    if ($null -eq $value) { return '' }
    if ($value -is [System.Collections.IEnumerable] -and $value -isnot [string]) {
        return (@($value | ForEach-Object { [string]$_ }) -join ';')
    }
    return [string]$value
}

function ConvertTo-AuditUtc {
    <#
        .SYNOPSIS
            A UTC timestamp, treating an Unspecified DateTime or a string with no zone as UTC.

        .DESCRIPTION
            Search-UnifiedAuditLog stores and reads -StartDate and -EndDate as UTC, and audit
            CreationTime has no zone on the Management Activity API
            (https://learn.microsoft.com/powershell/module/exchangepowershell/search-unifiedauditlog).
            ToUniversalTime() would read an Unspecified value as local time and shift the stamp.
    #>
    [CmdletBinding()]
    [OutputType([datetime])]
    param([Parameter(Mandatory)][AllowNull()]$Value)

    if ($null -eq $Value) { throw 'An audit timestamp is missing.' }
    if ($Value -is [datetime]) {
        if ($Value.Kind -eq [DateTimeKind]::Unspecified) { return [datetime]::SpecifyKind($Value, [DateTimeKind]::Utc) }
        return $Value.ToUniversalTime()
    }
    if ($Value -is [datetimeoffset]) { return $Value.UtcDateTime }

    $parsed = [datetime]::MinValue
    $styles = [System.Globalization.DateTimeStyles]::AdjustToUniversal -bor [System.Globalization.DateTimeStyles]::AssumeUniversal
    if (-not [datetime]::TryParse([string]$Value, [System.Globalization.CultureInfo]::InvariantCulture, $styles, [ref]$parsed)) {
        throw "Cannot read a UTC timestamp from '$Value'."
    }
    return $parsed
}

function Get-AuditRunStart {
    <#
        .SYNOPSIS
            Where an event collector starts: the latest timestamp already in its CSV, or
            -LookbackDays before now on the first run.
    #>
    [CmdletBinding()]
    [OutputType([datetime])]
    param(
        [Parameter(Mandatory)][string]$CsvPath,
        [Parameter(Mandatory)][string]$WatermarkColumn,
        [Parameter(Mandatory)][datetime]$End,
        [Parameter(Mandatory)][int]$LookbackDays
    )

    $watermark = Get-CsvWatermark -Path $CsvPath -Column $WatermarkColumn
    if ($watermark) { return (ConvertTo-AuditUtc $watermark) }
    return $End.AddDays(-$LookbackDays)
}

function New-AuditRow {
    <#
        .SYNOPSIS
            The common columns of one audit record, read from its AuditData (or the Management
            Activity event), with a fallback for each when the JSON lacks it.

        .DESCRIPTION
            Names are the common schema's
            (https://learn.microsoft.com/office/office-365-management-api/office-365-management-activity-api-schema).
            Workload-specific properties stay inside AuditData; nothing is flattened beyond these.
    #>
    [CmdletBinding()]
    param(
        [AllowNull()]$Data,
        [hashtable]$Fallback = @{},
        [Parameter(Mandatory)][AllowEmptyString()][string]$Json
    )

    function Pick([string]$Name) {
        $text = Get-ObjectText -Object $Data -Name $Name
        if ($text -ne '') { return $text }
        if ($Fallback.ContainsKey($Name) -and $null -ne $Fallback[$Name]) { return [string]$Fallback[$Name] }
        return ''
    }

    $created = Get-ObjectValue -Object $Data -Name 'CreationTime'
    if ($null -eq $created -or "$created" -eq '') { $created = $Fallback['CreationTime'] }

    [ordered]@{
        CreationTime   = if ($null -eq $created) { '' } else { ConvertTo-CsvTimestamp $created }
        RecordId       = Pick 'Id'
        RecordType     = Pick 'RecordType'
        Operation      = Pick 'Operation'
        Workload       = Pick 'Workload'
        UserId         = Pick 'UserId'
        ObjectId       = Pick 'ObjectId'
        ClientIP       = Pick 'ClientIP'
        ResultStatus   = Pick 'ResultStatus'
        OrganizationId = Pick 'OrganizationId'
        AuditData      = $Json
    }
}

#endregion

#region Source 1: Search-UnifiedAuditLog

function ConvertTo-AuditSearchRow {
    <#
        .SYNOPSIS
            One Search-UnifiedAuditLog record as an audit-search-cmdlet.csv row.

        .DESCRIPTION
            The record carries AuditData as JSON text. The cmdlet's own RecordType, CreationDate,
            UserIds and Identity are the fallback when the JSON lacks a field.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Record)

    $json = Get-ObjectText -Object $Record -Name 'AuditData'
    $data = $null
    if ($json -ne '') {
        try { $data = $json | ConvertFrom-Json -ErrorAction Stop } catch { $data = $null }
    }
    $fallback = @{
        CreationTime = Get-ObjectValue -Object $Record -Name 'CreationDate'
        Id           = Get-ObjectText -Object $Record -Name 'Identity'
        Operation    = Get-ObjectText -Object $Record -Name 'Operations'
        UserId       = Get-ObjectText -Object $Record -Name 'UserIds'
        ObjectId     = Get-ObjectText -Object $Record -Name 'ObjectIds'
    }
    $row = New-AuditRow -Data $data -Fallback $fallback -Json ($json -replace '\s*[\r\n]+\s*', '')
    # The cmdlet returns the record type by name; the JSON holds its number. Keep the name.
    $name = Get-ObjectText -Object $Record -Name 'RecordType'
    if ($name -ne '') { $row['RecordType'] = $name }
    [pscustomobject]$row
}

function Get-AuditSearchSession {
    <#
        .SYNOPSIS
            Every record of one Search-UnifiedAuditLog session for a date range.

        .DESCRIPTION
            ReturnLargeSet with a fresh SessionId and ResultSize 5000. The same command is used
            for every call of the session, because switching commands lowers the limit to 10,000.
            A call that returns nothing ends the session; so does a record whose
            AuditSearchRequestMetadata.moreRecordsAvailable is false.
            https://learn.microsoft.com/powershell/module/exchangepowershell/search-unifiedauditlog
            https://learn.microsoft.com/purview/audit-log-search-script
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][datetime]$Start,
        [Parameter(Mandatory)][datetime]$End
    )

    $sessionId = [guid]::NewGuid().ToString()
    $records = [System.Collections.Generic.List[object]]::new()
    while ($true) {
        $batch = @(Search-UnifiedAuditLog -StartDate $Start -EndDate $End -SessionId $sessionId `
                -SessionCommand ReturnLargeSet -ResultSize 5000 -Formatted -ErrorAction Stop | Where-Object { $null -ne $_ })
        if ($batch.Count -eq 0) { break }
        $records.AddRange($batch)

        $metadata = Get-ObjectValue -Object $batch[-1] -Name 'AuditSearchRequestMetadata'
        $more = Get-ObjectValue -Object $metadata -Name 'moreRecordsAvailable'
        if ($null -ne $more -and -not [System.Convert]::ToBoolean($more)) { break }
    }
    $records.ToArray()
}

function Get-AuditSearchSlice {
    <#
        .SYNOPSIS
            Every audit record in a date range, halving the range when a session reaches the
            50,000-record cap.

        .DESCRIPTION
            ReturnLargeSet pages up to 50,000 unsorted records. A session that gets there is not
            the full window, so its result is dropped and each half is read instead. A range down
            to -MinimumMinutes that still reaches 50,000 is returned with a warning in run.log.
            A record on a boundary can come back from both halves; the CSV key removes it.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][datetime]$Start,
        [Parameter(Mandatory)][datetime]$End,
        [Parameter(Mandatory)][string]$OutputPath,
        [string]$LogSource = 'audit-search-cmdlet',
        [int]$MinimumMinutes = 2
    )

    $records = @(Get-AuditSearchSession -Start $Start -End $End)
    if ($records.Count -lt 50000) { return $records }

    $minutes = ($End - $Start).TotalMinutes
    if ($minutes -le $MinimumMinutes) {
        Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $LogSource -Message (
            'The range {0:yyyy-MM-ddTHH:mm:ssZ} to {1:yyyy-MM-ddTHH:mm:ssZ} still reached the 50,000-record limit at {2} minutes. Its rows are incomplete.' -f $Start, $End, $minutes)
        return $records
    }
    $middle = $Start.AddTicks([long](($End - $Start).Ticks / 2))
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $LogSource -Message (
        'The range {0:yyyy-MM-ddTHH:mm:ssZ} to {1:yyyy-MM-ddTHH:mm:ssZ} reached the 50,000-record limit. Reading it as two halves.' -f $Start, $End)
    @(Get-AuditSearchSlice -Start $Start -End $middle -OutputPath $OutputPath -LogSource $LogSource -MinimumMinutes $MinimumMinutes)
    @(Get-AuditSearchSlice -Start $middle -End $End -OutputPath $OutputPath -LogSource $LogSource -MinimumMinutes $MinimumMinutes)
}

#endregion

#region Throttling

function Invoke-AuditThrottleRetry {
    <#
        .SYNOPSIS
            Runs a request and retries it while the service answers 429.

        .DESCRIPTION
            Waits for Retry-After when the error carries one, and backs off exponentially from
            30 seconds when it does not. It never retries at once.
            https://learn.microsoft.com/graph/throttling-limits#security-audit-log-query-service-limits
            Any other error, and a 429 after the last attempt, is rethrown.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][scriptblock]$Request,
        [int]$MaxAttempts = 5
    )

    $attempt = 0
    while ($true) {
        try {
            return (& $Request)
        }
        catch {
            $attempt++
            $throttled = $_.Exception.Message -match '\b429\b|Too ?Many ?Requests|AF429'
            if (-not $throttled -or $attempt -ge $MaxAttempts) { throw }

            $wait = 30 * [math]::Pow(2, $attempt - 1)
            try {
                $retryAfter = $_.Exception.Response.Headers.RetryAfter
                if ($retryAfter -and $retryAfter.Delta) { $wait = [math]::Ceiling($retryAfter.Delta.TotalSeconds) }
            }
            catch { Write-Verbose 'No Retry-After on the 429; backing off.' }
            Start-Sleep -Seconds $wait
        }
    }
}

#endregion

#region Source 2: Graph Audit Search API

function Test-GraphReadUri {
    <#
        .SYNOPSIS
            True when a Graph URL is relative or on a Microsoft Graph host, so a returned
            @odata.nextLink cannot send the session token elsewhere.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param([Parameter(Mandatory)][string]$Uri)

    if ($Uri -notmatch '^[a-z][a-z0-9+.-]*://') { return $true }
    $parsed = $null
    if (-not [uri]::TryCreate($Uri, [UriKind]::Absolute, [ref]$parsed)) { return $false }
    if ($parsed.Scheme -ne 'https') { return $false }
    return $parsed.Host -eq 'graph.microsoft.com' -or $parsed.Host -eq 'graph.microsoft.us'
}

function New-AuditGraphQuery {
    <#
        .SYNOPSIS
            Creates one audit log query and returns it.

        .DESCRIPTION
            POST /security/auditLog/queries. This creates the query object that records are then
            listed from; it does not change tenant data (D-007). The body is the documented
            auditLogQuery with a date range only.
            https://learn.microsoft.com/graph/api/security-auditcoreroot-post-auditlogqueries
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][datetime]$Start,
        [Parameter(Mandatory)][datetime]$End
    )

    $body = @{
        '@odata.type'       = '#microsoft.graph.security.auditLogQuery'
        displayName         = 'M365ReportLibrary unified audit log {0:yyyy-MM-ddTHH:mm:ssZ}' -f [datetime]::UtcNow
        filterStartDateTime = $Start.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        filterEndDateTime   = $End.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    } | ConvertTo-Json -Depth 5

    Invoke-AuditThrottleRetry -Request {
        Invoke-MgGraphRequest -Method POST -Uri '/v1.0/security/auditLog/queries' -Body $body -ContentType 'application/json' -ErrorAction Stop
    }
}

function Get-AuditGraphItem {
    <#
        .SYNOPSIS
            A Graph GET, retried while throttled.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Uri)

    if (-not (Test-GraphReadUri -Uri $Uri)) {
        throw "Refusing to request a page outside Microsoft Graph: $Uri"
    }
    Invoke-AuditThrottleRetry -Request { Invoke-MgGraphRequest -Method GET -Uri $Uri -ErrorAction Stop }
}

function Wait-AuditGraphQuery {
    <#
        .SYNOPSIS
            Polls an audit log query until it has succeeded, and returns it.

        .DESCRIPTION
            Status is notStarted, running, succeeded, failed or cancelled
            (https://learn.microsoft.com/graph/api/security-auditlogquery-get). failed and
            cancelled throw. A query still not succeeded after -MaxPolls polls throws, so a
            stuck query ends the run instead of hanging it.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$QueryId,
        [int]$PollSeconds = 15,
        [int]$MaxPolls = 240
    )

    for ($poll = 1; $poll -le $MaxPolls; $poll++) {
        $query = Get-AuditGraphItem -Uri "/v1.0/security/auditLog/queries/$QueryId"
        $status = [string](Get-ObjectValue -Object $query -Name 'status')
        if ($status -eq 'succeeded') { return $query }
        if ($status -in 'failed', 'cancelled') { throw "Audit log query $QueryId ended with status $status." }
        Start-Sleep -Seconds $PollSeconds
    }
    throw "Audit log query $QueryId did not succeed after $MaxPolls polls."
}

function Get-AuditGraphRecord {
    <#
        .SYNOPSIS
            Every record of a succeeded query, following @odata.nextLink until it is absent.

        .DESCRIPTION
            https://learn.microsoft.com/graph/api/security-auditlogquery-list-records. The paging
            of /records follows @odata.nextLink; the page does not confirm it separately
            (docs/candidates/unified-audit-log.md source 2), so a response without one is the end.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$QueryId)

    $next = "/v1.0/security/auditLog/queries/$QueryId/records"
    $visited = [System.Collections.Generic.HashSet[string]]::new()
    while ($next) {
        if (-not $visited.Add($next)) { throw "Graph returned a nextLink that was already requested: $next" }
        $page = Get-AuditGraphItem -Uri $next
        foreach ($item in @(Get-ObjectValue -Object $page -Name 'value')) {
            if ($null -ne $item) { $item }
        }
        $next = [string](Get-ObjectValue -Object $page -Name '@odata.nextLink')
    }
}

function ConvertTo-AuditGraphRow {
    <#
        .SYNOPSIS
            One Graph auditLogRecord as an audit-graph-records.csv row.

        .DESCRIPTION
            Property names are the auditLogRecord resource's: id, createdDateTime,
            auditLogRecordType, operation, organizationId, service, objectId, userId,
            userPrincipalName, clientIp, auditData.
            https://learn.microsoft.com/graph/api/security-auditlogquery-list-records
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Record,
        [Parameter(Mandatory)][string]$QueryId
    )

    $auditData = Get-ObjectValue -Object $Record -Name 'auditData'
    $json = if ($null -eq $auditData) { '' } else { $auditData | ConvertTo-Json -Depth 20 -Compress }
    $userId = Get-ObjectText -Object $Record -Name 'userPrincipalName'
    if ($userId -eq '') { $userId = Get-ObjectText -Object $Record -Name 'userId' }
    [pscustomobject][ordered]@{
        CreationTime   = ConvertTo-CsvTimestamp (Get-ObjectValue -Object $Record -Name 'createdDateTime')
        RecordId       = Get-ObjectText -Object $Record -Name 'id'
        RecordType     = Get-ObjectText -Object $Record -Name 'auditLogRecordType'
        Operation      = Get-ObjectText -Object $Record -Name 'operation'
        Workload       = Get-ObjectText -Object $Record -Name 'service'
        UserId         = $userId
        ObjectId       = Get-ObjectText -Object $Record -Name 'objectId'
        ClientIP       = Get-ObjectText -Object $Record -Name 'clientIp'
        ResultStatus   = Get-ObjectText -Object $auditData -Name 'ResultStatus'
        OrganizationId = Get-ObjectText -Object $Record -Name 'organizationId'
        QueryId        = $QueryId
        AuditData      = $json
    }
}

function Get-AuditGraphSlice {
    <#
        .SYNOPSIS
            The rows for a date range, one query per range, halving the range when the query went
            over its record limit.

        .DESCRIPTION
            Over the limit a query still reports succeeded, so isRecordCountLimitExceeded is read
            (https://learn.microsoft.com/graph/throttling-limits#security-audit-log-query-service-limits).
            approximateReturnedRecordCount is not compared with recordCountLimit. A range down to
            -MinimumMinutes that is still over the limit is returned with a warning.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][datetime]$Start,
        [Parameter(Mandatory)][datetime]$End,
        [Parameter(Mandatory)][string]$OutputPath,
        [string]$LogSource = 'audit-graph-records',
        [int]$MinimumMinutes = 60,
        [int]$PollSeconds = 15,
        [int]$MaxPolls = 240
    )

    $created = New-AuditGraphQuery -Start $Start -End $End
    $queryId = [string](Get-ObjectValue -Object $created -Name 'id')
    if ($queryId -eq '') { throw 'Graph created an audit log query with no id.' }

    $query = Wait-AuditGraphQuery -QueryId $queryId -PollSeconds $PollSeconds -MaxPolls $MaxPolls
    $exceeded = Get-ObjectValue -Object $query -Name 'isRecordCountLimitExceeded'
    $isOver = $null -ne $exceeded -and [System.Convert]::ToBoolean($exceeded)

    if ($isOver -and ($End - $Start).TotalMinutes -gt ($MinimumMinutes * 2)) {
        $middle = $Start.AddTicks([long](($End - $Start).Ticks / 2))
        Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $LogSource -Message (
            'Query {0} for {1:yyyy-MM-ddTHH:mm:ssZ} to {2:yyyy-MM-ddTHH:mm:ssZ} went over the record limit. Reading it as two halves.' -f $queryId, $Start, $End)
        Get-AuditGraphSlice -Start $Start -End $middle -OutputPath $OutputPath -LogSource $LogSource -MinimumMinutes $MinimumMinutes -PollSeconds $PollSeconds -MaxPolls $MaxPolls
        Get-AuditGraphSlice -Start $middle -End $End -OutputPath $OutputPath -LogSource $LogSource -MinimumMinutes $MinimumMinutes -PollSeconds $PollSeconds -MaxPolls $MaxPolls
        return
    }
    if ($isOver) {
        Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $LogSource -Message (
            'Query {0} is still over the record limit at {1} minutes. Its rows are incomplete.' -f $queryId, ($End - $Start).TotalMinutes)
    }

    foreach ($record in Get-AuditGraphRecord -QueryId $queryId) {
        ConvertTo-AuditGraphRow -Record $record -QueryId $queryId
    }
}

#endregion

#region Source 3: Office 365 Management Activity API

function Get-ActivityFeedBase {
    <#
        .SYNOPSIS
            https://{root}/api/v1.0/{tenant}/activity/feed for a cloud.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][ValidateSet('Commercial', 'GCC', 'GCCHigh')][string]$Environment,
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][hashtable]$Schema
    )

    $parsedTenant = [guid]::Empty
    if (-not [guid]::TryParse($TenantId, [ref]$parsedTenant)) { throw "TenantId '$TenantId' is not a GUID." }
    '{0}/api/v1.0/{1}/activity/feed' -f $Schema.ActivityFeedRoot[$Environment], $parsedTenant
}

function Test-ActivityFeedUri {
    <#
        .SYNOPSIS
            True when a URL is on the same host as the feed root, over https. A returned
            NextPageUri or contentUri on any other host is refused so the token is not sent there.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][string]$Base
    )

    $parsed = $null
    if (-not [uri]::TryCreate($Uri, [UriKind]::Absolute, [ref]$parsed)) { return $false }
    if ($parsed.Scheme -ne 'https') { return $false }
    return $parsed.Host -eq ([uri]$Base).Host
}

function Add-ActivityFeedPublisher {
    <#
        .SYNOPSIS
            Adds PublisherIdentifier to a URL that lacks it. Requests without it share one quota.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][string]$PublisherIdentifier
    )

    if ($Uri -match '[?&]PublisherIdentifier=') { return $Uri }
    $separator = if ($Uri.Contains('?')) { '&' } else { '?' }
    return '{0}{1}PublisherIdentifier={2}' -f $Uri, $separator, $PublisherIdentifier
}

function Get-ActivityFeedHeader {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Response, [Parameter(Mandatory)][string]$Name)

    $headers = Get-ObjectValue -Object $Response -Name 'Headers'
    $value = Get-ObjectValue -Object $headers -Name $Name
    if ($null -eq $value) { return $null }
    return [string](@($value)[0])
}

function Invoke-AuditActivityGet {
    <#
        .SYNOPSIS
            A GET to the Management Activity API with the bearer token, retried while throttled.

        .DESCRIPTION
            The token is read from the SecureString here and nowhere else. It is never logged.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][securestring]$AccessToken,
        [Parameter(Mandatory)][string]$Base,
        [Parameter(Mandatory)][string]$PublisherIdentifier
    )

    if (-not (Test-ActivityFeedUri -Uri $Uri -Base $Base)) {
        throw "Refusing to request a URL outside the Management Activity API: $Uri"
    }
    $target = Add-ActivityFeedPublisher -Uri $Uri -PublisherIdentifier $PublisherIdentifier
    $headers = @{ Authorization = 'Bearer ' + (ConvertFrom-SecureString -SecureString $AccessToken -AsPlainText) }
    Invoke-AuditThrottleRetry -Request { Invoke-WebRequest -Method Get -Uri $target -Headers $headers -ErrorAction Stop }
}

function Start-AuditActivitySubscription {
    <#
        .SYNOPSIS
            Starts the subscription for one content type.

        .DESCRIPTION
            POST {root}/subscriptions/start?contentType=... with no body. Listing and retrieving
            content needs the subscription. It does not change tenant data (D-007). A second start
            within 15 minutes is throttled, and stopping then starting a subscription does not
            return the content from the gap, so the collector starts only a type that is not
            enabled and never stops one.
            https://learn.microsoft.com/office/office-365-management-api/office-365-management-activity-api-reference#start-a-subscription
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Base,
        [Parameter(Mandatory)][string]$ContentType,
        [Parameter(Mandatory)][securestring]$AccessToken,
        [Parameter(Mandatory)][string]$PublisherIdentifier
    )

    $uri = '{0}/subscriptions/start?contentType={1}&PublisherIdentifier={2}' -f $Base, [uri]::EscapeDataString($ContentType), $PublisherIdentifier
    $headers = @{ Authorization = 'Bearer ' + (ConvertFrom-SecureString -SecureString $AccessToken -AsPlainText) }
    Invoke-AuditThrottleRetry -Request { Invoke-WebRequest -Method Post -Uri $uri -Headers $headers -ErrorAction Stop }
}

function Get-AuditActivitySubscription {
    <#
        .SYNOPSIS
            The content types that have an enabled subscription.
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [Parameter(Mandatory)][string]$Base,
        [Parameter(Mandatory)][securestring]$AccessToken,
        [Parameter(Mandatory)][string]$PublisherIdentifier
    )

    $response = Invoke-AuditActivityGet -Uri "$Base/subscriptions/list" -AccessToken $AccessToken -Base $Base -PublisherIdentifier $PublisherIdentifier
    $content = Get-ObjectText -Object $response -Name 'Content'
    if ($content -eq '') { return , [string[]]@() }
    $items = @($content | ConvertFrom-Json)
    return , [string[]]@($items | Where-Object { (Get-ObjectText -Object $_ -Name 'status') -eq 'enabled' } |
            ForEach-Object { Get-ObjectText -Object $_ -Name 'contentType' })
}

function Get-AuditActivityContent {
    <#
        .SYNOPSIS
            The content blobs listed for a content type and a window of at most 24 hours, paged
            through the NextPageUri response header until it is absent.

        .DESCRIPTION
            startTime and endTime select on contentCreated, not on the event time; startTime is
            inclusive and endTime exclusive, and they must be used together.
            https://learn.microsoft.com/office/office-365-management-api/office-365-management-activity-api-reference#list-available-content
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Base,
        [Parameter(Mandatory)][string]$ContentType,
        [Parameter(Mandatory)][datetime]$Start,
        [Parameter(Mandatory)][datetime]$End,
        [Parameter(Mandatory)][securestring]$AccessToken,
        [Parameter(Mandatory)][string]$PublisherIdentifier
    )

    if (($End - $Start).TotalHours -gt 24) { throw 'A content listing window must be 24 hours or less.' }

    $next = '{0}/subscriptions/content?contentType={1}&startTime={2:yyyy-MM-ddTHH:mm:ss}&endTime={3:yyyy-MM-ddTHH:mm:ss}' -f `
        $Base, [uri]::EscapeDataString($ContentType), $Start.ToUniversalTime(), $End.ToUniversalTime()
    $visited = [System.Collections.Generic.HashSet[string]]::new()
    while ($next) {
        if (-not $visited.Add($next)) { throw "The API returned a NextPageUri that was already requested: $next" }
        $response = Invoke-AuditActivityGet -Uri $next -AccessToken $AccessToken -Base $Base -PublisherIdentifier $PublisherIdentifier
        $content = Get-ObjectText -Object $response -Name 'Content'
        if ($content -ne '') {
            foreach ($blob in @($content | ConvertFrom-Json)) { if ($null -ne $blob) { $blob } }
        }
        $next = Get-ActivityFeedHeader -Response $response -Name 'NextPageUri'
    }
}

function Get-AuditActivityEvent {
    <#
        .SYNOPSIS
            The events of one content blob as audit-activity-feed.csv rows.

        .DESCRIPTION
            A blob holds 1 to N events, so a count of blobs is not a count of events.
            https://learn.microsoft.com/office/office-365-management-api/troubleshooting-the-office-365-management-activity-api
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Blob,
        [Parameter(Mandatory)][securestring]$AccessToken,
        [Parameter(Mandatory)][string]$Base,
        [Parameter(Mandatory)][string]$PublisherIdentifier
    )

    $uri = Get-ObjectText -Object $Blob -Name 'contentUri'
    if ($uri -eq '') { throw 'A content blob has no contentUri.' }
    $response = Invoke-AuditActivityGet -Uri $uri -AccessToken $AccessToken -Base $Base -PublisherIdentifier $PublisherIdentifier
    $content = Get-ObjectText -Object $response -Name 'Content'
    if ($content -eq '') { return }

    $contentCreated = ConvertTo-CsvTimestamp (Get-ObjectValue -Object $Blob -Name 'contentCreated')
    # ConvertFrom-Json turns date strings into dates, so the event's own text is kept from the
    # parser and written to AuditData as it arrived.
    $document = [System.Text.Json.JsonDocument]::Parse($content)
    try {
        foreach ($element in $document.RootElement.EnumerateArray()) {
            $raw = $element.GetRawText() -replace '\s*[\r\n]+\s*', ''
            $item = $raw | ConvertFrom-Json
            $row = New-AuditRow -Data $item -Json $raw
            $ordered = [ordered]@{
                ContentCreated = $contentCreated
                ContentType    = Get-ObjectText -Object $Blob -Name 'contentType'
                ContentId      = Get-ObjectText -Object $Blob -Name 'contentId'
            }
            foreach ($key in $row.Keys) { $ordered[$key] = $row[$key] }
            [pscustomobject]$ordered
        }
    }
    finally {
        $document.Dispose()
    }
}

#endregion
