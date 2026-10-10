#Requires -Version 7.0

# Dot-sourced by the Teams activity collectors. Report-specific: the shared connection
# and CSV functions come from shared/M365ReportLibrary.psm1. Helpers here follow the ones in
# reports/exchange-activity and reports/sharepoint-onedrive-activity.

function Get-TeamsSourceAvailability {
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

function Test-TeamsSourceSkipped {
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

    $availability = Get-TeamsSourceAvailability -Source $Source -Environment $Environment -Schema $Schema
    if ($availability.ShouldSkip) {
        Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $LogSource -Message (
            "Skipping $(Split-Path -Leaf $CsvPath). $($availability.Reason) $($availability.Reference)")
        Export-AppendCsv -Path $CsvPath -Column $Column
        return $true
    }
    if ($availability.Status -eq 'Unverified') {
        Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $LogSource -Message $availability.Reason
    }
    return $false
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

function Invoke-UsageReportCollector {
    <#
        .SYNOPSIS
            Reads one Graph usage report (a CSV behind a 302 redirect) and appends a snapshot
            stamped with the run date.

        .DESCRIPTION
            -Fetch is called with the path of a temporary file and must save the report there
            (Get-MgReport* -OutFile). Graph answers with a 302 to a preauthenticated download
            URL valid for a few minutes; the cmdlet follows it at once. The file is read and
            removed. A refusal writes the header only and logs the reason.
            https://learn.microsoft.com/graph/api/reportroot-getteamsuseractivityuserdetail
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)][string]$LogSource,
        [Parameter(Mandatory)][string]$CsvName,
        [Parameter(Mandatory)][string[]]$Column,
        [Parameter(Mandatory)][string[]]$KeyColumn,
        [Parameter(Mandatory)][scriptblock]$Fetch,
        [Parameter(Mandatory)][scriptblock]$MapRow,
        [Parameter(Mandatory)][string]$ReportName
    )

    $csvPath = Join-Path $OutputPath $CsvName
    $runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')
    $download = Join-Path ([System.IO.Path]::GetTempPath()) ('teams-activity-' + [guid]::NewGuid().ToString('N') + '.csv')
    $report = $null
    try {
        try {
            & $Fetch $download
            $report = if (Test-Path -LiteralPath $download) { @(Import-Csv -LiteralPath $download) } else { @() }
        }
        catch {
            Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $LogSource -Message (
                'The {0} report is unavailable to this sign-in or cloud ({1}). It needs Reports.Read.All and, for a delegated sign-in, a role such as Reports Reader. Writing the header only.' -f $ReportName, $_.Exception.Message)
            Export-AppendCsv -Path $csvPath -Column $Column
            return
        }
    }
    finally {
        Remove-Item -LiteralPath $download -Force -ErrorAction SilentlyContinue
    }

    $rows = @($report | ForEach-Object { & $MapRow $_ $runDate })
    $rows = @($rows | Where-Object { $null -ne $_ })

    $result = Export-AppendCsv -Path $csvPath -Rows $rows -Column $Column -KeyColumn $KeyColumn -PassThru
    Write-CollectorLog -OutputPath $OutputPath -Source $LogSource -Message (
        '{0}: {1} rows written, {2} skipped.' -f $CsvName, $result.Written, $result.Skipped)
}

function Get-ObjectText {
    <#
        .SYNOPSIS
            A property of a cmdlet result as text, or an empty string when the object has no
            such property. Few cmdlet pages document their output properties, so a missing one
            is normal and must not stop a run.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][AllowNull()]$Object,
        [Parameter(Mandatory)][string]$Name
    )

    if ($null -eq $Object) { return '' }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) { return '' }
    return [string]$property.Value
}

function Get-ObjectValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowNull()]$Object,
        [Parameter(Mandatory)][string]$Name
    )

    if ($null -eq $Object) { return $null }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    return $property.Value
}

function Get-CsvBoolean {
    <#
        .SYNOPSIS
            'True' or 'False' for a value that is set, an empty string for one that is not.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Position = 0)][AllowNull()]$Value)

    if ($null -eq $Value -or ($Value -is [string] -and [string]::IsNullOrWhiteSpace($Value))) { return '' }
    if ($Value -is [string]) {
        $parsed = $false
        if ([bool]::TryParse($Value, [ref]$parsed)) { return $parsed.ToString() }
        return $Value
    }
    return ([bool]$Value).ToString()
}

function Test-GraphReadUri {
    <#
        .SYNOPSIS
            True when a Graph read URL is relative or on a Microsoft Graph host.

        .DESCRIPTION
            A returned @odata.nextLink is requested as it is, but only on
            https://graph.microsoft.com or https://graph.microsoft.us
            (https://learn.microsoft.com/graph/deployments). Any other host is refused so the
            session token is not sent there.
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

function Get-GraphJsonValue {
    <#
        .SYNOPSIS
            Reads one key from a hashtable or object returned by Invoke-MgGraphRequest, or $null.

        .DESCRIPTION
            The name is one key, not a path, because Graph names such as @odata.nextLink contain dots.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowNull()]$Object,
        [Parameter(Mandatory)][string]$Name
    )

    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name)) { return $Object[$Name] }
        return $null
    }
    $property = $Object.PSObject.Properties[$Name]
    if ($property) { return $property.Value }
    return $null
}

function Get-GraphHttpStatus {
    <#
        .SYNOPSIS
            The HTTP status on a failed Graph read, or $null.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $ErrorRecord
    )

    $candidates = @($ErrorRecord.Exception, $ErrorRecord.Exception.InnerException)
    foreach ($exception in $candidates) {
        if ($null -eq $exception) { continue }
        $response = $exception.PSObject.Properties['Response']
        if (-not $response -or $null -eq $response.Value) { continue }
        $status = $response.Value.PSObject.Properties['StatusCode']
        if ($status -and $null -ne $status.Value) { return [int]$status.Value }
    }

    # PowerShell 7 and the Graph SDK: "Response status code does not indicate success: 429 (Too Many Requests)."
    if ($ErrorRecord.Exception.Message -match '\b(404|429|503)\b') {
        return [int]$Matches[1]
    }
    return $null
}

function Test-GraphThrottleStatus {
    <#
        .SYNOPSIS
            True when a failed read is HTTP 429 or 503.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)]
        $ErrorRecord
    )

    $status = Get-GraphHttpStatus -ErrorRecord $ErrorRecord
    return ($status -eq 429 -or $status -eq 503)
}

function Get-GraphRetryDelaySeconds {
    <#
        .SYNOPSIS
            Seconds to wait after HTTP 429 or 503.

        .DESCRIPTION
            Graph returns Retry-After on 429
            (https://learn.microsoft.com/graph/throttling).
            When the header is absent, wait doubles each attempt
            (https://learn.microsoft.com/graph/throttling-limits).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $ErrorRecord,

        [Parameter(Mandatory)]
        [int]$Attempt
    )

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

    if ($delay -lt 1) { return 1 }
    return $delay
}

function Invoke-ReadWithThrottleRetry {
    <#
        .SYNOPSIS
            Runs a read, and on HTTP 429 or 503 waits and tries again.

        .DESCRIPTION
            A 429 or 503 is throttling, not an empty result and not a refused permission.
            Four attempts, then the error is rethrown. 401 and 403 are not retried.
            https://learn.microsoft.com/graph/throttling
            https://learn.microsoft.com/graph/throttling-limits
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][scriptblock]$Action,
        [string]$OutputPath,
        [string]$Source,
        [int]$MaxAttempts = 4
    )

    for ($attempt = 1; ; $attempt++) {
        try {
            return (& $Action)
        }
        catch {
            $status = Get-GraphHttpStatus -ErrorRecord $_
            if (($status -ne 429 -and $status -ne 503) -or $attempt -ge $MaxAttempts) { throw }
            $delay = Get-GraphRetryDelaySeconds -ErrorRecord $_ -Attempt $attempt
            if ($OutputPath -and $Source) {
                Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $Source -Message (
                    'HTTP {0}. Waiting {1} seconds before attempt {2}.' -f $status, $delay, ($attempt + 1))
            }
            Start-Sleep -Seconds $delay
        }
    }
}

function Invoke-GraphGet {
    <#
        .SYNOPSIS
            One Graph read with GET, retrying HTTP 429 and 503. The only Invoke-MgGraphRequest
            call in this report.

        .DESCRIPTION
            Refuses a URL that is not relative or on a Microsoft Graph host, so the session
            token is not sent elsewhere. Only GET is sent.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [string]$OutputPath,
        [string]$LogSource
    )

    if (-not (Test-GraphReadUri -Uri $Uri)) {
        throw "Refusing to request a page outside Microsoft Graph: $Uri"
    }
    $requestUri = $Uri
    Invoke-ReadWithThrottleRetry -OutputPath $OutputPath -Source $LogSource -Action {
        Invoke-MgGraphRequest -Method GET -Uri $requestUri -ErrorAction Stop
    }
}

function Get-GraphPagedValue {
    <#
        .SYNOPSIS
            Every item of a Graph collection, following @odata.nextLink until it is absent.

        .DESCRIPTION
            https://learn.microsoft.com/graph/paging. The returned URL is requested as it is.
            For getAllSites the sample nextLink changes the path to oneDrive.getAllSites, so
            rebuilding /sites/getAllSites from a skiptoken would ask for the wrong page.
            The host must be graph.microsoft.com or graph.microsoft.us
            (https://learn.microsoft.com/graph/deployments). Only GET is sent.
            HTTP 429 and 503 are retried. Graph sends Retry-After on 429
            (https://learn.microsoft.com/graph/throttling).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [string]$OutputPath,
        [string]$LogSource
    )

    $next = $Uri
    $visited = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
    while (-not [string]::IsNullOrEmpty($next)) {
        if (-not (Test-GraphReadUri -Uri $next)) {
            throw "Refusing to request a page outside Microsoft Graph: $next"
        }
        if (-not $visited.Add($next)) {
            throw "Graph returned a nextLink that was already requested: $next"
        }
        $page = Invoke-GraphGet -Uri $next -OutputPath $OutputPath -LogSource $LogSource
        foreach ($item in @(Get-GraphJsonValue -Object $page -Name 'value')) {
            if ($null -ne $item) { $item }
        }
        $next = [string](Get-GraphJsonValue -Object $page -Name '@odata.nextLink')
    }
}

function Test-AuditSearchHasMoreRecords {
    <#
        .SYNOPSIS
            True when a Search-UnifiedAuditLog record says another page is expected.
    #>
    [CmdletBinding()]
    param($Record)

    if ($null -eq $Record) { return $false }

    $metaProperty = $Record.PSObject.Properties['AuditSearchRequestMetadata']
    if (-not $metaProperty -or $null -eq $metaProperty.Value) { return $false }

    $meta = $metaProperty.Value
    $flag = $null
    foreach ($key in @('moreRecordsAvailable', 'MoreRecordsAvailable')) {
        $value = Get-GraphJsonValue -Object $meta -Name $key
        if ($null -ne $value) { $flag = $value; break }
    }

    if ($null -eq $flag) { return $false }
    if ($flag -is [bool]) { return $flag }

    $parsed = $false
    if ([bool]::TryParse([string]$flag, [ref]$parsed)) { return $parsed }
    return $false
}

function ConvertTo-AuditQueryDate {
    <#
        .SYNOPSIS
            A Search-UnifiedAuditLog date in UTC.

        .DESCRIPTION
            -StartDate and -EndDate are stored in UTC. A value with no time zone is
            midnight UTC; converting it from the machine's local zone would shift the window.
            https://learn.microsoft.com/powershell/module/exchangepowershell/search-unifiedauditlog
    #>
    [CmdletBinding()]
    [OutputType([datetime])]
    param([Parameter(Mandatory)][datetime]$Value)

    if ($Value.Kind -eq [DateTimeKind]::Unspecified) {
        return [datetime]::SpecifyKind($Value, [DateTimeKind]::Utc)
    }
    return $Value.ToUniversalTime()
}

function Invoke-AuditSearch {
    <#
        .SYNOPSIS
            Searches the unified audit log over a range, in windows, paging each window
            with a ReturnLargeSet session.

        .DESCRIPTION
            Search-UnifiedAuditLog returns at most 100 records unless the same -SessionId
            is repeated with -SessionCommand ReturnLargeSet, which pages up to 50,000
            records a session (-ResultSize up to 5,000). A window that reaches 50,000 is
            incomplete and unsorted, so it is not returned: the caller gets the window in
            TruncatedWindow instead. The operations are filtered with -Operations;
            the Teams operations are named in
            https://learn.microsoft.com/purview/audit-log-activities#teams-activities.
            https://learn.microsoft.com/powershell/module/exchangepowershell/search-unifiedauditlog

        .OUTPUTS
            An object with Records (one object per record: Audit, the parsed AuditData, and
            RecordType) and TruncatedWindow (a Start/End window, or $null).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][datetime]$Start,
        [Parameter(Mandatory)][datetime]$End,
        [Parameter(Mandatory)][int]$WindowHours,
        [Parameter(Mandatory)][string[]]$Operation,
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)][string]$Source
    )

    $pageSize = 5000
    $sessionCap = 50000
    $nullPageRetries = 3
    $nullPageRetryDelayMs = 200

    $all = [System.Collections.Generic.List[object]]::new()
    $truncated = $null

    foreach ($window in Split-DateRange -Start $Start -End $End -WindowMinutes ($WindowHours * 60)) {
        $sessionId = 'teams-activity-{0:yyyyMMddHHmmss}-{1}' -f $window.Start, [guid]::NewGuid().ToString('N').Substring(0, 8)
        $collected = 0
        $nullTries = 0
        $windowRecords = [System.Collections.Generic.List[object]]::new()
        $windowTruncated = $false

        while ($collected -lt $sessionCap) {
            $search = @{
                StartDate      = $window.Start
                EndDate        = $window.End
                Operations     = $Operation
                SessionId      = $sessionId
                SessionCommand = 'ReturnLargeSet'
                ResultSize     = $pageSize
                # Without -Formatted, RecordType is an integer. The activities page uses
                # the display name, such as MicrosoftTeams.
                # https://learn.microsoft.com/powershell/module/exchangepowershell/search-unifiedauditlog
                Formatted      = $true
                ErrorAction    = 'Stop'
            }

            $raw = Search-UnifiedAuditLog @search

            # $null and an empty collection both mean "nothing this call". Retry briefly
            # at the start of a window, because the service often returns nothing while
            # the search is prepared, then treat the window as empty.
            $records = @($raw | Where-Object { $null -ne $_ })
            if ($records.Count -eq 0) {
                if ($collected -eq 0 -and $nullTries -lt $nullPageRetries) {
                    $nullTries++
                    Start-Sleep -Milliseconds $nullPageRetryDelayMs
                    continue
                }
                break
            }

            # ResultCount is the hit count across every iteration of the session, not the
            # size of this page. ReturnLargeSet stops at 50,000, so a count of exactly
            # 50,000 is already a capped window: the records are unsorted and the ones
            # past the cap were never returned.
            # https://learn.microsoft.com/powershell/module/exchangepowershell/search-unifiedauditlog
            $matched = 0
            $hasResultCount = $false
            $resultCountProperty = $records[0].PSObject.Properties['ResultCount']
            if ($resultCountProperty -and $null -ne $resultCountProperty.Value) {
                $hasResultCount = [int]::TryParse([string]$resultCountProperty.Value, [ref]$matched)
            }
            if ($hasResultCount -and $matched -ge $sessionCap) {
                $windowTruncated = $true
                break
            }

            $collected += $records.Count

            foreach ($record in $records) {
                $auditDataProperty = $record.PSObject.Properties['AuditData']
                if (-not $auditDataProperty -or [string]::IsNullOrWhiteSpace([string]$auditDataProperty.Value)) { continue }
                try {
                    $audit = [string]$auditDataProperty.Value | ConvertFrom-Json -ErrorAction Stop
                }
                catch {
                    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $Source -Message (
                        'Skipping a record whose AuditData is not valid JSON: {0}' -f $_.Exception.Message)
                    continue
                }
                $recordTypeProperty = $record.PSObject.Properties['RecordType']
                $windowRecords.Add([pscustomobject]@{
                        Audit      = $audit
                        RecordType = $(if ($recordTypeProperty) { [string]$recordTypeProperty.Value } else { '' })
                    })
            }

            if ($collected -ge $sessionCap) {
                $windowTruncated = $true
                break
            }

            # Repeat until the cmdlet returns nothing or the cap is hit. moreRecordsAvailable
            # says another iteration is expected even when a page is short.
            if (@($records | Where-Object { Test-AuditSearchHasMoreRecords -Record $_ }).Count -gt 0) { continue }
            if ($hasResultCount -and $matched -gt 0 -and $collected -ge $matched) { break }
            if ((-not $hasResultCount -or $matched -le 0) -and $records.Count -lt $pageSize) { break }
        }

        if ($windowTruncated) {
            $truncated = $window
            break
        }

        foreach ($item in $windowRecords) { $all.Add($item) }
    }

    [pscustomobject]@{ Records = $all.ToArray(); TruncatedWindow = $truncated }
}

function Get-GraphJsonPath {
    <#
        .SYNOPSIS
            Reads a dotted path such as quota.used from a Graph object, or $null.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowNull()]$Object,
        [Parameter(Mandatory)][string]$Path
    )

    $current = $Object
    foreach ($part in $Path.Split('.')) {
        if ($null -eq $current) { return $null }
        $current = Get-GraphJsonValue -Object $current -Name $part
    }
    return $current
}

function Group-TeamsEventCount {
    <#
        .SYNOPSIS
            Counts audit records per UTC day, workload, user and operation, for teams-audit-events.csv.

        .DESCRIPTION
            The day is the record's CreationTime in UTC. Days on or after -Before are left out:
            a day that holds a truncated window is incomplete, and writing it would move the
            resume point past events that were never returned. Records with no CreationTime,
            UserId or Operation are not counted.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Item,
        [AllowNull()][nullable[datetime]]$Before
    )

    $counts = @{}
    foreach ($entry in $Item) {
        $audit = $entry.Audit
        $created = ConvertTo-CsvTimestamp (Get-GraphAdditionalProperty -Object $audit -Name 'CreationTime')
        $userId = [string](Get-GraphAdditionalProperty -Object $audit -Name 'UserId')
        $operation = [string](Get-GraphAdditionalProperty -Object $audit -Name 'Operation')
        if (-not $created -or -not $userId -or -not $operation) { continue }

        $parsed = [datetime]::Parse($created, [cultureinfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal)
        if ($null -ne $Before -and $parsed.Date -ge $Before) { continue }

        $workload = [string](Get-GraphAdditionalProperty -Object $audit -Name 'Workload')
        $key = '{0}|{1}|{2}|{3}' -f $parsed.ToString('yyyy-MM-dd'), $workload, $userId, $operation
        if ($counts.ContainsKey($key)) { $counts[$key].EventCount++ }
        else {
            $counts[$key] = [pscustomobject]@{
                Date       = $parsed.ToString('yyyy-MM-dd')
                Workload   = $workload
                UserId     = $userId
                Operation  = $operation
                EventCount = 1
            }
        }
    }
    @($counts.Values | Sort-Object Date, Workload, UserId, Operation)
}

function ConvertTo-CallRecordRow {
    <#
        .SYNOPSIS
            The CSV rows for one call record: one per session, or one with empty session columns.

        .DESCRIPTION
            The callRecord, session and participantEndpoint property names are from
            https://learn.microsoft.com/graph/api/resources/callrecords-callrecord,
            .../callrecords-session and .../callrecords-participantendpoint. The endpoint's
            deprecated identity property stops returning data on 2026-06-30, so the user id is
            read from associatedIdentity. The platform is userAgent.platform
            (https://learn.microsoft.com/graph/api/resources/callrecords-clientuseragent).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Record,
        [AllowEmptyCollection()][object[]]$Session = @()
    )

    function Text($Object, [string]$Path) { $v = Get-GraphJsonPath -Object $Object -Path $Path; if ($null -eq $v) { '' } else { [string]$v } }

    $modalities = @(Get-GraphJsonValue -Object $Record -Name 'modalities') -join ';'
    $base = [ordered]@{
        CallRecordId         = Text $Record 'id'
        Version              = Text $Record 'version'
        Type                 = Text $Record 'type'
        Modalities           = $modalities
        StartDateTime        = ConvertTo-CsvTimestamp (Get-GraphJsonValue -Object $Record -Name 'startDateTime')
        EndDateTime          = ConvertTo-CsvTimestamp (Get-GraphJsonValue -Object $Record -Name 'endDateTime')
        LastModifiedDateTime = ConvertTo-CsvTimestamp (Get-GraphJsonValue -Object $Record -Name 'lastModifiedDateTime')
    }

    function Join-Row($Extra) {
        $row = [ordered]@{}
        foreach ($key in $base.Keys) { $row[$key] = $base[$key] }
        foreach ($key in $Extra.Keys) { $row[$key] = $Extra[$key] }
        [pscustomobject]$row
    }

    $sessions = @($Session | Where-Object { $null -ne $_ })
    if ($sessions.Count -eq 0) {
        Join-Row ([ordered]@{
                SessionId = ''; SessionStartDateTime = ''; SessionEndDateTime = ''
                CallerUserId = ''; CallerPlatform = ''; CalleeUserId = ''; CalleePlatform = ''
            })
        return
    }

    foreach ($s in $sessions) {
        Join-Row ([ordered]@{
                SessionId            = Text $s 'id'
                SessionStartDateTime = ConvertTo-CsvTimestamp (Get-GraphJsonValue -Object $s -Name 'startDateTime')
                SessionEndDateTime   = ConvertTo-CsvTimestamp (Get-GraphJsonValue -Object $s -Name 'endDateTime')
                CallerUserId         = Text $s 'caller.associatedIdentity.id'
                CallerPlatform       = Text $s 'caller.userAgent.platform'
                CalleeUserId         = Text $s 'callee.associatedIdentity.id'
                CalleePlatform       = Text $s 'callee.userAgent.platform'
            })
    }
}
