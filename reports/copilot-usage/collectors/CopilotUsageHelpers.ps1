#Requires -Version 7.0

# Dot-sourced by the Copilot usage collectors. Report-specific: the shared connection and
# CSV functions come from shared/M365ReportLibrary.psm1. The throttle, GET and audit-search
# helpers follow the ones in reports/teams-activity, which read the same kinds of source.

function Get-CopilotSourceAvailability {
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

function Test-CopilotSourceSkipped {
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

    $availability = Get-CopilotSourceAvailability -Source $Source -Environment $Environment -Schema $Schema
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

function Get-CopilotColumn {
    <#
        .SYNOPSIS
            The CSV columns of a report source: RunDate, then the columns of its header map.
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [Parameter(Mandatory)][hashtable]$Schema,
        [Parameter(Mandatory)][string]$MapKey
    )

    return , [string[]](@('RunDate') + @($Schema[$MapKey] | ForEach-Object { $_.Column }))
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
            the Copilot operation is named in https://learn.microsoft.com/purview/audit-copilot.
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
        $sessionId = 'copilot-usage-{0:yyyyMMddHHmmss}-{1}' -f $window.Start, [guid]::NewGuid().ToString('N').Substring(0, 8)
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
                # Without -Formatted, RecordType is an integer. The Purview Copilot page uses
                # the name, such as CopilotInteraction.
                # https://learn.microsoft.com/powershell/module/exchangepowershell/search-unifiedauditlog
                Formatted      = $true
                ErrorAction    = 'Stop'
            }

            # A 429 or 503 is throttling, not an empty window and not a missing role.
            # https://learn.microsoft.com/graph/throttling
            $raw = Invoke-ReadWithThrottleRetry -OutputPath $OutputPath -Source $Source -Action {
                Search-UnifiedAuditLog @search
            }

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

function Invoke-GraphGet {
    <#
        .SYNOPSIS
            One Graph read with GET, retrying HTTP 429 and 503. The only Invoke-MgGraphRequest
            call in this report.

        .DESCRIPTION
            Refuses a URL that is not relative or on a Microsoft Graph host, so the session
            token is not sent elsewhere. Only GET is sent. With -OutputFilePath the response
            body (the CSV stream of a usage report) is saved there instead of returned.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [string]$OutputFilePath,
        [string]$OutputPath,
        [string]$LogSource
    )

    if (-not (Test-GraphReadUri -Uri $Uri)) {
        throw "Refusing to request a page outside Microsoft Graph: $Uri"
    }
    $requestUri = $Uri
    $extra = @{}
    if ($OutputFilePath) { $extra['OutputFilePath'] = $OutputFilePath }
    Invoke-ReadWithThrottleRetry -OutputPath $OutputPath -Source $LogSource -Action {
        Invoke-MgGraphRequest -Method GET -Uri $requestUri -ErrorAction Stop @extra
    }
}

function Get-GraphPagedValue {
    <#
        .SYNOPSIS
            Every item of a Graph collection, following @odata.nextLink until it is absent.

        .DESCRIPTION
            https://learn.microsoft.com/graph/paging. The returned URL is requested as it is.
            The host must be graph.microsoft.com or graph.microsoft.us
            (https://learn.microsoft.com/graph/deployments). Only GET is sent.
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

function Invoke-CopilotReportCollector {
    <#
        .SYNOPSIS
            Reads one Copilot usage report API (a CSV stream) and appends a snapshot stamped
            with the run date.

        .DESCRIPTION
            v1.0 answers 200 OK with the CSV in the body, not a redirect and not a paged
            collection. The request is saved to a temporary file, read with Import-Csv and
            removed. Each column is read by the header name the Learn page gives
            (CopilotUsageSchema.psd1); a header the report does not carry reads as empty,
            because v1 carries fewer columns than v2. A refusal writes the header only and
            logs the reason; a 429 or 503 is retried and, if it persists, logged as
            throttling rather than as an empty report.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)][string]$LogSource,
        [Parameter(Mandatory)][string]$CsvName,
        [Parameter(Mandatory)][hashtable]$Schema,
        [Parameter(Mandatory)][string]$MapKey,
        [Parameter(Mandatory)][string[]]$KeyColumn,
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][string]$ReportName
    )

    $columns = Get-CopilotColumn -Schema $Schema -MapKey $MapKey
    $csvPath = Join-Path $OutputPath $CsvName
    $runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')
    $download = Join-Path ([System.IO.Path]::GetTempPath()) ('copilot-usage-' + [guid]::NewGuid().ToString('N') + '.csv')
    $report = $null
    try {
        Invoke-GraphGet -Uri $Uri -OutputFilePath $download -OutputPath $OutputPath -LogSource $LogSource | Out-Null
        $report = if (Test-Path -LiteralPath $download) { @(Import-Csv -LiteralPath $download) } else { @() }
    }
    catch {
        $status = Get-GraphHttpStatus -ErrorRecord $_
        if ($status -eq 429 -or $status -eq 503) {
            Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $LogSource -Message (
                'The {0} report is still throttled after retries ({1}). A 429 or 503 is not an empty report. Writing the header only.' -f $ReportName, $_.Exception.Message)
        }
        else {
            Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $LogSource -Message (
                'The {0} report is unavailable to this sign-in or cloud ({1}). It needs Reports.Read.All and, for a delegated sign-in, a role such as Reports Reader or AI Administrator. Writing the header only.' -f $ReportName, $_.Exception.Message)
        }
        Export-AppendCsv -Path $csvPath -Column $columns
        return
    }
    finally {
        Remove-Item -LiteralPath $download -Force -ErrorAction SilentlyContinue
    }

    $map = @($Schema[$MapKey])
    $rows = foreach ($line in $report) {
        $row = [ordered]@{ RunDate = $runDate }
        foreach ($entry in $map) {
            $row[$entry.Column] = Get-ReportColumnValue -Row $line -Header $entry.Header
        }
        [pscustomobject]$row
    }

    $result = Export-AppendCsv -Path $csvPath -Rows @($rows) -Column $columns -KeyColumn $KeyColumn -PassThru
    Write-CollectorLog -OutputPath $OutputPath -Source $LogSource -Message (
        '{0}: {1} rows written, {2} skipped.' -f $CsvName, $result.Written, $result.Skipped)
}

function Get-FirstValue {
    <#
        .SYNOPSIS
            The first non-empty value of a property found at any of several places in a record.

        .DESCRIPTION
            The Purview page lists AgentId and AgentName as common properties, but the schema
            example does not place them. Each is kept wherever the record puts it.
    #>
    [CmdletBinding()]
    param([AllowNull()]$Object, [Parameter(Mandatory)][string]$Name)

    if ($null -eq $Object) { return '' }
    $value = Get-GraphAdditionalProperty -Object $Object -Name $Name
    if ($null -eq $value) { return '' }
    return [string]$value
}
