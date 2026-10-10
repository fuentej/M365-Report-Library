#Requires -Version 7.0

# Dot-sourced by the Exchange activity collectors. Report-specific: the shared connection
# and CSV functions come from shared/M365ReportLibrary.psm1.

function Get-ExchangeSourceAvailability {
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

function Test-ExchangeSourceSkipped {
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

    $availability = Get-ExchangeSourceAvailability -Source $Source -Environment $Environment -Schema $Schema
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
            https://learn.microsoft.com/graph/api/reportroot-getmailboxusagedetail
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
    $download = Join-Path ([System.IO.Path]::GetTempPath()) ('exchange-activity-' + [guid]::NewGuid().ToString('N') + '.csv')
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

function Get-MailboxQuotaStatus {
    <#
        .SYNOPSIS
            The admin center's mailbox quota category from the usage report's byte columns.

        .DESCRIPTION
            Good is below the issue-warning quota; Warning is at or above it and below prohibit
            send; CantSend is at or above prohibit send and below prohibit send/receive;
            CantSendReceive is at or above prohibit send/receive. The boundary is "at or above".
            https://learn.microsoft.com/microsoft-365/admin/activity-reports/mailbox-usage
            Empty when the storage or a quota is missing or not a number.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [AllowNull()][string]$StorageUsed,
        [AllowNull()][string]$IssueWarning,
        [AllowNull()][string]$ProhibitSend,
        [AllowNull()][string]$ProhibitSendReceive
    )

    $numbers = foreach ($value in $StorageUsed, $IssueWarning, $ProhibitSend, $ProhibitSendReceive) {
        [decimal]$parsed = 0
        if (-not [decimal]::TryParse([string]$value, [System.Globalization.NumberStyles]::Integer, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$parsed)) {
            return ''
        }
        $parsed
    }
    $used, $warning, $send, $sendReceive = $numbers

    if ($used -ge $sendReceive) { return 'CantSendReceive' }
    if ($used -ge $send) { return 'CantSend' }
    if ($used -ge $warning) { return 'Warning' }
    return 'Good'
}

function ConvertTo-ByteCount {
    <#
        .SYNOPSIS
            The byte count in an Exchange size such as "1.5 GB (1,610,612,736 bytes)", or an
            empty string when the text has none.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Position = 0)][AllowNull()]$Value)

    if ($null -eq $Value) { return '' }
    $text = [string]$Value
    $match = [regex]::Match($text, '\(([\d,]+) bytes\)')
    if ($match.Success) { return $match.Groups[1].Value.Replace(',', '') }
    if ($text -match '^\d+$') { return $text }
    return ''
}

function Get-ObjectText {
    <#
        .SYNOPSIS
            A property of a cmdlet result as text, or an empty string when the object has no
            such property. The Exchange pages document few output properties, so a missing one
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

function Get-MessageTraceWindow {
    <#
        .SYNOPSIS
            The query windows for a message trace: at most 10 days each, none starting more
            than 90 days before now.

        .DESCRIPTION
            Get-MessageTraceV2 searches the last 90 days and returns 10 days per query
            (https://learn.microsoft.com/powershell/module/exchangepowershell/get-messagetracev2).
            The Graph message trace has the same limits
            (https://learn.microsoft.com/exchange/monitoring/trace-an-email-message/graph-api-message-trace).
            The newest window comes first so an interrupted run keeps the freshest data.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][datetime]$Start,
        [Parameter(Mandatory)][datetime]$End
    )

    $earliest = $End.AddDays(-90)
    if ($Start -lt $earliest) { $Start = $earliest }
    if ($End -le $Start) { return }

    $windows = [System.Collections.Generic.List[object]]::new()
    $cursor = $Start
    while ($cursor -lt $End) {
        $stop = $cursor.AddDays(10)
        if ($stop -gt $End) { $stop = $End }
        $windows.Add([pscustomobject]@{ Start = $cursor; End = $stop })
        $cursor = $stop
    }
    $windows.Reverse()
    $windows.ToArray()
}

function Invoke-MessageTraceV2Window {
    <#
        .SYNOPSIS
            Every Get-MessageTraceV2 row for one address and window, continuing past a full round.

        .DESCRIPTION
            The cmdlet has no page parameter. A round that returns ResultSize rows is not the
            rest of the window: the next round sets -EndDate to the last row's Received time and
            -StartingRecipientAddress to that row's RecipientAddress. Results run Received
            descending, then RecipientAddress ascending. Rows seen already are not returned
            twice, and a round that adds nothing ends the loop.
            https://learn.microsoft.com/powershell/module/exchangepowershell/get-messagetracev2
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateSet('Sender', 'Recipient')][string]$Role,
        [Parameter(Mandatory)][string]$Address,
        [Parameter(Mandatory)][datetime]$Start,
        [Parameter(Mandatory)][datetime]$End,
        [int]$ResultSize = 5000
    )

    $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $queryEnd = $End
    $startingRecipient = $null

    while ($true) {
        $query = @{
            StartDate   = $Start
            EndDate     = $queryEnd
            ResultSize  = $ResultSize
            ErrorAction = 'Stop'
        }
        if ($Role -eq 'Sender') { $query['SenderAddress'] = $Address } else { $query['RecipientAddress'] = $Address }
        if ($startingRecipient) { $query['StartingRecipientAddress'] = $startingRecipient }

        Wait-MessageTraceRateLimit
        $batch = @(Get-MessageTraceV2 @query | Where-Object { $null -ne $_ })

        $added = 0
        foreach ($row in $batch) {
            if ($seen.Add(('{0}|{1}|{2}' -f $row.MessageTraceId, $row.RecipientAddress, $row.Received))) {
                $added++
                $row
            }
        }

        if ($batch.Count -lt $ResultSize -or $added -eq 0) { break }

        $last = $batch[-1]
        $queryEnd = ([datetime]$last.Received).ToUniversalTime()
        $startingRecipient = [string]$last.RecipientAddress
    }
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

function Get-GraphPagedValue {
    <#
        .SYNOPSIS
            Every item of a Graph collection, following @odata.nextLink until it is absent.

        .DESCRIPTION
            https://learn.microsoft.com/graph/paging. The returned URL is requested as it is;
            a $skiptoken is never built by hand. Only GET is sent.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [switch]$RateLimited
    )

    $next = $Uri
    $visited = [System.Collections.Generic.HashSet[string]]::new()
    while ($next) {
        if (-not (Test-GraphReadUri -Uri $next)) {
            throw "Refusing to request a page outside Microsoft Graph: $next"
        }
        if (-not $visited.Add($next)) {
            throw "Graph returned a nextLink that was already requested: $next"
        }
        if ($RateLimited) { Wait-MessageTraceRateLimit }
        $page = Invoke-MgGraphRequest -Method GET -Uri $next -ErrorAction Stop
        foreach ($item in @(Get-GraphJsonValue -Object $page -Name 'value')) {
            if ($null -ne $item) { $item }
        }
        $next = [string](Get-GraphJsonValue -Object $page -Name '@odata.nextLink')
    }
}

function Wait-MessageTraceRateLimit {
    <#
        .SYNOPSIS
            Keeps a run under 100 message trace requests in any 5-minute window.

        .DESCRIPTION
            Both Get-MessageTraceV2 and the Graph message trace accept 100 requests per
            5-minute window. The cmdlet page names no retry, so the run waits for the window to
            clear instead of waiting to be refused. Stays at 95 to leave room for another caller.
            https://learn.microsoft.com/powershell/module/exchangepowershell/get-messagetracev2
            https://learn.microsoft.com/exchange/monitoring/trace-an-email-message/graph-api-message-trace
    #>
    [CmdletBinding()]
    param([int]$Limit = 95, [int]$WindowSeconds = 300)

    if (-not (Test-Path variable:script:TraceRequestTimes)) {
        $script:TraceRequestTimes = [System.Collections.Generic.Queue[datetime]]::new()
    }

    $now = [datetime]::UtcNow
    while ($script:TraceRequestTimes.Count -gt 0 -and ($now - $script:TraceRequestTimes.Peek()).TotalSeconds -ge $WindowSeconds) {
        [void]$script:TraceRequestTimes.Dequeue()
    }
    if ($script:TraceRequestTimes.Count -ge $Limit) {
        $wait = [math]::Ceiling($WindowSeconds - ($now - $script:TraceRequestTimes.Peek()).TotalSeconds)
        if ($wait -gt 0) { Start-Sleep -Seconds $wait }
        $script:TraceRequestTimes.Clear()
    }
    $script:TraceRequestTimes.Enqueue([datetime]::UtcNow)
}
