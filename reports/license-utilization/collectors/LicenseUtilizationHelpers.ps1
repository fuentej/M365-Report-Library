#Requires -Version 7.0

# Dot-sourced by the license-utilization collectors.

function Get-LicenseUtilizationSourceAvailability {
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
            Formats a signInActivity value, or an empty string when there is none.

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

    # Match a missing-licence refusal only. A permission name such as
    # LicenseAssignment.Read.All also contains "license", and treating that as a
    # missing licence would skip the source and exit successfully.
    # https://learn.microsoft.com/graph/api/resources/user
    return ($Message -match '(?i)NonPremium|RequiresPremium|AadPremium|premium\s+licen[cs]e|Entra ID P[12]\s+licen[cs]e')
}

function Get-LicenseUtilizationScope {
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

function Invoke-LicenseUtilizationSnapshot {
    <#
        .SYNOPSIS
            Runs one state collector: check availability, sign in, read, map, append.

        .DESCRIPTION
            Every source in this report is state, so the steps live here and each
            collector script only says what it reads and how a row is shaped.

              * A source the schema marks NotAvailable in this cloud writes the header only.
              * An Unverified source is attempted and logs a warning.
              * A missing licence is a logged skip: header only, no error and no exception.
              * HTTP 429 and 503 wait and try again. A read that stays throttled throws.
                It is not written up as a missing permission.
              * Any other failure is logged as an error, leaves a header-only file when
                nothing was collected, and throws so the run does not report success.

        .PARAMETER Source
            The schema key, which names the column list, the availability entry and the
            extra Graph scopes.

        .PARAMETER Fetch
            Reads the source. Called with no arguments; returns the objects, every page.

        .PARAMETER Map
            Shapes one object into rows. Called with the object and the RunDate.

        .PARAMETER Secondary
            A second CSV written from the same response: Source (the schema key of its
            columns), CsvName, KeyColumn and Map.

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
        [hashtable]$Secondary,

        [ValidateSet('Commercial', 'GCC', 'GCCHigh')]
        [string]$Environment = 'Commercial',

        [string]$AppId,
        [string]$CertificateThumbprint,
        [string]$TenantId,
        [string]$Organization,
        [string]$SchemaPath = (Join-Path $PSScriptRoot 'LicenseUtilizationSchema.psd1'),
        [switch]$SkipConnect
    )

    $schema = Import-PowerShellDataFile -LiteralPath $SchemaPath
    $columns = Get-SchemaColumn -Schema $schema -Key $Source
    $csvPath = Join-Path $OutputPath $CsvName
    $log = [System.IO.Path]::GetFileNameWithoutExtension($CsvName)
    $runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

    if (-not (Test-Path -LiteralPath $OutputPath)) {
        New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
    }

    function Write-HeaderOnly {
        Export-AppendCsv -Path $csvPath -Column $columns
        if ($Secondary) {
            Export-AppendCsv -Path (Join-Path $OutputPath $Secondary.CsvName) -Column (Get-SchemaColumn -Schema $schema -Key $Secondary.Source)
        }
    }

    $availability = Get-LicenseUtilizationSourceAvailability -Source $Source -Environment $Environment -Schema $schema
    if ($availability.ShouldSkip) {
        Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $log -Message (
            "Skipping $CsvName. $($availability.Reason) $($availability.Reference)")
        Write-HeaderOnly
        return
    }
    if ($availability.Status -eq 'Unverified') {
        Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $log -Message $availability.Reason
    }

    if (-not $SkipConnect) {
        Connect-M365Service -Service Graph -Environment $Environment -AppId $AppId `
            -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization `
            -Scopes (Get-LicenseUtilizationScope -Schema $schema -Source $Source)
    }

    $items = $null
    try {
        # A 429 or 503 is throttling, not a missing permission and not an empty snapshot.
        # Graph sends Retry-After on 429. Four attempts, then the run stops.
        # https://learn.microsoft.com/graph/throttling
        for ($attempt = 1; ; $attempt++) {
            try {
                $items = @(& $Fetch)
                break
            }
            catch {
                $status = Get-LicenseGraphHttpStatus -ErrorRecord $_
                if (($status -ne 429 -and $status -ne 503) -or $attempt -ge 4) { throw }
                $delay = Get-LicenseRetryDelaySeconds -ErrorRecord $_ -Attempt $attempt
                Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $log -Message (
                    'HTTP {0}. Waiting {1} seconds before attempt {2}.' -f $status, $delay, ($attempt + 1))
                Start-Sleep -Seconds $delay
            }
        }
    }
    catch {
        $message = $_.Exception.Message
        $status = Get-LicenseGraphHttpStatus -ErrorRecord $_
        if ($status -eq 429 -or $status -eq 503) {
            Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $log -Message (
                "$Description is still throttled after retries ($message). A 429 or 503 is not an empty report. It needs a later run, not a different role. Writing the header only.")
            Write-HeaderOnly
            throw "$Description is still throttled after retries ($message). A 429 or 503 is not an empty report."
        }
        if (Test-LicenseError -Message $message) {
            Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $log -Message (
                "Skipping ${CsvName}: ${Description} is not licensed in this tenant ($message). It needs $License. Writing the header only.")
        }
        else {
            Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $log -Message (
                "$Description is unavailable to this sign-in ($message). It needs $License. Writing the header only.")
            Write-HeaderOnly
            throw "$Description is unavailable ($message)."
        }
        Write-HeaderOnly
        return
    }

    $rows = foreach ($item in $items) {
        if ($null -ne $item) { & $Map $item $runDate }
    }

    $result = Export-AppendCsv -Path $csvPath -Rows @($rows) -Column $columns -KeyColumn $KeyColumn -PassThru
    Write-CollectorLog -OutputPath $OutputPath -Source $log -Message (
        '{0}: {1} rows written, {2} skipped.' -f $CsvName, $result.Written, $result.Skipped)

    if ($Secondary) {
        $secondaryRows = foreach ($item in $items) {
            if ($null -ne $item) { & $Secondary.Map $item $runDate }
        }
        $secondaryResult = Export-AppendCsv -Path (Join-Path $OutputPath $Secondary.CsvName) -Rows @($secondaryRows) `
            -Column (Get-SchemaColumn -Schema $schema -Key $Secondary.Source) -KeyColumn $Secondary.KeyColumn -PassThru
        Write-CollectorLog -OutputPath $OutputPath -Source $log -Message (
            '{0}: {1} rows written, {2} skipped.' -f $Secondary.CsvName, $secondaryResult.Written, $secondaryResult.Skipped)
    }
}

#region Usage reports

function Get-NormalizedHeader {
    <#
        .SYNOPSIS
            A header with case, spaces and punctuation removed, for matching.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory, Position = 0)][AllowEmptyString()][string]$Name)

    return ($Name -replace '[^A-Za-z0-9]', '').ToLowerInvariant()
}

function ConvertTo-UsageColumnName {
    <#
        .SYNOPSIS
            The CSV column for a usage-report header: "Outlook (Windows)" is OutlookWindows.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory, Position = 0)][string]$Header)

    $words = $Header -split '[^A-Za-z0-9]+' | Where-Object { $_ }
    return (($words | ForEach-Object { $_.Substring(0, 1).ToUpperInvariant() + $_.Substring(1) }) -join '')
}

function Get-UsageColumn {
    <#
        .SYNOPSIS
            The column order of a usage-report CSV: RunDate, then one column per header.
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param([Parameter(Mandatory)][string[]]$Header)

    $columns = @('RunDate') + @($Header | ForEach-Object { ConvertTo-UsageColumnName $_ })
    return , [string[]]$columns
}

function Get-SchemaColumn {
    <#
        .SYNOPSIS
            The column order of a CSV: a schema list, or a usage report's derived columns.
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param([Parameter(Mandatory)][hashtable]$Schema, [Parameter(Mandatory)][string]$Key)

    if ($Schema.UsageReports.ContainsKey($Key)) {
        return , (Get-UsageColumn -Header $Schema.UsageReports[$Key])
    }
    return , [string[]]$Schema[$Key]
}

function Get-LicenseGraphHttpStatus {
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

function Get-LicenseRetryDelaySeconds {
    <#
        .SYNOPSIS
            Seconds to wait after HTTP 429 or 503.

        .DESCRIPTION
            Graph returns Retry-After on 429
            (https://learn.microsoft.com/graph/throttling).
            When the header is absent, the wait doubles each attempt and is capped at 60 seconds
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

function Get-GraphReportCsv {
    <#
        .SYNOPSIS
            Downloads one usage report as CSV and returns its rows.

        .DESCRIPTION
            A GET. Graph answers the report functions with a 302 to a preauthenticated
            URL that is valid for a few minutes, so the file is read at once. The
            copilot function answers 200 with the CSV in the body. Either way the
            response is saved to a temporary file, read and removed.
            https://learn.microsoft.com/graph/api/reportroot-getoffice365activeuserdetail
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Uri)

    $download = Join-Path ([System.IO.Path]::GetTempPath()) ('license-usage-' + [guid]::NewGuid().ToString('N') + '.csv')
    try {
        Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputFilePath $download -ErrorAction Stop
        if (-not (Test-Path -LiteralPath $download)) {
            throw "The usage report did not download a file from $Uri."
        }
        $headerLine = Get-Content -LiteralPath $download -TotalCount 1
        $headerText = if ($null -eq $headerLine) { '' } else { [string]$headerLine }
        # Every user-detail CSV names this column. A 302 body, an empty file, or an
        # error page does not, and must not be stored as an empty report.
        # https://learn.microsoft.com/graph/api/reportroot-getoffice365activeuserdetail
        if ((Get-NormalizedHeader $headerText) -notlike '*userprincipalname*') {
            throw "The usage report download from $Uri is not the CSV. It has no User Principal Name column."
        }
        @(Import-Csv -LiteralPath $download)
    }
    finally {
        Remove-Item -LiteralPath $download -Force -ErrorAction SilentlyContinue
    }
}

function Get-GraphReportJsonPage {
    <#
        .SYNOPSIS
            One page of a usage report requested as JSON. A GET.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Uri)

    Invoke-MgGraphRequest -Method GET -Uri $Uri -ErrorAction Stop
}

function Get-GraphReportJson {
    <#
        .SYNOPSIS
            Every row of a usage report requested as JSON, following @odata.nextLink
            until it is absent.

        .DESCRIPTION
            One JSON page is not the full set (the default page is 200). A next link on
            another host is refused rather than followed. Each item's details are
            flattened beside it, one object per detail, so the rows match the CSV form.
            https://learn.microsoft.com/graph/api/reportroot-getm365appuserdetail
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Uri)

    $hostName = ([uri]$Uri).Host
    $next = $Uri
    $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
    while (-not [string]::IsNullOrEmpty($next)) {
        if (-not $seen.Add($next)) {
            throw "The usage report next link was already followed ($next). Refusing to request it again."
        }
        if (([uri]$next).Host -ne $hostName) {
            throw "The next link points at $(([uri]$next).Host), not $hostName. Refusing to follow it."
        }

        $page = Get-GraphReportJsonPage -Uri $next
        foreach ($item in @(Get-GraphAdditionalProperty -Object $page -Name 'value')) {
            if ($null -eq $item) { continue }
            $details = @(Get-GraphAdditionalProperty -Object $item -Name 'details')
            if ($details.Count -eq 0) { $details = @($null) }
            foreach ($detail in $details) {
                $flat = [ordered]@{}
                foreach ($source in @($item, $detail)) {
                    if ($null -eq $source) { continue }
                    $pairs = if ($source -is [System.Collections.IDictionary]) { $source.GetEnumerator() } else { $source.PSObject.Properties | ForEach-Object { [pscustomobject]@{ Key = $_.Name; Value = $_.Value } } }
                    foreach ($pair in $pairs) {
                        if ($pair.Key -ne 'details') { $flat[[string]$pair.Key] = $pair.Value }
                    }
                }
                [pscustomobject]$flat
            }
        }

        $next = [string](Get-GraphAdditionalProperty -Object $page -Name '@odata.nextLink')
    }
}

function Test-NamesConcealed {
    <#
        .SYNOPSIS
            True when the last report-settings.csv snapshot says usage reports conceal names.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param([Parameter(Mandatory)][string]$OutputPath)

    $latest = @(Get-CsvLatestSnapshot -Path (Join-Path $OutputPath 'report-settings.csv'))
    return [bool]($latest | Where-Object { $_ -and [string]$_.DisplayConcealedNames -eq 'True' })
}

function Invoke-LicenseUsageReport {
    <#
        .SYNOPSIS
            Runs one usage-report collector (sources 5a to 5g).

        .PARAMETER ReportPath
            The function and its arguments below /v1.0, such as
            reports/getEmailActivityUserDetail(period='D30').

        .PARAMETER Json
            Ask for JSON and follow its paging, instead of the CSV download.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][string]$CsvName,
        [Parameter(Mandatory)][string]$ReportPath,
        [Parameter(Mandatory)][string]$Description,
        [Parameter(Mandatory)][string]$OutputPath,
        # The period argument, such as D30. Written into ReportPeriod when the CSV
        # omits that column, so two periods collected on the same run date stay distinct.
        [string]$ReportPeriod,
        # What the source needs, written into run.log when the call is refused.
        # Copilot's /copilot page lists different roles than the other usage reports.
        [string]$License = 'Reports.Read.All and, when signed in, a role such as Reports Reader (Global Reader and Usage Summary Reports Reader see tenant-level data only)',
        [switch]$Json,

        [ValidateSet('Commercial', 'GCC', 'GCCHigh')]
        [string]$Environment = 'Commercial',

        [string]$AppId,
        [string]$CertificateThumbprint,
        [string]$TenantId,
        [string]$Organization,
        [string]$SchemaPath = (Join-Path $PSScriptRoot 'LicenseUtilizationSchema.psd1'),
        [switch]$SkipConnect
    )

    $schema = Import-PowerShellDataFile -LiteralPath $SchemaPath
    $headers = [string[]]$schema.UsageReports[$Source]
    $base = (Get-M365ServiceEndpoint -Service Graph -Environment $Environment).ResourceEndpoint
    $uri = '{0}/v1.0/{1}' -f $base, $ReportPath
    if ($Json) { $uri += '?$format=application/json' }

    $fetch = if ($Json) { { Get-GraphReportJson -Uri $uri } } else { { Get-GraphReportCsv -Uri $uri } }

    $map = {
        param($row, $runDate)

        $lookup = @{}
        foreach ($property in $row.PSObject.Properties) { $lookup[(Get-NormalizedHeader $property.Name)] = $property.Value }

        $out = [ordered]@{ RunDate = $runDate }
        foreach ($header in $headers) {
            $value = $lookup[(Get-NormalizedHeader $header)]
            $out[(ConvertTo-UsageColumnName $header)] = if ($null -eq $value) { '' } else { [string]$value }
        }
        # getOffice365ActiveUserDetail's documented CSV stops at Assigned Products and
        # does not include Report Period. The other user-detail CSVs do, as a day count
        # (30, not D30). When the column is missing, store that day count from -Period
        # so a second period on the same RunDate is not skipped as a duplicate.
        # https://learn.microsoft.com/graph/api/reportroot-getoffice365activeuserdetail
        if ([string]::IsNullOrWhiteSpace([string]$out['ReportPeriod']) -and -not [string]::IsNullOrWhiteSpace($ReportPeriod)) {
            $out['ReportPeriod'] = ($ReportPeriod -replace '^(?i)D(?=\d)', '')
        }
        # A row with no user is not a user; the report appends nothing for it.
        if ([string]::IsNullOrWhiteSpace($out['UserPrincipalName'])) { return }
        [pscustomobject]$out
    }

    # Names concealed means the user column holds a hash that cannot be joined to
    # user-licenses.csv. The rows are still written; the report must not read them as
    # zero usage.
    if ((Test-NamesConcealed -OutputPath $OutputPath)) {
        Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source ([System.IO.Path]::GetFileNameWithoutExtension($CsvName)) -Message (
            'report-settings.csv says usage reports conceal user names. The rows cannot be joined to users on UserPrincipalName.')
    }

    Invoke-LicenseUtilizationSnapshot -Source $Source -CsvName $CsvName -Description $Description `
        -License $License `
        -KeyColumn @('RunDate', 'UserPrincipalName', 'ReportPeriod') -OutputPath $OutputPath -Environment $Environment `
        -AppId $AppId -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization `
        -SchemaPath $SchemaPath -SkipConnect:$SkipConnect -Fetch $fetch -Map $map
}

#endregion
