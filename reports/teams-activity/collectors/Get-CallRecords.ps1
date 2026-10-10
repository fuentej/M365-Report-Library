#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes call-records.csv: Teams calls and meetings with the platform each session endpoint used, resuming from the last start time collected.

    .DESCRIPTION
        Source 6: Microsoft Graph GET /communications/callRecords
        (https://learn.microsoft.com/graph/api/callrecords-cloudcommunications-list-callrecords), then each
        record with its sessions expanded
        (https://learn.microsoft.com/graph/api/callrecords-callrecord-get). The list does not include
        sessions. Both the list and the session collection are followed through @odata.nextLink until it
        is absent; the list page size is 60 by default, so one page is not the full set.

        Event source. The query is startDateTime ge the latest StartDateTime already in the file and
        startDateTime lt the end of the window, so a re-run resumes where the last one stopped. The first
        run reaches back -LookbackDays; Graph keeps records for 30 days, and an older or not-yet-available
        record answers 404, which is logged and skipped. The first version of a record can take 150
        minutes to appear after the call ends
        (https://learn.microsoft.com/graph/callrecords-api-faq), so the window ends -DelayMinutes before
        now (default 180) to avoid advancing the resume point past calls that are not readable yet.
        Later versions of a record carry the same id with a higher Version; Version is part of the row
        key, so a re-read version appends its own rows and a reader keeps the highest Version per
        CallRecordId.

        Available in Commercial, GCC and GCC High. Needs CallRecords.Read.All as an APPLICATION
        permission; delegated is not supported, so use -AppId and -CertificateThumbprint. This collector
        sees calls and meetings only, not chat or channel messages, and not live event streamers.

    .EXAMPLE
        ./Get-CallRecords.ps1 -OutputPath ./out -AppId $appId -CertificateThumbprint $thumbprint -TenantId $tenantId
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$OutputPath,

    [ValidateSet('Commercial', 'GCC', 'GCCHigh')]
    [string]$Environment = 'Commercial',

    [string]$AppId,
    [string]$CertificateThumbprint,
    [string]$TenantId,
    [string]$Organization,

    [ValidateRange(1, 30)]
    [int]$LookbackDays = 30,

    [ValidateRange(0, 1440)]
    [int]$DelayMinutes = 180,

    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Imported by path without -Force: a forced re-import builds a fresh module session
# state, which would discard anything the caller has arranged around the module - and
# that is exactly how the tests put mocks in front of the tenant cmdlets.
Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

. (Join-Path $PSScriptRoot 'TeamsActivityHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'TeamsActivitySchema.psd1')
$columns = $schema.CallRecords
$source = 'call-records'
$csvPath = Join-Path $OutputPath 'call-records.csv'

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

if (Test-TeamsSourceSkipped -Source 'CallRecords' -LogSource $source -CsvPath $csvPath -Column $columns `
        -OutputPath $OutputPath -Environment $Environment -Schema $schema) {
    return
}

if (-not $SkipConnect) {
    Connect-M365Service -Service Graph -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization `
        -Scopes (@(Get-DefaultGraphScope) + 'CallRecords.Read.All')
}

$now = [datetime]::UtcNow
$end = $now.AddMinutes(-$DelayMinutes)
$watermark = Get-CsvWatermark -Path $csvPath -Column 'StartDateTime'
$start = if ($null -ne $watermark) { $watermark } else { $now.AddDays(-$LookbackDays) }

if ($end -le $start) {
    Write-CollectorLog -OutputPath $OutputPath -Source $source -Message 'Nothing new to collect: the window is empty.'
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}

$format = 'yyyy-MM-ddTHH:mm:ssZ'
$filter = 'startDateTime ge {0} and startDateTime lt {1}' -f $start.ToString($format), $end.ToString($format)
Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'Reading call records from {0} to {1}.' -f $start.ToString($format), $end.ToString($format))

try {
    $listed = @(Get-GraphPagedValue -Uri ('/v1.0/communications/callRecords?$filter=' + [uri]::EscapeDataString($filter)) -OutputPath $OutputPath -LogSource $source)
}
catch {
    Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
        'Call records are unavailable to this sign-in ({0}). They need the application permission CallRecords.Read.All. Writing the header only.' -f $_.Exception.Message)
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}

$rows = [System.Collections.Generic.List[object]]::new()
foreach ($item in $listed) {
    $id = [string](Get-GraphJsonValue -Object $item -Name 'id')
    if (-not $id) { continue }
    try {
        $recordId = $id
        $record = Invoke-ReadWithThrottleRetry -OutputPath $OutputPath -Source $source -Action {
            Invoke-MgGraphRequest -Method GET -Uri ('/v1.0/communications/callRecords/{0}?$expand=sessions' -f $recordId) -ErrorAction Stop
        }
        $sessions = @(Get-GraphJsonValue -Object $record -Name 'sessions')
        $more = [string](Get-GraphJsonValue -Object $record -Name 'sessions@odata.nextLink')
        if ($more) {
            $sessions += @(Get-GraphPagedValue -Uri $more -OutputPath $OutputPath -LogSource $source)
        }
    }
    catch {
        # A 404 is a record that is not readable yet or is older than 30 days, not an absent call.
        # https://learn.microsoft.com/graph/callrecords-api-faq
        if ((Get-GraphHttpStatus -ErrorRecord $_) -eq 404) {
            Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
                'Call record {0} answered 404 (not available yet, or past the 30-day window). Skipping it.' -f $id)
            continue
        }
        throw
    }
    foreach ($row in (ConvertTo-CallRecordRow -Record $record -Session $sessions)) { $rows.Add($row) }
}

$result = Export-AppendCsv -Path $csvPath -Rows $rows.ToArray() -Column $columns `
    -KeyColumn @('CallRecordId', 'Version', 'SessionId') -PassThru
Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'call-records.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
