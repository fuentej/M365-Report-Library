#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes audit-graph-records.csv: unified audit log records from the Microsoft Graph Audit
        Search API, appended from where the last run stopped.

    .DESCRIPTION
        Source 2: POST /security/auditLog/queries creates a query, GET /security/auditLog/queries/{id}
        is polled until its status is succeeded, and GET /security/auditLog/queries/{id}/records is
        read through @odata.nextLink
        (https://learn.microsoft.com/graph/api/security-auditcoreroot-post-auditlogqueries,
        https://learn.microsoft.com/graph/api/security-auditlogquery-get,
        https://learn.microsoft.com/graph/api/security-auditlogquery-list-records).
        The create call is a POST that makes the query object; it does not change tenant data. Decision
        D-007 allows it. It is the only POST this collector sends.

        Event source. A run starts at the latest CreationTime already in the file, or -LookbackDays back
        on the first run, and sends one query per -SliceMinutes slice, oldest first. A query that
        succeeded but went over its record limit (isRecordCountLimitExceeded) is read again as two
        halves. A query that is still over the limit at -SliceMinutes is not written. A throttled
        call (429) waits for Retry-After, or backs off from 30 seconds when there is none, and
        never retries at once. A tenant gets at least 200 submissions per rolling 24 hours
        (https://learn.microsoft.com/graph/throttling-limits#security-audit-log-query-service-limits),
        so keep -LookbackDays x slices per day well under that.

        Not available in GCC High: the API pages mark US Government L4 unsupported. There the CSV gets
        its header only and the reason goes to run.log. Permissions: AuditLogsQuery.Read.All (all
        workloads) to create and list records. Reading one query object lists ThreatIntelligence.Read.All
        as the least privileged permission, so grant it too. The Purview role is not stated on the API
        pages (docs/candidates/unified-audit-log.md source 2).

    .EXAMPLE
        ./Get-AuditGraphRecords.ps1 -OutputPath ./out
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

    # How far back the first run looks, in days. Later runs resume from the file.
    [ValidateRange(1, 180)]
    [int]$LookbackDays = 7,

    # The range each query covers. The default is one day.
    [ValidateRange(60, 10080)]
    [int]$SliceMinutes = 1440,

    # Seconds between polls of a running query, and how many polls before giving up on it.
    [int]$PollSeconds = 15,
    [int]$MaxPolls = 240,

    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Imported by path without -Force: a forced re-import builds a fresh module session
# state, which would discard anything the caller has arranged around the module - and
# that is exactly how the tests put mocks in front of the tenant cmdlets.
Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

. (Join-Path $PSScriptRoot 'UnifiedAuditLogHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'UnifiedAuditLogSchema.psd1')
$columns = $schema.AuditGraphRecords
$source = 'audit-graph-records'
$csvPath = Join-Path $OutputPath 'audit-graph-records.csv'

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

if (Test-AuditSourceSkipped -Source 'AuditGraphRecords' -LogSource $source -CsvPath $csvPath -Column $columns `
        -OutputPath $OutputPath -Environment $Environment -Schema $schema) {
    return
}

if (-not $SkipConnect) {
    Connect-M365Service -Service Graph -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization `
        -Scopes @('AuditLogsQuery.Read.All', 'ThreatIntelligence.Read.All')
}

$end = [datetime]::UtcNow
$start = Get-AuditRunStart -CsvPath $csvPath -WatermarkColumn 'CreationTime' -End $end -LookbackDays $LookbackDays
Export-AppendCsv -Path $csvPath -Column $columns
$total = [pscustomobject]@{ Written = 0; Skipped = 0 }

try {
    foreach ($slice in Split-DateRange -Start $start -End $end -WindowMinutes $SliceMinutes) {
        $rows = @(Get-AuditGraphSlice -Start $slice.Start -End $slice.End -OutputPath $OutputPath -LogSource $source `
                -PollSeconds $PollSeconds -MaxPolls $MaxPolls)
        $result = Export-AppendCsv -Path $csvPath -Rows $rows -Column $columns -KeyColumn @('RecordId') -PassThru
        $total.Written += $result.Written
        $total.Skipped += $result.Skipped
    }
}
catch {
    Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
        'The Graph audit log query failed ({0}). It needs AuditLogsQuery.Read.All (and ThreatIntelligence.Read.All to read the query). Slices already written are kept; the next run resumes from the latest CreationTime.' -f $_.Exception.Message)
    throw
}

Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'audit-graph-records.csv: {0} rows written, {1} skipped. Window {2:yyyy-MM-ddTHH:mm:ssZ} to {3:yyyy-MM-ddTHH:mm:ssZ}.' -f $total.Written, $total.Skipped, $start, $end)
