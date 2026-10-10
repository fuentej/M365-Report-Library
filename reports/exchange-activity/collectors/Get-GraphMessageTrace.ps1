#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes graph-message-trace.csv: messages sent and received per recipient, from the Graph message trace, appended from where the last run stopped.

    .DESCRIPTION
        Source 11: GET /beta/admin/exchange/tracing/messageTraces
        (https://learn.microsoft.com/exchange/monitoring/trace-an-email-message/graph-api-message-trace).
        Event source. A run starts at the latest ReceivedDateTime already in the file, or -LookbackDays
        back on the first run, and never earlier than 90 days. The filter is
        receivedDateTime ge <start> and receivedDateTime le <end> in ISO 8601 UTC, at most 10 days per
        query. The default page is 1000; $top=5000 asks for the most. Paging follows the returned
        @odata.nextLink until it is absent; no $skiptoken is built. One object is one recipient, not one
        distinct message, and there is no read count. The list is on /beta.

        Needs the application permission ExchangeMessageTrace.Read.All (no delegated permission is
        listed) and a service principal for app 8bd644d1-64a1-4d4b-ae52-2e0cbf64e373 in the tenant.
        Until that provisioning finishes the API returns 401, and a 401 is not an empty trace: this
        script logs it and keeps the file as it was. Over 100 requests in 5 minutes Graph refuses; the run
        waits to stay under it. Cloud availability is UNVERIFIED in every cloud (the page has no
        national-cloud table).

    .EXAMPLE
        ./Get-GraphMessageTrace.ps1 -OutputPath ./out
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

    # How far back the first run looks, in days (1 to 90). Later runs resume from the file.
    [ValidateRange(1, 90)]
    [int]$LookbackDays = 10,

    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Imported by path without -Force: a forced re-import builds a fresh module session
# state, which would discard anything the caller has arranged around the module - and
# that is exactly how the tests put mocks in front of the tenant cmdlets.
Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

. (Join-Path $PSScriptRoot 'ExchangeActivityHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'ExchangeActivitySchema.psd1')
$columns = $schema.GraphMessageTrace
$source = 'graph-message-trace'
$csvPath = Join-Path $OutputPath 'graph-message-trace.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

if (Test-ExchangeSourceSkipped -Source 'GraphMessageTrace' -LogSource $source -CsvPath $csvPath -Column $columns `
        -OutputPath $OutputPath -Environment $Environment -Schema $schema) {
    return
}

if (-not $SkipConnect) {
    Connect-M365Service -Service Graph -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization
}

$end = [datetime]::UtcNow
$watermark = Get-CsvWatermark -Path $csvPath -Column 'ReceivedDateTime'
$start = if ($watermark) { $watermark } else { $end.AddDays(-$LookbackDays) }

Export-AppendCsv -Path $csvPath -Column $columns
$total = [pscustomobject]@{ Written = 0; Skipped = 0 }

try {
    foreach ($window in Get-MessageTraceWindow -Start $start -End $end) {
        $filter = 'receivedDateTime ge {0:yyyy-MM-ddTHH:mm:ssZ} and receivedDateTime le {1:yyyy-MM-ddTHH:mm:ssZ}' -f $window.Start.ToUniversalTime(), $window.End.ToUniversalTime()
        $uri = '/beta/admin/exchange/tracing/messageTraces?$filter=' + [uri]::EscapeDataString($filter) + '&$top=5000'

        $rows = foreach ($item in Get-GraphPagedValue -Uri $uri -RateLimited) {
            [pscustomobject]@{
                ReceivedDateTime = ConvertTo-CsvTimestamp (Get-GraphJsonValue -Object $item -Name 'receivedDateTime')
                Id               = [string](Get-GraphJsonValue -Object $item -Name 'id')
                SenderAddress    = [string](Get-GraphJsonValue -Object $item -Name 'senderAddress')
                RecipientAddress = [string](Get-GraphJsonValue -Object $item -Name 'recipientAddress')
                Status           = [string](Get-GraphJsonValue -Object $item -Name 'status')
                Size             = [string](Get-GraphJsonValue -Object $item -Name 'size')
            }
        }
        $result = Export-AppendCsv -Path $csvPath -Rows @($rows) -Column $columns -KeyColumn @('Id', 'RecipientAddress') -PassThru
        $total.Written += $result.Written
        $total.Skipped += $result.Skipped
    }
}
catch {
    $message = $_.Exception.Message
    if ($message -match '401|Unauthorized|[Ss]ervice principal') {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            'The message trace API refused the sign-in ({0}). If the service principal for app 8bd644d1-64a1-4d4b-ae52-2e0cbf64e373 was just created, provisioning can take hours; a 401 is not an empty trace. The file is unchanged.' -f $message)
    }
    else {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            'The Graph message trace failed ({0}). It needs the application permission ExchangeMessageTrace.Read.All. Rows already written are kept; the next run resumes from the latest ReceivedDateTime.' -f $message)
    }
    return
}

Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'graph-message-trace.csv: {0} rows written, {1} skipped. Window {2:yyyy-MM-ddTHH:mm:ssZ} to {3:yyyy-MM-ddTHH:mm:ssZ}.' -f $total.Written, $total.Skipped, $start, $end)
