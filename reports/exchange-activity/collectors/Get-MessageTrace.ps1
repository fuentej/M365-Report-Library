#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes message-trace.csv: messages sent and received per mailbox, from Exchange message trace, appended from where the last run stopped.

    .DESCRIPTION
        Source 8: Get-MessageTraceV2 -SenderAddress <addr> and -RecipientAddress <addr>
        (https://learn.microsoft.com/powershell/module/exchangepowershell/get-messagetracev2),
        a send and receive count the library makes itself. It is the fallback where the Graph usage
        reports are not available (GCC High). There is no read count.

        Event source. A run starts at the latest Received already in message-trace.csv, or
        -LookbackDays back on the first run, and never earlier than 90 days. Each query covers at
        most 10 days. The cmdlet has no page parameter: a round that returns 5000 rows is not the
        rest of the window, so the next round sets -EndDate to the last row's Received time and
        -StartingRecipientAddress to that row's RecipientAddress
        (https://learn.microsoft.com/exchange/monitoring/trace-an-email-message/message-trace-modern-eac).
        Each row is one recipient, not one distinct message. A message that already returned 1000
        recipient rows is queried again with -MessageTraceId, which the cmdlet requires for a
        message sent to more than 1000 recipients. The run stays under 100 requests in 5 minutes
        by waiting, and an unfinished run leaves message-trace.pending so the next run repeats
        the original start instead of skipping mailboxes behind a newer watermark.

        Start and end are passed as dates with a time, in UTC, not as date-only values (a date-only
        value uses the session's regional short date).

        The cmdlet pages defer to the permissions page for the role. App-only sign-in needs
        Exchange.ManageAsApp plus an Exchange role. Organization Management and Recipient
        Management are listed for managing recipients
        (https://learn.microsoft.com/exchange/permissions-exo/feature-permissions); the least
        privileged role that can run this cmdlet is UNVERIFIED until Get-ManagementRole is run
        in the tenant. Availability in GCC and GCC High is UNVERIFIED at cmdlet level: the
        service description says only that remote PowerShell is available
        (https://learn.microsoft.com/office365/servicedescriptions/office-365-platform-service-description/office-365-us-government/exchange-online-for-us-government-environments).

    .EXAMPLE
        ./Get-MessageTrace.ps1 -OutputPath ./out
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

    # Read at most this many mailboxes. For a trial run; 0 reads them all.
    [int]$MailboxLimit = 0,

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
$columns = $schema.MessageTrace
$source = 'message-trace'
$csvPath = Join-Path $OutputPath 'message-trace.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

if (Test-ExchangeSourceSkipped -Source 'MessageTrace' -LogSource $source -CsvPath $csvPath -Column $columns `
        -OutputPath $OutputPath -Environment $Environment -Schema $schema) {
    return
}

$connectedHere = -not $SkipConnect
if ($connectedHere) {
    Connect-M365Service -Service ExchangeOnline -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization
}

try {
    try {
        $mailboxes = @(Get-EXOMailbox -ResultSize Unlimited -PropertySets Minimum -ErrorAction Stop)
    }
    catch {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            'Get-EXOMailbox is unavailable to this sign-in ({0}). Writing the header only.' -f $_.Exception.Message)
        Export-AppendCsv -Path $csvPath -Column $columns
        return
    }
    if ($MailboxLimit -gt 0) { $mailboxes = @($mailboxes | Select-Object -First $MailboxLimit) }

    $end = [datetime]::UtcNow
    # An unfinished run keeps this file. The CSV watermark is the newest row already written,
    # which can belong to a mailbox that finished while a later mailbox did not. Resuming from
    # that timestamp would drop the unfinished mailbox's older messages.
    $pendingPath = Join-Path $OutputPath 'message-trace.pending'
    if (Test-Path -LiteralPath $pendingPath) {
        $start = ConvertTo-ExchangeUtc ((Get-Content -LiteralPath $pendingPath -Raw).Trim())
    }
    else {
        $watermark = Get-CsvWatermark -Path $csvPath -Column 'Received'
        $start = if ($watermark) { ConvertTo-ExchangeUtc $watermark } else { $end.AddDays(-$LookbackDays) }
    }
    Set-Content -LiteralPath $pendingPath -Value ($start.ToString('yyyy-MM-ddTHH:mm:ssZ')) -Encoding ascii -NoNewline

    $total = [pscustomobject]@{ Written = 0; Skipped = 0 }
    Export-AppendCsv -Path $csvPath -Column $columns

    try {
        foreach ($mailbox in $mailboxes) {
            $address = [string]$mailbox.PrimarySmtpAddress
            if ([string]::IsNullOrWhiteSpace($address)) { continue }

            $rows = [System.Collections.Generic.List[object]]::new()
            foreach ($window in Get-MessageTraceWindow -Start $start -End $end) {
                foreach ($role in 'Sender', 'Recipient') {
                    foreach ($trace in Invoke-MessageTraceV2Window -Role $role -Address $address -Start $window.Start -End $window.End) {
                        $rows.Add([pscustomobject]@{
                                Received         = ConvertTo-CsvTimestamp $trace.Received
                                MessageTraceId   = [string]$trace.MessageTraceId
                                SenderAddress    = [string]$trace.SenderAddress
                                RecipientAddress = [string]$trace.RecipientAddress
                                Status           = [string]$trace.Status
                            })
                    }
                }
            }
            $result = Export-AppendCsv -Path $csvPath -Rows $rows.ToArray() -Column $columns -KeyColumn @('MessageTraceId', 'RecipientAddress') -PassThru
            $total.Written += $result.Written
            $total.Skipped += $result.Skipped
        }
        Remove-Item -LiteralPath $pendingPath -Force -ErrorAction SilentlyContinue
        Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
            'message-trace.csv: {0} rows written, {1} skipped. Window {2:yyyy-MM-ddTHH:mm:ssZ} to {3:yyyy-MM-ddTHH:mm:ssZ}.' -f $total.Written, $total.Skipped, $start, $end)
    }
    catch {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            'Get-MessageTraceV2 failed ({0}). Rows already written are kept. The next run resumes from {1:yyyy-MM-ddTHH:mm:ssZ}, not from the newest Received, because a later mailbox may not have been read. Message trace needs a role such as Help Desk or View-Only Recipients (UNVERIFIED which is least).' -f $_.Exception.Message, $start)
    }
}
finally {
    if ($connectedHere) {
        Disconnect-ExchangeOnline -Confirm:$false
    }
}
