#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes mailbox-change-events.csv: audit events for inbox rule creation or change,
        mailbox forwarding changes and mailbox permission changes, appended from where the
        last run stopped.

    .DESCRIPTION
        Source: Search-UnifiedAuditLog in Exchange Online PowerShell
        (https://learn.microsoft.com/powershell/module/exchangepowershell/search-unifiedauditlog).
        Operations: New-InboxRule, Set-InboxRule, UpdateInboxRules (Outlook client),
        Set-Mailbox (forwarding is the ForwardingSmtpAddress parameter, kept in the
        Parameters column), Add-MailboxPermission and Remove-MailboxPermission
        (https://learn.microsoft.com/purview/audit-log-activities#exchange-mailbox-activities).

        Mail flow rule changes are Exchange admin audit records, searched with
        -RecordType ExchangeAdmin; the activities page publishes no fixed operation list for
        them, so none is invented. That search is not filtered by operation at the service,
        and its records are kept only when the Operation contains 'TransportRule' (this
        collector's filter, not Microsoft's). Pass -SkipExchangeAdmin to leave it out.

        A call with no -SessionCommand returns 100 records at most, so each window is paged
        with the same -SessionId and -SessionCommand ReturnLargeSet until nothing comes
        back. ReturnLargeSet is unsorted and stops at 50,000: a window that reaches the cap
        is not written, the log names its exact -StartDate and -EndDate, and the collector
        stops so it can be re-run with a smaller -WindowHours. -StartDate and -EndDate are
        UTC.

        Needs auditing on and the Audit Reader role group (View-Only Audit Logs) or Audit
        Manager. Set-Mailbox records are visible only to unrestricted admins. Audit
        (Standard) retains 180 days; one year for Exchange records of users with E5 or an
        Audit (Premium) add-on; ten years needs the 10-year add-on and a retention policy.
        Exchange records are typically searchable 60 to 90 minutes after the event.

    .PARAMETER SkipExchangeAdmin
        Do not run the -RecordType ExchangeAdmin search for mail flow rule changes.

    .EXAMPLE
        ./Get-MailboxChangeEvents.ps1 -OutputPath ./out -LookbackDays 90 -WindowHours 6
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

    [datetime]$StartDate,
    [datetime]$EndDate,

    [ValidateRange(1, 180)]
    [int]$LookbackDays = 90,

    [ValidateRange(1, 24)]
    [int]$WindowHours = 24,

    [switch]$SkipExchangeAdmin,

    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Imported by path without -Force: a forced re-import builds a fresh module session
# state, which would discard anything the caller has arranged around the module - and
# that is exactly how the tests put mocks in front of the tenant cmdlets.
Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

. (Join-Path $PSScriptRoot 'MailboxExfiltrationHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'MailboxExfiltrationSchema.psd1')
$columns = $schema.AuditEvents
$source = 'mailbox-change-events'
$csvPath = Join-Path $OutputPath 'mailbox-change-events.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-MailboxSourceAvailability -Source 'MailboxChangeEvents' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping mailbox-change-events.csv. $($availability.Reason) $($availability.Reference)")
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}
if ($availability.Status -eq 'Unverified') {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message $availability.Reason
}

$connectedHere = -not $SkipConnect
if ($connectedHere) {
    Connect-M365Service -Service ExchangeOnline -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization
}

try {
    try {
        $watermark = Get-CsvWatermark -Path $csvPath -Column 'CreationTime'

        $start = if ($PSBoundParameters.ContainsKey('StartDate')) { $StartDate.ToUniversalTime() }
        elseif ($null -ne $watermark) { $watermark }
        else { [datetime]::UtcNow.AddDays(-$LookbackDays) }

        $end = if ($PSBoundParameters.ContainsKey('EndDate')) { $EndDate.ToUniversalTime() } else { [datetime]::UtcNow }

        if ($end -le $start) {
            if ($PSBoundParameters.ContainsKey('StartDate') -or $PSBoundParameters.ContainsKey('EndDate')) {
                throw ('The requested range is empty: the end ({0}) is not later than the start ({1}).' -f
                    (ConvertTo-CsvTimestamp $end), (ConvertTo-CsvTimestamp $start))
            }

            Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
                'Nothing new to collect: the watermark ({0}) is already at or after the end of the range.' -f (ConvertTo-CsvTimestamp $start))
            Export-AppendCsv -Path $csvPath -Column $columns
            return
        }

        Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
            'Searching the unified audit log from {0} to {1} in {2}-hour windows (watermark: {3}).' -f
            (ConvertTo-CsvTimestamp $start), (ConvertTo-CsvTimestamp $end), $WindowHours,
            $(if ($null -eq $watermark) { 'none' } else { ConvertTo-CsvTimestamp $watermark }))

        $found = Invoke-AuditSearch -Start $start -End $end -WindowHours $WindowHours -Operation $schema.MailboxChangeOperations `
            -OutputPath $OutputPath -Source $source

        if ($null -eq $found.TruncatedWindow -and -not $SkipExchangeAdmin) {
            $admin = Invoke-AuditSearch -Start $start -End $end -WindowHours $WindowHours -RecordType 'ExchangeAdmin' `
                -OutputPath $OutputPath -Source $source
            $found = [pscustomobject]@{
                Records         = @($found.Records) + @($admin.Records | Where-Object { [string]$_.Audit.Operation -like '*TransportRule*' })
                TruncatedWindow = $admin.TruncatedWindow
            }
        }
    }
    catch {
        # A mistaken range is the caller's error, not a refusal by the service.
        if ($_.Exception.Message -like 'The requested range is empty*') { throw }

        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            'Search-UnifiedAuditLog is unavailable to this sign-in ({0}). It needs auditing turned on and the Audit Reader role group. Writing the header only.' -f $_.Exception.Message)
        Export-AppendCsv -Path $csvPath -Column $columns
        return
    }

    $rows = @($found.Records | ForEach-Object { ConvertTo-AuditRow -Item $_ } | Where-Object { $true })
    $result = Export-AppendCsv -Path $csvPath -Rows $rows -Column $columns -KeyColumn 'Id' -PassThru
    Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
        'mailbox-change-events.csv: {0} rows written, {1} skipped as already collected.' -f $result.Written, $result.Skipped)

    if ($null -ne $found.TruncatedWindow) {
        $windowStart = ConvertTo-CsvTimestamp $found.TruncatedWindow.Start
        $windowEnd = ConvertTo-CsvTimestamp $found.TruncatedWindow.End
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            "The window $windowStart to $windowEnd matches 50,000 or more records; ReturnLargeSet is unsorted, so writing it would move the watermark past events that were never returned. Re-run this window with -StartDate $windowStart -EndDate $windowEnd and a smaller -WindowHours.")
        throw ('The unified audit log window {0} to {1} reached the 50,000-record session cap. Re-run with -StartDate {0} -EndDate {1} and a smaller -WindowHours.' -f $windowStart, $windowEnd)
    }
}
finally {
    if ($connectedHere) {
        Disconnect-ExchangeOnline -Confirm:$false
    }
}
