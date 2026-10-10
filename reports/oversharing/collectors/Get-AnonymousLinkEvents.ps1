#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes anonymous-link-events.csv: audit events for Anyone links created, updated, used
        and removed, appended from where the last run stopped.

    .DESCRIPTION
        Source: Search-UnifiedAuditLog in Exchange Online PowerShell
        (https://learn.microsoft.com/powershell/module/exchangepowershell/search-unifiedauditlog).
        Operations: AnonymousLinkCreated, AnonymousLinkUpdated, AnonymousLinkUsed and
        AnonymousLinkRemoved
        (https://learn.microsoft.com/purview/audit-log-activities#sharing-and-access-request-activities).

        A call with no -SessionCommand returns at most 100 records. -ResultSize defaults to
        100 and its maximum is 5,000, which is one page, not the session cap. Each window is
        paged by repeating the same -SessionId and -SessionCommand ReturnLargeSet until a call
        returns no records or AuditSearchRequestMetadata.moreRecordsAvailable is false.
        ReturnLargeSet is unsorted and stops at 50,000 results; a window that reaches 50,000
        is not the full window and is not written. The collector names its exact -StartDate and
        -EndDate in run.log and stops, so it can be re-run with a smaller -WindowHours. Do not
        switch -SessionCommand on one session (output drops to 10,000). -HighCompleteness is
        preview and not in every tenant, so it is not used; without it the cmdlet page says
        results can be missing.

        -StartDate and -EndDate are UTC: a value with no time zone is UTC, and a date with no
        time is midnight UTC. Appends from the latest CreationTime already collected.

        Role: View-Only Audit Logs or Audit Logs in the Microsoft Purview portal (Audit Reader
        holds the view-only role, Audit Manager holds both); the audit cmdlets also need the
        Exchange admin center Audit Logs or View-Only Audit Logs role
        (https://learn.microsoft.com/purview/audit-get-started#step-2-assign-permissions-to-search-the-audit-log).
        Audit (Standard) is enough. Retention is 180 days for records generated on or after 17
        October 2023, one year only for Exchange, SharePoint, OneDrive and Microsoft Entra
        records of users with E5 or the audit add-on, and a custom retention policy can be
        shorter (https://learn.microsoft.com/purview/audit-log-retention-policies).
        -LookbackDays accepts up to 365 so a first run can cover that one-year retention.
        Learn
        recommends the Management Activity API for programmatic export; its availability in
        GCC and GCC High was not checked, and this collector does not use it.

        Who created a link is the UserId on the AnonymousLinkCreated record; sharing-link
        permissions on an item (item-permissions.csv) do not name the creator.

    .EXAMPLE
        ./Get-AnonymousLinkEvents.ps1 -OutputPath ./out -LookbackDays 30 -WindowHours 1
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

    [ValidateRange(1, 365)]
    [int]$LookbackDays = 90,

    [ValidateRange(1, 24)]
    [int]$WindowHours = 24,

    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Imported by path without -Force: a forced re-import builds a fresh module session
# state, which would discard anything the caller has arranged around the module - and
# that is exactly how the tests put mocks in front of the tenant cmdlets.
Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

. (Join-Path $PSScriptRoot 'OversharingHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'OversharingSchema.psd1')
$columns = $schema.AuditEvents
$source = 'anonymous-link-events'
$csvPath = Join-Path $OutputPath 'anonymous-link-events.csv'

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-OversharingSourceAvailability -Source 'AnonymousLinkEvents' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping anonymous-link-events.csv. $($availability.Reason) $($availability.Reference)")
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

        $start = if ($PSBoundParameters.ContainsKey('StartDate')) { ConvertTo-AuditQueryDate $StartDate }
        elseif ($null -ne $watermark) { ConvertTo-AuditQueryDate $watermark }
        else { ConvertTo-AuditQueryDate ([datetime]::UtcNow.AddDays(-$LookbackDays)) }

        $end = if ($PSBoundParameters.ContainsKey('EndDate')) { ConvertTo-AuditQueryDate $EndDate } else { ConvertTo-AuditQueryDate ([datetime]::UtcNow) }

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

        $found = Invoke-AuditSearch -Start $start -End $end -WindowHours $WindowHours -Operation $schema.AnonymousLinkOperations `
            -OutputPath $OutputPath -Source $source
    }
    catch {
        # A mistaken range is the caller's error, not a refusal by the service.
        if ($_.Exception.Message -like 'The requested range is empty*') { throw }

        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            'Search-UnifiedAuditLog is unavailable to this sign-in ({0}). It needs auditing turned on and the Audit Reader role group. Writing the header only.' -f $_.Exception.Message)
        Export-AppendCsv -Path $csvPath -Column $columns
        return
    }

    $rows = @($found.Records | ForEach-Object { ConvertTo-SharingAuditRow -Item $_ } | Where-Object { $true })
    $result = Export-AppendCsv -Path $csvPath -Rows $rows -Column $columns -KeyColumn 'Id' -PassThru
    Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
        'anonymous-link-events.csv: {0} rows written, {1} skipped as already collected.' -f $result.Written, $result.Skipped)

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
