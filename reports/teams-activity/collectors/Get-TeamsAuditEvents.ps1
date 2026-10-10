#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes teams-audit-events.csv: Teams meeting, call, message and chat-created events from the unified audit log, counted per UTC day, workload, user and operation, resuming from the last day collected.

    .DESCRIPTION
        Source 8: Search-UnifiedAuditLog in Exchange Online PowerShell
        (https://learn.microsoft.com/powershell/module/exchangepowershell/search-unifiedauditlog).
        The operations are the ones the contract lists from
        https://learn.microsoft.com/purview/audit-log-activities#teams-activities (MeetingDetail,
        MeetingParticipantDetail, CallParticipantDetail, MessageSent, ChatCreated).

        The records are counted here, by the library: the counts will not match the usage report counts
        and must not be compared with them. MessageSent is in public preview and is generated for chat
        only when guests, federated or anonymous users are present, so it is not a complete chat count.
        ChatCreated is logged only for chats created through a Graph API call. MeetingParticipantDetail
        already includes recorded or transcribed calls, so it is not added to CallParticipantDetail.
        CallParticipantDetail, MessageSent and ChatCreated are UNVERIFIED in GCC and GCC High, and the
        collector logs a warning there.

        Event source. The query resumes the day after the latest Date in the file, and only whole UTC
        days before today are written, so a day is never counted twice or half counted. The first run
        reaches back -LookbackDays. Each window is paged with the same -SessionId and
        -SessionCommand ReturnLargeSet until the call returns nothing. A window that reaches the
        50,000-record session cap is incomplete and unsorted: that day and every later one are not
        written, an error is logged, and the run stops so the resume point never moves past events that
        were not returned. Re-run it with -StartDate, -EndDate and a smaller -WindowHours.
        -HighCompleteness is preview and not in every tenant, so it is not used.

        Needs the View-Only Audit Logs (or Audit Logs) role and auditing turned on. Audit (Standard)
        keeps Teams records for 180 days. Available in Commercial, GCC and GCC High.

    .EXAMPLE
        ./Get-TeamsAuditEvents.ps1 -OutputPath ./out
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

    # 180 days is the Audit (Standard) retention for Teams records.
    # https://learn.microsoft.com/purview/audit-log-retention-policies
    [ValidateRange(1, 180)]
    [int]$LookbackDays = 30,

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

. (Join-Path $PSScriptRoot 'TeamsActivityHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'TeamsActivitySchema.psd1')
$columns = $schema.TeamsAuditEvents
$source = 'teams-audit-events'
$csvPath = Join-Path $OutputPath 'teams-audit-events.csv'

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

if (Test-TeamsSourceSkipped -Source 'TeamsAuditEvents' -LogSource $source -CsvPath $csvPath -Column $columns `
        -OutputPath $OutputPath -Environment $Environment -Schema $schema) {
    return
}

if ($Environment -ne 'Commercial') {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        'Whether {0} are logged in {1} is UNVERIFIED; the collector will search for them anyway.' -f
        ($schema.UnverifiedAuditOperations -join ', '), $Environment)
}

$connectedHere = -not $SkipConnect
if ($connectedHere) {
    Connect-M365Service -Service ExchangeOnline -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization
}

try {
    $today = [datetime]::SpecifyKind([datetime]::UtcNow.Date, [DateTimeKind]::Utc)
    try {
        $watermark = Get-CsvWatermark -Path $csvPath -Column 'Date'

        # The latest Date in the file is a whole day, so the next run starts the day after.
        $start = if ($PSBoundParameters.ContainsKey('StartDate')) { ConvertTo-AuditQueryDate $StartDate }
        elseif ($null -ne $watermark) { ConvertTo-AuditQueryDate $watermark.AddDays(1) }
        else { ConvertTo-AuditQueryDate $today.AddDays(-$LookbackDays) }

        # Whole days only: today is still being written.
        $end = if ($PSBoundParameters.ContainsKey('EndDate')) { ConvertTo-AuditQueryDate $EndDate } else { $today }

        if ($end -le $start) {
            if ($PSBoundParameters.ContainsKey('StartDate') -or $PSBoundParameters.ContainsKey('EndDate')) {
                throw ('The requested range is empty: the end ({0}) is not later than the start ({1}).' -f
                    (ConvertTo-CsvTimestamp $end), (ConvertTo-CsvTimestamp $start))
            }

            Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
                'Nothing new to collect: the last day collected ({0}) is the last whole day.' -f
                $(if ($null -eq $watermark) { 'none' } else { $watermark.ToString('yyyy-MM-dd') }))
            Export-AppendCsv -Path $csvPath -Column $columns
            return
        }

        Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
            'Searching the unified audit log from {0} to {1} in {2}-hour windows (last day collected: {3}).' -f
            (ConvertTo-CsvTimestamp $start), (ConvertTo-CsvTimestamp $end), $WindowHours,
            $(if ($null -eq $watermark) { 'none' } else { $watermark.ToString('yyyy-MM-dd') }))

        $found = Invoke-AuditSearch -Start $start -End $end -WindowHours $WindowHours -Operation $schema.TeamsAuditOperations `
            -OutputPath $OutputPath -Source $source
    }
    catch {
        # A mistaken range is the caller's error, not a refusal by the service.
        if ($_.Exception.Message -like 'The requested range is empty*') { throw }

        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            'Search-UnifiedAuditLog is unavailable to this sign-in ({0}). It needs auditing turned on and the View-Only Audit Logs role. Writing the header only.' -f $_.Exception.Message)
        Export-AppendCsv -Path $csvPath -Column $columns
        return
    }

    # A truncated window makes its own day, and every later one, incomplete.
    $cutoff = $null
    if ($null -ne $found.TruncatedWindow) { $cutoff = $found.TruncatedWindow.Start.Date }

    $rows = Group-TeamsEventCount -Item $found.Records -Before $cutoff
    $result = Export-AppendCsv -Path $csvPath -Rows $rows -Column $columns -KeyColumn @('Date', 'Workload', 'UserId', 'Operation') -PassThru
    Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
        'teams-audit-events.csv: {0} rows written, {1} skipped as already collected.' -f $result.Written, $result.Skipped)

    if ($null -ne $found.TruncatedWindow) {
        $windowStart = ConvertTo-CsvTimestamp $found.TruncatedWindow.Start
        $windowEnd = ConvertTo-CsvTimestamp $found.TruncatedWindow.End
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            "The window $windowStart to $windowEnd matches 50,000 or more records; ReturnLargeSet is unsorted, so writing it would move the resume point past events that were never returned. Re-run it with -StartDate $($found.TruncatedWindow.Start.Date.ToString('yyyy-MM-dd')) -EndDate $windowEnd and a smaller -WindowHours.")
        throw ('The unified audit log window {0} to {1} reached the 50,000-record session cap. Re-run with -StartDate {2} and a smaller -WindowHours.' -f $windowStart, $windowEnd, $found.TruncatedWindow.Start.Date.ToString('yyyy-MM-dd'))
    }
}
finally {
    if ($connectedHere) {
        Disconnect-ExchangeOnline -Confirm:$false
    }
}
