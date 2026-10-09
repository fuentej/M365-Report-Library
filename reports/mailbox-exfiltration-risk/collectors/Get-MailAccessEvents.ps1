#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes mail-access-events.csv: audit events for mail items accessed and for mail
        sent, sent as, or sent on behalf, appended from where the last run stopped.

    .DESCRIPTION
        Source: Search-UnifiedAuditLog in Exchange Online PowerShell
        (https://learn.microsoft.com/powershell/module/exchangepowershell/search-unifiedauditlog).
        Operations: MailItemsAccessed (read or protocol access, not a send), Send (sent,
        replied or forwarded mail), SendAs and SendOnBehalf
        (https://learn.microsoft.com/purview/audit-log-activities#exchange-mailbox-activities).

        MailItemsAccessed is Audit (Standard) and is on by default for users with Office
        365 E3/E5 or Microsoft 365 E3/E5
        (https://learn.microsoft.com/purview/audit-log-investigate-accounts). Its
        SensitivityLabel property is an Audit (Premium) property and is not collected here;
        a missing Premium licence is not the same as zero access events. Accesses in one
        bind record are aggregated over about two minutes, and duplicate binds inside an
        hour are filtered.

        Paging, the 50,000-record cap, the UTC dates, the roles and the retention are as in
        Get-MailboxChangeEvents.ps1: each window is paged with ReturnLargeSet, and a window
        that reaches the cap is not written but named in the log so it can be re-run with a
        smaller -WindowHours.

    .EXAMPLE
        ./Get-MailAccessEvents.ps1 -OutputPath ./out -LookbackDays 30 -WindowHours 1
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
$source = 'mail-access-events'
$csvPath = Join-Path $OutputPath 'mail-access-events.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-MailboxSourceAvailability -Source 'MailAccessEvents' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping mail-access-events.csv. $($availability.Reason) $($availability.Reference)")
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

        $found = Invoke-AuditSearch -Start $start -End $end -WindowHours $WindowHours -Operation $schema.MailAccessOperations `
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

    $rows = @($found.Records | ForEach-Object { ConvertTo-AuditRow -Item $_ } | Where-Object { $true })
    $result = Export-AppendCsv -Path $csvPath -Rows $rows -Column $columns -KeyColumn 'Id' -PassThru
    Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
        'mail-access-events.csv: {0} rows written, {1} skipped as already collected.' -f $result.Written, $result.Skipped)

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
