#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes sharing-events.csv: SharePoint and OneDrive sharing activity from the
        unified audit log.

    .DESCRIPTION
        Source: Search-UnifiedAuditLog
        (https://learn.microsoft.com/powershell/module/exchangepowershell/search-unifiedauditlog),
        reached through Exchange Online PowerShell.

        The operations collected are the sharing and access request activities listed at
        https://learn.microsoft.com/purview/audit-log-activities#sharing-and-access-request-activities
        and enumerated in GuestAccessSchema.psd1.

        A -SessionId with -SessionCommand ReturnLargeSet returns up to 50,000 unsorted
        results per session, so the range is walked in windows (-WindowHours) and each
        window gets its own session. A window that matches more than 50,000 records is
        not written: the results are unsorted, so appending them would move the watermark
        past events that were never returned. The collector logs the exact -StartDate and
        -EndDate of that window and stops so it can be re-run with a smaller -WindowHours.

        A $null page is retried a few times; the service often returns nothing on the
        first call while the search is still being prepared. An empty collection means
        the session is done.

        Retention is 180 days in Audit (Standard)
        (https://learn.microsoft.com/purview/audit-log-retention-policies).

    .EXAMPLE
        ./Get-SharingEvents.ps1 -OutputPath ./out -LookbackDays 90 -WindowHours 6
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

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'GuestAccessSchema.psd1')
$columns = $schema.SharingEvents
$operations = $schema.SharingOperations
$source = 'sharing-events'
$csvPath = Join-Path $OutputPath 'sharing-events.csv'

# ReturnLargeSet caps a session at 50,000 records, returned in pages of -ResultSize.
$pageSize = 5000
$sessionCap = 50000
# $null (as opposed to an empty collection) usually means the search is not ready yet.
$nullPageRetries = 3

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$connectedHere = -not $SkipConnect
if ($connectedHere) {
    Connect-M365Service -Service ExchangeOnline -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization
}

try {
    $watermark = Get-CsvWatermark -Path $csvPath -Column 'CreationTime'

    $start = if ($PSBoundParameters.ContainsKey('StartDate')) { $StartDate.ToUniversalTime() }
    elseif ($null -ne $watermark) { $watermark }
    else { [datetime]::UtcNow.AddDays(-$LookbackDays) }

    $end = if ($PSBoundParameters.ContainsKey('EndDate')) { $EndDate.ToUniversalTime() } else { [datetime]::UtcNow }

    if ($end -le $start) {
        if ($PSBoundParameters.ContainsKey('StartDate') -or $PSBoundParameters.ContainsKey('EndDate')) {
            # An inverted or empty range asked for explicitly is a mistake, and collecting
            # nothing while reporting success would hide it.
            throw ('The requested range is empty: the end ({0}) is not later than the start ({1}).' -f
                (ConvertTo-CsvTimestamp $end), (ConvertTo-CsvTimestamp $start))
        }

        # Resuming from a watermark that already sits at or past now: nothing has happened
        # since the last run. Make sure the file exists, say so, and stop.
        Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
            'Nothing new to collect: the watermark ({0}) is already at or after the end of the range.' -f (ConvertTo-CsvTimestamp $start))
        Export-AppendCsv -Path $csvPath -Column $columns
        return
    }

    Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
        'Searching the unified audit log for {0} operations from {1} to {2} in {3}-hour windows (watermark: {4}).' -f
        $operations.Count, (ConvertTo-CsvTimestamp $start), (ConvertTo-CsvTimestamp $end), $WindowHours,
        $(if ($null -eq $watermark) { 'none' } else { ConvertTo-CsvTimestamp $watermark }))

    $rows = [System.Collections.Generic.List[object]]::new()
    $truncatedWindow = $null

    foreach ($window in Split-DateRange -Start $start -End $end -WindowMinutes ($WindowHours * 60)) {
        if ($null -ne $truncatedWindow) { break }

        $sessionId = 'guest-access-{0:yyyyMMddHHmmss}-{1}' -f $window.Start, [guid]::NewGuid().ToString('N').Substring(0, 8)
        $collected = 0
        $nullTries = 0
        $windowRows = [System.Collections.Generic.List[object]]::new()
        $windowTruncated = $false

        while ($collected -lt $sessionCap) {
            $raw = $null
            try {
                $raw = Search-UnifiedAuditLog -StartDate $window.Start -EndDate $window.End `
                    -Operations $operations -SessionId $sessionId -SessionCommand ReturnLargeSet `
                    -ResultSize $pageSize -ErrorAction Stop
            }
            catch {
                Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
                    'Search-UnifiedAuditLog is unavailable to this sign-in ({0}). It needs auditing turned on and the View-Only Audit Logs or Audit Logs role. Writing the header only.' -f $_.Exception.Message)
                Export-AppendCsv -Path $csvPath -Column $columns
                return
            }

            # $null = not ready. @() = this session has nothing (more). Do not wrap $null
            # in @(): that is a one-element array and looks like a page of data.
            if ($null -eq $raw) {
                if ($nullTries -lt $nullPageRetries) {
                    $nullTries++
                    Start-Sleep -Seconds 1
                    continue
                }
                break
            }

            $records = @($raw | Where-Object { $null -ne $_ })
            if ($records.Count -eq 0) { break }

            $resultCountProperty = $records[0].PSObject.Properties['ResultCount']
            if ($resultCountProperty -and $null -ne $resultCountProperty.Value) {
                $matched = 0
                if ([int]::TryParse([string]$resultCountProperty.Value, [ref]$matched) -and $matched -gt $sessionCap) {
                    $windowTruncated = $true
                    break
                }
            }

            $collected += $records.Count

            foreach ($record in $records) {
                $audit = $null
                $auditDataProperty = $record.PSObject.Properties['AuditData']
                if ($auditDataProperty -and -not [string]::IsNullOrWhiteSpace([string]$auditDataProperty.Value)) {
                    try {
                        $audit = [string]$auditDataProperty.Value | ConvertFrom-Json -ErrorAction Stop
                    }
                    catch {
                        Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
                            'Skipping a record whose AuditData is not valid JSON: {0}' -f $_.Exception.Message)
                        continue
                    }
                }
                if ($null -eq $audit) { continue }

                $windowRows.Add([pscustomobject]@{
                        CreationTime          = ConvertTo-CsvTimestamp (Get-GraphAdditionalProperty -Object $audit -Name 'creationTime')
                        Id                    = [string](Get-GraphAdditionalProperty -Object $audit -Name 'id')
                        Operation             = [string](Get-GraphAdditionalProperty -Object $audit -Name 'operation')
                        UserId                = [string](Get-GraphAdditionalProperty -Object $audit -Name 'userId')
                        Workload              = [string](Get-GraphAdditionalProperty -Object $audit -Name 'workload')
                        SiteUrl               = [string](Get-GraphAdditionalProperty -Object $audit -Name 'siteUrl')
                        ObjectId              = [string](Get-GraphAdditionalProperty -Object $audit -Name 'objectId')
                        SourceFileName        = [string](Get-GraphAdditionalProperty -Object $audit -Name 'sourceFileName')
                        TargetUserOrGroupName = [string](Get-GraphAdditionalProperty -Object $audit -Name 'targetUserOrGroupName')
                        TargetUserOrGroupType = [string](Get-GraphAdditionalProperty -Object $audit -Name 'targetUserOrGroupType')
                    })
            }

            if ($collected -ge $sessionCap) {
                $windowTruncated = $true
                break
            }

            if ($records.Count -lt $pageSize) { break }
        }

        if ($windowTruncated) {
            $truncatedWindow = $window
            break
        }

        foreach ($row in $windowRows) { $rows.Add($row) }
    }

    $result = Export-AppendCsv -Path $csvPath -Rows $rows.ToArray() -Column $columns -KeyColumn 'Id' -PassThru
    Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
        'sharing-events.csv: {0} rows written, {1} skipped as already collected.' -f $result.Written, $result.Skipped)

    if ($null -ne $truncatedWindow) {
        $windowStart = ConvertTo-CsvTimestamp $truncatedWindow.Start
        $windowEnd = ConvertTo-CsvTimestamp $truncatedWindow.End
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            "The window $windowStart to $windowEnd matches more than 50,000 records; ReturnLargeSet is unsorted, so writing it would move the watermark past events that were never returned. Re-run this window with -StartDate $windowStart -EndDate $windowEnd and a smaller -WindowHours.")
        throw ('The unified audit log window {0} to {1} exceeded the 50,000-record session cap. Re-run with -StartDate {0} -EndDate {1} and a smaller -WindowHours.' -f $windowStart, $windowEnd)
    }
}
finally {
    if ($connectedHere) {
        Disconnect-ExchangeOnline -Confirm:$false
    }
}
