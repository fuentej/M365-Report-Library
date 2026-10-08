#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes group-creation-events.csv: Microsoft 365 group and Team creation events,
        with the creator, from the unified audit log.

    .DESCRIPTION
        Source: Search-UnifiedAuditLog
        (https://learn.microsoft.com/powershell/module/exchangepowershell/search-unifiedauditlog),
        reached through Exchange Online PowerShell. The operations are AddGroup ("Added
        group", no trailing period) and TeamCreated ("Created team")
        (https://learn.microsoft.com/purview/audit-log-activities). AddGroup covers Microsoft
        365 groups and security groups created in the Microsoft 365 admin center or the
        Azure portal; it does not replace TeamCreated. Creation over time for every group
        comes from createdDateTime in groups.csv, so this file only adds the creator.

        TargetDisplayName and TargetGroupId are read from the TeamName and TeamGuid audit
        properties (https://learn.microsoft.com/purview/audit-log-detailed-properties), which
        a Teams record carries. Learn documents no group-name property for AddGroup, so for
        those records the two columns stay empty and ObjectId holds what the record gives.

        A -SessionId with -SessionCommand ReturnLargeSet returns up to 50,000 unsorted
        results per session, so the range is walked in windows (-WindowHours) and each
        window gets its own session. A window that matches more than 50,000 records is
        not written: the results are unsorted, so appending them would move the watermark
        past events that were never returned. The collector logs the exact -StartDate and
        -EndDate of that window and stops so it can be re-run with a smaller -WindowHours.

        Retention is 180 days in Audit (Standard); TeamCreated is a Teams workload, so the
        one-year Entra default does not cover it
        (https://learn.microsoft.com/purview/audit-log-retention-policies). Audit (Standard)
        is available in all three clouds; that AddGroup and TeamCreated are recorded in GCC
        and GCC High is UNVERIFIED, so those clouds are attempted with a warning.

        Needs auditing turned on and the View-Only Audit Logs or Audit Logs role.

    .EXAMPLE
        ./Get-GroupCreationEvents.ps1 -OutputPath ./out -LookbackDays 90 -WindowHours 6
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

. (Join-Path $PSScriptRoot 'TeamsGroupsHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'TeamsGroupsSchema.psd1')
$columns = $schema.GroupCreationEvents
$operations = $schema.CreationOperations
$source = 'group-creation-events'
$csvPath = Join-Path $OutputPath 'group-creation-events.csv'

# ReturnLargeSet caps a session at 50,000 records, returned in pages of -ResultSize.
$pageSize = 5000
$sessionCap = 50000
# A function or cmdlet that outputs nothing assigns $null, same as "not ready".
# Retry a few times, briefly, then treat the window as empty.
$nullPageRetries = 3
$nullPageRetryDelayMs = 200

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-TeamsGroupsSourceAvailability -Source 'GroupCreationEvents' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping group-creation-events.csv. $($availability.Reason) $($availability.Reference)")
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
        'Searching the unified audit log for {0} creation operations from {1} to {2} in {3}-hour windows (watermark: {4}).' -f
        $operations.Count, (ConvertTo-CsvTimestamp $start), (ConvertTo-CsvTimestamp $end), $WindowHours,
        $(if ($null -eq $watermark) { 'none' } else { ConvertTo-CsvTimestamp $watermark }))

    $rows = [System.Collections.Generic.List[object]]::new()
    $truncatedWindow = $null

    function Test-AuditSearchHasMoreRecords {
        <#
            .SYNOPSIS
                True when a Search-UnifiedAuditLog record says another page is expected.
        #>
        param($Record)

        if ($null -eq $Record) { return $false }

        $metaProperty = $Record.PSObject.Properties['AuditSearchRequestMetadata']
        if (-not $metaProperty -or $null -eq $metaProperty.Value) { return $false }

        $meta = $metaProperty.Value
        $flag = $null
        if ($meta -is [System.Collections.IDictionary]) {
            foreach ($key in @('moreRecordsAvailable', 'MoreRecordsAvailable')) {
                if ($meta.Contains($key)) { $flag = $meta[$key]; break }
            }
        }
        else {
            foreach ($name in @('moreRecordsAvailable', 'MoreRecordsAvailable')) {
                $property = $meta.PSObject.Properties[$name]
                if ($property) { $flag = $property.Value; break }
            }
        }

        if ($null -eq $flag) { return $false }
        if ($flag -is [bool]) { return $flag }

        $parsed = $false
        if ([bool]::TryParse([string]$flag, [ref]$parsed)) { return $parsed }
        return $false
    }

    foreach ($window in Split-DateRange -Start $start -End $end -WindowMinutes ($WindowHours * 60)) {
        if ($null -ne $truncatedWindow) { break }

        $sessionId = 'teams-groups-{0:yyyyMMddHHmmss}-{1}' -f $window.Start, [guid]::NewGuid().ToString('N').Substring(0, 8)
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

            # Do not wrap $null in @(): that is a one-element array and looks like data.
            # $null and an empty collection both mean "nothing this call". Retry a few
            # times (the service often returns nothing while the search is prepared),
            # then treat the window as empty.
            $records = @($raw | Where-Object { $null -ne $_ })
            if ($records.Count -eq 0) {
                if ($collected -eq 0 -and $nullTries -lt $nullPageRetries) {
                    $nullTries++
                    Start-Sleep -Milliseconds $nullPageRetryDelayMs
                    continue
                }
                break
            }

            # ResultCount is the hit count across every iteration of this session, not
            # the size of this page.
            # https://learn.microsoft.com/powershell/module/exchangepowershell/search-unifiedauditlog
            $resultCountProperty = $records[0].PSObject.Properties['ResultCount']
            $matched = 0
            $hasResultCount = $false
            if ($resultCountProperty -and $null -ne $resultCountProperty.Value) {
                if ([int]::TryParse([string]$resultCountProperty.Value, [ref]$matched)) {
                    $hasResultCount = $true
                }
            }
            if ($hasResultCount -and $matched -gt $sessionCap) {
                $windowTruncated = $true
                break
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
                        ObjectId              = [string](Get-GraphAdditionalProperty -Object $audit -Name 'objectId')
                        TargetDisplayName     = [string](Get-GraphAdditionalProperty -Object $audit -Name 'teamName')
                        TargetGroupId         = [string](Get-GraphAdditionalProperty -Object $audit -Name 'teamGuid')
                    })
            }

            if ($collected -ge $sessionCap) {
                $windowTruncated = $true
                break
            }

            # Repeat until the cmdlet returns nothing or the session cap is hit.
            # A page shorter than -ResultSize is not the end: ResultCount is the
            # hit count across iterations, and moreRecordsAvailable says another
            # iteration is still expected.
            # https://learn.microsoft.com/powershell/module/exchangepowershell/search-unifiedauditlog
            $moreRecords = $false
            foreach ($record in $records) {
                if (Test-AuditSearchHasMoreRecords -Record $record) {
                    $moreRecords = $true
                    break
                }
            }
            if ($moreRecords) { continue }

            $reportedTotalReached = $hasResultCount -and $matched -gt 0 -and $collected -ge $matched
            $shortPageWithoutTotal = (-not $hasResultCount -or $matched -le 0) -and $records.Count -lt $pageSize
            if ($reportedTotalReached -or $shortPageWithoutTotal) { break }
        }

        if ($windowTruncated) {
            $truncatedWindow = $window
            break
        }

        foreach ($row in $windowRows) { $rows.Add($row) }
    }

    $result = Export-AppendCsv -Path $csvPath -Rows $rows.ToArray() -Column $columns -KeyColumn 'Id' -PassThru
    Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
        'group-creation-events.csv: {0} rows written, {1} skipped as already collected.' -f $result.Written, $result.Skipped)

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
