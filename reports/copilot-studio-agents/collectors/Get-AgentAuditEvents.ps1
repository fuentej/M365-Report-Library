#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes agent-audit-events.csv: Copilot Studio agent authoring events (create,
        delete, publish, share, authentication and component changes) from the unified
        audit log.

    .DESCRIPTION
        Source: Search-UnifiedAuditLog -Operations, reached through Exchange Online
        PowerShell
        (https://learn.microsoft.com/powershell/module/exchangepowershell/search-unifiedauditlog).
        The operations are the authoring event labels on
        https://learn.microsoft.com/microsoft-copilot-studio/admin-logging-copilot-studio,
        listed in CopilotStudioSchema.psd1. Authoring events are not returned by
        -RecordType CopilotInteraction, which is the usage record type and is not
        collected here.

        The cmdlet returns 100 records unless the same -SessionId is repeated with
        -SessionCommand ReturnLargeSet, which pages up to 50,000 records a session
        (-ResultSize up to 5,000). The range is walked in windows (-WindowHours), one
        session each. A window that reaches 50,000 is not written: the results are
        unsorted, so appending them would move the watermark past events never returned.
        The collector logs the exact -StartDate and -EndDate of that window and stops so
        it can be re-run with a smaller -WindowHours.

        Retention is 180 days. Audit (Premium)'s one-year default policy does not cover
        Copilot Studio (https://learn.microsoft.com/purview/audit-log-retention-policies).
        CreationTime is UTC. Appends from the latest CreationTime already collected.

        Needs the Audit Reader role group (View-Only Audit Logs). Whether Copilot Studio
        audit events are recorded in GCC and GCC High is UNVERIFIED, so those clouds are
        attempted with a warning, as is whether pay-as-you-go billing is required.

    .EXAMPLE
        ./Get-AgentAuditEvents.ps1 -OutputPath ./out -LookbackDays 90 -WindowHours 6
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

. (Join-Path $PSScriptRoot 'CopilotStudioHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'CopilotStudioSchema.psd1')
$columns = $schema.AgentAuditEvents
$operations = $schema.AuditOperations
$source = 'agent-audit-events'
$csvPath = Join-Path $OutputPath 'agent-audit-events.csv'

# ReturnLargeSet caps a session at 50,000 records, returned in pages of -ResultSize.
$pageSize = 5000
$sessionCap = 50000
# The service often returns nothing while a search is prepared. Retry briefly, then
# treat the window as empty.
$nullPageRetries = 3
$nullPageRetryDelayMs = 200

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-CopilotStudioSourceAvailability -Source 'AgentAuditEvents' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping agent-audit-events.csv. $($availability.Reason) $($availability.Reference)")
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
            throw ('The requested range is empty: the end ({0}) is not later than the start ({1}).' -f
                (ConvertTo-CsvTimestamp $end), (ConvertTo-CsvTimestamp $start))
        }

        Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
            'Nothing new to collect: the watermark ({0}) is already at or after the end of the range.' -f (ConvertTo-CsvTimestamp $start))
        Export-AppendCsv -Path $csvPath -Column $columns
        return
    }

    Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
        'Searching the unified audit log for {0} Copilot Studio authoring operations from {1} to {2} in {3}-hour windows (watermark: {4}).' -f
        $operations.Count, (ConvertTo-CsvTimestamp $start), (ConvertTo-CsvTimestamp $end), $WindowHours,
        $(if ($null -eq $watermark) { 'none' } else { ConvertTo-CsvTimestamp $watermark }))

    $rows = [System.Collections.Generic.List[object]]::new()
    $truncatedWindow = $null

    foreach ($window in Split-DateRange -Start $start -End $end -WindowMinutes ($WindowHours * 60)) {
        $sessionId = 'copilot-studio-{0:yyyyMMddHHmmss}-{1}' -f $window.Start, [guid]::NewGuid().ToString('N').Substring(0, 8)
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
                    'Search-UnifiedAuditLog is unavailable to this sign-in ({0}). It needs auditing turned on and the Audit Reader role group. Writing the header only.' -f $_.Exception.Message)
                Export-AppendCsv -Path $csvPath -Column $columns
                return
            }

            $records = @($raw | Where-Object { $null -ne $_ })
            if ($records.Count -eq 0) {
                if ($collected -eq 0 -and $nullTries -lt $nullPageRetries) {
                    $nullTries++
                    Start-Sleep -Milliseconds $nullPageRetryDelayMs
                    continue
                }
                break
            }

            # ResultCount is the hit count across every iteration of the session, not
            # the size of this page.
            $matched = 0
            $hasResultCount = $false
            $resultCountProperty = $records[0].PSObject.Properties['ResultCount']
            if ($resultCountProperty -and $null -ne $resultCountProperty.Value) {
                $hasResultCount = [int]::TryParse([string]$resultCountProperty.Value, [ref]$matched)
            }
            if ($hasResultCount -and $matched -gt $sessionCap) {
                $windowTruncated = $true
                break
            }

            $collected += $records.Count

            foreach ($record in $records) {
                $auditDataProperty = $record.PSObject.Properties['AuditData']
                if (-not $auditDataProperty -or [string]::IsNullOrWhiteSpace([string]$auditDataProperty.Value)) { continue }
                try {
                    $audit = [string]$auditDataProperty.Value | ConvertFrom-Json -ErrorAction Stop
                }
                catch {
                    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
                        'Skipping a record whose AuditData is not valid JSON: {0}' -f $_.Exception.Message)
                    continue
                }

                $windowRows.Add([pscustomobject]@{
                        CreationTime     = ConvertTo-CsvTimestamp (Get-GraphAdditionalProperty -Object $audit -Name 'CreationTime')
                        Id               = [string](Get-GraphAdditionalProperty -Object $audit -Name 'Id')
                        Operation        = [string](Get-GraphAdditionalProperty -Object $audit -Name 'Operation')
                        UserId           = [string](Get-GraphAdditionalProperty -Object $audit -Name 'UserKey')
                        ResultStatus     = [string](Get-GraphAdditionalProperty -Object $audit -Name 'ResultStatus')
                        BotId            = [string](Get-GraphAdditionalProperty -Object $audit -Name 'BotId')
                        BotSchemaName    = [string](Get-GraphAdditionalProperty -Object $audit -Name 'BotSchemaName')
                        BotComponentId   = [string](Get-GraphAdditionalProperty -Object $audit -Name 'BotComponentId')
                        BotComponentType = [string](Get-GraphAdditionalProperty -Object $audit -Name 'BotComponentType')
                    })
            }

            if ($collected -ge $sessionCap) {
                $windowTruncated = $true
                break
            }

            # Repeat until the cmdlet returns nothing or the cap is hit. A page shorter
            # than -ResultSize is not the end on its own; the reported total is.
            if ($hasResultCount -and $matched -gt 0 -and $collected -ge $matched) { break }
            if ((-not $hasResultCount -or $matched -le 0) -and $records.Count -lt $pageSize) { break }
        }

        if ($windowTruncated) {
            $truncatedWindow = $window
            break
        }

        foreach ($row in $windowRows) { $rows.Add($row) }
    }

    $result = Export-AppendCsv -Path $csvPath -Rows $rows.ToArray() -Column $columns -KeyColumn 'Id' -PassThru
    Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
        'agent-audit-events.csv: {0} rows written, {1} skipped as already collected.' -f $result.Written, $result.Skipped)

    if ($null -ne $truncatedWindow) {
        $windowStart = ConvertTo-CsvTimestamp $truncatedWindow.Start
        $windowEnd = ConvertTo-CsvTimestamp $truncatedWindow.End
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
