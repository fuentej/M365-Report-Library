#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes copilot-accessed-resources.csv: the resources Microsoft 365
        Copilot accessed, appended each run.

    .DESCRIPTION
        Source: CopilotInteraction records from Search-UnifiedAuditLog, reached
        through Exchange Online PowerShell, flattening
        CopilotEventData.AccessedResources into one row per resource, with the
        resource's sensitivity label ID, the status of Copilot's access, and
        the policy that blocked or restricted it when there was one. An
        interaction that touched no resource keeps one row with the resource
        columns empty.

        Search-UnifiedAuditLog is used rather than the Graph audit log query
        API because RecordType CopilotInteraction is documented against the
        unified audit log (https://learn.microsoft.com/purview/audit-copilot).

        Each run resumes from the watermark already in
        copilot-accessed-resources.csv (or -LookbackDays ago on a first run)
        and appends. Rows are keyed on RecordId plus ResourceId, so a window
        that overlaps the previous run's does not duplicate rows.

        A -SessionId with -SessionCommand ReturnLargeSet returns up to 50,000
        unsorted results per session, so the range is walked in windows
        (-WindowHours) and each window gets its own session. A window that
        matches more than 50,000 records is not written: the results are
        unsorted, so appending them would move the watermark past events that
        were never returned. The collector logs the exact range and stops so
        it can be re-run with a smaller -WindowHours.

        Read-only: the only tenant cmdlet it calls is Search-UnifiedAuditLog.

    .PARAMETER LookbackDays
        How far back to go when there is no watermark to resume from.

    .PARAMETER SkipConnect
        Use an existing Exchange Online PowerShell session instead of signing
        in.

    .EXAMPLE
        ./Get-CopilotAccessedResources.ps1 -OutputPath ./out -LookbackDays 90 -WindowHours 6

    .LINK
        https://learn.microsoft.com/purview/audit-copilot
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

    [ValidateRange(1, 180)]
    [int]$LookbackDays = 30,

    [ValidateRange(1, 24)]
    [int]$WindowHours = 24,

    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

. (Join-Path $PSScriptRoot 'PurviewIpHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'PurviewIpSchema.psd1')
$columns = $schema.CopilotAccessedResources
$source = 'copilot-accessed-resources'
$csvPath = Join-Path $OutputPath 'copilot-accessed-resources.csv'

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

$availability = Get-PurviewSourceAvailability -Source 'CopilotAuditRecords' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping copilot-accessed-resources.csv. $($availability.Reason) $($availability.Reference)")
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

    $start = if ($null -ne $watermark) { $watermark } else { [datetime]::UtcNow.AddDays(-$LookbackDays) }
    $end = [datetime]::UtcNow

    if ($end -le $start) {
        Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
            'Nothing new to collect: the watermark ({0}) is already at or after now.' -f (ConvertTo-CsvTimestamp $start))
        Export-AppendCsv -Path $csvPath -Column $columns
        return
    }

    Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
        'Searching the unified audit log for CopilotInteraction records from {0} to {1} in {2}-hour windows (watermark: {3}).' -f
        (ConvertTo-CsvTimestamp $start), (ConvertTo-CsvTimestamp $end), $WindowHours,
        $(if ($null -eq $watermark) { 'none' } else { ConvertTo-CsvTimestamp $watermark }))

    $rows = [System.Collections.Generic.List[object]]::new()
    $truncatedWindow = $null

    foreach ($window in (Split-DateRange -Start $start -End $end -WindowMinutes ($WindowHours * 60))) {
        if ($null -ne $truncatedWindow) { break }

        $sessionId = 'purview-ip-{0:yyyyMMddHHmmss}-{1}' -f $window.Start, [guid]::NewGuid().ToString('N').Substring(0, 8)
        $collected = 0
        $nullTries = 0
        $windowRows = [System.Collections.Generic.List[object]]::new()
        $windowTruncated = $false

        while ($collected -lt $sessionCap) {
            $raw = $null
            try {
                $raw = Search-UnifiedAuditLog -StartDate $window.Start -EndDate $window.End `
                    -RecordType 'CopilotInteraction' -SessionId $sessionId -SessionCommand ReturnLargeSet `
                    -ResultSize $pageSize -ErrorAction Stop
            }
            catch {
                Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
                    'Search-UnifiedAuditLog is unavailable to this sign-in ({0}). It needs auditing turned on and the Audit Reader / View-Only Audit Logs role. Writing the header only.' -f $_.Exception.Message)
                Export-AppendCsv -Path $csvPath -Column $columns
                return
            }

            # Do not wrap $null in @(): that is a one-element array and looks
            # like data. $null and an empty collection both mean "nothing this
            # call". Retry a few times (the service often returns nothing
            # while the search is prepared), then treat the window as empty.
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
                $auditData = Get-PurviewProperty $record 'AuditData'
                if ($auditData -is [string] -and -not [string]::IsNullOrWhiteSpace($auditData)) {
                    try { $auditData = $auditData | ConvertFrom-Json -ErrorAction Stop }
                    catch {
                        Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
                            'Skipping a record with unparseable AuditData: {0}' -f $_.Exception.Message)
                        continue
                    }
                }
                if ($null -eq $auditData) { continue }

                $creationTime = Get-PurviewProperty $auditData 'CreationTime' (Get-PurviewProperty $record 'CreationDate')
                $creationTimestamp = if ($creationTime) { ConvertTo-CsvTimestamp $creationTime } else { '' }

                $eventData = Get-PurviewProperty $auditData 'CopilotEventData'
                $accessedResources = @(Get-PurviewProperty $eventData 'AccessedResources' @())

                $common = @{
                    RecordId     = [string](Get-PurviewProperty $auditData 'Id' (Get-PurviewProperty $record 'Identity'))
                    CreationTime = $creationTimestamp
                    EventDate    = if ($creationTimestamp) { $creationTimestamp.Substring(0, 10) } else { '' }
                    Operation    = Get-PurviewProperty $auditData 'Operation' (Get-PurviewProperty $record 'Operations')
                    RecordType   = Get-PurviewProperty $record 'RecordType' 'CopilotInteraction'
                    Workload     = Get-PurviewProperty $auditData 'Workload'
                    UserId       = Get-PurviewProperty $auditData 'UserId' (Get-PurviewProperty $record 'UserIds')
                    UserKey      = Get-PurviewProperty $auditData 'UserKey'
                    UserType     = Get-PurviewProperty $auditData 'UserType'
                    AppHost      = Get-PurviewProperty $eventData 'AppHost'
                    AppIdentity  = Get-PurviewProperty $auditData 'AppIdentity'
                    AgentId      = Get-PurviewProperty $auditData 'AgentId'
                    AgentName    = Get-PurviewProperty $auditData 'AgentName'
                    ThreadId     = Get-PurviewProperty $eventData 'ThreadId'
                }

                if ($accessedResources.Count -eq 0) {
                    $windowRows.Add([pscustomobject]$common)
                    continue
                }

                foreach ($resource in $accessedResources) {
                    # PolicyDetails has been seen both as one object and as an
                    # array of them, and an empty array means no policy at all.
                    $policy = Get-PurviewPolicyDetail (Get-PurviewProperty $resource 'PolicyDetails')
                    $status = [string](Get-PurviewProperty $resource 'Status' '')

                    $row = $common.Clone()
                    $row['ResourceId'] = Get-PurviewProperty $resource 'Id' (Get-PurviewProperty $resource 'ID')
                    $row['ResourceName'] = Get-PurviewProperty $resource 'Name'
                    $row['ResourceType'] = Get-PurviewProperty $resource 'Type'
                    $row['ResourceAction'] = Get-PurviewProperty $resource 'Action'
                    $row['SiteUrl'] = Get-PurviewProperty $resource 'SiteUrl'
                    $row['ListItemUniqueId'] = Get-PurviewProperty $resource 'listItemUniqueId' (Get-PurviewProperty $resource 'ListItemUniqueId')
                    $row['SensitivityLabelId'] = Get-PurviewProperty $resource 'SensitivityLabelId'
                    $row['Status'] = $status
                    $row['XpiaDetected'] = Get-PurviewProperty $resource 'XPIADetected'
                    $row['PolicyId'] = $policy.PolicyId
                    $row['PolicyName'] = $policy.PolicyName
                    $row['PolicyRules'] = $policy.Rules
                    # A policy is named only when access was blocked or
                    # restricted; a non-success status says the same thing on
                    # its own.
                    $row['AccessBlocked'] = $policy.HasPolicy -or ($status -and $status -ne 'success')

                    $windowRows.Add([pscustomobject]$row)
                }
            }

            if ($collected -ge $sessionCap) {
                $windowTruncated = $true
                break
            }

            # A page shorter than -ResultSize is not the end of the session. The
            # cmdlet is repeated until it returns nothing, or until ResultCount
            # says every hit is already in hand. Stopping on the first short page
            # drops every later hit in the same window.
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
}
catch {
    Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
        "Copilot audit records are unavailable to this sign-in ({0}). Writing the header only." -f $_.Exception.Message)
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}
finally {
    if ($connectedHere) {
        Disconnect-ExchangeOnline -Confirm:$false
    }
}

$result = Export-AppendCsv -Path $csvPath -Rows $rows.ToArray() -Column $columns -KeyColumn @('RecordId', 'ResourceId') -PassThru
Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'copilot-accessed-resources.csv: {0} rows written, {1} skipped as already collected.' -f $result.Written, $result.Skipped)

if ($null -ne $truncatedWindow) {
    $windowStart = ConvertTo-CsvTimestamp $truncatedWindow.Start
    $windowEnd = ConvertTo-CsvTimestamp $truncatedWindow.End
    Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
        "The window $windowStart to $windowEnd matches more than 50,000 records; ReturnLargeSet is unsorted, so writing it would move the watermark past events that were never returned. Re-run this window with a smaller -WindowHours.")
    throw ('The unified audit log window {0} to {1} exceeded the 50,000-record session cap. Re-run with a smaller -WindowHours.' -f $windowStart, $windowEnd)
}
