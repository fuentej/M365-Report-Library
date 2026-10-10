#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes copilot-audit-events.csv: one row per Copilot interaction record in the unified
        audit log, appended from the last exported timestamp.

    .DESCRIPTION
        Source 6 of docs/candidates/copilot-usage.md: Search-UnifiedAuditLog -Operations
        CopilotInteraction in Exchange Online PowerShell
        (https://learn.microsoft.com/powershell/module/exchangepowershell/search-unifiedauditlog;
        record properties: https://learn.microsoft.com/purview/audit-copilot). Each window is
        paged with the same -SessionId and -SessionCommand ReturnLargeSet until the call
        returns nothing. A window that reaches the 50,000-record session cap is incomplete and
        unsorted: it and every later one are not written, an error is logged and the run
        stops, so the resume point never moves past events that were not returned. Re-run it
        with -StartDate, -EndDate and a smaller -WindowHours.

        -Operations CopilotInteraction does not return TeamCopilotInteraction (Facilitator),
        ConnectedAIAppInteraction or AIAppInteraction records. AppIdentity is not a cmdlet
        parameter, so it is read from each record. AgentId and AgentName are kept wherever
        the record places them.

        A record is not one prompt: it typically holds a prompt and a response, and can hold
        one prompt with several responses. MessageCount and PromptMessageCount say how many.
        Message text is not in the record (a message is an id and isPrompt) and is not
        collected. Counts built from these rows are not the official active-user counts and
        must not be set beside them as the same measure.

        Available in Commercial; UNVERIFIED in GCC and GCC High (no page names the cmdlet for
        those clouds), where the collector asks anyway, warns, and logs any refusal. Needs the
        View-Only Audit Logs or Audit Logs role and auditing turned on. Audit (Standard) keeps
        Copilot records for 180 days. Records can arrive late, so a record older than the
        watermark that is ingested after a run is not collected by the next one.

    .EXAMPLE
        ./Get-CopilotAuditEvents.ps1 -OutputPath ./out -LookbackDays 7
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

    # 180 days is the Audit (Standard) retention for Copilot records.
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

. (Join-Path $PSScriptRoot 'CopilotUsageHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'CopilotUsageSchema.psd1')
$columns = [string[]]$schema.CopilotAuditEvents
$source = 'copilot-audit-events'
$csvPath = Join-Path $OutputPath 'copilot-audit-events.csv'

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

if (Test-CopilotSourceSkipped -Source 'CopilotAuditEvents' -LogSource $source -CsvPath $csvPath -Column $columns `
        -OutputPath $OutputPath -Environment $Environment -Schema $schema) {
    return
}

$connectedHere = -not $SkipConnect
if ($connectedHere) {
    Connect-M365Service -Service ExchangeOnline -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization
}

try {
    $watermark = Get-CsvWatermark -Path $csvPath -Column 'CreationTime'

    $start = if ($PSBoundParameters.ContainsKey('StartDate')) { ConvertTo-AuditQueryDate $StartDate }
    elseif ($null -ne $watermark) { ConvertTo-AuditQueryDate $watermark }
    else { ConvertTo-AuditQueryDate ([datetime]::UtcNow.AddDays(-$LookbackDays)) }

    $end = if ($PSBoundParameters.ContainsKey('EndDate')) { ConvertTo-AuditQueryDate $EndDate } else { [datetime]::UtcNow }

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
        'Searching the unified audit log for CopilotInteraction from {0} to {1} in {2}-hour windows (watermark: {3}).' -f
        (ConvertTo-CsvTimestamp $start), (ConvertTo-CsvTimestamp $end), $WindowHours,
        $(if ($null -eq $watermark) { 'none' } else { ConvertTo-CsvTimestamp $watermark }))

    try {
        $found = Invoke-AuditSearch -Start $start -End $end -WindowHours $WindowHours -Operation @('CopilotInteraction') `
            -OutputPath $OutputPath -Source $source
    }
    catch {
        $status = Get-GraphHttpStatus -ErrorRecord $_
        if ($status -eq 429 -or $status -eq 503) {
            # A throttle is not a missing role and not an empty log. Do not write a header
            # and return, which would look like a finished run with nothing to collect.
            # https://learn.microsoft.com/graph/throttling
            Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
                'Search-UnifiedAuditLog is still throttled after retries ({0}). A 429 or 503 is not an empty audit result.' -f $_.Exception.Message)
            throw
        }
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            'Search-UnifiedAuditLog is unavailable to this sign-in or cloud ({0}). It needs auditing turned on and the View-Only Audit Logs role. Writing the header only.' -f $_.Exception.Message)
        Export-AppendCsv -Path $csvPath -Column $columns
        return
    }

    $rows = foreach ($entry in $found.Records) {
        $audit = $entry.Audit
        $event = Get-GraphAdditionalProperty -Object $audit -Name 'CopilotEventData'

        $messages = @(Get-GraphAdditionalProperty -Object $event -Name 'Messages' | Where-Object { $null -ne $_ })
        $prompts = @($messages | Where-Object { [string](Get-GraphAdditionalProperty -Object $_ -Name 'isPrompt') -eq 'True' })
        $resources = @(Get-GraphAdditionalProperty -Object $event -Name 'AccessedResources' | Where-Object { $null -ne $_ })
        $plugins = @(Get-GraphAdditionalProperty -Object $event -Name 'AISystemPlugin' | Where-Object { $null -ne $_ })

        $recordType = if ($entry.RecordType) { $entry.RecordType } else { Get-FirstValue $audit 'RecordType' }

        [pscustomobject]@{
            CreationTime          = ConvertTo-CsvTimestamp (Get-GraphAdditionalProperty -Object $audit -Name 'CreationTime')
            Id                    = Get-FirstValue $audit 'Id'
            UserId                = Get-FirstValue $audit 'UserId'
            Operation             = Get-FirstValue $audit 'Operation'
            RecordType            = $recordType
            Workload              = Get-FirstValue $audit 'Workload'
            AppHost               = Get-FirstValue $event 'AppHost'
            AppIdentity           = $(foreach ($where in $event, $audit) { $v = Get-FirstValue $where 'AppIdentity'; if ($v) { $v; break } })
            AgentId               = $(foreach ($where in $event, $audit) { $v = Get-FirstValue $where 'AgentId'; if ($v) { $v; break } })
            AgentName             = $(foreach ($where in $event, $audit) { $v = Get-FirstValue $where 'AgentName'; if ($v) { $v; break } })
            MessageCount          = $messages.Count
            PromptMessageCount    = $prompts.Count
            AccessedResourceCount = $resources.Count
            PluginIds             = ($plugins | ForEach-Object { Get-FirstValue $_ 'ID' } | Where-Object { $_ }) -join ';'
        }
    }

    $result = Export-AppendCsv -Path $csvPath -Rows @($rows) -Column $columns -KeyColumn 'Id' -PassThru
    Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
        'copilot-audit-events.csv: {0} rows written, {1} skipped as already collected.' -f $result.Written, $result.Skipped)

    if ($null -ne $found.TruncatedWindow) {
        $windowStart = ConvertTo-CsvTimestamp $found.TruncatedWindow.Start
        $windowEnd = ConvertTo-CsvTimestamp $found.TruncatedWindow.End
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            "The window $windowStart to $windowEnd matches 50,000 or more records; ReturnLargeSet is unsorted, so writing it would move the resume point past events that were never returned. Re-run it with -StartDate $windowStart -EndDate $windowEnd and a smaller -WindowHours.")
        throw ('The unified audit log window {0} to {1} reached the 50,000-record session cap. Re-run with -StartDate {0} and a smaller -WindowHours.' -f $windowStart, $windowEnd)
    }
}
finally {
    if ($connectedHere) {
        Disconnect-ExchangeOnline -Confirm:$false
    }
}
