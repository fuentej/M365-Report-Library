#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes copilot-interactions.csv: one row of metadata per Copilot interaction (prompt or
        response) from the interaction export API, appended per user from the last exported
        timestamp. The prompt and response text is never stored.

    .DESCRIPTION
        Source 8 of docs/candidates/copilot-usage.md:
        GET /copilot/users/{id}/interactionHistory/getAllEnterpriseInteractions, v1.0
        (https://learn.microsoft.com/microsoft-365-copilot/extensibility/api/ai-services/interaction-export/aiinteractionhistory-getallenterpriseinteractions),
        one call per user. The users come from users.csv (the shared Entra users collector;
        run it first) or -UserId.

        Only the interaction's id, session and request ids, appClass, interactionType,
        conversationType, createdDateTime, locale and context types are read. body,
        attachments, links and mentions carry prompt and response content and are not read
        into the file, because no proposed page needs them. This is the most sensitive source
        in the report: the application permission AiEnterpriseInteraction.Read.All lets the
        caller read every user's prompts.

        The createdDateTime filter carries both a lower and an upper bound, as the page
        requires. $top is 100, the recommended size. @odata.nextLink is followed when a
        response carries one. The export supports six appClass values (Word, Excel, Teams,
        BizChat, WebChat and CoworkChat); Outlook, PowerPoint, OneNote and Loop are not in it,
        so this is not the per-app count of the usage reports. It does not retrieve
        interactions in Copilot Studio agents, and it also returns interactions for deleted
        users and deleted interactions. Delta is not supported.

        A 429 waits Retry-After and is retried; if it persists the run stops, logs that a 429
        is not an empty history, and keeps what it has. A user the API refuses (for example
        one with no Copilot licence) is logged and skipped; the next user is read. The
        resume point is per user, so a skipped user's history is not passed over.

        Available in all three clouds. Needs the application permission
        AiEnterpriseInteraction.Read.All (delegated is not supported) and, for each user, a
        Microsoft 365 Copilot licence with the "Microsoft Copilot with Graph-grounded chat"
        service plan.

    .EXAMPLE
        ./Get-CopilotInteractions.ps1 -OutputPath ./out -LookbackDays 7
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

    # Read these users instead of the users in users.csv.
    [string[]]$UserId,

    [datetime]$StartDate,
    [datetime]$EndDate,

    # The page states no retention for this API.
    [ValidateRange(1, 365)]
    [int]$LookbackDays = 7,

    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Imported by path without -Force: see Get-CopilotUsageUserDetail.ps1.
Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

. (Join-Path $PSScriptRoot 'CopilotUsageHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'CopilotUsageSchema.psd1')
$columns = [string[]]$schema.CopilotInteractions
$source = 'copilot-interactions'
$csvPath = Join-Path $OutputPath 'copilot-interactions.csv'
$usersPath = Join-Path $OutputPath 'users.csv'

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

if (Test-CopilotSourceSkipped -Source 'CopilotInteractions' -LogSource $source -CsvPath $csvPath -Column $columns `
        -OutputPath $OutputPath -Environment $Environment -Schema $schema) {
    return
}

$users = @(if ($UserId) { $UserId }
    else { Get-CsvLatestSnapshot -Path $usersPath | ForEach-Object { $_.Id } | Where-Object { $_ } | Sort-Object -Unique })

if ($users.Count -eq 0) {
    Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
        'users.csv holds no users and no -UserId was given, so there is nobody to read interactions for. Run the shared users collector first. Writing the header only.')
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}

if (-not $SkipConnect) {
    Connect-M365Service -Service Graph -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization `
        -Scopes @('AiEnterpriseInteraction.Read.All')
}

$end = if ($PSBoundParameters.ContainsKey('EndDate')) { $EndDate.ToUniversalTime() } else { [datetime]::UtcNow }
$explicitStart = $PSBoundParameters.ContainsKey('StartDate')
$firstStart = if ($explicitStart) { $StartDate.ToUniversalTime() } else { $end.AddDays(-$LookbackDays) }

if ($explicitStart -and $end -le $firstStart) {
    throw ('The requested range is empty: the end ({0}) is not later than the start ({1}).' -f
        (ConvertTo-CsvTimestamp $end), (ConvertTo-CsvTimestamp $firstStart))
}

# The resume point is per user: one user's history that could not be read must not be
# passed over because another user's was.
$latestByUser = @{}
if (Test-Path -LiteralPath $csvPath) {
    foreach ($existing in @(Import-Csv -LiteralPath $csvPath)) {
        if ([string]::IsNullOrWhiteSpace($existing.CreatedDateTime)) { continue }
        $stamp = (ConvertTo-CsvTimestamp $existing.CreatedDateTime)
        if (-not $latestByUser.ContainsKey($existing.UserId) -or $stamp -gt $latestByUser[$existing.UserId]) {
            $latestByUser[$existing.UserId] = $stamp
        }
    }
}

Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'Reading interactions for {0} users up to {1} (first start: {2}).' -f
    $users.Count, (ConvertTo-CsvTimestamp $end), (ConvertTo-CsvTimestamp $firstStart))

$rows = [System.Collections.Generic.List[object]]::new()
$failed = 0
$stopped = $null

foreach ($user in $users) {
    $from = if (-not $explicitStart -and $latestByUser.ContainsKey($user)) { $latestByUser[$user] } else { ConvertTo-CsvTimestamp $firstStart }
    $filter = 'createdDateTime gt {0} and createdDateTime lt {1}' -f $from, (ConvertTo-CsvTimestamp $end)
    $uri = 'v1.0/copilot/users/{0}/interactionHistory/getAllEnterpriseInteractions?$top=100&$filter={1}' -f
        [uri]::EscapeDataString($user), [uri]::EscapeDataString($filter)

    try {
        $items = @(Get-GraphPagedValue -Uri $uri -OutputPath $OutputPath -LogSource $source)
    }
    catch {
        $status = Get-GraphHttpStatus -ErrorRecord $_
        if ($status -eq 429 -or $status -eq 503) {
            $stopped = $_.Exception.Message
            Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
                'Still throttled after retries ({0}). A 429 or 503 is not an empty history. Stopping; the users read so far are kept and the next run resumes per user.' -f $stopped)
            break
        }
        $failed++
        Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
            'Skipping user {0}: the interaction export refused or failed ({1}). It needs AiEnterpriseInteraction.Read.All as an application permission and a Microsoft 365 Copilot licence for the user.' -f $user, $_.Exception.Message)
        continue
    }

    foreach ($item in $items) {
        # Only metadata is read. body, attachments, links and mentions are left alone.
        $contexts = @(Get-GraphJsonValue -Object $item -Name 'contexts' | Where-Object { $null -ne $_ })
        $rows.Add([pscustomobject]@{
                CreatedDateTime  = ConvertTo-CsvTimestamp (Get-GraphJsonValue -Object $item -Name 'createdDateTime')
                Id               = [string](Get-GraphJsonValue -Object $item -Name 'id')
                UserId           = $user
                SessionId        = [string](Get-GraphJsonValue -Object $item -Name 'sessionId')
                RequestId        = [string](Get-GraphJsonValue -Object $item -Name 'requestId')
                AppClass         = [string](Get-GraphJsonValue -Object $item -Name 'appClass')
                InteractionType  = [string](Get-GraphJsonValue -Object $item -Name 'interactionType')
                ConversationType = [string](Get-GraphJsonValue -Object $item -Name 'conversationType')
                Locale           = [string](Get-GraphJsonValue -Object $item -Name 'locale')
                ContextCount     = $contexts.Count
                ContextTypes     = ($contexts | ForEach-Object { [string](Get-GraphJsonValue -Object $_ -Name 'contextType') } | Where-Object { $_ }) -join ';'
            })
    }
}

$result = Export-AppendCsv -Path $csvPath -Rows $rows.ToArray() -Column $columns -KeyColumn 'UserId', 'Id' -PassThru
Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'copilot-interactions.csv: {0} rows written, {1} skipped as already collected, {2} users skipped.' -f $result.Written, $result.Skipped, $failed)

if ($null -ne $stopped) {
    throw "Reading Copilot interactions stopped because Graph kept throttling ($stopped). Re-run to resume."
}
if ($failed -eq $users.Count) {
    throw "The interaction export refused or failed for every one of the $($users.Count) users. See run.log."
}
