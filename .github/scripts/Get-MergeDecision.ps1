#Requires -Version 7.0

<#
    .SYNOPSIS
        Decides whether a pull request is ready to auto-merge.

    .DESCRIPTION
        Pure decision logic for the auto-merge workflow (.github/workflows/auto-merge.yml).
        It never calls GitHub itself - the workflow gathers the head commit message and the
        check runs / commit statuses for that commit and hands them to Get-MergeDecision,
        which returns a Decision of 'Merge', 'Wait' or 'Skip' and a human-readable Reason.

        The workflow's own check (job name $SelfCheckName) is excluded from $Check before
        this function is called; it is also filtered again here as a safety net so a caller
        that forgets never counts the workflow's own report as "another check".

    .PARAMETER CommitMessage
        The full head commit message. Only its first line is examined: it must be exactly
        'Verdict: Ready to merge' (Cursor's review automation writes that line, or
        'Verdict: Not ready', as the first line of an empty commit it pushes with Joshua's
        GitHub account).

    .PARAMETER Check
        Every check run and commit status reported on the head commit, as objects with a
        Name and a State.         State is the check run's status while it is not completed, and its
        conclusion once it is. Commit statuses pass their state through.
        Only 'success' counts as the other check this decision requires. 'pending',
        'queued', 'in_progress', 'waiting' and 'requested' are still running.
        'neutral' and 'skipped' neither pass nor block — a check marked skipped or
        neutral is ignored, the way BlogKB's and AzureFoundry's auto-merge rules
        already treat them. Every other value blocks the merge. At least one check
        must still report 'success': a head commit whose other checks are only
        'neutral' and/or 'skipped' does not merge.

    .PARAMETER SelfCheckName
        The name of this workflow's own check run, excluded from $Check before it is
        evaluated so the workflow never waits on or counts itself.

    .PARAMETER BaseRepositoryFullName
        The full name (owner/repo) of the repository the pull request targets - the
        workflow's own repository.

    .PARAMETER HeadRepositoryFullName
        The full name (owner/repo) of the pull request's head repository
        (pull_request.head.repo.full_name). A pull request whose head repository differs
        from $BaseRepositoryFullName is from a fork and is never merged, regardless of
        verdict or checks. An empty string means the head repository is missing (a fork
        that was since deleted) and is treated the same way.

    .EXAMPLE
        Get-MergeDecision -CommitMessage "Verdict: Ready to merge`n`nAll good." `
            -Check @(@{ Name = 'Pester'; State = 'success' }) `
            -BaseRepositoryFullName 'owner/repo' -HeadRepositoryFullName 'owner/repo'
#>

Set-StrictMode -Version Latest

$script:PassingStates = @('success')
$script:PendingStates = @('pending', 'queued', 'in_progress', 'waiting', 'requested')
$script:NonBlockingStates = @('success', 'neutral', 'skipped')

function Get-MergeDecision {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$CommitMessage,

        [Parameter(Mandatory)]
        [AllowNull()]
        [AllowEmptyCollection()]
        [object[]]$Check,

        [Parameter(Mandatory)]
        [string]$BaseRepositoryFullName,

        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$HeadRepositoryFullName,

        [string]$SelfCheckName = 'auto-merge'
    )

    if ([string]::IsNullOrEmpty($HeadRepositoryFullName)) {
        return [pscustomobject]@{
            Decision = 'Skip'
            Reason   = 'The pull request has no head repository (a deleted fork); refusing to merge.'
        }
    }

    if ($HeadRepositoryFullName -cne $BaseRepositoryFullName) {
        return [pscustomobject]@{
            Decision = 'Skip'
            Reason   = "The pull request's head repository is '$HeadRepositoryFullName', not '$BaseRepositoryFullName'; refusing to merge a fork."
        }
    }

    $firstLine = ($CommitMessage -split "`r?`n", 2)[0]

    if ($firstLine -cne 'Verdict: Ready to merge') {
        return [pscustomobject]@{
            Decision = 'Skip'
            Reason   = "The head commit's first line is '$firstLine', not 'Verdict: Ready to merge'."
        }
    }

    $others = @(@($Check) | Where-Object {
            $null -ne $_ -and $_.Name -ne $SelfCheckName
        })

    if ($others.Count -eq 0) {
        return [pscustomobject]@{
            Decision = 'Skip'
            Reason   = 'Ready to merge, but no check run or commit status other than this workflow''s own reported on the head commit.'
        }
    }

    $failed = @($others | Where-Object { $script:NonBlockingStates -notcontains $_.State -and $script:PendingStates -notcontains $_.State })
    if ($failed.Count -gt 0) {
        $names = ($failed | ForEach-Object { $_.Name }) -join ', '
        return [pscustomobject]@{
            Decision = 'Skip'
            Reason   = "Ready to merge, but not success: $names."
        }
    }

    $pending = @($others | Where-Object { $script:PendingStates -contains $_.State })
    if ($pending.Count -gt 0) {
        $names = ($pending | ForEach-Object { $_.Name }) -join ', '
        return [pscustomobject]@{
            Decision = 'Wait'
            Reason   = "Ready to merge, waiting on: $names."
        }
    }

    $succeeded = @($others | Where-Object { $script:PassingStates -contains $_.State })
    if ($succeeded.Count -eq 0) {
        return [pscustomobject]@{
            Decision = 'Skip'
            Reason   = 'Ready to merge, but no check run or commit status other than this workflow''s own actually succeeded (neutral and skipped do not count).'
        }
    }

    # skipped and neutral are ignored, so the all-clear reason is only true when
    # every remaining check actually succeeded.
    $ignored = @($others | Where-Object { $_.State -eq 'neutral' -or $_.State -eq 'skipped' })
    if ($ignored.Count -gt 0) {
        $names = ($ignored | ForEach-Object { $_.Name }) -join ', '
        return [pscustomobject]@{
            Decision = 'Merge'
            Reason   = "Ready to merge: at least one other check succeeded. Ignored neutral or skipped checks: $names."
        }
    }

    return [pscustomobject]@{
        Decision = 'Merge'
        Reason   = 'Ready to merge and every other check passed.'
    }
}

function Get-CompleteGitHubPageItems {
    <#
        .SYNOPSIS
            The items from one already-concatenated GitHub list response.

        .DESCRIPTION
            List check runs and the combined commit status both default to 30
            results per page and both return total_count for the full set. A
            caller hands this the concatenated pages. Duplicate ids, from a
            page that was requested twice, are collapsed. The result is thrown
            away when the distinct count does not equal total_count, so a
            decision cannot treat a missing page as "no more checks".
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Json,

        [Parameter(Mandatory)]
        [string]$ItemsProperty,

        [Parameter(Mandatory)]
        [string]$Label
    )

    if ([string]::IsNullOrWhiteSpace($Json)) {
        throw "The $Label payload was empty."
    }

    $document = $Json | ConvertFrom-Json
    if ($null -eq $document.total_count) {
        throw "The $Label payload has no total_count."
    }

    $items = @($document.$ItemsProperty | Where-Object { $null -ne $_ })
    $unique = [System.Collections.Generic.List[object]]::new()
    $seen = @{}
    foreach ($item in $items) {
        $id = [string]$item.id
        if ($seen.ContainsKey($id)) { continue }
        $seen[$id] = $true
        $unique.Add($item)
    }

    if ($unique.Count -ne [int]$document.total_count) {
        throw "The $Label payload is incomplete: total_count is $($document.total_count) but $($unique.Count) distinct results were returned."
    }

    # A bare array returned from a function is enumerated. Hand back one object
    # so the caller can read .Items without losing a page of checks.
    return [pscustomobject]@{
        Items = $unique.ToArray()
    }
}

function ConvertFrom-GitHubCheckPayload {
    <#
        .SYNOPSIS
            Normalises check-run and combined-status documents into Name/State pairs.

        .DESCRIPTION
            Check runs contribute their status until they complete, then their
            conclusion. Commit statuses contribute their state. Both documents
            must be complete pages; see Get-CompleteGitHubPageItems.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$CheckRunJson,

        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$StatusJson
    )

    $runs = @(
        (Get-CompleteGitHubPageItems -Json $CheckRunJson -ItemsProperty 'check_runs' -Label 'check runs').Items
    )
    $statuses = @(
        (Get-CompleteGitHubPageItems -Json $StatusJson -ItemsProperty 'statuses' -Label 'commit statuses').Items
    )

    $checks = [System.Collections.Generic.List[object]]::new()
    foreach ($run in $runs) {
        if ($null -eq $run) { continue }
        $state = if ([string]$run.status -ne 'completed') { [string]$run.status } else { [string]$run.conclusion }
        $checks.Add([pscustomobject]@{ Name = [string]$run.name; State = $state })
    }
    foreach ($status in $statuses) {
        if ($null -eq $status) { continue }
        $checks.Add([pscustomobject]@{ Name = [string]$status.context; State = [string]$status.state })
    }

    return [pscustomobject]@{
        Items = $checks.ToArray()
    }
}

function Test-WorkflowRunsOnPush {
    <#
        .SYNOPSIS
            Whether a workflow file's trigger block includes push.

        .DESCRIPTION
            Reads the workflow text. The top-level key is on, including the
            quoted forms. A push mentioned only inside a job is not a trigger.
            Parser failures are the caller's to surface; this returns false
            only when the trigger block was read and does not list push.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Yaml
    )

    $inTrigger = $false
    foreach ($line in ($Yaml -split "`r?`n")) {
        if ($line -match '^\s*(#|$)') { continue }

        $isTopLevel = $line -match '^[^ \t#]'
        if ($isTopLevel) {
            if ($inTrigger) { return $false }

            if ($line -match '^(?:"on"|''on''|on)\s*:\s*(.*)$') {
                $rest = ($Matches[1] -replace '\s+#.*$', '').Trim()
                if ([string]::IsNullOrWhiteSpace($rest) -or $rest -eq '|' -or $rest -eq '>') {
                    $inTrigger = $true
                    continue
                }

                return [bool]($rest -cmatch '(^|[\s\[''"])push([\s\]''",]|$)')
            }

            continue
        }

        if (-not $inTrigger) { continue }
        if ($line -cmatch '^\s+-\s+[''"]?push[''"]?\s*(#.*)?$') { return $true }
        if ($line -cmatch '^\s+[''"]?push[''"]?\s*:') { return $true }
    }

    return $false
}
