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
        Name and a State. State is normalised by the caller to one of:
          success   - 'success', 'neutral', 'skipped'
          pending   - 'pending', 'queued', 'in_progress'
          failure   - anything else ('failure', 'error', 'timed_out', 'cancelled',
                      'action_required', 'stale', ...)

    .PARAMETER SelfCheckName
        The name of this workflow's own check run, excluded from $Check before it is
        evaluated so the workflow never waits on or counts itself.

    .EXAMPLE
        Get-MergeDecision -CommitMessage "Verdict: Ready to merge`n`nAll good." `
            -Check @(@{ Name = 'Pester'; State = 'success' })
#>

Set-StrictMode -Version Latest

$script:PassingStates = @('success', 'neutral', 'skipped')
$script:PendingStates = @('pending', 'queued', 'in_progress')

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

        [string]$SelfCheckName = 'auto-merge'
    )

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

    $failed = @($others | Where-Object { $script:PassingStates -notcontains $_.State -and $script:PendingStates -notcontains $_.State })
    if ($failed.Count -gt 0) {
        $names = ($failed | ForEach-Object { $_.Name }) -join ', '
        return [pscustomobject]@{
            Decision = 'Skip'
            Reason   = "Ready to merge, but failed: $names."
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

    return [pscustomobject]@{
        Decision = 'Merge'
        Reason   = 'Ready to merge and every other check passed.'
    }
}
