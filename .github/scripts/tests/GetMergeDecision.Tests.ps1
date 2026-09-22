#Requires -Version 7.0

BeforeAll {
    $script:ScriptPath = Join-Path $PSScriptRoot '../Get-MergeDecision.ps1'
    . $script:ScriptPath

    $script:ReadyMessage = "Verdict: Ready to merge`n`nEverything checked out."
    $script:NotReadyMessage = "Verdict: Not ready`n`nStill has open comments."
}

Describe 'Get-MergeDecision' {
    It 'merges when the verdict is ready and every other check passed' {
        $result = Get-MergeDecision -CommitMessage $script:ReadyMessage -Check @(
            @{ Name = 'Pester'; State = 'success' }
            @{ Name = 'lint'; State = 'success' }
        )

        $result.Decision | Should -Be 'Merge'
    }

    It 'does not merge when the verdict is Not ready' {
        $result = Get-MergeDecision -CommitMessage $script:NotReadyMessage -Check @(
            @{ Name = 'Pester'; State = 'success' }
        )

        $result.Decision | Should -Not -Be 'Merge'
        $result.Decision | Should -Be 'Skip'
    }

    It 'does not merge when the verdict is ready but a check failed' {
        $result = Get-MergeDecision -CommitMessage $script:ReadyMessage -Check @(
            @{ Name = 'Pester'; State = 'failure' }
            @{ Name = 'lint'; State = 'success' }
        )

        $result.Decision | Should -Not -Be 'Merge'
        $result.Decision | Should -Be 'Skip'
    }

    It 'does not merge when the verdict is ready but there are no other checks' {
        $result = Get-MergeDecision -CommitMessage $script:ReadyMessage -Check @()

        $result.Decision | Should -Not -Be 'Merge'
        $result.Decision | Should -Be 'Skip'
    }

    It 'waits when the verdict is ready but a check is still pending' {
        $result = Get-MergeDecision -CommitMessage $script:ReadyMessage -Check @(
            @{ Name = 'Pester'; State = 'pending' }
            @{ Name = 'lint'; State = 'success' }
        )

        $result.Decision | Should -Be 'Wait'
    }

    It 'ignores its own check when deciding there are no other checks' {
        $result = Get-MergeDecision -CommitMessage $script:ReadyMessage -Check @(
            @{ Name = 'auto-merge'; State = 'in_progress' }
        )

        $result.Decision | Should -Be 'Skip'
        $result.Reason | Should -Match 'no check run or commit status other than'
    }

    It 'ignores a differently-named self check when given -SelfCheckName' {
        $result = Get-MergeDecision -CommitMessage $script:ReadyMessage -Check @(
            @{ Name = 'my-auto-merge'; State = 'in_progress' }
            @{ Name = 'Pester'; State = 'success' }
        ) -SelfCheckName 'my-auto-merge'

        $result.Decision | Should -Be 'Merge'
    }

    It 'treats a queued check as pending, not failed' {
        $result = Get-MergeDecision -CommitMessage $script:ReadyMessage -Check @(
            @{ Name = 'Pester'; State = 'queued' }
        )

        $result.Decision | Should -Be 'Wait'
    }

    It 'does not merge when the only other check was skipped' {
        $result = Get-MergeDecision -CommitMessage $script:ReadyMessage -Check @(
            @{ Name = 'optional-job'; State = 'skipped' }
        )

        $result.Decision | Should -Be 'Skip'
    }

    It 'does not merge when another check completed neutral' {
        $result = Get-MergeDecision -CommitMessage $script:ReadyMessage -Check @(
            @{ Name = 'Pester'; State = 'success' }
            @{ Name = 'review'; State = 'neutral' }
        )

        $result.Decision | Should -Be 'Skip'
        $result.Reason | Should -Match 'not success: review'
    }

    It 'waits when a check run is waiting rather than treating it as a failure' {
        $result = Get-MergeDecision -CommitMessage $script:ReadyMessage -Check @(
            @{ Name = 'gate'; State = 'waiting' }
            @{ Name = 'Pester'; State = 'success' }
        )

        $result.Decision | Should -Be 'Wait'
    }

    It 'waits when a check run has been requested but has not started' {
        $result = Get-MergeDecision -CommitMessage $script:ReadyMessage -Check @(
            @{ Name = 'Pester'; State = 'requested' }
        )

        $result.Decision | Should -Be 'Wait'
    }

    It 'requires the first line to match exactly, not just contain the verdict' {
        $result = Get-MergeDecision -CommitMessage "Some prefix. Verdict: Ready to merge" -Check @(
            @{ Name = 'Pester'; State = 'success' }
        )

        $result.Decision | Should -Be 'Skip'
    }

    It 'is case-sensitive about the verdict line' {
        $result = Get-MergeDecision -CommitMessage "verdict: ready to merge" -Check @(
            @{ Name = 'Pester'; State = 'success' }
        )

        $result.Decision | Should -Be 'Skip'
    }
}
