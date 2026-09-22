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

Describe 'ConvertFrom-GitHubCheckPayload' {
    It 'does not merge when a failed check is past the first page of results' {
        $checkRuns = @'
{"total_count":2,"check_runs":[{"id":1,"name":"Pester","status":"completed","conclusion":"success"},{"id":2,"name":"lint","status":"completed","conclusion":"failure"}]}
'@
        $statuses = '{"total_count":0,"statuses":[]}'

        $checks = @((ConvertFrom-GitHubCheckPayload -CheckRunJson $checkRuns -StatusJson $statuses).Items)
        $result = Get-MergeDecision -CommitMessage $script:ReadyMessage -Check $checks

        $result.Decision | Should -Be 'Skip'
        $result.Reason | Should -Match 'lint'
    }

    It 'refuses a check-run page that is shorter than total_count' {
        $checkRuns = '{"total_count":31,"check_runs":[{"id":1,"name":"Pester","status":"completed","conclusion":"success"}]}'
        $statuses = '{"total_count":0,"statuses":[]}'

        { ConvertFrom-GitHubCheckPayload -CheckRunJson $checkRuns -StatusJson $statuses } |
            Should -Throw '*incomplete*'
    }

    It 'refuses a commit-status page that is shorter than total_count' {
        $checkRuns = '{"total_count":0,"check_runs":[]}'
        $statuses = '{"total_count":2,"statuses":[{"id":9,"context":"ci/local","state":"success"}]}'

        { ConvertFrom-GitHubCheckPayload -CheckRunJson $checkRuns -StatusJson $statuses } |
            Should -Throw '*incomplete*'
    }

    It 'collapses a check run that was returned on two pages and still sees a later failure' {
        $checkRuns = @'
{"total_count":2,"check_runs":[{"id":1,"name":"Pester","status":"completed","conclusion":"success"},{"id":1,"name":"Pester","status":"completed","conclusion":"success"},{"id":2,"name":"lint","status":"completed","conclusion":"failure"}]}
'@
        $statuses = '{"total_count":0,"statuses":[]}'

        $checks = @((ConvertFrom-GitHubCheckPayload -CheckRunJson $checkRuns -StatusJson $statuses).Items)
        @($checks).Count | Should -Be 2
        (Get-MergeDecision -CommitMessage $script:ReadyMessage -Check $checks).Decision | Should -Be 'Skip'
    }

    It 'maps an unfinished check run to its status and a commit status to its state' {
        $checkRuns = '{"total_count":1,"check_runs":[{"id":1,"name":"gate","status":"waiting","conclusion":null}]}'
        $statuses = '{"total_count":1,"statuses":[{"id":4,"context":"ci/jenkins","state":"pending"}]}'

        $checks = @((ConvertFrom-GitHubCheckPayload -CheckRunJson $checkRuns -StatusJson $statuses).Items)
        ($checks | Where-Object { $_.Name -eq 'gate' }).State | Should -Be 'waiting'
        ($checks | Where-Object { $_.Name -eq 'ci/jenkins' }).State | Should -Be 'pending'

        (Get-MergeDecision -CommitMessage $script:ReadyMessage -Check $checks).Decision | Should -Be 'Wait'
    }
}

Describe 'Test-WorkflowRunsOnPush' {
    It 'sees a push trigger written as a block' {
        $yaml = @"
name: deploy
on:
  push:
    branches: [main]
  pull_request:
"@
        Test-WorkflowRunsOnPush -Yaml $yaml | Should -BeTrue
    }

    It 'sees a push trigger written on one line' {
        Test-WorkflowRunsOnPush -Yaml "on: [push, pull_request]`n" | Should -BeTrue
        Test-WorkflowRunsOnPush -Yaml "on: push`n" | Should -BeTrue
    }

    It 'does not treat pull_request and workflow_dispatch as a push trigger' {
        $yaml = @"
name: tests
on:
  pull_request:
  workflow_dispatch:
"@
        Test-WorkflowRunsOnPush -Yaml $yaml | Should -BeFalse
    }

    It 'does not treat the word push inside a job as a trigger' {
        $yaml = @"
name: tests
on:
  pull_request:
jobs:
  pester:
    steps:
      - run: echo push:
"@
        Test-WorkflowRunsOnPush -Yaml $yaml | Should -BeFalse
    }

    It 'reads a quoted on key' {
        $yaml = @"
'on':
  push:
"@
        Test-WorkflowRunsOnPush -Yaml $yaml | Should -BeTrue
    }
}
