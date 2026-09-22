#Requires -Version 7.0

<#
    Throwaway test for BRO-251: proves a failing Pester test fails the tests.yml CI job
    (not just leaves it green). This file is never meant to reach main - it lives only on
    a throwaway branch pushed to confirm the red run, then discarded.
#>

Describe 'Deliberate failure to prove CI goes red' {
    It 'fails on purpose' {
        1 | Should -Be 2
    }
}
