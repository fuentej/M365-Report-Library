#Requires -Version 7.0

BeforeAll {
    $script:CandidatePath = Join-Path $PSScriptRoot '../../docs/candidates/identity-posture.md'
    $script:Candidate = Get-Content -LiteralPath $script:CandidatePath -Raw
}

Describe 'identity posture source candidate' {
    It 'exists for BRO-315 and covers every starting question' {
        $script:Candidate | Should -Match 'BRO-315'
        $script:Candidate | Should -Match 'Which users have which authentication methods'
        $script:Candidate | Should -Match 'Which Conditional Access policies exist'
        $script:Candidate | Should -Match 'Who holds privileged directory roles'
        $script:Candidate | Should -Match 'Which accounts are disabled, stale or have never signed in'
        $script:Candidate | Should -Match 'Legacy authentication sign-ins over time'
        $script:Candidate | Should -Match 'Which users Entra flags as risky'
        $pages = $script:Candidate -split '## Proposed report pages', 2
        $pageSection = ($pages[1] -split '## Starting questions', 2)[0]
        $pageCount = @($pageSection -split '\r?\n' | Where-Object { $_ -match '^\| [1-8] \|' }).Count
        $pageCount | Should -BeGreaterOrEqual 5
        $pageCount | Should -BeLessOrEqual 8
    }

    It 'names the signInActivity permission, role, page cap and empty timestamps' {
        $script:Candidate | Should -Not -Match 'Role: UNVERIFIED'
        $script:Candidate | Should -Match 'User\.Read\.All'
        $script:Candidate | Should -Match 'AuditLog\.Read\.All'
        $script:Candidate | Should -Match 'Reports Reader is the least privileged role'
        $script:Candidate | Should -Match 'caps a page at 500'
        $script:Candidate | Should -Match 'not together with any other filterable property'
        $script:Candidate | Should -Match '0001-01-01T00:00:00Z'
        $script:Candidate | Should -Match 'not backfilled'
        $script:Candidate | Should -Match 'before April 2020'
    }

    It 'marks signInActivity available in GCC and unverified as a property in GCC High' {
        $signInRow = ($script:Candidate -split '\r?\n' | Where-Object { $_ -match '^\| 5 \| Last sign-in' })
        $signInRow | Should -Not -BeNullOrEmpty
        $signInRow | Should -Match '\[Available\]\(https://learn\.microsoft\.com/graph/deployments\)'
        $signInRow | Should -Match 'UNVERIFIED'
        $signInRow | Should -Not -Match 'UNVERIFIED \(the page excerpt'
    }

    It 'records the PIM license and the active-assignment API that does not need it' {
        $script:Candidate | Should -Not -Match 'PIM license requirement not found'
        $script:Candidate | Should -Match 'Microsoft Entra ID P2 or Microsoft Entra ID Governance'
        $script:Candidate | Should -Match 'GET /roleManagement/directory/roleAssignments`'
        $script:Candidate | Should -Match 'Get-MgRoleManagementDirectoryRoleAssignment'
        $script:Candidate | Should -Match 'RoleManagement\.Read\.Directory'
        $script:Candidate | Should -Match 'assignmentType` is `Assigned` or `Activated`'
        $script:Candidate | Should -Match 'eligible assignments are removed'
    }

    It 'defines privileged with the documented isPrivileged filter' {
        $script:Candidate | Should -Not -Match 'which was not read in full'
        $script:Candidate | Should -Match 'isPrivileged eq true'
        $script:Candidate | Should -Match 'Get-MgBetaRoleManagementDirectoryRoleDefinition'
    }

    It 'records Conditional Access license, read role and state values' {
        $script:Candidate | Should -Match 'Global Secure Access Administrator'
        $script:Candidate | Should -Match 'enabledForReportingButNotEnforced'
        $script:Candidate | Should -Match 'Microsoft Entra ID P1'
        $script:Candidate | Should -Match 'Microsoft 365 Business Premium'
        $conditionalAccessRow = ($script:Candidate -split '\r?\n' | Where-Object { $_ -match '^\| 3 \|' })
        $conditionalAccessRow | Should -Not -Match 'None named on the page'
    }

    It 'keeps non-interactive legacy sign-ins on the beta filter and pages sign-ins' {
        $script:Candidate | Should -Match 'signInEventTypes/any\(t: t eq ''nonInteractiveUser''\)'
        $script:Candidate | Should -Match 'Get-MgBetaAuditLogSignIn'
        $script:Candidate | Should -Match 'interactive in nature'
        $script:Candidate | Should -Match 'Maximum and default page size is 1,000'
        $script:Candidate | Should -Match 'Policy\.Read\.ConditionalAccess'
        $script:Candidate | Should -Match 'other clients'
    }

    It 'does not treat risky users as licensed by P1' {
        $script:Candidate | Should -Not -Match 'Sources 2, 5, 6 and 7 need Entra ID P1 or P2'
        $script:Candidate | Should -Match 'P1 does not license this API'
        $script:Candidate | Should -Match 'riskLevel` is `low`, `medium`, `high`, `hidden`, `none` or `unknownFutureValue`'
        $script:Candidate | Should -Match 'Maximum page size with `\$top` is 500'
    }

    It 'puts a Learn link on every Available or NotAvailable cell in the source table' {
        $sources = $script:Candidate -split '## Sources', 2
        $sourceSection = ($sources[1] -split 'Consolidated notes:', 2)[0]
        $rows = @($sourceSection -split '\r?\n' | Where-Object { $_ -match '^\| \d' })
        $rows.Count | Should -BeGreaterThan 0
        foreach ($row in $rows) {
            $cells = $row -split '\|' | Select-Object -Skip 1
            foreach ($cell in $cells) {
                if ($cell -match 'Available|NotAvailable') {
                    $cell | Should -Match 'https://learn\.microsoft\.com/'
                }
            }
        }
    }
}
