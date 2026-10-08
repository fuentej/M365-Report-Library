#Requires -Version 7.0

BeforeAll {
    $script:CandidatePath = Join-Path $PSScriptRoot '../../docs/candidates/teams-groups-lifecycle.md'
    $script:Candidate = Get-Content -LiteralPath $script:CandidatePath -Raw
}

Describe 'Teams and Groups lifecycle source list' {
    It 'reads archived state from Get team, which returns isArchived' {
        $script:Candidate | Should -Match 'GET /teams/\{team-id\}'
        $script:Candidate | Should -Match 'isArchived'
        $script:Candidate | Should -Not -Match 'how to read it was not verified'
        $script:Candidate | Should -Match 'Team\.ReadBasic\.All'
        $script:Candidate | Should -Match 'TeamsGCCH'
        $script:Candidate | Should -Match 'only teams the caller owns or belongs to'
    }

    It 'uses Group.Read.All to read group properties' {
        $script:Candidate | Should -Match '`Group\.Read\.All` reads group properties'
        $script:Candidate | Should -Match 'disableNesting'
        $script:Candidate | Should -Not -Match 'least privileged read permission is UNVERIFIED'
    }

    It 'records the list-owners permission and an Entra role from that page' {
        $script:Candidate | Should -Match 'GroupMember\.Read\.All'
        $script:Candidate | Should -Match 'Directory Readers'
        $script:Candidate | Should -Not -Match 'Entra role: not read'
    }

    It 'limits the 30-day restore window to Microsoft 365 and security groups' {
        $script:Candidate | Should -Match 'Distribution groups are permanently deleted immediately'
        $script:Candidate | Should -Match 'securityEnabled'
    }

    It 'records TeamCreated separately from AddGroup' {
        $script:Candidate | Should -Match 'TeamCreated'
        $script:Candidate | Should -Match 'AddGroup'
        $script:Candidate | Should -Match 'Teams workload'
    }

    It 'records the directory audit permission and Reports Reader role' {
        $script:Candidate | Should -Match 'AuditLog\.Read\.All'
        $script:Candidate | Should -Match 'Reports Reader'
        $script:Candidate | Should -Not -Match 'Entra directory audits: not verified here'
    }

    It 'does not treat resourceProvisioningOptions as present on every team' {
        $script:Candidate | Should -Match 'unused old teams do not have that value set'
    }

    It 'records the expiration license and the v1.0 expirationDateTime property' {
        $script:Candidate | Should -Match 'Microsoft Entra ID P1 or P2'
        $script:Candidate | Should -Match 'expirationDateTime'
        $script:Candidate | Should -Not -Match 'License needed for the policy: UNVERIFIED'
        $script:Candidate | Should -Not -Match 'graph/templates/terraform'
    }

    It 'records both documented guest-count column names' {
        $script:Candidate | Should -Match 'External Member Count'
        $script:Candidate | Should -Match 'Guest Count'
        $script:Candidate | Should -Match 'Owner Principal Name'
    }

    It 'records page size, the unbounded last activity date, and detail-row roles' {
        $script:Candidate | Should -Match '@odata\.nextLink'
        $script:Candidate | Should -Match '100 objects by default'
        $script:Candidate | Should -Match 'regardless of the D7/D30/D90/D180 window'
        $script:Candidate | Should -Match 'Global Reader and Usage Summary Reports Reader'
        $script:Candidate | Should -Not -Match 'audit records are the fallback'
    }
}
