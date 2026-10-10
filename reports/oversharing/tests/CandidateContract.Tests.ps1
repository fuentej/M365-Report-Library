#Requires -Version 7.0

BeforeAll {
    $script:Doc = Get-Content -LiteralPath (Join-Path $PSScriptRoot '../../../docs/candidates/oversharing.md') -Raw
    $script:Schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot '../collectors/OversharingSchema.psd1')
}

Describe 'The collectors follow the contract doc' {
    It 'has a collector for each of the eleven source rows' {
        $script:Schema.SourceAvailability.Keys.Count | Should -Be 11
        foreach ($script in 'Get-Sites', 'Get-ItemSharingPermissions', 'Get-SitePermissionBreadth', 'Get-EveryoneItemExposure', 'Get-SharingLinkActivity',
            'Get-EeeuActivity', 'Get-LabeledFileSites', 'Get-SiteSharingSettings', 'Get-AnonymousLinkEvents', 'Get-SharingEvents', 'Get-AuditLogStatus') {
            Test-Path -LiteralPath (Join-Path $PSScriptRoot "../collectors/$script.ps1") | Should -BeTrue
        }
    }

    It 'copies the audit operations listed in the doc' {
        foreach ($operation in $script:Schema.AnonymousLinkOperations + $script:Schema.SharingOperations) {
            $script:Doc | Should -Match ([regex]::Escape($operation))
        }
        $script:Schema.AnonymousLinkOperations.Count | Should -Be 4
        foreach ($operation in 'AnonymousLinkCreated', 'AnonymousLinkUpdated', 'AnonymousLinkUsed', 'AnonymousLinkRemoved') {
            $script:Schema.AnonymousLinkOperations | Should -Contain $operation
        }
        foreach ($operation in 'CompanyLinkCreated', 'SecureLinkCreated', 'SharingSet', 'SharingInvitationCreated', 'SharingInheritanceBroken', 'AddedToGroup') {
            $script:Schema.SharingOperations | Should -Contain $operation
        }
    }

    It 'reads availability from the doc: rows 3b, 4, 5c are UNVERIFIED where the doc says so, and nothing is NotAvailable' {
        $row3b = @($script:Doc -split "`n" | Where-Object { $_ -match '^\| 3b \|' })[0]
        $cells = $row3b -split '\|'
        $cells[-3].Trim() | Should -Match '^UNVERIFIED'
        $cells[-2].Trim() | Should -Match '^UNVERIFIED'
        $script:Schema.SourceAvailability.EveryoneItemExposure.Commercial.Status | Should -Be 'Available'
        $script:Schema.SourceAvailability.EveryoneItemExposure.GCC.Status | Should -Be 'Unverified'
        $script:Schema.SourceAvailability.EveryoneItemExposure.GCCHigh.Status | Should -Be 'Unverified'
        foreach ($cloud in 'Commercial', 'GCC', 'GCCHigh') {
            $script:Schema.SourceAvailability.SiteSharingSettings[$cloud].Status | Should -Be 'Unverified'
        }
        $script:Doc | Should -Not -Match '\[NotAvailable\]'
    }
}
