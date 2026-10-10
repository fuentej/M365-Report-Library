#Requires -Version 7.0

BeforeDiscovery {
    $script:Files = @(
        @{ File = 'sites.csv'; Key = 'Sites' }
        @{ File = 'item-permissions.csv'; Key = 'ItemPermissions' }
        @{ File = 'site-permission-breadth.csv'; Key = 'SitePermissionBreadth' }
        @{ File = 'everyone-item-exposure.csv'; Key = 'EveryoneItemExposure' }
        @{ File = 'sharing-link-activity.csv'; Key = 'SharingLinkActivity' }
        @{ File = 'eeeu-activity.csv'; Key = 'EeeuActivity' }
        @{ File = 'labeled-file-sites.csv'; Key = 'LabeledFileSites' }
        @{ File = 'site-sharing-settings.csv'; Key = 'SiteSharingSettings' }
        @{ File = 'anonymous-link-events.csv'; Key = 'AuditEvents' }
        @{ File = 'sharing-events.csv'; Key = 'AuditEvents' }
        @{ File = 'audit-log-status.csv'; Key = 'AuditLogStatus' }
    )
}

BeforeAll {
    $script:Samples = Join-Path $PSScriptRoot '../samples'
    $script:Schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot '../collectors/OversharingSchema.psd1')
    Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')
}

Describe 'The sample CSVs match the schema' {
    It '<File> has the columns the schema lists and holds rows' -ForEach $script:Files {
        $path = Join-Path $script:Samples $File
        (Get-CsvHeaderColumn -Path $path) -join ',' | Should -Be ($script:Schema[$Key] -join ',')
        @(Import-Csv -LiteralPath $path).Count | Should -BeGreaterThan 0
    }

    It 'has one sample per source row of the contract, eleven in all' {
        @(Get-ChildItem -LiteralPath $script:Samples -Filter '*.csv').Count | Should -Be 11
    }

    It 'uses only example.com addresses and no secret sharing URL' {
        $text = (Get-ChildItem -LiteralPath $script:Samples -Recurse -Filter '*.csv' | Get-Content -Raw) -join "`n"
        $domains = [regex]::Matches($text, '[A-Za-z0-9._-]+@([A-Za-z0-9.-]+)') | ForEach-Object { $_.Groups[1].Value }
        ($domains | Where-Object { $_ -ne 'example.com' }) | Should -BeNullOrEmpty
        $text | Should -Not -Match 'microsoftonline|/:[a-z]:/'
    }
}
