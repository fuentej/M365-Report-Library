#Requires -Version 7.0

<#
    The committed sample set is what the later Power BI report will be built against, so
    its shape matters as much as the collectors'.
#>

BeforeAll {
    $script:Root = Join-Path $PSScriptRoot '../../..' | Resolve-Path | Select-Object -ExpandProperty Path
    $script:Report = Join-Path $script:Root 'reports/teams-groups-lifecycle'
    $script:Samples = Join-Path $script:Report 'samples'

    Import-Module (Join-Path $script:Root 'shared/M365ReportLibrary.psm1') -Force

    $script:Schema = Import-PowerShellDataFile -LiteralPath (Join-Path $script:Report 'collectors/TeamsGroupsSchema.psd1')

    $script:Files = @{
        'groups.csv'                   = $script:Schema.Groups
        'group-owners.csv'             = $script:Schema.GroupOwners
        'deleted-groups.csv'           = $script:Schema.DeletedGroups
        'group-lifecycle-policies.csv' = $script:Schema.GroupLifecyclePolicies
        'group-lifecycle-coverage.csv' = $script:Schema.GroupLifecycleCoverage
        'team-activity.csv'            = $script:Schema.TeamActivity
        'group-activity.csv'           = $script:Schema.GroupActivity
        'team-archive-status.csv'      = $script:Schema.TeamArchiveStatus
        'group-creation-events.csv'    = $script:Schema.GroupCreationEvents
        'users.csv'                    = (Get-EntraUserCsvColumn)
    }
}

AfterAll {
    Remove-Module M365ReportLibrary -Force -ErrorAction SilentlyContinue
}

Describe 'Every sample file has its collector''s columns' {
    It '<Name> has the schema''s columns in order' -ForEach @(
        'groups.csv', 'group-owners.csv', 'deleted-groups.csv', 'group-lifecycle-policies.csv', 'group-lifecycle-coverage.csv'
        'team-activity.csv', 'group-activity.csv', 'team-archive-status.csv', 'group-creation-events.csv', 'users.csv'
    ) {
        $path = Join-Path $script:Samples $_
        Test-Path -LiteralPath $path | Should -BeTrue
        (Get-CsvHeaderColumn -Path $path) -join ',' | Should -Be ($script:Files[$_] -join ',')
        @(Import-Csv -LiteralPath $path).Count | Should -BeGreaterThan 0
    }
}

Describe 'The GCC High samples are header-only' {
    It '<Name> holds the header and no rows' -ForEach @(
        @{ Name = 'team-activity.csv'; Key = 'TeamActivity' }
        @{ Name = 'group-activity.csv'; Key = 'GroupActivity' }
    ) {
        $path = Join-Path $script:Samples "gcchigh/$Name"
        (Get-Content -LiteralPath $path).Count | Should -Be 1
        (Get-CsvHeaderColumn -Path $path) -join ',' | Should -Be ($script:Schema[$Key] -join ',')
    }
}

Describe 'The sample data is fake' {
    It 'uses no address outside example.com' {
        $addresses = foreach ($file in Get-ChildItem -LiteralPath $script:Samples -Recurse -Filter '*.csv') {
            foreach ($row in Import-Csv -LiteralPath $file.FullName) {
                foreach ($property in $row.PSObject.Properties) {
                    if ($property.Value -match '@') { $property.Value }
                }
            }
        }
        $outside = @($addresses | Where-Object { $_ -notmatch '@([a-z0-9-]+\.)*example\.(com|onmicrosoft\.com)$' } | Sort-Object -Unique)
        $outside -join '; ' | Should -BeNullOrEmpty
    }
}

Describe 'The sample data covers the cases the report pages ask about' {
    BeforeAll {
        $script:Groups = @(Import-Csv -LiteralPath (Join-Path $script:Samples 'groups.csv'))
        $script:Latest = ($script:Groups.RunDate | Sort-Object -Descending | Select-Object -First 1)
        $script:Owners = @(Import-Csv -LiteralPath (Join-Path $script:Samples 'group-owners.csv') | Where-Object RunDate -EQ $script:Latest)
        $script:Archive = @(Import-Csv -LiteralPath (Join-Path $script:Samples 'team-archive-status.csv') | Where-Object RunDate -EQ $script:Latest)
    }

    It 'holds several snapshots, so a trend is possible' {
        ($script:Groups.RunDate | Sort-Object -Unique).Count | Should -BeGreaterOrEqual 3
    }

    It 'has Teams and groups that are not Teams' {
        @($script:Groups | Where-Object { $_.RunDate -eq $script:Latest -and $_.IsTeam -eq 'True' }).Count | Should -BeGreaterThan 0
        @($script:Groups | Where-Object { $_.RunDate -eq $script:Latest -and $_.IsTeam -eq 'False' }).Count | Should -BeGreaterThan 0
    }

    It 'has ownerless, single-owner and unknown-owner groups' {
        @($script:Owners | Where-Object OwnerListStatus -EQ 'None').Count | Should -BeGreaterThan 0
        @($script:Owners | Where-Object OwnerListStatus -EQ 'Unknown').Count | Should -BeGreaterThan 0
        $single = $script:Owners | Where-Object OwnerListStatus -EQ 'Listed' | Group-Object GroupId | Where-Object Count -EQ 1
        @($single).Count | Should -BeGreaterThan 0
    }

    It 'has archived and active teams' {
        @($script:Archive | Where-Object IsArchived -EQ 'True').Count | Should -BeGreaterThan 0
        @($script:Archive | Where-Object IsArchived -EQ 'False').Count | Should -BeGreaterThan 0
    }

    It 'purges soft-deleted groups 30 days after deletion' {
        foreach ($row in Import-Csv -LiteralPath (Join-Path $script:Samples 'deleted-groups.csv')) {
            ([datetime]$row.PurgeDateTime - [datetime]$row.DeletedDateTime).TotalDays | Should -Be 30
        }
    }

    It 'uses only the two documented creation operations' {
        $operations = (Import-Csv -LiteralPath (Join-Path $script:Samples 'group-creation-events.csv')).Operation | Sort-Object -Unique
        $operations | Should -Be @('AddGroup', 'TeamCreated')
    }
}

Describe 'New-SampleData.ps1 is deterministic' {
    It 'regenerates the committed files byte for byte' {
        $temp = Join-Path ([System.IO.Path]::GetTempPath()) ('teams-groups-samples-' + [guid]::NewGuid().ToString('N'))
        try {
            & (Join-Path $script:Report 'New-SampleData.ps1') -OutputPath $temp

            foreach ($file in Get-ChildItem -LiteralPath $script:Samples -Recurse -Filter '*.csv') {
                $relative = [System.IO.Path]::GetRelativePath($script:Samples, $file.FullName)
                $regenerated = Join-Path $temp $relative
                Test-Path -LiteralPath $regenerated | Should -BeTrue -Because "$relative should be regenerated"
                (Get-FileHash -LiteralPath $regenerated).Hash | Should -Be (Get-FileHash -LiteralPath $file.FullName).Hash -Because $relative
            }
        }
        finally {
            Remove-Item -LiteralPath $temp -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}
