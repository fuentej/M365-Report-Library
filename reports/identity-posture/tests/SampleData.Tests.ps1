#Requires -Version 7.0

<#
    The committed sample set is what the later Power BI report will be built against, so
    its shape matters as much as the collectors'.
#>

BeforeAll {
    $script:Root = Join-Path $PSScriptRoot '../../..' | Resolve-Path | Select-Object -ExpandProperty Path
    $script:Report = Join-Path $script:Root 'reports/identity-posture'
    $script:Samples = Join-Path $script:Report 'samples'

    Import-Module (Join-Path $script:Root 'shared/M365ReportLibrary.psm1') -Force

    $script:Schema = Import-PowerShellDataFile -LiteralPath (Join-Path $script:Report 'collectors/IdentityPostureSchema.psd1')

    $script:Files = @{
        'authentication-methods.csv'      = $script:Schema.AuthenticationMethods
        'conditional-access-policies.csv' = $script:Schema.ConditionalAccessPolicies
        'role-assignments-active.csv'     = $script:Schema.ActiveRoleAssignments
        'role-assignments-eligible.csv'   = $script:Schema.EligibleRoleAssignments
        'role-assignments.csv'            = $script:Schema.RoleAssignments
        'user-signin-activity.csv'        = $script:Schema.UserSignInActivity
        'signins.csv'                     = $script:Schema.SignIns
        'risky-users.csv'                 = $script:Schema.RiskyUsers
        'users.csv'                       = (Get-EntraUserCsvColumn)
    }
}

AfterAll {
    Remove-Module M365ReportLibrary -Force -ErrorAction SilentlyContinue
}

Describe 'Every sample file has its collector''s columns' {
    It '<_> has the schema''s columns in order' -ForEach @(
        'authentication-methods.csv', 'conditional-access-policies.csv', 'role-assignments-active.csv'
        'role-assignments-eligible.csv', 'role-assignments.csv', 'user-signin-activity.csv'
        'signins.csv', 'risky-users.csv', 'users.csv'
    ) {
        $path = Join-Path $script:Samples $_
        Test-Path -LiteralPath $path | Should -BeTrue
        (Get-CsvHeaderColumn -Path $path) -join ',' | Should -Be ($script:Files[$_] -join ',')
        @(Import-Csv -LiteralPath $path).Count | Should -BeGreaterThan 0
    }
}

Describe 'The unlicensed samples are header-only' {
    # A tenant without Entra ID P2 or Governance gets a header and a run.log warning
    # from these collectors, not a file of zeros.
    It '<Name> holds the header and no rows' -ForEach @(
        @{ Name = 'role-assignments-active.csv'; Key = 'ActiveRoleAssignments' }
        @{ Name = 'role-assignments-eligible.csv'; Key = 'EligibleRoleAssignments' }
        @{ Name = 'risky-users.csv'; Key = 'RiskyUsers' }
    ) {
        $path = Join-Path $script:Samples "unlicensed/$Name"
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

    It 'uses only documentation addresses for sign-in IPs' {
        # 203.0.113.0/24 is TEST-NET-3 (RFC 5737).
        foreach ($row in Import-Csv -LiteralPath (Join-Path $script:Samples 'signins.csv')) {
            $row.IpAddress | Should -Match '^203\.0\.113\.\d+$'
        }
    }
}

Describe 'The sample data uses the values the Learn pages document' {
    It 'uses only the three Conditional Access states' {
        $states = (Import-Csv -LiteralPath (Join-Path $script:Samples 'conditional-access-policies.csv')).State | Sort-Object -Unique
        $states | Should -Be @('disabled', 'enabled', 'enabledForReportingButNotEnforced')
    }

    It 'uses only documented riskLevel and riskState values' {
        $rows = @(Import-Csv -LiteralPath (Join-Path $script:Samples 'risky-users.csv'))
        $levels = 'low', 'medium', 'high', 'hidden', 'none', 'unknownFutureValue'
        $states = 'none', 'confirmedSafe', 'remediated', 'dismissed', 'atRisk', 'confirmedCompromised', 'unknownFutureValue'
        foreach ($row in $rows) {
            $levels | Should -Contain $row.RiskLevel
            $states | Should -Contain $row.RiskState
        }
    }

    It 'uses only Assigned or Activated, and Inherited, Direct or Group' {
        foreach ($row in Import-Csv -LiteralPath (Join-Path $script:Samples 'role-assignments-active.csv')) {
            'Assigned', 'Activated' | Should -Contain $row.AssignmentType
            'Inherited', 'Direct', 'Group' | Should -Contain $row.MemberType
        }
    }

    It 'uses only the legacy clientAppUsed values' {
        foreach ($row in Import-Csv -LiteralPath (Join-Path $script:Samples 'signins.csv')) {
            $script:Schema.LegacyClientAppValues | Should -Contain $row.ClientAppUsed
        }
    }
}

Describe 'The sample data covers the cases the report pages ask about' {
    BeforeAll {
        $script:Users = @(Import-Csv -LiteralPath (Join-Path $script:Samples 'users.csv'))
        $script:Latest = ($script:Users.RunDate | Sort-Object -Descending | Select-Object -First 1)
        $script:Activity = @(Import-Csv -LiteralPath (Join-Path $script:Samples 'user-signin-activity.csv') | Where-Object RunDate -EQ $script:Latest)
        $script:Registration = @(Import-Csv -LiteralPath (Join-Path $script:Samples 'authentication-methods.csv') | Where-Object RunDate -EQ $script:Latest)
    }

    It 'holds several snapshots, so a trend is possible' {
        ($script:Users.RunDate | Sort-Object -Unique).Count | Should -BeGreaterOrEqual 2
    }

    It 'has members, guests and a disabled account' {
        @($script:Users | Where-Object { $_.RunDate -eq $script:Latest -and $_.UserType -eq 'Guest' }).Count | Should -BeGreaterThan 0
        @($script:Users | Where-Object { $_.RunDate -eq $script:Latest -and $_.AccountEnabled -eq 'False' }).Count | Should -BeGreaterThan 0
    }

    It 'has empty sign-in cells, and never a placeholder date' {
        @($script:Activity | Where-Object { [string]::IsNullOrEmpty($_.LastSignInDateTime) }).Count | Should -BeGreaterThan 0
        @($script:Activity | Where-Object { [string]::IsNullOrEmpty($_.LastNonInteractiveSignInDateTime) }).Count | Should -BeGreaterThan 0
        (Get-Content -LiteralPath (Join-Path $script:Samples 'user-signin-activity.csv') -Raw) | Should -Not -Match '0001-01-01'
    }

    It 'does not list a disabled user in the registration report' {
        $disabled = @($script:Users | Where-Object { $_.RunDate -eq $script:Latest -and $_.AccountEnabled -eq 'False' }).Id
        $script:Registration.UserId | Where-Object { $disabled -contains $_ } | Should -BeNullOrEmpty
    }

    It 'has an administrator without MFA that is fixed by the later snapshot, and a user with no multifactor method' {
        $all = @(Import-Csv -LiteralPath (Join-Path $script:Samples 'authentication-methods.csv'))
        @($all | Where-Object { $_.IsAdmin -eq 'True' -and $_.IsMfaRegistered -eq 'False' -and $_.RunDate -ne $script:Latest }).Count | Should -BeGreaterThan 0
        @($script:Registration | Where-Object { $_.IsMfaRegistered -eq 'False' }).Count | Should -BeGreaterThan 0
    }

    It 'refers only to users that users.csv lists' {
        $ids = $script:Users.Id
        foreach ($name in 'authentication-methods.csv', 'user-signin-activity.csv', 'risky-users.csv') {
            foreach ($row in Import-Csv -LiteralPath (Join-Path $script:Samples $name)) {
                $id = if ($row.PSObject.Properties['UserId']) { $row.UserId } else { $row.Id }
                $ids | Should -Contain $id -Because "$name row $id"
            }
        }
        foreach ($row in Import-Csv -LiteralPath (Join-Path $script:Samples 'role-assignments-active.csv')) {
            $ids | Should -Contain $row.PrincipalId
        }
    }

    It 'has a permanent and a time-bound assignment' {
        $rows = @(Import-Csv -LiteralPath (Join-Path $script:Samples 'role-assignments-active.csv'))
        @($rows | Where-Object { [string]::IsNullOrEmpty($_.EndDateTime) }).Count | Should -BeGreaterThan 0
        @($rows | Where-Object { -not [string]::IsNullOrEmpty($_.EndDateTime) }).Count | Should -BeGreaterThan 0
    }

    It 'has a Conditional Access policy that changed state between snapshots' {
        $rows = @(Import-Csv -LiteralPath (Join-Path $script:Samples 'conditional-access-policies.csv'))
        $changed = $rows | Group-Object Id | Where-Object { @($_.Group.State | Sort-Object -Unique).Count -gt 1 }
        @($changed).Count | Should -BeGreaterThan 0
    }
}
