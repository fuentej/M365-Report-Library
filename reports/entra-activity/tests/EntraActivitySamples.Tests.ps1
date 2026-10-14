#Requires -Version 7.0

BeforeAll {
    $script:Root = Join-Path $PSScriptRoot '../../..' | Resolve-Path | Select-Object -ExpandProperty Path
    $script:Samples = Join-Path $script:Root 'reports/entra-activity/samples'
    $script:Interactive = @(Import-Csv -LiteralPath (Join-Path $script:Samples 'signins-interactive.csv'))
    $script:NonInteractive = @(Import-Csv -LiteralPath (Join-Path $script:Samples 'signins-noninteractive.csv'))
    $script:SignIns = @($script:Interactive) + @($script:NonInteractive)
    foreach ($row in $script:SignIns) {
        # The sample people work in US Eastern time, UTC-4 for all of September 2026.
        $local = [DateTimeOffset]::Parse($row.CreatedDateTime).UtcDateTime.AddHours(-4)
        $row | Add-Member -NotePropertyName Local -NotePropertyValue $local -Force
    }
    $script:LegacyClients = 'IMAP4', 'POP3', 'Authenticated SMTP', 'Exchange ActiveSync', 'Other clients'
    $script:SprayIp = '203.0.113.250'
}

Describe 'The sign-in samples behave like a real tenant' {
    It 'holds at least 2,000 sign-ins over 30 days for 40 or more people' {
        $script:Interactive.Count | Should -BeGreaterOrEqual 2000
        $dates = $script:SignIns | ForEach-Object { [DateTimeOffset]::Parse($_.CreatedDateTime).UtcDateTime }
        $span = ($dates | Measure-Object -Maximum).Maximum - ($dates | Measure-Object -Minimum).Minimum
        $span.TotalDays | Should -BeGreaterThan 28
        $span.TotalDays | Should -BeLessOrEqual 30
        @($script:SignIns | Where-Object UserId | Select-Object -ExpandProperty UserId -Unique).Count | Should -BeGreaterOrEqual 40
    }

    It 'has most sign-ins in local business hours on weekdays' {
        $normal = @($script:SignIns | Where-Object IpAddress -NE $script:SprayIp)
        $business = @($normal | Where-Object { $_.Local.DayOfWeek -notin 'Saturday', 'Sunday' -and $_.Local.Hour -ge 8 -and $_.Local.Hour -lt 18 })
        $business.Count / $normal.Count | Should -BeGreaterThan 0.85
    }

    It 'has a few odd-hour sign-ins outside the spray' {
        $odd = @($script:SignIns | Where-Object { $_.IpAddress -ne $script:SprayIp -and $_.Local.Hour -lt 5 })
        $odd.Count | Should -BeGreaterOrEqual 5
        $odd.Count | Should -BeLessThan 40
    }

    It 'has one password spray: one address, failed password, 50 or more accounts' {
        $failures = @($script:Interactive | Where-Object ErrorCode -EQ '50126')
        $byAddress = @($failures | Group-Object IpAddress | ForEach-Object {
                [pscustomobject]@{ Ip = $_.Name; Accounts = @($_.Group.UserPrincipalName | Sort-Object -Unique).Count }
            })
        $sprayers = @($byAddress | Where-Object Accounts -GE 50)
        $sprayers.Count | Should -Be 1
        $sprayers[0].Ip | Should -Be $script:SprayIp
        # Every sign-in from that address is a failure, all inside one hour.
        $fromSprayIp = @($script:SignIns | Where-Object IpAddress -EQ $script:SprayIp)
        @($fromSprayIp | Where-Object ErrorCode -NE '50126').Count | Should -Be 0
        $times = $fromSprayIp.Local | Sort-Object
        ($times[-1] - $times[0]).TotalMinutes | Should -BeLessThan 60
    }

    It 'has legacy protocol sign-ins' {
        $legacy = @($script:SignIns | Where-Object { $_.ClientAppUsed -in $script:LegacyClients })
        $legacy.Count | Should -BeGreaterOrEqual 20
        @($legacy.ClientAppUsed | Sort-Object -Unique).Count | Should -BeGreaterOrEqual 3
        @($script:Interactive | Where-Object { $_.ClientAppUsed -in $script:LegacyClients }).Count | Should -BeGreaterThan 0
        @($script:NonInteractive | Where-Object { $_.ClientAppUsed -in $script:LegacyClients }).Count | Should -BeGreaterThan 0
    }

    It 'never uses error code 1024' {
        @($script:SignIns | Where-Object ErrorCode -EQ '1024').Count | Should -Be 0
    }

    It 'keeps every Conditional Access row pointing at a sign-in that exists' {
        $ids = $script:Interactive.Id
        foreach ($row in Import-Csv -LiteralPath (Join-Path $script:Samples 'signin-conditional-access.csv')) {
            $row.SignInId | Should -BeIn $ids
        }
    }
}
