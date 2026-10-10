#Requires -Version 7.0

# Static scan of this report's scripts. Nothing here touches a tenant.

BeforeAll {
    $script:ScanRoot = Join-Path $PSScriptRoot '../collectors'
    $script:AllowedVerbs = @('Connect', 'Disconnect', 'Get', 'Search')

    $script:Calls = foreach ($file in Get-ChildItem -LiteralPath $script:ScanRoot -Filter '*.ps1') {
        $tokens = $null
        $errors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$errors)
        $ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.CommandAst] }, $true) | ForEach-Object {
            [pscustomobject]@{ File = $file.Name; Name = $_.GetCommandName(); Ast = $_ }
        }
    }
}

Describe 'Only read-only tenant commands are called' {
    It 'calls only Connect, Disconnect, Get and Search commands on Graph and Exchange Online' {
        $offenders = $script:Calls | Where-Object {
            $_.Name -and $_.Name -match '^[A-Za-z]+-(Mg[A-Za-z0-9]*|UnifiedAuditLog)$' -and
            $_.Name -ne 'Invoke-MgGraphRequest' -and
            $script:AllowedVerbs -notcontains $_.Name.Substring(0, $_.Name.IndexOf('-'))
        } | ForEach-Object { '{0}: {1}' -f $_.File, $_.Name }
        $offenders -join '; ' | Should -BeNullOrEmpty
    }

    It 'covers the tenant cmdlets it expects' {
        $names = $script:Calls.Name | Sort-Object -Unique
        foreach ($expected in 'Get-MgReportTeamUserActivityUserDetail', 'Get-MgReportTeamUserActivityCount',
            'Get-MgReportTeamDeviceUsageUserDetail', 'Get-MgAdminReportSetting', 'Search-UnifiedAuditLog', 'Invoke-MgGraphRequest') {
            $names | Should -Contain $expected
        }
    }

    It 'sends a raw Graph request only from the one Invoke-GraphGet helper, and only as GET' {
        $raw = @($script:Calls | Where-Object { $_.Name -in @('Invoke-RestMethod', 'Invoke-WebRequest', 'curl', 'wget') })
        $raw.Count | Should -Be 0

        $graph = @($script:Calls | Where-Object { $_.Name -eq 'Invoke-MgGraphRequest' })
        $graph.Count | Should -Be 1
        $graph[0].File | Should -Be 'TeamsActivityHelpers.ps1'
        foreach ($call in $graph) {
            $call.Ast.Extent.Text | Should -Match '-Method GET\b' -Because ('{0} line {1}' -f $call.File, $call.Ast.Extent.StartLineNumber)
        }
    }

    It 'passes no -Method other than GET to any command' {
        $offenders = foreach ($call in $script:Calls) {
            foreach ($element in @($call.Ast.CommandElements) | Where-Object { $_ -is [System.Management.Automation.Language.CommandParameterAst] -and $_.ParameterName -eq 'Method' }) {
                if ($call.Ast.Extent.Text -notmatch '-Method GET\b') { '{0}:{1}' -f $call.File, $element.Extent.StartLineNumber }
            }
        }
        $offenders -join '; ' | Should -BeNullOrEmpty
    }

    It 'never reads message bodies or the content of chat and channel messages' {
        $text = (Get-ChildItem -LiteralPath $script:ScanRoot -Filter '*.ps1' | ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw }) -join "`n"
        $text | Should -Not -Match 'getAllMessages'
        $text | Should -Not -Match '\$select=[^'']*(subject|body)'
    }
}
