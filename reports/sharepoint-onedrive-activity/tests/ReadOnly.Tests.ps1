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
    It 'calls only Connect, Disconnect, Get and Search commands on SharePoint Online, Exchange Online and Graph' {
        $offenders = $script:Calls | Where-Object {
            $_.Name -and $_.Name -match '^[A-Za-z]+-((EXO|Mg|SPO)[A-Za-z0-9]*|UnifiedAuditLog|ExchangeOnline)$' -and
            $_.Name -ne 'Invoke-MgGraphRequest' -and
            $script:AllowedVerbs -notcontains $_.Name.Substring(0, $_.Name.IndexOf('-'))
        } | ForEach-Object { '{0}: {1}' -f $_.File, $_.Name }
        $offenders -join '; ' | Should -BeNullOrEmpty
    }

    It 'covers the tenant cmdlets it expects' {
        $names = $script:Calls.Name | Sort-Object -Unique
        foreach ($expected in 'Get-SPOSite', 'Get-SPOTenant', 'Search-UnifiedAuditLog', 'Connect-SPOService', 'Get-MgAdminReportSetting',
            'Get-MgReportSharePointSiteUsageDetail', 'Get-MgReportOneDriveUsageAccountDetail',
            'Get-MgReportSharePointSiteUsageStorage', 'Get-MgReportOneDriveUsageStorage',
            'Get-MgReportSharePointActivityUserDetail', 'Get-MgReportOneDriveActivityUserDetail', 'Invoke-MgGraphRequest') {
            $names | Should -Contain $expected
        }
    }

    It 'sends a raw HTTP request only from the one Graph paging helper, and only as GET' {
        $raw = @($script:Calls | Where-Object { $_.Name -in @('Invoke-RestMethod', 'Invoke-WebRequest', 'curl', 'wget') })
        $raw.Count | Should -Be 0

        $graph = @($script:Calls | Where-Object { $_.Name -eq 'Invoke-MgGraphRequest' })
        $graph.Count | Should -Be 1
        $graph[0].File | Should -Be 'SharePointOneDriveHelpers.ps1'
        $methods = @($graph[0].Ast.CommandElements | Where-Object { $_ -is [System.Management.Automation.Language.CommandParameterAst] -and $_.ParameterName -eq 'Method' })
        $methods.Count | Should -Be 1
        $graph[0].Ast.Extent.Text | Should -Match '-Method GET\b'
    }

    It 'passes no -Method other than GET to any command' {
        $offenders = foreach ($call in $script:Calls) {
            foreach ($element in @($call.Ast.CommandElements) | Where-Object { $_ -is [System.Management.Automation.Language.CommandParameterAst] -and $_.ParameterName -eq 'Method' }) {
                if ($call.Ast.Extent.Text -notmatch '-Method GET\b') { '{0}:{1}' -f $call.File, $element.Extent.StartLineNumber }
            }
        }
        $offenders -join '; ' | Should -BeNullOrEmpty
    }

    It 'never changes a tenant setting' {
        $text = (Get-ChildItem -LiteralPath $script:ScanRoot -Filter '*.ps1' | ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw }) -join "`n"
        $text | Should -Not -Match '(Set|New|Remove|Update|Add|Start)-(SPO|Mg)[A-Za-z]+' -Because 'a collector only reads'
    }

    It 'never imports the shared module with -Force' {
        $text = (Get-ChildItem -LiteralPath $script:ScanRoot -Filter '*.ps1' | ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw }) -join "`n"
        $text | Should -Not -Match 'Import-Module[^\r\n]*-Force'
    }
}
