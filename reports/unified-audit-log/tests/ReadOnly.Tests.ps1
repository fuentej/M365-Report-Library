#Requires -Version 7.0

# Static scan of this report's scripts. Nothing here touches a tenant.
#
# Decision D-007 (2026-10-10) allows exactly two POSTs, both of which start a read and change no
# tenant data: the Graph audit log query create and the Management Activity subscription start.
# This file pins both to their documented paths and to the one function that sends each.

BeforeAll {
    $script:ScanRoot = Join-Path $PSScriptRoot '../collectors'
    $script:AllowedVerbs = @('Connect', 'Disconnect', 'Get', 'Search')

    $script:Defined = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $script:Calls = foreach ($file in Get-ChildItem -LiteralPath $script:ScanRoot -Filter '*.ps1') {
        $tokens = $null
        $errors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$errors)
        foreach ($function in $ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
            [void]$script:Defined.Add($function.Name)
        }
        $ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.CommandAst] }, $true) | ForEach-Object {
            $enclosing = $_.Parent
            while ($enclosing -and $enclosing -isnot [System.Management.Automation.Language.FunctionDefinitionAst]) { $enclosing = $enclosing.Parent }
            [pscustomobject]@{
                File = $file.Name; Name = $_.GetCommandName(); Ast = $_
                Function = if ($enclosing) { $enclosing.Name } else { '' }
            }
        }
    }

    function Get-MethodValue {
        param($Call)
        $elements = @($Call.Ast.CommandElements)
        for ($i = 0; $i -lt $elements.Count - 1; $i++) {
            if ($elements[$i] -is [System.Management.Automation.Language.CommandParameterAst] -and $elements[$i].ParameterName -eq 'Method') {
                return $elements[$i + 1].Extent.Text.Trim("'", '"')
            }
        }
        return $null
    }
}

Describe 'Only read-only tenant commands are called' {
    It 'calls only Connect, Disconnect, Get and Search commands on the tenant' {
        $offenders = $script:Calls | Where-Object {
            $_.Name -and -not $script:Defined.Contains($_.Name) -and $_.Name -ne 'Invoke-MgGraphRequest' -and
            $_.Name -match '^[A-Za-z]+-(UnifiedAuditLog|AdminAuditLogConfig|ExchangeOnline|IPPSSession|Mg[A-Za-z0-9]*|EXO[A-Za-z0-9]*|Mailbox|TransportRule|[A-Za-z]*Retention[A-Za-z]*)$' -and
            $script:AllowedVerbs -notcontains $_.Name.Substring(0, $_.Name.IndexOf('-'))
        } | ForEach-Object { '{0}: {1}' -f $_.File, $_.Name }
        $offenders -join '; ' | Should -BeNullOrEmpty
    }

    It 'covers the tenant cmdlets it expects' {
        $names = $script:Calls.Name | Sort-Object -Unique
        foreach ($expected in 'Search-UnifiedAuditLog', 'Get-AdminAuditLogConfig', 'Get-UnifiedAuditLogRetentionPolicy',
            'Invoke-MgGraphRequest', 'Invoke-WebRequest') {
            $names | Should -Contain $expected
        }
    }

    It 'sends no raw HTTP request other than through Invoke-WebRequest and Invoke-MgGraphRequest' {
        $raw = @($script:Calls | Where-Object { $_.Name -in @('Invoke-RestMethod', 'curl', 'wget', 'Start-BitsTransfer') })
        $raw.Count | Should -Be 0
    }
}

Describe 'The two POSTs' {
    It 'sends one Graph POST, to the audit log query collection, from New-AuditGraphQuery only' {
        $graph = @($script:Calls | Where-Object { $_.Name -eq 'Invoke-MgGraphRequest' })
        $posts = @($graph | Where-Object { (Get-MethodValue $_) -eq 'POST' })
        $posts.Count | Should -Be 1
        $posts[0].Function | Should -Be 'New-AuditGraphQuery'
        $posts[0].Ast.Extent.Text | Should -Match "-Uri '/v1\.0/security/auditLog/queries'"

        $others = @($graph | Where-Object { (Get-MethodValue $_) -ne 'POST' })
        $others.Count | Should -BeGreaterThan 0
        foreach ($call in $others) { Get-MethodValue $call | Should -Be 'GET' }
    }

    It 'sends one Management Activity POST, to subscriptions/start, from Start-AuditActivitySubscription only' {
        $web = @($script:Calls | Where-Object { $_.Name -eq 'Invoke-WebRequest' })
        $posts = @($web | Where-Object { (Get-MethodValue $_) -eq 'Post' })
        $posts.Count | Should -Be 1
        $posts[0].Function | Should -Be 'Start-AuditActivitySubscription'

        $function = ($script:Calls | Where-Object { $_.Function -eq 'Start-AuditActivitySubscription' } | Select-Object -First 1).Ast
        $text = (Get-Content -LiteralPath (Join-Path $script:ScanRoot 'UnifiedAuditLogHelpers.ps1') -Raw)
        $text | Should -Match "'\{0\}/subscriptions/start\?contentType=\{1\}&PublisherIdentifier=\{2\}'"

        $gets = @($web | Where-Object { (Get-MethodValue $_) -ne 'Post' })
        foreach ($call in $gets) {
            Get-MethodValue $call | Should -Be 'Get'
            $call.Function | Should -Be 'Invoke-AuditActivityGet'
        }
    }

    It 'never stops a subscription' {
        $text = (Get-ChildItem -LiteralPath $script:ScanRoot -Filter '*.ps1' | ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw }) -join "`n"
        $text | Should -Not -Match 'subscriptions/stop'
    }

    It 'passes no -Method other than GET or POST to any command' {
        $offenders = foreach ($call in $script:Calls) {
            $method = Get-MethodValue $call
            if ($null -ne $method -and $method -notin 'GET', 'POST', 'Get', 'Post') { '{0}:{1}' -f $call.File, $method }
        }
        $offenders -join '; ' | Should -BeNullOrEmpty
    }
}

Describe 'The token' {
    It 'is read from the SecureString in two request functions and nowhere else' {
        $readers = @($script:Calls | Where-Object { $_.Name -eq 'ConvertFrom-SecureString' } | ForEach-Object { $_.Function } | Sort-Object -Unique)
        $readers | Should -Be @('Invoke-AuditActivityGet', 'Start-AuditActivitySubscription')
    }

    It 'is never named on a logging line' {
        foreach ($file in Get-ChildItem -LiteralPath $script:ScanRoot -Filter '*.ps1') {
            $lines = Get-Content -LiteralPath $file.FullName | Where-Object { $_ -match 'Write-CollectorLog|Write-Verbose|Write-Host|Write-Warning|Write-Output' }
            ($lines -join "`n") | Should -Not -Match 'AccessToken|Bearer' -Because $file.Name
        }
    }
}
