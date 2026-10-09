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
    It 'calls only Connect, Disconnect, Get and Search commands on Exchange Online and Graph' {
        $offenders = $script:Calls | Where-Object {
            $_.Name -and ($_.Name -match '^[A-Za-z]+-(EXO|Mg)' -or $_.Name -match '-(Mailbox|MailboxPermission|RecipientPermission|InboxRule|TransportRule|AcceptedDomain|OrganizationConfig|AdminAuditLogConfig|UnifiedAuditLog|ExchangeOnline)$') -and
            $script:AllowedVerbs -notcontains $_.Name.Substring(0, $_.Name.IndexOf('-'))
        } | ForEach-Object { '{0}: {1}' -f $_.File, $_.Name }
        $offenders -join '; ' | Should -BeNullOrEmpty
    }

    It 'covers the tenant cmdlets it expects' {
        $names = $script:Calls.Name | Sort-Object -Unique
        foreach ($expected in 'Get-EXOMailbox', 'Get-AcceptedDomain', 'Get-InboxRule', 'Get-TransportRule', 'Get-MailboxPermission',
            'Get-RecipientPermission', 'Get-OrganizationConfig', 'Get-AdminAuditLogConfig', 'Search-UnifiedAuditLog',
            'Get-MgOauth2PermissionGrant', 'Get-MgServicePrincipal', 'Get-MgServicePrincipalAppRoleAssignedTo') {
            $names | Should -Contain $expected
        }
    }

    It 'reaches no service through a raw HTTP call' {
        $offenders = $script:Calls | Where-Object { $_.Name -in @('Invoke-RestMethod', 'Invoke-WebRequest', 'Invoke-MgGraphRequest', 'curl', 'wget') } |
            ForEach-Object { '{0}:{1} {2}' -f $_.File, $_.Ast.Extent.StartLineNumber, $_.Name }
        $offenders -join '; ' | Should -BeNullOrEmpty
    }

    It 'passes no -Method other than GET to any command' {
        $offenders = foreach ($call in $script:Calls) {
            $elements = @($call.Ast.CommandElements)
            foreach ($element in $elements | Where-Object { $_ -is [System.Management.Automation.Language.CommandParameterAst] -and $_.ParameterName -eq 'Method' }) {
                '{0}:{1}' -f $call.File, $element.Extent.StartLineNumber
            }
        }
        $offenders -join '; ' | Should -BeNullOrEmpty
    }
}
