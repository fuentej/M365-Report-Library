#Requires -Version 7.0

# Static scan of this report's scripts. Nothing here touches a tenant.

BeforeDiscovery {
    $script:ScanRoot = Join-Path $PSScriptRoot '../collectors'
}

BeforeAll {
    $script:ScanRoot = Join-Path $PSScriptRoot '../collectors'
    $script:AllowedVerbs = @('Connect', 'Disconnect', 'Get', 'Search')
    # Az and Exchange cmdlets this report calls; any other tenant cmdlet must be added
    # here on purpose, with the reason.
    $script:TenantCmdlets = @(
        'Connect-AzAccount', 'Get-AzAccessToken', 'Connect-M365Service', 'Disconnect-ExchangeOnline', 'Search-UnifiedAuditLog'
    )

    $script:Calls = foreach ($file in Get-ChildItem -LiteralPath $script:ScanRoot -Filter '*.ps1') {
        $tokens = $null
        $errors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$errors)
        $ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.CommandAst] }, $true) | ForEach-Object {
            [pscustomobject]@{
                File     = $file.Name
                Name     = $_.GetCommandName()
                Ast      = $_
                Function = $(
                    $parent = $_.Parent
                    while ($parent -and $parent -isnot [System.Management.Automation.Language.FunctionDefinitionAst]) { $parent = $parent.Parent }
                    if ($parent) { $parent.Name } else { '' }
                )
            }
        }
    }

    function Get-MethodArgument {
        param($CommandAst)
        $elements = @($CommandAst.CommandElements)
        for ($i = 0; $i -lt $elements.Count; $i++) {
            $element = $elements[$i]
            if ($element -isnot [System.Management.Automation.Language.CommandParameterAst]) { continue }
            if ($element.ParameterName -ne 'Method') { continue }
            $value = if ($null -ne $element.Argument) { $element.Argument.Extent.Text } else { $elements[$i + 1].Extent.Text }
            return ($value -replace "['`"]", '')
        }
        return $null
    }
}

Describe 'Only read-only tenant commands are called' {
    It 'calls only Connect, Disconnect, Get and Search tenant cmdlets' {
        $offenders = $script:Calls | Where-Object {
            $_.Name -and $_.Name -in $script:TenantCmdlets -and
            $script:AllowedVerbs -notcontains $_.Name.Substring(0, $_.Name.IndexOf('-'))
        } | ForEach-Object { '{0} {1}' -f $_.File, $_.Name }
        $offenders -join '; ' | Should -BeNullOrEmpty
    }

    It 'covers the tenant cmdlets it expects' {
        $names = $script:Calls.Name | Sort-Object -Unique
        $names | Should -Contain 'Connect-AzAccount'
        $names | Should -Contain 'Get-AzAccessToken'
        $names | Should -Contain 'Search-UnifiedAuditLog'
        $names | Should -Contain 'Invoke-RestMethod'
    }
}

Describe 'Raw HTTP requests are GET, except the one inventory query' {
    It 'passes -Method Get to every Invoke-RestMethod outside Invoke-InventoryQuery' {
        $offenders = $script:Calls | Where-Object { $_.Name -eq 'Invoke-RestMethod' -and $_.Function -ne 'Invoke-InventoryQuery' } |
            Where-Object { (Get-MethodArgument $_.Ast) -ne 'Get' } |
            ForEach-Object { '{0}:{1}' -f $_.File, $_.Ast.Extent.StartLineNumber }
        $offenders -join '; ' | Should -BeNullOrEmpty
    }

    It 'uses POST only in Invoke-InventoryQuery, against the resourcequery endpoint, which reads and changes nothing' {
        $posts = @($script:Calls | Where-Object { $_.Name -eq 'Invoke-RestMethod' -and (Get-MethodArgument $_.Ast) -eq 'Post' })
        $posts.Count | Should -Be 1
        $posts[0].Function | Should -Be 'Invoke-InventoryQuery'
        (Get-Content -LiteralPath (Join-Path $script:ScanRoot 'CopilotStudioHelpers.ps1') -Raw) | Should -Match '/resourcequery/resources/query'
    }

    It 'uses no other raw HTTP command' {
        $offenders = $script:Calls | Where-Object { $_.Name -in @('Invoke-WebRequest', 'Invoke-MgGraphRequest', 'curl', 'wget') } |
            ForEach-Object { '{0}:{1} {2}' -f $_.File, $_.Ast.Extent.StartLineNumber, $_.Name }
        $offenders -join '; ' | Should -BeNullOrEmpty
    }
}

Describe 'No credentials are accepted for the delegated collectors' {
    It 'gives the inventory and Dataverse collectors no -AppId or -CertificateThumbprint' {
        foreach ($name in 'Get-PowerPlatformEnvironments', 'Get-CopilotStudioAgents', 'Get-AgentConnectors', 'Get-AgentComponents', 'Get-AgentModifications') {
            $command = Get-Command (Join-Path $script:ScanRoot "$name.ps1")
            $command.Parameters.Keys | Should -Not -Contain 'AppId'
            $command.Parameters.Keys | Should -Not -Contain 'CertificateThumbprint'
        }
    }

    It 'gives the audit collector the library''s app-only parameters' {
        $command = Get-Command (Join-Path $script:ScanRoot 'Get-AgentAuditEvents.ps1')
        $command.Parameters.Keys | Should -Contain 'AppId'
        $command.Parameters.Keys | Should -Contain 'CertificateThumbprint'
    }
}
