#Requires -Version 7.0

# Static scan of this report's scripts. Nothing here touches a tenant.

BeforeAll {
    $script:ScanRoot = Join-Path $PSScriptRoot '../collectors'
    $script:AllowedVerbs = @('Connect', 'Disconnect', 'Get', 'Search')
    # Tenant commands this report calls whose verb is not Get, Search, Connect or Disconnect.
    # Each is allowed on purpose, with the reason:
    #   Start-SPODataAccessGovernanceInsight   creates a Data access governance report. It
    #       reads the tenant's permission state and writes nothing to it; the report is the
    #       only way Microsoft offers this data
    #       (https://learn.microsoft.com/sharepoint/powershell-for-data-access-governance).
    #   Export-SPODataAccessGovernanceInsight  downloads that report as a CSV to a local folder.
    #   Invoke-MgGraphRequest                  a GET only, in one helper function (checked below);
    #       needed to follow a getAllSites @odata.nextLink exactly as returned.
    $script:AllowedOtherVerbCommands = @(
        'Start-SPODataAccessGovernanceInsight', 'Export-SPODataAccessGovernanceInsight', 'Invoke-MgGraphRequest'
    )
    # Commands that change a tenant and must never appear.
    # Start-SPOAuditDataCollectionForActivityInsights turns audit data collection on.
    $script:Forbidden = @(
        'Start-SPOAuditDataCollectionForActivityInsights', 'Set-SPOTenant', 'Set-SPOSite', 'Remove-SPOSite',
        'Invoke-RestMethod', 'Invoke-WebRequest', 'curl', 'wget'
    )
    $script:TenantPattern = '^(Connect|Disconnect|Get|Search|Start|Stop|Set|New|Remove|Add|Update|Export|Invoke)-(Mg|SPO|EXO|AdminAuditLog|UnifiedAuditLog|Label)'

    $script:Calls = foreach ($file in Get-ChildItem -LiteralPath $script:ScanRoot -Filter '*.ps1') {
        $tokens = $null
        $errors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$errors)
        if ($errors.Count -gt 0) { throw ('{0} does not parse: {1}' -f $file.Name, $errors[0].Message) }
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
    It 'calls only Connect, Disconnect, Get and Search tenant commands, plus the allowed ones' {
        $offenders = $script:Calls | Where-Object { $_.Name -and $_.Name -match $script:TenantPattern } | Where-Object {
            $verb = $_.Name.Substring(0, $_.Name.IndexOf('-'))
            $script:AllowedVerbs -notcontains $verb -and $script:AllowedOtherVerbCommands -notcontains $_.Name
        } | Where-Object { $_.Name -notlike 'Export-AppendCsv' } | ForEach-Object { '{0} {1}' -f $_.File, $_.Name }
        $offenders -join '; ' | Should -BeNullOrEmpty
    }

    It 'never calls a command that changes the tenant, turns on data collection or makes a raw HTTP request' {
        $offenders = $script:Calls | Where-Object { $_.Name -in $script:Forbidden } |
            ForEach-Object { '{0}:{1} {2}' -f $_.File, $_.Ast.Extent.StartLineNumber, $_.Name }
        $offenders -join '; ' | Should -BeNullOrEmpty
    }

    It 'covers the tenant commands it expects' {
        $names = $script:Calls.Name | Sort-Object -Unique
        foreach ($expected in 'Connect-SPOService', 'Get-SPOSite', 'Get-SPOTenant', 'Start-SPODataAccessGovernanceInsight',
            'Get-SPODataAccessGovernanceInsight', 'Export-SPODataAccessGovernanceInsight', 'Get-SPOAuditDataCollectionStatusForActivityInsights',
            'Search-UnifiedAuditLog', 'Get-AdminAuditLogConfig', 'Invoke-MgGraphRequest', 'Get-Label') {
            $names | Should -Contain $expected
        }
    }
}

Describe 'Raw Graph requests are GET and sit in one helper function' {
    It 'passes -Method GET to every Invoke-MgGraphRequest' {
        $offenders = $script:Calls | Where-Object { $_.Name -eq 'Invoke-MgGraphRequest' } |
            Where-Object { (Get-MethodArgument $_.Ast) -ne 'GET' } |
            ForEach-Object { '{0}:{1}' -f $_.File, $_.Ast.Extent.StartLineNumber }
        $offenders -join '; ' | Should -BeNullOrEmpty
    }

    It 'calls Invoke-MgGraphRequest only from Invoke-GraphGet' {
        $calls = @($script:Calls | Where-Object { $_.Name -eq 'Invoke-MgGraphRequest' })
        $calls.Count | Should -Be 1
        $calls[0].File | Should -Be 'OversharingHelpers.ps1'
        $calls[0].Function | Should -Be 'Invoke-GraphGet'
    }
}

Describe 'Sign-in' {
    It 'asks only for read permissions' {
        $scopes = $script:Calls | Where-Object { $_.Name -eq 'Connect-M365Service' } | ForEach-Object { $_.Ast.Extent.Text } | Out-String
        $scopes | Should -Not -Match 'ReadWrite'
    }
}
