#Requires -Version 7.0

# Static scan of this report's scripts. Nothing here touches a tenant.

BeforeAll {
    $script:ScanRoot = Join-Path $PSScriptRoot '../collectors'
    $script:AllowedVerbs = @('Connect', 'Disconnect', 'Get')
    # Tenant cmdlets this report calls; any other one must be added here on purpose.
    $script:TenantCmdlets = @(
        'Connect-M365Service', 'Get-MgUser', 'Get-MgSubscribedSku', 'Get-MgUserLicenseDetail', 'Get-MgAdminReportSetting'
    )

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
    It 'calls only Connect, Disconnect and Get tenant cmdlets' {
        $offenders = $script:Calls | Where-Object {
            $_.Name -and $_.Name -match '^[A-Za-z]+-Mg' -and
            $script:AllowedVerbs -notcontains $_.Name.Substring(0, $_.Name.IndexOf('-'))
        } | Where-Object { $_.Name -ne 'Invoke-MgGraphRequest' } | ForEach-Object { '{0} {1}' -f $_.File, $_.Name }
        $offenders -join '; ' | Should -BeNullOrEmpty
    }

    It 'covers the tenant cmdlets it expects' {
        $names = $script:Calls.Name | Sort-Object -Unique
        foreach ($expected in $script:TenantCmdlets + 'Invoke-MgGraphRequest') { $names | Should -Contain $expected }
    }
}

Describe 'Raw Graph requests are GET and sit in two helper functions' {
    It 'passes -Method GET to every Invoke-MgGraphRequest' {
        $offenders = $script:Calls | Where-Object { $_.Name -eq 'Invoke-MgGraphRequest' } |
            Where-Object { (Get-MethodArgument $_.Ast) -ne 'GET' } |
            ForEach-Object { '{0}:{1}' -f $_.File, $_.Ast.Extent.StartLineNumber }
        $offenders -join '; ' | Should -BeNullOrEmpty
    }

    It 'calls Invoke-MgGraphRequest only from Get-GraphReportCsv and Get-GraphReportJsonPage' {
        $calls = @($script:Calls | Where-Object { $_.Name -eq 'Invoke-MgGraphRequest' })
        $calls.Count | Should -Be 2
        $calls.File | Sort-Object -Unique | Should -Be 'LicenseUtilizationHelpers.ps1'
        ($calls.Function | Sort-Object) -join ',' | Should -Be 'Get-GraphReportCsv,Get-GraphReportJsonPage'
    }

    It 'uses no other raw HTTP command' {
        $offenders = $script:Calls | Where-Object { $_.Name -in @('Invoke-RestMethod', 'Invoke-WebRequest', 'curl', 'wget') } |
            ForEach-Object { '{0}:{1} {2}' -f $_.File, $_.Ast.Extent.StartLineNumber, $_.Name }
        $offenders -join '; ' | Should -BeNullOrEmpty
    }
}
