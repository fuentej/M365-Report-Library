#Requires -Version 7.0

<#
    Nothing in this repository may change a tenant. These tests read every script in the
    library - not the test scripts, which stub the cmdlets - and check that the only
    tenant commands they call are Get, Search, Connect or Disconnect, and that no Graph
    request asks for a method other than GET.
#>

BeforeAll {
    $script:Root = Join-Path $PSScriptRoot '../../..' | Resolve-Path | Select-Object -ExpandProperty Path

    # Nouns belonging to a service this library talks to. A command is a "tenant command"
    # when its noun starts with Mg or EXO (the Microsoft Graph PowerShell SDK and the
    # Exchange Online REST cmdlets) or is one of the Exchange Online and Security &
    # Compliance nouns below.
    $script:TenantNounPrefixes = @('Mg', 'EXO')
    $script:TenantNouns = @(
        'ExchangeOnline', 'IPPSSession', 'UnifiedAuditLog', 'Mailbox', 'MailboxPermission'
        'MailboxFolderPermission', 'RecipientPermission', 'Recipient', 'OrganizationConfig'
        'SharingPolicy', 'AcceptedDomain', 'DistributionGroup', 'UnifiedGroup', 'TransportRule'
        'ComplianceSearch', 'ComplianceSearchAction', 'RetentionCompliancePolicy'
        'DlpCompliancePolicy', 'Label', 'LabelPolicy', 'ProtectionAlert', 'AdminAuditLogConfig'
        'SPOSite', 'SPOUser', 'SPOTenant', 'SPOExternalUser'
    )
    $script:AllowedVerbs = @('Get', 'Search', 'Connect', 'Disconnect')

    function Get-LibraryScript {
        <#
            .SYNOPSIS
                Every script the library ships, leaving out the test scripts.
        #>
        Get-ChildItem -LiteralPath $script:Root -Recurse -File -Include '*.ps1', '*.psm1' |
            Where-Object { $_.FullName -notmatch '[\\/]tests[\\/]' } |
            Sort-Object FullName
    }

    function Get-CommandCall {
        <#
            .SYNOPSIS
                Every command invocation in a script, as name plus the AST that called it.
        #>
        param([Parameter(Mandatory)][string]$Path)

        $tokens = $null
        $errors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors)

        if ($errors.Count -gt 0) {
            throw ("{0} does not parse: {1}" -f $Path, ($errors[0].Message))
        }

        foreach ($command in $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.CommandAst] }, $true)) {
            $name = $command.GetCommandName()
            if ([string]::IsNullOrWhiteSpace($name)) { continue }
            [pscustomobject]@{ Name = $name; Ast = $command; Path = $Path }
        }
    }

    function Test-TenantCommand {
        param([Parameter(Mandatory)][string]$Name)

        $dash = $Name.IndexOf('-')
        if ($dash -lt 1) { return $false }

        $noun = $Name.Substring($dash + 1)
        foreach ($prefix in $script:TenantNounPrefixes) {
            if ($noun.StartsWith($prefix, [System.StringComparison]::Ordinal) -and $noun.Length -gt $prefix.Length) {
                return $true
            }
        }
        return $script:TenantNouns -contains $noun
    }

    $script:Scripts = @(Get-LibraryScript)
    $script:Calls = @($script:Scripts | ForEach-Object { Get-CommandCall -Path $_.FullName })
    $script:TenantCalls = @($script:Calls | Where-Object { Test-TenantCommand -Name $_.Name })
}

Describe 'Every script in the library parses' {
    It 'finds scripts to check' {
        $script:Scripts.Count | Should -BeGreaterThan 5
    }

    It 'reads a meaningful number of tenant commands' {
        # Guards the checks below against silently passing because the scan found nothing.
        $script:TenantCalls.Count | Should -BeGreaterThan 5
    }
}

Describe 'Only read-only tenant commands are called' {
    It 'calls no Graph, Exchange Online or Security & Compliance cmdlet outside <AllowedVerbs>' {
        $offenders = $script:TenantCalls | Where-Object {
            $verb = $_.Name.Substring(0, $_.Name.IndexOf('-'))
            $script:AllowedVerbs -notcontains $verb
        } | ForEach-Object {
            '{0}:{1} {2}' -f (Split-Path $_.Path -Leaf), $_.Ast.Extent.StartLineNumber, $_.Name
        }

        $offenders -join '; ' | Should -BeNullOrEmpty
    }

    It 'lists the tenant commands it checked' {
        # Not an assertion about behaviour so much as a record in the test output of what
        # the scan actually covered.
        $names = ($script:TenantCalls.Name | Sort-Object -Unique)
        $names | Should -Not -BeNullOrEmpty
        $names | Should -Contain 'Get-MgUser'
        $names | Should -Contain 'Get-MgUserMemberOfAsGroup'
        $names | Should -Contain 'Search-UnifiedAuditLog'
        $names | Should -Contain 'Connect-MgGraph'
        $names | Should -Contain 'Disconnect-ExchangeOnline'
    }
}

Describe 'No Graph request uses a method other than GET' {
    It 'passes no -Method other than GET to a tenant command' {
        $offenders = foreach ($call in $script:TenantCalls) {
            $elements = @($call.Ast.CommandElements)
            for ($i = 0; $i -lt $elements.Count; $i++) {
                $element = $elements[$i]
                if ($element -isnot [System.Management.Automation.Language.CommandParameterAst]) { continue }
                if ($element.ParameterName -ne 'Method') { continue }

                $value = if ($null -ne $element.Argument) { $element.Argument.Extent.Text }
                elseif ($i + 1 -lt $elements.Count) { $elements[$i + 1].Extent.Text }
                else { '<none>' }

                if (($value -replace "['`"]", '') -ne 'GET') {
                    '{0}:{1} {2} -Method {3}' -f (Split-Path $call.Path -Leaf), $element.Extent.StartLineNumber, $call.Name, $value
                }
            }
        }

        $offenders -join '; ' | Should -BeNullOrEmpty
    }

    It 'reaches no service through a raw HTTP call' {
        # Invoke-RestMethod and friends would slip past the verb rule above, so they are
        # not allowed in the library at all.
        $offenders = $script:Calls |
            Where-Object { $_.Name -in @('Invoke-RestMethod', 'Invoke-WebRequest', 'Invoke-MgGraphRequest', 'curl', 'wget') } |
            ForEach-Object { '{0}:{1} {2}' -f (Split-Path $_.Path -Leaf), $_.Ast.Extent.StartLineNumber, $_.Name }

        $offenders -join '; ' | Should -BeNullOrEmpty
    }
}
