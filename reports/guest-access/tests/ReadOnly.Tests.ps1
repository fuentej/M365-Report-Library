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
        'DlpCompliancePolicy', 'DlpComplianceRule', 'Label', 'LabelPolicy', 'AutoSensitivityLabelPolicy'
        'ComplianceTag', 'DlpSensitiveInformationType', 'ActivityExplorerData', 'ContentExplorerData'
        'ProtectionAlert', 'AdminAuditLogConfig'
        'SPOSite', 'SPOUser', 'SPOTenant', 'SPOExternalUser'
    )
    # Export-ActivityExplorerData and Export-ContentExplorerData are read-only despite
    # the verb: both only return activity/content records, documented at
    # https://learn.microsoft.com/powershell/module/exchangepowershell/export-activityexplorerdata
    # and https://learn.microsoft.com/powershell/module/exchangepowershell/export-contentexplorerdata.
    # 'Export' is safe to allow globally here because a command only reaches this list
    # at all when its noun is already one of the curated $script:TenantNouns above.
    $script:AllowedVerbs = @('Get', 'Search', 'Connect', 'Disconnect', 'Export')

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

    function Get-CommandCallFromText {
        <#
            .SYNOPSIS
                Every command invocation in a snippet, in the same shape as Get-CommandCall.
        #>
        param([Parameter(Mandatory)][string]$Text)

        $tokens = $null
        $errors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseInput($Text, [ref]$tokens, [ref]$errors)

        if ($errors.Count -gt 0) {
            throw ("snippet does not parse: {0}" -f $errors[0].Message)
        }

        foreach ($command in $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.CommandAst] }, $true)) {
            $name = $command.GetCommandName()
            if ([string]::IsNullOrWhiteSpace($name)) { continue }
            [pscustomobject]@{ Name = $name; Ast = $command; Path = '<snippet>' }
        }
    }

    function Get-InvokedCommandName {
        <#
            .SYNOPSIS
                The command name with a module qualifier removed.
        #>
        param([Parameter(Mandatory)][string]$Name)

        $slash = $Name.LastIndexOf('\')
        if ($slash -ge 0 -and $slash -lt ($Name.Length - 1)) {
            return $Name.Substring($slash + 1)
        }
        return $Name
    }

    function Test-ForcedImportModule {
        <#
            .SYNOPSIS
                True when the call is Import-Module and -Force is bound.

            .DESCRIPTION
                ParameterName is the text the author wrote, so -Fo and -For are not equal to
                Force. StaticParameterBinder resolves those abbreviations the way the engine
                does at runtime. A module-qualified call such as
                Microsoft.PowerShell.Core\Import-Module is still Import-Module.
        #>
        param([Parameter(Mandatory)]$Call)

        if ((Get-InvokedCommandName -Name $Call.Name) -ne 'Import-Module') {
            return $false
        }

        $bound = [System.Management.Automation.Language.StaticParameterBinder]::BindCommand($Call.Ast, $true)
        return @($bound.BoundParameters.Keys) -contains 'Force'
    }

    function Test-SharedModuleImport {
        <#
            .SYNOPSIS
                True when an Import-Module call names M365ReportLibrary.psm1.
        #>
        param([Parameter(Mandatory)]$Call)

        if ((Get-InvokedCommandName -Name $Call.Name) -ne 'Import-Module') {
            return $false
        }
        return $Call.Ast.Extent.Text -match 'M365ReportLibrary\.psm1'
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

Describe 'The shared module is never imported with -Force' {
    It 'imports the shared module at least once, so the check below is not vacuous' {
        $imports = @($script:Calls | Where-Object { Test-SharedModuleImport -Call $_ })
        $imports.Count | Should -BeGreaterThan 0
    }

    It 'passes no -Force switch to Import-Module' {
        # -Force removes the loaded module and imports it again, discarding any Pester mock
        # a caller installed against it (-ModuleName M365ReportLibrary) before the collector
        # runs. See the "Adding a report" section of the root README.md.
        # https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/import-module
        $offenders = foreach ($call in $script:Calls) {
            if (-not (Test-ForcedImportModule -Call $call)) { continue }
            '{0}:{1} Import-Module -Force' -f (Split-Path $call.Path -Leaf), $call.Ast.Extent.StartLineNumber
        }

        $offenders -join '; ' | Should -BeNullOrEmpty
    }

    It 'treats an abbreviated or module-qualified -Force as a forced import' {
        # These bind -Force at runtime. A check of ParameterName -eq 'Force', or of the
        # command name -eq 'Import-Module', misses both.
        $samples = @(
            'Import-Module (Join-Path $PSScriptRoot ''../../../shared/M365ReportLibrary.psm1'') -Fo'
            'Import-Module (Join-Path $PSScriptRoot ''../../../shared/M365ReportLibrary.psm1'') -For'
            'Microsoft.PowerShell.Core\Import-Module (Join-Path $PSScriptRoot ''../../../shared/M365ReportLibrary.psm1'') -Force'
        )

        foreach ($sample in $samples) {
            $forced = @(Get-CommandCallFromText -Text $sample | Where-Object { Test-ForcedImportModule -Call $_ })
            $forced.Count | Should -Be 1
            (Test-SharedModuleImport -Call $forced[0]) | Should -BeTrue
        }
    }

    It 'records the convention, the mock reason, and the already-loaded trade-off' {
        $readme = Get-Content -LiteralPath (Join-Path $script:Root 'README.md') -Raw
        $start = $readme.IndexOf('## Adding a report')
        $start | Should -BeGreaterThan -1
        $section = ($readme.Substring($start) -replace '\s+', ' ')
        $section | Should -Match 'never with `-Force`'
        $section | Should -Match 'drops any Pester mock'
        $section | Should -Match 'already imported the module keeps that copy'
        $section | Should -Match 'Scripts under `tests/` are outside this rule'
    }

    It 'does not treat a plain import, an ambiguous -F, or -Function as -Force' {
        $samples = @(
            'Import-Module (Join-Path $PSScriptRoot ''../../../shared/M365ReportLibrary.psm1'')'
            'Import-Module $m -F'
            'Import-Module $m -Function Get-Thing'
        )

        foreach ($sample in $samples) {
            $forced = @(Get-CommandCallFromText -Text $sample | Where-Object { Test-ForcedImportModule -Call $_ })
            $forced.Count | Should -Be 0
        }
    }
}

Describe 'Only read-only tenant commands are called' {
    It 'calls no Graph, Exchange Online or Security & Compliance cmdlet outside <AllowedVerbs>' {
        # Invoke-MgGraphRequest in the oversharing, Exchange activity, SharePoint and OneDrive,
        # Teams activity, Copilot usage and Entra activity helpers is the one Invoke-verb call here:
        # each follows @odata.nextLink as returned and sends headers the SDK page iterator
        # drops. The next Describe holds every -Method to GET, and each report's
        # ReadOnly.Tests.ps1 holds that helper file to one GET function.
        # The unified audit log helpers (UnifiedAuditLogHelpers.ps1) also call it, for the Graph
        # audit query GETs and for the one POST described in the next Describe.
        $offenders = $script:TenantCalls | Where-Object {
            -not ($_.Name -eq 'Invoke-MgGraphRequest' -and (Split-Path $_.Path -Leaf) -in @('OversharingHelpers.ps1', 'ExchangeActivityHelpers.ps1', 'SharePointOneDriveHelpers.ps1', 'TeamsActivityHelpers.ps1', 'CopilotUsageHelpers.ps1', 'EntraActivityHelpers.ps1', 'UnifiedAuditLogHelpers.ps1', 'LicenseUtilizationHelpers.ps1'))
        } | Where-Object {
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

                # Decision D-007: the unified audit log helpers send the one POST that creates a Graph
                # audit log query (a read, no tenant change). reports/unified-audit-log/tests/ReadOnly.Tests.ps1
                # holds it to that one path and to the one function that sends it.
                if ((Split-Path $call.Path -Leaf) -eq 'UnifiedAuditLogHelpers.ps1' -and ($value -replace "['`"]", '') -eq 'POST' -and $call.Name -eq 'Invoke-MgGraphRequest') { continue }

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
        #
        # One exception: the Copilot Studio agents report reads the Power Platform
        # inventory API and the Dataverse Web API, neither of which has a PowerShell
        # cmdlet. Its helper file is exempt here and held to a stricter rule in
        # reports/copilot-studio-agents/tests/ReadOnly.Tests.ps1 (GET everywhere, and one
        # POST that carries an inventory query specification).
        #
        # The oversharing report reads Microsoft Graph sites, drives and item permissions
        # with GET through Invoke-MgGraphRequest; its helper file is exempt here and held
        # to one GET function in reports/oversharing/tests/ReadOnly.Tests.ps1.
        #
        # The Exchange activity report reads the Graph message trace (/beta, paged by the
        # @odata.nextLink as returned) with GET through the one Invoke-MgGraphRequest call in
        # its helper file, which reports/exchange-activity/tests/ReadOnly.Tests.ps1 holds to
        # one GET call.
        #
        # The SharePoint and OneDrive activity report reads getAllSites, the drive lists and the
        # site activity intervals with GET through the one Invoke-MgGraphRequest call in its
        # helper file, which reports/sharepoint-onedrive-activity/tests/ReadOnly.Tests.ps1 holds
        # to one GET call.
        #
        # The Teams activity report reads call records (list, record with sessions, and their
        # nextLinks) with GET through the one Invoke-MgGraphRequest call in its helper file
        # (Invoke-GraphGet), which reports/teams-activity/tests/ReadOnly.Tests.ps1 holds to one
        # GET call.
        #
        # The Copilot usage report reads usage reports and interaction history with GET
        # through the one Invoke-MgGraphRequest call in its helper file (Invoke-GraphGet),
        # which reports/copilot-usage/tests/ReadOnly.Tests.ps1 holds to one GET call.
        #
        # The Entra activity report reads sign-ins and directory audits the same way, so
        # Prefer: include-unknown-enum-members is on every page, including each nextLink.
        # reports/entra-activity/tests/ReadOnly.Tests.ps1 holds that helper to one GET call.
        #
        # The unified audit log report reads the Graph Audit Search API through Invoke-MgGraphRequest
        # and the Office 365 Management Activity API (which has no PowerShell cmdlet) through
        # Invoke-WebRequest, both in UnifiedAuditLogHelpers.ps1. Between them they send two POSTs, each
        # of which starts a read and changes no tenant data (decision D-007): the audit log query create
        # and the subscription start. reports/unified-audit-log/tests/ReadOnly.Tests.ps1 pins each POST
        # to its path and to the one function that sends it, holds every other call to GET, and checks
        # that the token is read in two request functions only and never logged.
        #
        # The license utilization usage reports are read with GET through Invoke-MgGraphRequest.
        # That helper file is exempt here and held to GET in two functions in
        # reports/license-utilization/tests/ReadOnly.Tests.ps1.
        $offenders = $script:Calls |
            Where-Object { (Split-Path $_.Path -Leaf) -notin @('CopilotStudioHelpers.ps1', 'OversharingHelpers.ps1', 'ExchangeActivityHelpers.ps1', 'SharePointOneDriveHelpers.ps1', 'TeamsActivityHelpers.ps1', 'CopilotUsageHelpers.ps1', 'EntraActivityHelpers.ps1', 'UnifiedAuditLogHelpers.ps1', 'LicenseUtilizationHelpers.ps1') } |
            Where-Object { $_.Name -in @('Invoke-RestMethod', 'Invoke-WebRequest', 'Invoke-MgGraphRequest', 'curl', 'wget') } |
            ForEach-Object { '{0}:{1} {2}' -f (Split-Path $_.Path -Leaf), $_.Ast.Extent.StartLineNumber, $_.Name }

        $offenders -join '; ' | Should -BeNullOrEmpty
    }
}

Describe 'The test workflow runs every report that has tests' {
    It 'includes the Entra activity tests' {
        # A merge of main can replace this line with only the other report's path.
        # The Entra tests then never run, so a failure there does not fail CI.
        $workflow = Get-Content -LiteralPath (Join-Path $script:Root '.github/workflows/tests.yml') -Raw
        $workflow | Should -Match ([regex]::Escape('./reports/entra-activity/tests'))
        $workflow | Should -Match ([regex]::Escape('./reports/copilot-usage/tests'))
    }
}
