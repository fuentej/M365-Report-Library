#Requires -Version 7.0

BeforeAll {
    $script:Root = Join-Path $PSScriptRoot '../../..' | Resolve-Path | Select-Object -ExpandProperty Path
    $script:Collectors = Join-Path $script:Root 'reports/entra-activity/collectors'
    $script:Samples = Join-Path $script:Root 'reports/entra-activity/samples'

    . (Join-Path $script:Root 'shared/tests/TenantCmdletStubs.ps1')

    # Not in the shared stubs: the beta sign-in cmdlet (Microsoft.Graph.Beta.Reports) and
    # Get-MgContext. Like the shared stubs these are global and throw if they leak.
    function global:Get-MgBetaAuditLogSignIn {
        [CmdletBinding()]
        param([switch]$All, [string]$Filter, [int]$Top)
        throw 'Get-MgBetaAuditLogSignIn was called for real. Mock it in the test.'
    }
    function global:Get-MgContext {
        [CmdletBinding()]
        param()
        throw 'Get-MgContext was called for real. Mock it in the test.'
    }

    Import-Module (Join-Path $script:Root 'shared/M365ReportLibrary.psm1') -Force

    $script:Schema = Import-PowerShellDataFile -LiteralPath (Join-Path $script:Collectors 'EntraActivitySchema.psd1')

    function New-TestFolder {
        $path = Join-Path ([System.IO.Path]::GetTempPath()) ('entra-activity-' + [guid]::NewGuid().ToString('N'))
        New-Item -Path $path -ItemType Directory -Force | Out-Null
        return $path
    }

    function Get-HeaderText {
        param([Parameter(Mandatory)][string]$Path)
        return ((Get-CsvHeaderColumn -Path $Path) -join ',')
    }

    function Get-LogText {
        param([Parameter(Mandatory)][string]$Folder)
        $log = Join-Path $Folder 'run.log'
        if (Test-Path -LiteralPath $log) { return (Get-Content -LiteralPath $log -Raw) }
        return ''
    }

    function Invoke-CollectorScript {
        param([Parameter(Mandatory)][string]$Name, [hashtable]$Arguments = @{})
        & (Join-Path $script:Collectors $Name) @Arguments
    }

    # signIn: https://learn.microsoft.com/graph/api/resources/signin#properties
    # List (v1.0 and beta): https://learn.microsoft.com/graph/api/signin-list
    # appliedConditionalAccessPolicy:
    # https://learn.microsoft.com/graph/api/resources/appliedconditionalaccesspolicy
    function New-MockSignIn {
        param(
            [string]$Id = 'signin-1',
            [datetime]$At = [datetime]'2026-09-10T06:05:00Z',
            [bool]$Interactive = $true,
            [int]$ErrorCode = 0,
            [bool]$WithPolicies = $true,
            [string]$UserId = 'user-1',
            [string]$Upn = 'casey.chaudhry@example.com'
        )

        $policies = if ($WithPolicies) {
            @(
                [pscustomobject]@{
                    Id                      = 'policy-1'
                    DisplayName             = 'Require MFA for all users'
                    EnforcedGrantControls   = @('mfa')
                    EnforcedSessionControls = @()
                    Result                  = 'success'
                }
                [pscustomobject]@{
                    Id                      = 'policy-2'
                    DisplayName             = 'Block legacy authentication'
                    EnforcedGrantControls   = @('block')
                    EnforcedSessionControls = @('signInFrequency')
                    Result                  = 'notApplied'
                }
            )
        }
        else { @() }

        [pscustomobject]@{
            Id                              = $Id
            CreatedDateTime                 = $At
            UserId                          = $UserId
            UserPrincipalName               = $Upn
            AppId                           = '00000002-0000-0ff1-ce00-000000000000'
            AppDisplayName                  = 'Office 365 Exchange Online'
            ResourceDisplayName             = 'Office 365 Exchange Online'
            IPAddress                       = '203.0.113.10'
            ClientAppUsed                   = 'Browser'
            IsInteractive                   = $Interactive
            ConditionalAccessStatus         = 'success'
            RiskDetail                      = 'none'
            RiskLevelAggregated             = 'none'
            RiskLevelDuringSignIn           = 'none'
            RiskState                       = 'none'
            Location                        = [pscustomobject]@{ City = 'Tampa'; State = 'Florida'; CountryOrRegion = 'US' }
            DeviceDetail                    = [pscustomobject]@{ OperatingSystem = 'Windows 11'; Browser = 'Edge 128.0'; IsCompliant = $true; IsManaged = $true }
            Status                          = [pscustomobject]@{ ErrorCode = $ErrorCode; FailureReason = 'Other.'; AdditionalDetails = 'Detail text' }
            AppliedConditionalAccessPolicies = $policies
            SignInEventTypes                = $(if ($Interactive) { @('interactiveUser') } else { @('nonInteractiveUser') })
        }
    }

    # directoryAudit: https://learn.microsoft.com/graph/api/resources/directoryaudit
    # List: https://learn.microsoft.com/graph/api/directoryaudit-list
    function New-MockAudit {
        param(
            [string]$Id = 'Directory_audit-1',
            [datetime]$At = [datetime]'2026-09-10T09:00:00Z',
            [string]$Category = 'Policy',
            [string]$Operation = 'Update',
            [string]$Result = 'success'
        )

        [pscustomobject]@{
            Id                  = $Id
            ActivityDateTime    = $At
            ActivityDisplayName = 'Update Conditional Access policy'
            Category            = $Category
            OperationType       = $Operation
            Result              = $Result
            ResultReason        = ''
            LoggedByService     = 'Conditional Access'
            CorrelationId       = 'dddddddd-0000-4000-8000-000000000001'
            InitiatedBy         = [pscustomobject]@{
                User = [pscustomobject]@{ Id = 'user-9'; UserPrincipalName = 'admin.alvarez@example.com' }
                App  = $null
            }
            TargetResources     = @(
                [pscustomobject]@{
                    Id                 = 'policy-1'
                    Type               = 'Policy'
                    DisplayName        = 'Require MFA for all users'
                    UserPrincipalName  = $null
                    ModifiedProperties = @([pscustomobject]@{ DisplayName = 'State'; OldValue = '"disabled"'; NewValue = '"enabled"' })
                }
            )
        }
    }

    # One row per tenant-reading collector: the script, CSV, schema key, the cmdlet it
    # reads through, and the builder for a mock object.
    $script:EventCases = @(
        @{ Script = 'Get-InteractiveSignIns.ps1'; Csv = 'signins-interactive.csv'; Key = 'SignIns'; Cmdlet = 'Get-MgAuditLogSignIn'; Builder = 'New-MockSignIn' }
        @{ Script = 'Get-NonInteractiveSignIns.ps1'; Csv = 'signins-noninteractive.csv'; Key = 'SignIns'; Cmdlet = 'Get-MgBetaAuditLogSignIn'; Builder = 'New-MockSignIn' }
        @{ Script = 'Get-SignInConditionalAccess.ps1'; Csv = 'signin-conditional-access.csv'; Key = 'SignInConditionalAccess'; Cmdlet = 'Get-MgAuditLogSignIn'; Builder = 'New-MockSignIn' }
        @{ Script = 'Get-DirectoryAudits.ps1'; Csv = 'directory-audits.csv'; Key = 'DirectoryAudits'; Cmdlet = 'Get-MgAuditLogDirectoryAudit'; Builder = 'New-MockAudit' }
    )
}

AfterAll {
    Remove-Module M365ReportLibrary -Force -ErrorAction SilentlyContinue
}

Describe 'Each collector writes the columns of its sample file' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
    }
    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It '<Csv>' -ForEach $script:EventCases {
        $builder = $Builder
        Mock -CommandName $Cmdlet -MockWith { & $builder }

        Invoke-CollectorScript $Script @{ OutputPath = $script:folder; LookbackDays = 1 }

        $produced = Join-Path $script:folder $Csv
        @(Import-Csv -LiteralPath $produced).Count | Should -BeGreaterThan 0
        Get-HeaderText -Path $produced | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples $Csv))
        Get-HeaderText -Path $produced | Should -Be ($script:Schema[$Key] -join ',')
    }

    It 'retention-reference.csv' {
        Invoke-CollectorScript 'Get-RetentionReference.ps1' @{ OutputPath = $script:folder } 3>$null

        $produced = Join-Path $script:folder 'retention-reference.csv'
        @(Import-Csv -LiteralPath $produced).Count | Should -Be 3
        Get-HeaderText -Path $produced | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples 'retention-reference.csv'))
        Get-HeaderText -Path $produced | Should -Be ($script:Schema.RetentionReference -join ',')
    }

    It 'every sample file has rows and only the columns of its schema' {
        foreach ($case in @(
                @{ Csv = 'signins-interactive.csv'; Key = 'SignIns' }
                @{ Csv = 'signins-noninteractive.csv'; Key = 'SignIns' }
                @{ Csv = 'signin-conditional-access.csv'; Key = 'SignInConditionalAccess' }
                @{ Csv = 'directory-audits.csv'; Key = 'DirectoryAudits' }
                @{ Csv = 'retention-reference.csv'; Key = 'RetentionReference' }
            )) {
            $path = Join-Path $script:Samples $case.Csv
            Get-HeaderText -Path $path | Should -Be ($script:Schema[$case.Key] -join ',')
            @(Import-Csv -LiteralPath $path).Count | Should -BeGreaterThan 0
        }
    }

    It 'uses example.com and documentation IP addresses in every sample' {
        $text = Get-ChildItem -LiteralPath $script:Samples -Filter '*.csv' | Get-Content -Raw
        $text -join '' | Should -Not -Match '@(?!example\.com)[a-z0-9.-]+\.[a-z]{2,}'
        $addresses = [regex]::Matches(($text -join ''), '\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b').Value | Sort-Object -Unique
        foreach ($address in $addresses) { $address | Should -BeLike '203.0.113.*' }
    }
}

Describe 'Each connection targets the endpoints of its -Environment' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-MgGraph -ModuleName M365ReportLibrary -MockWith { }
        Mock Get-MgAuditLogSignIn -MockWith { @() }
        Mock Get-MgBetaAuditLogSignIn -MockWith { @() }
        Mock Get-MgAuditLogDirectoryAudit -MockWith { @() }
    }
    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    # https://learn.microsoft.com/graph/deployments: GCC calls the global service
    # (https://graph.microsoft.com); GCC High is the USGov environment.
    It '<Script> signs in to <Graph> for <Environment>' -ForEach @(
        foreach ($script in 'Get-InteractiveSignIns.ps1', 'Get-NonInteractiveSignIns.ps1', 'Get-SignInConditionalAccess.ps1', 'Get-DirectoryAudits.ps1') {
            @{ Script = $script; Environment = 'Commercial'; Graph = 'Global' }
            @{ Script = $script; Environment = 'GCC'; Graph = 'Global' }
            @{ Script = $script; Environment = 'GCCHigh'; Graph = 'USGov' }
        }
    ) {
        Invoke-CollectorScript $Script @{ OutputPath = $script:folder; Environment = $Environment; LookbackDays = 1 }

        $expected = $Graph
        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter { $Environment -eq $expected }
    }

    It 'resolves the Graph resource endpoint the contract names for each cloud' -ForEach @(
        @{ Environment = 'Commercial'; Endpoint = 'https://graph.microsoft.com' }
        @{ Environment = 'GCC'; Endpoint = 'https://graph.microsoft.com' }
        @{ Environment = 'GCCHigh'; Endpoint = 'https://graph.microsoft.us' }
    ) {
        (Get-M365ServiceEndpoint -Service Graph -Environment $Environment).ResourceEndpoint | Should -Be $Endpoint
    }

    It 'reads sign-ins with AuditLog.Read.All and does not request a Conditional Access permission' -ForEach @(
        @{ Script = 'Get-InteractiveSignIns.ps1' }
        @{ Script = 'Get-NonInteractiveSignIns.ps1' }
        @{ Script = 'Get-DirectoryAudits.ps1' }
    ) {
        Invoke-CollectorScript $Script @{ OutputPath = $script:folder; LookbackDays = 1 }

        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter {
            $Scopes -contains 'AuditLog.Read.All' -and $Scopes -notcontains 'Policy.Read.All'
        }
    }

    It 'asks for a Conditional Access read permission, not the write one, to read policy detail' {
        Invoke-CollectorScript 'Get-SignInConditionalAccess.ps1' @{ OutputPath = $script:folder; LookbackDays = 1 }

        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter {
            $Scopes -contains 'AuditLog.Read.All' -and $Scopes -contains 'Policy.Read.All' -and $Scopes -notcontains 'Policy.ReadWrite.ConditionalAccess'
        }
    }

    It 'signs in app-only without scopes when given a certificate' {
        Invoke-CollectorScript 'Get-DirectoryAudits.ps1' @{
            OutputPath = $script:folder; LookbackDays = 1; AppId = 'app-1'; CertificateThumbprint = 'AB12'; TenantId = 'tenant-1'
        }

        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter {
            $ClientId -eq 'app-1' -and $CertificateThumbprint -eq 'AB12' -and $TenantId -eq 'tenant-1' -and -not $Scopes
        }
    }

    It 'the retention reference makes no sign-in at all' {
        Invoke-CollectorScript 'Get-RetentionReference.ps1' @{ OutputPath = $script:folder; Environment = 'GCCHigh' } 3>$null

        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 0 -Exactly
    }
}

Describe 'Paging is followed' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
        $script:oneDay = @{ StartDate = [datetime]'2026-09-10T00:00:00Z'; EndDate = [datetime]'2026-09-11T00:00:00Z' }
    }
    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    # 1,250 objects is more than the 1,000-row sign-in page. Without -All the mock
    # returns the first 100, which is what a single request returns.
    # https://learn.microsoft.com/graph/paging
    It '<Cmdlet> keeps every object when the service returns more than one page (<Csv>)' -ForEach @(
        @{ Script = 'Get-InteractiveSignIns.ps1'; Csv = 'signins-interactive.csv'; Cmdlet = 'Get-MgAuditLogSignIn'; Builder = 'New-MockSignIn'; Rows = 1250 }
        @{ Script = 'Get-NonInteractiveSignIns.ps1'; Csv = 'signins-noninteractive.csv'; Cmdlet = 'Get-MgBetaAuditLogSignIn'; Builder = 'New-MockSignIn'; Rows = 1250 }
        @{ Script = 'Get-DirectoryAudits.ps1'; Csv = 'directory-audits.csv'; Cmdlet = 'Get-MgAuditLogDirectoryAudit'; Builder = 'New-MockAudit'; Rows = 1250 }
        @{ Script = 'Get-SignInConditionalAccess.ps1'; Csv = 'signin-conditional-access.csv'; Cmdlet = 'Get-MgAuditLogSignIn'; Builder = 'New-MockSignIn'; Rows = 2500 }
    ) {
        $builder = $Builder
        Mock -CommandName $Cmdlet -MockWith {
            $limit = if ($All) { 1250 } else { 100 }
            1..$limit | ForEach-Object { & $builder -Id "id-$_" }
        }

        Invoke-CollectorScript $Script ($script:oneDay + @{ OutputPath = $script:folder })

        @(Import-Csv -LiteralPath (Join-Path $script:folder $Csv)).Count | Should -Be $Rows
        Should -Invoke -CommandName $Cmdlet -Times 1 -Exactly -ParameterFilter { $All }
    }

    It 'reads every page of the non-interactive stream too when the Conditional Access collector is asked to' {
        Mock Get-MgAuditLogSignIn -MockWith {
            $limit = if ($All) { 1100 } else { 100 }
            1..$limit | ForEach-Object { New-MockSignIn -Id "i-$_" -WithPolicies $false }
        }
        Mock Get-MgBetaAuditLogSignIn -MockWith {
            $limit = if ($All) { 1100 } else { 100 }
            1..$limit | ForEach-Object { New-MockSignIn -Id "n-$_" -Interactive $false -WithPolicies $false }
        }

        Invoke-CollectorScript 'Get-SignInConditionalAccess.ps1' ($script:oneDay + @{ OutputPath = $script:folder; IncludeNonInteractive = $true })

        @(Import-Csv -LiteralPath (Join-Path $script:folder 'signin-conditional-access.csv')).Count | Should -Be 2200
        Should -Invoke Get-MgBetaAuditLogSignIn -Times 1 -Exactly -ParameterFilter { $All }
    }

    It 'does not call the beta API from the Conditional Access collector unless asked' {
        Mock Get-MgAuditLogSignIn -MockWith { New-MockSignIn }
        Mock Get-MgBetaAuditLogSignIn -MockWith { throw 'beta must not be called' }

        Invoke-CollectorScript 'Get-SignInConditionalAccess.ps1' ($script:oneDay + @{ OutputPath = $script:folder })

        Should -Invoke Get-MgBetaAuditLogSignIn -Times 0 -Exactly
    }

    It 'filters each sign-in window on createdDateTime and each audit window on activityDateTime' {
        Mock Get-MgAuditLogSignIn -MockWith { @() }
        Mock Get-MgAuditLogDirectoryAudit -MockWith { @() }

        Invoke-CollectorScript 'Get-InteractiveSignIns.ps1' ($script:oneDay + @{ OutputPath = $script:folder })
        Invoke-CollectorScript 'Get-DirectoryAudits.ps1' ($script:oneDay + @{ OutputPath = $script:folder })

        Should -Invoke Get-MgAuditLogSignIn -Times 1 -Exactly -ParameterFilter {
            $Filter -eq 'createdDateTime ge 2026-09-10T00:00:00Z and createdDateTime lt 2026-09-11T00:00:00Z'
        }
        Should -Invoke Get-MgAuditLogDirectoryAudit -Times 1 -Exactly -ParameterFilter {
            $Filter -eq 'activityDateTime ge 2026-09-10T00:00:00Z and activityDateTime lt 2026-09-11T00:00:00Z'
        }
    }

    It 'filters the beta stream to nonInteractiveUser and never to ne interactiveUser' {
        Mock Get-MgBetaAuditLogSignIn -MockWith { @() }

        Invoke-CollectorScript 'Get-NonInteractiveSignIns.ps1' ($script:oneDay + @{ OutputPath = $script:folder })

        Should -Invoke Get-MgBetaAuditLogSignIn -Times 1 -Exactly -ParameterFilter {
            $Filter -eq "(createdDateTime ge 2026-09-10T00:00:00Z and createdDateTime lt 2026-09-11T00:00:00Z) and signInEventTypes/any(t: t eq 'nonInteractiveUser')" -and
            $Filter -notmatch "ne 'interactiveUser'"
        }
    }

    It 'queries one window per day so no request asks for an unbounded span' {
        Mock Get-MgAuditLogSignIn -MockWith { @() }

        Invoke-CollectorScript 'Get-InteractiveSignIns.ps1' @{
            OutputPath = $script:folder; StartDate = [datetime]'2026-09-10T00:00:00Z'; EndDate = [datetime]'2026-09-13T00:00:00Z'
        }

        Should -Invoke Get-MgAuditLogSignIn -Times 3 -Exactly
    }
}

Describe 'What each collector keeps' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
        $script:oneDay = @{ StartDate = [datetime]'2026-09-10T00:00:00Z'; EndDate = [datetime]'2026-09-11T00:00:00Z' }
    }
    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'keeps every error code, including 0 and codes that are not 5 or 6 digits, and the failure text' {
        Mock Get-MgAuditLogSignIn -MockWith {
            New-MockSignIn -Id 'ok' -ErrorCode 0
            New-MockSignIn -Id 'odd' -ErrorCode 1024
            New-MockSignIn -Id 'mfa' -ErrorCode 50058
        }

        Invoke-CollectorScript 'Get-InteractiveSignIns.ps1' ($script:oneDay + @{ OutputPath = $script:folder })

        $rows = Import-Csv -LiteralPath (Join-Path $script:folder 'signins-interactive.csv')
        ($rows.ErrorCode | Sort-Object) | Should -Be @('0', '1024', '50058')
        $rows[0].FailureReason | Should -Be 'Other.'
        $rows[0].AdditionalDetails | Should -Be 'Detail text'
    }

    It 'does not drop a sign-in that has no user, such as a 50058' {
        Mock Get-MgAuditLogSignIn -MockWith { New-MockSignIn -Id 'nouser' -ErrorCode 50058 -UserId '' -Upn '' }

        Invoke-CollectorScript 'Get-InteractiveSignIns.ps1' ($script:oneDay + @{ OutputPath = $script:folder })

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:folder 'signins-interactive.csv'))
        $rows.Count | Should -Be 1
        $rows[0].UserId | Should -BeNullOrEmpty
    }

    It 'flattens location, device and status, and writes the timestamp as UTC' {
        Mock Get-MgAuditLogSignIn -MockWith { New-MockSignIn }

        Invoke-CollectorScript 'Get-InteractiveSignIns.ps1' ($script:oneDay + @{ OutputPath = $script:folder })

        $row = Import-Csv -LiteralPath (Join-Path $script:folder 'signins-interactive.csv')
        $row.CreatedDateTime | Should -Be '2026-09-10T06:05:00Z'
        $row.City | Should -Be 'Tampa'
        $row.CountryOrRegion | Should -Be 'US'
        $row.DeviceBrowser | Should -Be 'Edge 128.0'
        $row.DeviceIsCompliant | Should -Be 'True'
        $row.SignInEventTypes | Should -Be 'interactiveUser'
    }

    It 'tags non-interactive rows with their event type' {
        Mock Get-MgBetaAuditLogSignIn -MockWith { New-MockSignIn -Interactive $false }

        Invoke-CollectorScript 'Get-NonInteractiveSignIns.ps1' ($script:oneDay + @{ OutputPath = $script:folder })

        $row = Import-Csv -LiteralPath (Join-Path $script:folder 'signins-noninteractive.csv')
        $row.IsInteractive | Should -Be 'False'
        $row.SignInEventTypes | Should -Be 'nonInteractiveUser'
    }

    It 'writes one Conditional Access row per applied policy and records that the detail was readable' {
        Mock Get-MgContext -MockWith { [pscustomobject]@{ Scopes = @('AuditLog.Read.All', 'Policy.Read.All') } }
        Mock Get-MgAuditLogSignIn -MockWith { New-MockSignIn }

        Invoke-CollectorScript 'Get-SignInConditionalAccess.ps1' ($script:oneDay + @{ OutputPath = $script:folder })

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:folder 'signin-conditional-access.csv'))
        $rows.Count | Should -Be 2
        ($rows.PolicyDisplayName | Sort-Object) | Should -Be @('Block legacy authentication', 'Require MFA for all users')
        $rows.PolicyDetailReadable | Should -Not -Contain 'False'
        ($rows | Where-Object PolicyId -eq 'policy-2').EnforcedSessionControls | Should -Be 'signInFrequency'
        $rows[0].ConditionalAccessStatus | Should -Be 'success'
    }

    It 'keeps the status of a sign-in with no policy detail and says the detail was not readable' {
        # With AuditLog.Read.All alone appliedConditionalAccessPolicies is omitted without
        # an error. https://learn.microsoft.com/graph/api/signin-list#permissions
        Mock Get-MgContext -MockWith { [pscustomobject]@{ Scopes = @('AuditLog.Read.All') } }
        Mock Get-MgAuditLogSignIn -MockWith { New-MockSignIn -WithPolicies $false }

        Invoke-CollectorScript 'Get-SignInConditionalAccess.ps1' ($script:oneDay + @{ OutputPath = $script:folder })

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:folder 'signin-conditional-access.csv'))
        $rows.Count | Should -Be 1
        $rows[0].ConditionalAccessStatus | Should -Be 'success'
        $rows[0].PolicyId | Should -BeNullOrEmpty
        $rows[0].PolicyDetailReadable | Should -Be 'False'
    }

    It 'treats an unreadable session as detail not readable' {
        Mock Get-MgAuditLogSignIn -MockWith { New-MockSignIn -WithPolicies $false }

        Invoke-CollectorScript 'Get-SignInConditionalAccess.ps1' ($script:oneDay + @{ OutputPath = $script:folder })

        (Import-Csv -LiteralPath (Join-Path $script:folder 'signin-conditional-access.csv')).PolicyDetailReadable | Should -Be 'False'
    }

    It 'keeps every directory audit category, operation type and result as returned, including Policy and a non-GUID id' {
        Mock Get-MgAuditLogDirectoryAudit -MockWith {
            New-MockAudit -Id 'SSGM_b662f17a-4e4d-4e1c-9248-cdec180024b2_MCDC4_88453290' -Category 'Policy'
            New-MockAudit -Id 'a2' -Category 'SomeNewCategory' -Operation 'Reset' -Result 'timeout'
            New-MockAudit -Id 'a3' -Category 'UserManagement' -Operation 'Add' -Result 'failure'
        }

        Invoke-CollectorScript 'Get-DirectoryAudits.ps1' ($script:oneDay + @{ OutputPath = $script:folder })

        $rows = Import-Csv -LiteralPath (Join-Path $script:folder 'directory-audits.csv')
        $rows.Id | Should -Contain 'SSGM_b662f17a-4e4d-4e1c-9248-cdec180024b2_MCDC4_88453290'
        ($rows.Category | Sort-Object) | Should -Be @('Policy', 'SomeNewCategory', 'UserManagement')
        $rows.OperationType | Should -Contain 'Reset'
        ($rows.Result | Sort-Object) | Should -Be @('failure', 'success', 'timeout')
        $rows[0].TargetResourceTypes | Should -Be 'Policy'
        $rows[0].InitiatedByUserPrincipalName | Should -Be 'admin.alvarez@example.com'
        $rows[0].ModifiedProperties | Should -Match 'State'
    }
}

Describe 'An event collector resumes from its watermark' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
    }
    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It '<Script> starts at the latest <Column> already exported and does not repeat that row' -ForEach @(
        @{ Script = 'Get-InteractiveSignIns.ps1'; Csv = 'signins-interactive.csv'; Cmdlet = 'Get-MgAuditLogSignIn'; Builder = 'New-MockSignIn'; Column = 'CreatedDateTime'; Name = 'createdDateTime' }
        @{ Script = 'Get-NonInteractiveSignIns.ps1'; Csv = 'signins-noninteractive.csv'; Cmdlet = 'Get-MgBetaAuditLogSignIn'; Builder = 'New-MockSignIn'; Column = 'CreatedDateTime'; Name = 'createdDateTime' }
        @{ Script = 'Get-SignInConditionalAccess.ps1'; Csv = 'signin-conditional-access.csv'; Cmdlet = 'Get-MgAuditLogSignIn'; Builder = 'New-MockSignIn'; Column = 'CreatedDateTime'; Name = 'createdDateTime' }
        @{ Script = 'Get-DirectoryAudits.ps1'; Csv = 'directory-audits.csv'; Cmdlet = 'Get-MgAuditLogDirectoryAudit'; Builder = 'New-MockAudit'; Column = 'ActivityDateTime'; Name = 'activityDateTime' }
    ) {
        $builder = $Builder
        Mock -CommandName $Cmdlet -MockWith { & $builder -Id 'row-1' -At ([datetime]'2026-09-10T06:05:00Z') }
        Invoke-CollectorScript $Script @{
            OutputPath = $script:folder; StartDate = [datetime]'2026-09-10T00:00:00Z'; EndDate = [datetime]'2026-09-10T12:00:00Z'
        }

        Mock -CommandName $Cmdlet -MockWith {
            & $builder -Id 'row-1' -At ([datetime]'2026-09-10T06:05:00Z')
            & $builder -Id 'row-2' -At ([datetime]'2026-09-10T08:00:00Z')
        }
        Invoke-CollectorScript $Script @{ OutputPath = $script:folder; EndDate = [datetime]'2026-09-10T12:00:00Z' }

        $expected = "$Name ge 2026-09-10T06:05:00Z"
        Should -Invoke -CommandName $Cmdlet -ParameterFilter { $Filter.Contains($expected) }
        $csv = Join-Path $script:folder $Csv
        @(Import-Csv -LiteralPath $csv | Where-Object { $_.$Column -eq '2026-09-10T06:05:00Z' }).Count | Should -BeGreaterThan 0
        $keys = @(Import-Csv -LiteralPath $csv | ForEach-Object { if ($_.PSObject.Properties['SignInId']) { $_.SignInId } else { $_.Id } }) | Sort-Object -Unique
        $keys | Should -Be @('row-1', 'row-2')
    }

    It 'keeps complete windows and stops, with an error, when a later window fails' {
        Mock Get-MgAuditLogSignIn -MockWith {
            if ($Filter -match 'createdDateTime ge 2026-09-11') { throw 'Request timed out.' }
            New-MockSignIn -Id 'signin-1'
        }

        { Invoke-CollectorScript 'Get-InteractiveSignIns.ps1' @{
                OutputPath = $script:folder; StartDate = [datetime]'2026-09-10T00:00:00Z'; EndDate = [datetime]'2026-09-12T00:00:00Z'
            } 3>$null } | Should -Throw '*part-way*'

        @(Import-Csv -LiteralPath (Join-Path $script:folder 'signins-interactive.csv')).Count | Should -Be 1
        Get-LogText -Folder $script:folder | Should -Match 'Request timed out'
    }

    It 'rejects an empty range that was asked for explicitly' {
        Mock Get-MgAuditLogSignIn -MockWith { @() }
        { Invoke-CollectorScript 'Get-InteractiveSignIns.ps1' @{
                OutputPath = $script:folder; StartDate = [datetime]'2026-09-12T00:00:00Z'; EndDate = [datetime]'2026-09-10T00:00:00Z'
            } } | Should -Throw '*range is empty*'
    }

    It 'looks back -LookbackDays on a first run' {
        Mock Get-MgAuditLogDirectoryAudit -MockWith { @() }

        Invoke-CollectorScript 'Get-DirectoryAudits.ps1' @{ OutputPath = $script:folder; LookbackDays = 2 }

        $first = [datetime]::UtcNow.AddDays(-2).ToString('yyyy-MM-dd')
        Should -Invoke Get-MgAuditLogDirectoryAudit -ParameterFilter { $Filter.StartsWith("activityDateTime ge $first") }
    }
}

Describe 'A state source stamps the run date' {
    BeforeEach {
        $script:folder = New-TestFolder
    }
    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'stamps every retention row with the UTC run date and appends a block per run date, never repeating one' {
        Invoke-CollectorScript 'Get-RetentionReference.ps1' @{ OutputPath = $script:folder } 3>$null
        Invoke-CollectorScript 'Get-RetentionReference.ps1' @{ OutputPath = $script:folder } 3>$null

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:folder 'retention-reference.csv'))
        $rows.Count | Should -Be 3
        ($rows.RunDate | Sort-Object -Unique) | Should -Be ([datetime]::UtcNow.ToString('yyyy-MM-dd'))
    }

    It 'writes the documented windows: 7 days on Free, 30 on P1 and P2, 90 for P2 risky sign-ins' {
        Invoke-CollectorScript 'Get-RetentionReference.ps1' @{ OutputPath = $script:folder } 3>$null

        $rows = Import-Csv -LiteralPath (Join-Path $script:folder 'retention-reference.csv')
        $free = $rows | Where-Object LicenseLevel -eq 'Free'
        $p2 = $rows | Where-Object LicenseLevel -eq 'P2'
        $free.SignInRetentionDays | Should -Be '7'
        $free.AuditRetentionDays | Should -Be '7'
        ($rows | Where-Object LicenseLevel -eq 'P1').SignInRetentionDays | Should -Be '30'
        $p2.AuditRetentionDays | Should -Be '30'
        $p2.RiskySignInRetentionDays | Should -Be '90'
    }
}

Describe 'A source Microsoft documents as NotAvailable writes the header only' {
    BeforeEach {
        $script:folder = New-TestFolder
        # The contract marks no source NotAvailable today, so the tests point the
        # collectors at a copy of the schema that does.
        $text = Get-Content -LiteralPath (Join-Path $script:Collectors 'EntraActivitySchema.psd1') -Raw
        $script:notAvailable = Join-Path $script:folder 'schema.psd1'
        $changed = $text -replace "Status = '(Available|Unverified)'", "Status = 'NotAvailable'"
        Set-Content -LiteralPath $script:notAvailable -Value $changed -Encoding utf8
        Mock Connect-MgGraph -ModuleName M365ReportLibrary -MockWith { }
        Mock Get-MgAuditLogSignIn -MockWith { throw 'Get-MgAuditLogSignIn must not be called' }
        Mock Get-MgBetaAuditLogSignIn -MockWith { throw 'Get-MgBetaAuditLogSignIn must not be called' }
        Mock Get-MgAuditLogDirectoryAudit -MockWith { throw 'Get-MgAuditLogDirectoryAudit must not be called' }
    }
    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It '<Script> skips, signs in to nothing and logs why' -ForEach @(
        @{ Script = 'Get-InteractiveSignIns.ps1'; Csv = 'signins-interactive.csv'; Key = 'SignIns' }
        @{ Script = 'Get-NonInteractiveSignIns.ps1'; Csv = 'signins-noninteractive.csv'; Key = 'SignIns' }
        @{ Script = 'Get-SignInConditionalAccess.ps1'; Csv = 'signin-conditional-access.csv'; Key = 'SignInConditionalAccess' }
        @{ Script = 'Get-DirectoryAudits.ps1'; Csv = 'directory-audits.csv'; Key = 'DirectoryAudits' }
        @{ Script = 'Get-RetentionReference.ps1'; Csv = 'retention-reference.csv'; Key = 'RetentionReference' }
    ) {
        Invoke-CollectorScript $Script @{ OutputPath = $script:folder; SchemaPath = $script:notAvailable } 3>$null

        $path = Join-Path $script:folder $Csv
        (Get-Content -LiteralPath $path).Count | Should -Be 1
        Get-HeaderText -Path $path | Should -Be ($script:Schema[$Key] -join ',')
        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 0 -Exactly
        Get-LogText -Folder $script:folder | Should -Match 'documented as unavailable'
    }
}

Describe 'An UNVERIFIED source is attempted with a warning' {
    BeforeEach {
        $script:folder = New-TestFolder
        $text = Get-Content -LiteralPath (Join-Path $script:Collectors 'EntraActivitySchema.psd1') -Raw
        $script:unverified = Join-Path $script:folder 'schema.psd1'
        Set-Content -LiteralPath $script:unverified -Value ($text -replace "Status = 'Available'", "Status = 'Unverified'") -Encoding utf8
        Mock Connect-M365Service -MockWith { }
    }
    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'attempts an unverified sign-in source and warns' {
        Mock Get-MgAuditLogSignIn -MockWith { New-MockSignIn }

        Invoke-CollectorScript 'Get-InteractiveSignIns.ps1' @{
            OutputPath = $script:folder; SchemaPath = $script:unverified
            StartDate = [datetime]'2026-09-10T00:00:00Z'; EndDate = [datetime]'2026-09-11T00:00:00Z'
        } 3>$null

        Should -Invoke Get-MgAuditLogSignIn -Times 1 -Exactly -ParameterFilter { $All }
        @(Import-Csv -LiteralPath (Join-Path $script:folder 'signins-interactive.csv')).Count | Should -Be 1
        Get-LogText -Folder $script:folder | Should -Match 'UNVERIFIED'
    }

    It 'warns that log retention is UNVERIFIED in every cloud and still writes the windows' -ForEach @(
        @{ Environment = 'Commercial' }
        @{ Environment = 'GCC' }
        @{ Environment = 'GCCHigh' }
    ) {
        Invoke-CollectorScript 'Get-RetentionReference.ps1' @{ OutputPath = $script:folder; Environment = $Environment } 3>$null

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:folder 'retention-reference.csv'))
        $rows.Count | Should -Be 3
        $rows.Environment | Sort-Object -Unique | Should -Be $Environment
        $rows.RetentionStatus | Sort-Object -Unique | Should -Be 'UNVERIFIED'
        Get-LogText -Folder $script:folder | Should -Match 'UNVERIFIED'
    }

    It 'does not warn about an Available source' {
        Mock Get-MgAuditLogDirectoryAudit -MockWith { New-MockAudit }

        Invoke-CollectorScript 'Get-DirectoryAudits.ps1' @{ OutputPath = $script:folder; LookbackDays = 1 }

        Get-LogText -Folder $script:folder | Should -Not -Match 'UNVERIFIED'
    }
}

Describe 'A refusal is logged, not hidden' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
    }
    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'a missing licence is a logged skip with the header only, saying "not licensed"' {
        Mock Get-MgAuditLogSignIn -MockWith { throw 'Neither tenant is B2C or tenant doesn''t have premium license' }

        Invoke-CollectorScript 'Get-InteractiveSignIns.ps1' @{ OutputPath = $script:folder; LookbackDays = 1 } 3>$null

        (Get-Content -LiteralPath (Join-Path $script:folder 'signins-interactive.csv')).Count | Should -Be 1
        Get-LogText -Folder $script:folder | Should -Match 'not licensed'
    }

    It 'any other failure is logged and does not report success' {
        Mock Get-MgAuditLogDirectoryAudit -MockWith { throw 'Insufficient privileges to complete the operation.' }

        { Invoke-CollectorScript 'Get-DirectoryAudits.ps1' @{ OutputPath = $script:folder; LookbackDays = 1 } 3>$null } |
            Should -Throw '*failed*'

        Get-LogText -Folder $script:folder | Should -Match 'Insufficient privileges'
    }
}

Describe 'Run-All.ps1' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
        # The shared users collector signs in from inside the module.
        Mock Connect-MgGraph -ModuleName M365ReportLibrary -MockWith { }
        Mock Get-MgUser -ModuleName M365ReportLibrary -MockWith { @() }
        Mock Get-MgContext -MockWith { $null }
        Mock Get-MgAuditLogSignIn -MockWith { New-MockSignIn }
        Mock Get-MgBetaAuditLogSignIn -MockWith { New-MockSignIn -Interactive $false }
        Mock Get-MgAuditLogDirectoryAudit -MockWith { New-MockAudit }
    }
    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'runs every collector into one folder, with the shared users collector first' {
        Invoke-CollectorScript 'Run-All.ps1' @{ OutputPath = $script:folder; StartDate = [datetime]'2026-09-10T00:00:00Z'; EndDate = [datetime]'2026-09-11T00:00:00Z' } 3>$null

        foreach ($csv in 'users.csv', 'signins-interactive.csv', 'signins-noninteractive.csv', 'signin-conditional-access.csv', 'directory-audits.csv', 'retention-reference.csv') {
            Test-Path -LiteralPath (Join-Path $script:folder $csv) | Should -BeTrue -Because $csv
        }
        Should -Invoke Get-MgUser -ModuleName M365ReportLibrary -Times 1
    }

    It 'continues past a failing collector and then fails the run' {
        Mock Get-MgAuditLogDirectoryAudit -MockWith { throw 'Insufficient privileges to complete the operation.' }

        { Invoke-CollectorScript 'Run-All.ps1' @{ OutputPath = $script:folder; StartDate = [datetime]'2026-09-10T00:00:00Z'; EndDate = [datetime]'2026-09-11T00:00:00Z' } 3>$null } |
            Should -Throw '*1 collector(s) stopped*'

        Test-Path -LiteralPath (Join-Path $script:folder 'retention-reference.csv') | Should -BeTrue
    }
}
