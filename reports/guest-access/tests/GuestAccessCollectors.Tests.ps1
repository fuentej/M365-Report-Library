#Requires -Version 7.0

BeforeAll {
    $script:Root = Join-Path $PSScriptRoot '../../..' | Resolve-Path | Select-Object -ExpandProperty Path
    $script:Collectors = Join-Path $script:Root 'reports/guest-access/collectors'
    $script:Samples = Join-Path $script:Root 'reports/guest-access/samples'

    . (Join-Path $script:Root 'shared/tests/TenantCmdletStubs.ps1')
    Import-Module (Join-Path $script:Root 'shared/M365ReportLibrary.psm1') -Force

    $script:Schema = Import-PowerShellDataFile -LiteralPath (Join-Path $script:Collectors 'GuestAccessSchema.psd1')

    function New-TestFolder {
        $path = Join-Path ([System.IO.Path]::GetTempPath()) ('guest-access-' + [guid]::NewGuid().ToString('N'))
        New-Item -Path $path -ItemType Directory -Force | Out-Null
        return $path
    }

    function New-MockGuest {
        param(
            [string]$Id = 'guest-1',
            [string]$Domain = 'partner1.example.com',
            [string]$State = 'Accepted',
            [bool]$Enabled = $true,
            [hashtable]$SignInActivity = @{
                lastSignInDateTime               = '2026-08-20T09:00:00Z'
                lastNonInteractiveSignInDateTime = '2026-08-21T09:00:00Z'
                lastSuccessfulSignInDateTime     = '2026-08-20T09:00:00Z'
            }
        )

        $additional = @{}
        if ($null -ne $SignInActivity) { $additional['signInActivity'] = $SignInActivity }

        [pscustomobject]@{
            Id                              = $Id
            DisplayName                     = "Avery Abara ($Domain)"
            Mail                            = "avery.abara@$Domain"
            UserPrincipalName               = "avery.abara_$Domain#EXT#@example.onmicrosoft.com"
            UserType                        = 'Guest'
            CreatedDateTime                 = [datetime]'2026-02-03T04:05:06Z'
            CreationType                    = 'Invitation'
            ExternalUserState               = $State
            ExternalUserStateChangeDateTime = [datetime]'2026-02-04T04:05:06Z'
            AccountEnabled                  = $Enabled
            AdditionalProperties            = $additional
        }
    }

    function New-MockDirectoryAudit {
        param([string]$Id = 'audit-1', [datetime]$When = [datetime]'2026-08-01T10:00:00Z')

        [pscustomobject]@{
            Id                  = $Id
            ActivityDateTime    = $When
            ActivityDisplayName = 'Invite external user'
            Result              = 'success'
            InitiatedBy         = [pscustomobject]@{
                AdditionalProperties = @{
                    user = @{ userPrincipalName = 'blair.bergstrom@example.com' }
                    app  = @{ displayName = 'Microsoft Entra admin center' }
                }
            }
            TargetResources     = @(
                [pscustomobject]@{
                    AdditionalProperties = @{
                        type              = 'User'
                        id                = 'guest-1'
                        userPrincipalName = 'avery.abara_partner1.example.com#EXT#@example.onmicrosoft.com'
                    }
                }
            )
        }
    }

    function New-MockSignIn {
        param([string]$Id = 'signin-1', [string]$UserId = 'guest-1', [datetime]$When = [datetime]'2026-08-15T11:00:00Z')

        [pscustomobject]@{
            Id                      = $Id
            CreatedDateTime         = $When
            UserId                  = $UserId
            UserPrincipalName       = 'avery.abara_partner1.example.com#EXT#@example.onmicrosoft.com'
            AppDisplayName          = 'Microsoft Teams'
            ResourceDisplayName     = 'Microsoft Teams Services'
            IPAddress               = '203.0.113.10'
            ClientAppUsed           = 'Browser'
            IsInteractive           = $true
            ConditionalAccessStatus = 'success'
            Location                = [pscustomobject]@{ City = 'Dublin'; CountryOrRegion = 'IE' }
            Status                  = [pscustomobject]@{ ErrorCode = 0 }
        }
    }

    function New-MockAuditRecord {
        param([string]$Id = 'share-1', [string]$Operation = 'SharingSet')

        # The unified audit log returns the detail as a JSON string in AuditData.
        $auditData = @{
            CreationTime          = '2026-08-10T12:00:00'
            Id                    = $Id
            Operation             = $Operation
            UserId                = 'blair.bergstrom@example.com'
            Workload              = 'SharePoint'
            SiteUrl               = 'https://example.sharepoint.com/sites/northwind/'
            ObjectId              = 'https://example.sharepoint.com/sites/northwind/Shared Documents/Pricing Model.xlsx'
            SourceFileName        = 'Pricing Model.xlsx'
            TargetUserOrGroupName = 'avery.abara@partner1.example.com'
            TargetUserOrGroupType = 'Guest'
        } | ConvertTo-Json -Compress

        [pscustomobject]@{ RecordType = 'SharePointSharingOperation'; AuditData = $auditData }
    }

    function New-MockMembership {
        param([string]$Id = 'group-1', [string]$Name = 'Project Northwind', [string[]]$Provisioning = @('Team'))

        [pscustomobject]@{
            Id                   = $Id
            AdditionalProperties = @{
                '@odata.type'                = '#microsoft.graph.group'
                displayName                  = $Name
                visibility                   = 'Private'
                resourceProvisioningOptions  = $Provisioning
            }
        }
    }

    function Set-TestGuestsCsv {
        <#
            .SYNOPSIS
                Seeds a guests.csv that the sign-in and membership collectors read.
        #>
        param([Parameter(Mandatory)][string]$OutputPath, [string]$Id = 'guest-1')

        $row = [pscustomobject]@{
            RunDate                          = '2026-08-01'
            Id                               = $Id
            DisplayName                      = 'Avery Abara'
            Mail                             = 'avery.abara@partner1.example.com'
            UserPrincipalName                = 'avery.abara_partner1.example.com#EXT#@example.onmicrosoft.com'
            ExternalDomain                   = 'partner1.example.com'
            CreatedDateTime                  = '2026-02-03T04:05:06Z'
            CreationType                     = 'Invitation'
            ExternalUserState                = 'Accepted'
            ExternalUserStateChangeDateTime  = '2026-02-04T04:05:06Z'
            AccountEnabled                   = $true
            LastSignInDateTime               = '2026-08-20T09:00:00Z'
            LastNonInteractiveSignInDateTime = '2026-08-21T09:00:00Z'
            LastSuccessfulSignInDateTime     = '2026-08-20T09:00:00Z'
        }

        Export-AppendCsv -Path (Join-Path $OutputPath 'guests.csv') -Rows @($row) -Column $script:Schema.Guests
    }

    function Get-HeaderText {
        param([Parameter(Mandatory)][string]$Path)
        return ((Get-CsvHeaderColumn -Path $Path) -join ',')
    }
}

AfterAll {
    Remove-Module M365ReportLibrary -Force -ErrorAction SilentlyContinue
}

Describe 'Collector output matches the committed sample files' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'guests.csv' {
        Mock Get-MgUser -MockWith { New-MockGuest }

        & (Join-Path $script:Collectors 'Get-Guests.ps1') -OutputPath $script:folder

        $produced = Join-Path $script:folder 'guests.csv'
        @(Import-Csv -LiteralPath $produced).Count | Should -BeGreaterThan 0
        Get-HeaderText -Path $produced | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples 'guests.csv'))
        Get-HeaderText -Path $produced | Should -Be ($script:Schema.Guests -join ',')
    }

    It 'guest-invitations.csv' {
        Mock Get-MgAuditLogDirectoryAudit -MockWith { New-MockDirectoryAudit }

        & (Join-Path $script:Collectors 'Get-GuestInvitations.ps1') -OutputPath $script:folder

        $produced = Join-Path $script:folder 'guest-invitations.csv'
        @(Import-Csv -LiteralPath $produced).Count | Should -BeGreaterThan 0
        Get-HeaderText -Path $produced | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples 'guest-invitations.csv'))
    }

    It 'guest-signins.csv' {
        Set-TestGuestsCsv -OutputPath $script:folder
        Mock Get-MgAuditLogSignIn -MockWith { New-MockSignIn }

        & (Join-Path $script:Collectors 'Get-GuestSignIns.ps1') -OutputPath $script:folder `
            -StartDate ([datetime]'2026-08-15T00:00:00Z') -EndDate ([datetime]'2026-08-16T00:00:00Z')

        $produced = Join-Path $script:folder 'guest-signins.csv'
        @(Import-Csv -LiteralPath $produced).Count | Should -BeGreaterThan 0
        Get-HeaderText -Path $produced | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples 'guest-signins.csv'))
    }

    It 'sharing-events.csv' {
        Mock Search-UnifiedAuditLog -MockWith { New-MockAuditRecord }

        & (Join-Path $script:Collectors 'Get-SharingEvents.ps1') -OutputPath $script:folder `
            -StartDate ([datetime]'2026-08-10T00:00:00Z') -EndDate ([datetime]'2026-08-11T00:00:00Z')

        $produced = Join-Path $script:folder 'sharing-events.csv'
        @(Import-Csv -LiteralPath $produced).Count | Should -BeGreaterThan 0
        Get-HeaderText -Path $produced | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples 'sharing-events.csv'))
    }

    It 'guest-memberships.csv' {
        Set-TestGuestsCsv -OutputPath $script:folder
        Mock Get-MgUserMemberOf -MockWith { New-MockMembership }

        & (Join-Path $script:Collectors 'Get-GuestMemberships.ps1') -OutputPath $script:folder

        $produced = Join-Path $script:folder 'guest-memberships.csv'
        @(Import-Csv -LiteralPath $produced).Count | Should -BeGreaterThan 0
        Get-HeaderText -Path $produced | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples 'guest-memberships.csv'))
    }

    It 'users.csv' {
        Mock Get-MgUser -ModuleName M365ReportLibrary -MockWith {
            [pscustomobject]@{
                Id                   = 'user-1'; DisplayName = 'Avery Abara'
                UserPrincipalName    = 'avery.abara@example.com'; Mail = 'avery.abara@example.com'
                UserType             = 'Member'; AccountEnabled = $true
                CreatedDateTime      = [datetime]'2025-01-02T03:04:05Z'
                Department           = 'Engineering'; JobTitle = 'Specialist'; City = 'Seattle'; Country = 'US'
                Manager              = $null
                AdditionalProperties = @{}
            }
        }

        Invoke-EntraUserCollector -OutputPath $script:folder -SkipConnect

        $produced = Join-Path $script:folder 'users.csv'
        Get-HeaderText -Path $produced | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples 'users.csv'))
    }
}

Describe 'Get-Guests.ps1' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'reads the home domain and the sign-in activity' {
        Mock Get-MgUser -MockWith { New-MockGuest -Domain 'partner7.example.com' }

        & (Join-Path $script:Collectors 'Get-Guests.ps1') -OutputPath $script:folder

        $row = Import-Csv -LiteralPath (Join-Path $script:folder 'guests.csv')
        $row.ExternalDomain | Should -Be 'partner7.example.com'
        $row.LastSignInDateTime | Should -Be '2026-08-20T09:00:00Z'
        $row.LastNonInteractiveSignInDateTime | Should -Be '2026-08-21T09:00:00Z'
        $row.LastSuccessfulSignInDateTime | Should -Be '2026-08-20T09:00:00Z'
    }

    It 'falls back to filtering userType in the script when Graph rejects the combination' {
        Mock Get-MgUser -MockWith { throw "Combination of 'signInActivity' with '`$filter' is not supported." } `
            -ParameterFilter { -not [string]::IsNullOrEmpty($Filter) -and $Property -contains 'signInActivity' }
        Mock Get-MgUser -MockWith {
            @(
                New-MockGuest -Id 'guest-1'
                [pscustomobject]@{
                    Id = 'member-1'; DisplayName = 'A Member'; Mail = 'a@example.com'
                    UserPrincipalName = 'a@example.com'; UserType = 'Member'
                    CreatedDateTime = [datetime]'2025-01-01T00:00:00Z'; CreationType = $null
                    ExternalUserState = $null; ExternalUserStateChangeDateTime = $null
                    AccountEnabled = $true; AdditionalProperties = @{}
                }
            )
        } -ParameterFilter { [string]::IsNullOrEmpty($Filter) }

        & (Join-Path $script:Collectors 'Get-Guests.ps1') -OutputPath $script:folder -WarningAction SilentlyContinue

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:folder 'guests.csv'))
        $rows.Count | Should -Be 1
        $rows[0].Id | Should -Be 'guest-1'
    }

    It 'leaves the sign-in columns empty and says why when signInActivity is unavailable' {
        Mock Get-MgUser -MockWith { throw 'Neither tenant is B2C or tenant doesn''t have premium license' } `
            -ParameterFilter { $Property -contains 'signInActivity' }
        Mock Get-MgUser -MockWith { New-MockGuest -SignInActivity $null } `
            -ParameterFilter { $Property -notcontains 'signInActivity' }

        & (Join-Path $script:Collectors 'Get-Guests.ps1') -OutputPath $script:folder -WarningAction SilentlyContinue

        $row = Import-Csv -LiteralPath (Join-Path $script:folder 'guests.csv')
        $row.LastSignInDateTime | Should -BeNullOrEmpty
        $row.LastNonInteractiveSignInDateTime | Should -BeNullOrEmpty
        $row.LastSuccessfulSignInDateTime | Should -BeNullOrEmpty

        Get-Content -LiteralPath (Join-Path $script:folder 'run.log') -Raw |
            Should -Match 'needs Entra ID P1 or P2'
    }

    It 'appends a second snapshot rather than rewriting the file' {
        Mock Get-MgUser -MockWith { @(New-MockGuest -Id 'guest-1'; New-MockGuest -Id 'guest-2') }

        & (Join-Path $script:Collectors 'Get-Guests.ps1') -OutputPath $script:folder
        & (Join-Path $script:Collectors 'Get-Guests.ps1') -OutputPath $script:folder

        # Same day, same guests: the composite RunDate+Id key keeps a re-run idempotent.
        $rows = @(Import-Csv -LiteralPath (Join-Path $script:folder 'guests.csv'))
        $rows.Count | Should -Be 2
        (Get-Content -LiteralPath (Join-Path $script:folder 'guests.csv') | Select-String -Pattern '^"RunDate"').Count |
            Should -Be 1
    }
}

Describe 'Event collectors resume from the watermark' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
        # A recent watermark keeps the number of query windows small.
        $script:watermark = [datetime]::UtcNow.AddHours(-30).ToString('yyyy-MM-ddTHH:mm:00Z')
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'Get-GuestInvitations.ps1 starts at the latest ActivityDateTime already collected' {
        Export-AppendCsv -Path (Join-Path $script:folder 'guest-invitations.csv') -Column $script:Schema.GuestInvitations -Rows @(
            [pscustomobject]@{
                ActivityDateTime = $script:watermark; Id = 'audit-0'
                ActivityDisplayName = 'Invite external user'; Result = 'success'
                InitiatedByUserPrincipalName = 'blair.bergstrom@example.com'
                InitiatedByAppDisplayName = 'Microsoft Entra admin center'
                TargetUserId = 'guest-0'; TargetUserPrincipalName = 'old@example.com'
            }
        )

        Mock Get-MgAuditLogDirectoryAudit -MockWith { }

        & (Join-Path $script:Collectors 'Get-GuestInvitations.ps1') -OutputPath $script:folder

        Should -Invoke Get-MgAuditLogDirectoryAudit -Times 1 -Exactly -ParameterFilter {
            $Filter -like "*activityDateTime ge $($script:watermark)*"
        }
    }

    It 'Get-GuestSignIns.ps1 starts at the latest CreatedDateTime already collected' {
        Set-TestGuestsCsv -OutputPath $script:folder
        Export-AppendCsv -Path (Join-Path $script:folder 'guest-signins.csv') -Column $script:Schema.GuestSignIns -Rows @(
            [pscustomobject]@{
                CreatedDateTime = $script:watermark; Id = 'signin-0'; UserId = 'guest-1'
                UserPrincipalName = 'avery@partner1.example.com'; AppDisplayName = 'Microsoft Teams'
                ResourceDisplayName = 'Microsoft Teams Services'; IpAddress = '203.0.113.1'
                City = 'Dublin'; CountryOrRegion = 'IE'; ClientAppUsed = 'Browser'
                IsInteractive = $true; ErrorCode = '0'; ConditionalAccessStatus = 'success'
            }
        )

        Mock Get-MgAuditLogSignIn -MockWith { }

        & (Join-Path $script:Collectors 'Get-GuestSignIns.ps1') -OutputPath $script:folder

        Should -Invoke Get-MgAuditLogSignIn -Times 1 -ParameterFilter {
            $Filter -like "*createdDateTime ge $($script:watermark)*"
        }
    }

    It 'Get-SharingEvents.ps1 starts at the latest CreationTime already collected' {
        Export-AppendCsv -Path (Join-Path $script:folder 'sharing-events.csv') -Column $script:Schema.SharingEvents -Rows @(
            [pscustomobject]@{
                CreationTime = $script:watermark; Id = 'share-0'; Operation = 'SharingSet'
                UserId = 'blair.bergstrom@example.com'; Workload = 'SharePoint'
                SiteUrl = 'https://example.sharepoint.com/sites/northwind/'
                ObjectId = 'https://example.sharepoint.com/sites/northwind/Shared Documents/x.docx'
                SourceFileName = 'x.docx'; TargetUserOrGroupName = 'avery@partner1.example.com'
                TargetUserOrGroupType = 'Guest'
            }
        )

        Mock Search-UnifiedAuditLog -MockWith { }

        & (Join-Path $script:Collectors 'Get-SharingEvents.ps1') -OutputPath $script:folder

        Should -Invoke Search-UnifiedAuditLog -Times 1 -ParameterFilter {
            $StartDate -eq [datetime]::Parse($script:watermark, [cultureinfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AdjustToUniversal -bor [System.Globalization.DateTimeStyles]::AssumeUniversal)
        }
    }

    It '<CollectorScript> refuses a range that ends at or before it starts' -ForEach @(
        @{ CollectorScript = 'Get-GuestInvitations.ps1' }
        @{ CollectorScript = 'Get-GuestSignIns.ps1' }
        @{ CollectorScript = 'Get-SharingEvents.ps1' }
    ) {
        # Collecting nothing and reporting success would hide the mistake.
        Set-TestGuestsCsv -OutputPath $script:folder
        Mock Get-MgAuditLogDirectoryAudit -MockWith { }
        Mock Get-MgAuditLogSignIn -MockWith { }
        Mock Search-UnifiedAuditLog -MockWith { }

        {
            & (Join-Path $script:Collectors $CollectorScript) -OutputPath $script:folder `
                -StartDate ([datetime]'2026-08-16T00:00:00Z') -EndDate ([datetime]'2026-08-15T00:00:00Z')
        } | Should -Throw '*range is empty*'

        Should -Invoke Get-MgAuditLogDirectoryAudit -Times 0 -Exactly
        Should -Invoke Get-MgAuditLogSignIn -Times 0 -Exactly
        Should -Invoke Search-UnifiedAuditLog -Times 0 -Exactly
    }

    It 'Get-GuestInvitations.ps1 stops quietly when the watermark is already up to date' {
        # A resumed run with nothing new is not an error: leave the file alone, say so.
        $future = [datetime]::UtcNow.AddHours(2).ToString('yyyy-MM-ddTHH:mm:00Z')
        Export-AppendCsv -Path (Join-Path $script:folder 'guest-invitations.csv') -Column $script:Schema.GuestInvitations -Rows @(
            [pscustomobject]@{
                ActivityDateTime = $future; Id = 'audit-0'
                ActivityDisplayName = 'Invite external user'; Result = 'success'
                InitiatedByUserPrincipalName = 'blair.bergstrom@example.com'
                InitiatedByAppDisplayName = 'Microsoft Entra admin center'
                TargetUserId = 'guest-0'; TargetUserPrincipalName = 'old@example.com'
            }
        )

        Mock Get-MgAuditLogDirectoryAudit -MockWith { }

        { & (Join-Path $script:Collectors 'Get-GuestInvitations.ps1') -OutputPath $script:folder } | Should -Not -Throw

        Should -Invoke Get-MgAuditLogDirectoryAudit -Times 0 -Exactly
        @(Import-Csv -LiteralPath (Join-Path $script:folder 'guest-invitations.csv')).Count | Should -Be 1
        Get-Content -LiteralPath (Join-Path $script:folder 'run.log') -Raw | Should -Match 'Nothing new to collect'
    }

    It 'Get-GuestInvitations.ps1 skips an event it already holds' {
        Mock Get-MgAuditLogDirectoryAudit -MockWith { New-MockDirectoryAudit -Id 'audit-1' }

        & (Join-Path $script:Collectors 'Get-GuestInvitations.ps1') -OutputPath $script:folder
        & (Join-Path $script:Collectors 'Get-GuestInvitations.ps1') -OutputPath $script:folder

        @(Import-Csv -LiteralPath (Join-Path $script:folder 'guest-invitations.csv')).Count | Should -Be 1
    }
}

Describe 'A source that is unavailable or unlicensed leaves a header-only CSV and a log line' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'guests.csv when Graph refuses to list users' {
        Mock Get-MgUser -MockWith { throw 'Insufficient privileges to complete the operation.' }

        & (Join-Path $script:Collectors 'Get-Guests.ps1') -OutputPath $script:folder -WarningAction SilentlyContinue

        $csv = Join-Path $script:folder 'guests.csv'
        (Get-Content -LiteralPath $csv).Count | Should -Be 1
        Get-HeaderText -Path $csv | Should -Be ($script:Schema.Guests -join ',')
        Get-Content -LiteralPath (Join-Path $script:folder 'run.log') -Raw | Should -Match 'Writing the header only'
    }

    It 'guest-invitations.csv when the directory audit log is refused' {
        Mock Get-MgAuditLogDirectoryAudit -MockWith { throw 'Neither tenant is B2C or tenant doesn''t have premium license' }

        & (Join-Path $script:Collectors 'Get-GuestInvitations.ps1') -OutputPath $script:folder -WarningAction SilentlyContinue

        $csv = Join-Path $script:folder 'guest-invitations.csv'
        (Get-Content -LiteralPath $csv).Count | Should -Be 1
        Get-HeaderText -Path $csv | Should -Be ($script:Schema.GuestInvitations -join ',')
        Get-Content -LiteralPath (Join-Path $script:folder 'run.log') -Raw | Should -Match 'AuditLog.Read.All'
    }

    It 'guest-signins.csv when the sign-in log needs a licence the tenant does not have' {
        Set-TestGuestsCsv -OutputPath $script:folder
        Mock Get-MgAuditLogSignIn -MockWith { throw 'Neither tenant is B2C or tenant doesn''t have premium license' }

        & (Join-Path $script:Collectors 'Get-GuestSignIns.ps1') -OutputPath $script:folder `
            -StartDate ([datetime]'2026-08-15T00:00:00Z') -EndDate ([datetime]'2026-08-16T00:00:00Z') -WarningAction SilentlyContinue

        $csv = Join-Path $script:folder 'guest-signins.csv'
        (Get-Content -LiteralPath $csv).Count | Should -Be 1
        Get-HeaderText -Path $csv | Should -Be ($script:Schema.GuestSignIns -join ',')
        Get-Content -LiteralPath (Join-Path $script:folder 'run.log') -Raw | Should -Match 'Entra ID P1 or P2'
    }

    It 'guest-signins.csv when guests.csv has not been collected yet' {
        & (Join-Path $script:Collectors 'Get-GuestSignIns.ps1') -OutputPath $script:folder -WarningAction SilentlyContinue

        (Get-Content -LiteralPath (Join-Path $script:folder 'guest-signins.csv')).Count | Should -Be 1
        Get-Content -LiteralPath (Join-Path $script:folder 'run.log') -Raw | Should -Match 'Run Get-Guests.ps1 first'
    }

    It 'sharing-events.csv when auditing is off or the role is missing' {
        Mock Search-UnifiedAuditLog -MockWith { throw 'The term ''Search-UnifiedAuditLog'' is not recognized' }

        & (Join-Path $script:Collectors 'Get-SharingEvents.ps1') -OutputPath $script:folder `
            -StartDate ([datetime]'2026-08-10T00:00:00Z') -EndDate ([datetime]'2026-08-11T00:00:00Z') -WarningAction SilentlyContinue

        $csv = Join-Path $script:folder 'sharing-events.csv'
        (Get-Content -LiteralPath $csv).Count | Should -Be 1
        Get-HeaderText -Path $csv | Should -Be ($script:Schema.SharingEvents -join ',')
        Get-Content -LiteralPath (Join-Path $script:folder 'run.log') -Raw | Should -Match 'View-Only Audit Logs'
    }

    It 'guest-memberships.csv when group membership cannot be read for any guest' {
        Set-TestGuestsCsv -OutputPath $script:folder
        Mock Get-MgUserMemberOf -MockWith { throw 'Insufficient privileges to complete the operation.' }

        & (Join-Path $script:Collectors 'Get-GuestMemberships.ps1') -OutputPath $script:folder -WarningAction SilentlyContinue

        $csv = Join-Path $script:folder 'guest-memberships.csv'
        (Get-Content -LiteralPath $csv).Count | Should -Be 1
        Get-HeaderText -Path $csv | Should -Be ($script:Schema.GuestMemberships -join ',')
        Get-Content -LiteralPath (Join-Path $script:folder 'run.log') -Raw | Should -Match 'GroupMember.Read.All'
    }
}

Describe 'Collectors connect to the cloud they were asked for' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
        Mock Get-MgUser -MockWith { }
        Mock Get-MgAuditLogDirectoryAudit -MockWith { }
        Mock Get-MgAuditLogSignIn -MockWith { }
        Mock Get-MgUserMemberOf -MockWith { }
        Mock Search-UnifiedAuditLog -MockWith { }
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    # The data keys are deliberately not named Service or Environment: inside a
    # -ParameterFilter those names are the mocked call's own bound parameters.
    It '<CollectorScript> asks for the <ExpectedService> service in <Cloud>' -ForEach @(
        @{ CollectorScript = 'Get-Guests.ps1'; ExpectedService = 'Graph'; Cloud = 'Commercial' }
        @{ CollectorScript = 'Get-Guests.ps1'; ExpectedService = 'Graph'; Cloud = 'GCC' }
        @{ CollectorScript = 'Get-Guests.ps1'; ExpectedService = 'Graph'; Cloud = 'GCCHigh' }
        @{ CollectorScript = 'Get-GuestInvitations.ps1'; ExpectedService = 'Graph'; Cloud = 'GCCHigh' }
        @{ CollectorScript = 'Get-GuestSignIns.ps1'; ExpectedService = 'Graph'; Cloud = 'GCCHigh' }
        @{ CollectorScript = 'Get-GuestMemberships.ps1'; ExpectedService = 'Graph'; Cloud = 'GCC' }
        @{ CollectorScript = 'Get-SharingEvents.ps1'; ExpectedService = 'ExchangeOnline'; Cloud = 'Commercial' }
        @{ CollectorScript = 'Get-SharingEvents.ps1'; ExpectedService = 'ExchangeOnline'; Cloud = 'GCCHigh' }
    ) {
        Set-TestGuestsCsv -OutputPath $script:folder

        & (Join-Path $script:Collectors $CollectorScript) -OutputPath $script:folder -Environment $Cloud -WarningAction SilentlyContinue

        Should -Invoke Connect-M365Service -Times 1 -Exactly -ParameterFilter {
            $Service -eq $ExpectedService -and $Environment -eq $Cloud
        }
    }

    It 'reuses the session when -SkipConnect is given' {
        & (Join-Path $script:Collectors 'Get-Guests.ps1') -OutputPath $script:folder -SkipConnect -WarningAction SilentlyContinue

        Should -Invoke Connect-M365Service -Times 0 -Exactly
    }
}

Describe 'Get-SharingEvents.ps1' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'searches with a session id and ReturnLargeSet, for the sharing operations only' {
        Mock Search-UnifiedAuditLog -MockWith { New-MockAuditRecord }

        & (Join-Path $script:Collectors 'Get-SharingEvents.ps1') -OutputPath $script:folder `
            -StartDate ([datetime]'2026-08-10T00:00:00Z') -EndDate ([datetime]'2026-08-11T00:00:00Z')

        Should -Invoke Search-UnifiedAuditLog -Times 1 -Exactly -ParameterFilter {
            $SessionCommand -eq 'ReturnLargeSet' -and
            -not [string]::IsNullOrEmpty($SessionId) -and
            $Operations -contains 'AnonymousLinkUsed' -and
            $Operations -contains 'SecureLinkCreated' -and
            $Operations.Count -eq 11
        }
    }

    It 'splits the range into windows so no one search covers too much' {
        Mock Search-UnifiedAuditLog -MockWith { }

        & (Join-Path $script:Collectors 'Get-SharingEvents.ps1') -OutputPath $script:folder `
            -StartDate ([datetime]'2026-08-10T00:00:00Z') -EndDate ([datetime]'2026-08-13T00:00:00Z') -WindowHours 24

        Should -Invoke Search-UnifiedAuditLog -Times 3 -Exactly
    }

    It 'reads the columns out of each record''s AuditData' {
        Mock Search-UnifiedAuditLog -MockWith { New-MockAuditRecord -Id 'share-9' -Operation 'AnonymousLinkUsed' }

        & (Join-Path $script:Collectors 'Get-SharingEvents.ps1') -OutputPath $script:folder `
            -StartDate ([datetime]'2026-08-10T00:00:00Z') -EndDate ([datetime]'2026-08-11T00:00:00Z')

        $row = Import-Csv -LiteralPath (Join-Path $script:folder 'sharing-events.csv')
        $row.Id | Should -Be 'share-9'
        $row.Operation | Should -Be 'AnonymousLinkUsed'
        $row.SourceFileName | Should -Be 'Pricing Model.xlsx'
        $row.TargetUserOrGroupType | Should -Be 'Guest'
        $row.CreationTime | Should -Be '2026-08-10T12:00:00Z'
    }
}

Describe 'Get-GuestMemberships.ps1' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
        Set-TestGuestsCsv -OutputPath $script:folder
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'marks a group as a Team only when resourceProvisioningOptions says so' {
        Mock Get-MgUserMemberOf -MockWith {
            @(
                New-MockMembership -Id 'group-1' -Name 'Project Northwind' -Provisioning @('Team')
                New-MockMembership -Id 'group-2' -Name 'Budget Planning' -Provisioning @()
            )
        }

        & (Join-Path $script:Collectors 'Get-GuestMemberships.ps1') -OutputPath $script:folder

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:folder 'guest-memberships.csv'))
        ($rows | Where-Object GroupId -eq 'group-1').IsTeam | Should -Be 'True'
        ($rows | Where-Object GroupId -eq 'group-2').IsTeam | Should -Be 'False'
    }

    It 'ignores directory objects that are not groups' {
        Mock Get-MgUserMemberOf -MockWith {
            @(
                New-MockMembership -Id 'group-1'
                [pscustomobject]@{
                    Id                   = 'role-1'
                    AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.directoryRole'; displayName = 'Guest Inviter' }
                }
            )
        }

        & (Join-Path $script:Collectors 'Get-GuestMemberships.ps1') -OutputPath $script:folder

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:folder 'guest-memberships.csv'))
        $rows.Count | Should -Be 1
        $rows[0].GroupId | Should -Be 'group-1'
    }
}

Describe 'Run-All.ps1' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
        Mock Get-MgUser -MockWith { New-MockGuest }
        # Invoke-EntraUserCollector runs inside the module, so its own calls have to be
        # mocked there.
        Mock Get-MgUser -ModuleName M365ReportLibrary -MockWith { }
        Mock Connect-MgGraph -ModuleName M365ReportLibrary -MockWith { }
        Mock Get-MgAuditLogDirectoryAudit -MockWith { New-MockDirectoryAudit }
        Mock Get-MgAuditLogSignIn -MockWith { New-MockSignIn }
        Mock Get-MgUserMemberOf -MockWith { New-MockMembership }
        Mock Search-UnifiedAuditLog -MockWith { New-MockAuditRecord }
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'writes every CSV of the report plus run.log' {
        & (Join-Path $script:Collectors 'Run-All.ps1') -OutputPath $script:folder `
            -StartDate ([datetime]::UtcNow.AddHours(-2)) -EndDate ([datetime]::UtcNow) -WarningAction SilentlyContinue

        foreach ($name in @('users.csv', 'guests.csv', 'guest-invitations.csv', 'guest-signins.csv', 'sharing-events.csv', 'guest-memberships.csv', 'run.log')) {
            Test-Path -LiteralPath (Join-Path $script:folder $name) | Should -BeTrue -Because "Run-All.ps1 should produce $name"
        }

        Get-Content -LiteralPath (Join-Path $script:folder 'run.log') -Raw | Should -Match 'Finished\.'
    }

    It 'keeps going when one collector fails outright' {
        Mock Get-MgAuditLogDirectoryAudit -MockWith { throw 'boom' }

        & (Join-Path $script:Collectors 'Run-All.ps1') -OutputPath $script:folder `
            -StartDate ([datetime]::UtcNow.AddHours(-2)) -EndDate ([datetime]::UtcNow) -WarningAction SilentlyContinue

        Test-Path -LiteralPath (Join-Path $script:folder 'guest-memberships.csv') | Should -BeTrue
        Get-Content -LiteralPath (Join-Path $script:folder 'run.log') -Raw | Should -Match 'Finished\.'
    }
}
