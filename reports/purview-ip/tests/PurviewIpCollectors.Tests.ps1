#Requires -Version 7.0

BeforeAll {
    $script:Root = Join-Path $PSScriptRoot '../../..' | Resolve-Path | Select-Object -ExpandProperty Path
    $script:Collectors = Join-Path $script:Root 'reports/purview-ip/collectors'
    $script:Samples = Join-Path $script:Root 'reports/purview-ip/samples'

    . (Join-Path $script:Root 'shared/tests/TenantCmdletStubs.ps1')
    Import-Module (Join-Path $script:Root 'shared/M365ReportLibrary.psm1') -Force
    . (Join-Path $script:Collectors 'PurviewIpHelpers.ps1')

    $script:Schema = Import-PowerShellDataFile -LiteralPath (Join-Path $script:Collectors 'PurviewIpSchema.psd1')

    function New-TestFolder {
        $path = Join-Path ([System.IO.Path]::GetTempPath()) ('purview-ip-' + [guid]::NewGuid().ToString('N'))
        New-Item -Path $path -ItemType Directory -Force | Out-Null
        return $path
    }

    function Get-HeaderText {
        param([Parameter(Mandatory)][string]$Path)
        return ((Get-CsvHeaderColumn -Path $Path) -join ',')
    }

    function New-MockLabel {
        param(
            [string]$Guid = 'lbl-1',
            [string]$Name = 'General',
            [string]$ParentId = '',
            [int]$Priority = 1,
            [bool]$Disabled = $false,
            [bool]$IsDefault = $false,
            [bool]$EncryptionEnabled = $false
        )
        [pscustomobject]@{
            Guid              = $Guid
            Name              = $Name
            DisplayName       = $Name
            ParentId          = $ParentId
            Priority          = $Priority
            Disabled          = $Disabled
            Workload          = 'SharePoint,Exchange'
            IsDefault         = $IsDefault
            EncryptionEnabled = $EncryptionEnabled
            Comment           = ''
            WhenCreatedUTC    = [datetime]'2025-01-01T00:00:00Z'
            WhenChangedUTC    = [datetime]'2025-06-01T00:00:00Z'
        }
    }

    function New-MockDlpPolicy {
        param(
            [string]$Guid = 'dlp-1',
            [string]$Name = 'Global DLP',
            [string[]]$EnforcementPlanes = @(),
            [string[]]$Locations = @()
        )
        [pscustomobject]@{
            Guid                     = $Guid
            Name                     = $Name
            Priority                 = 0
            Mode                     = 'Enable'
            Enabled                  = $true
            Locations                = $Locations
            EnforcementPlanes        = $EnforcementPlanes
            ExchangeLocation         = @()
            SharePointLocation       = @()
            OneDriveLocation         = @()
            TeamsLocation            = @()
            EndpointDlpLocation      = @()
            PowerBIDlpLocation       = @()
            ThirdPartyAppDlpLocation = @()
            Comment                  = ''
            WhenCreatedUTC           = [datetime]'2025-01-01T00:00:00Z'
            WhenChangedUTC           = [datetime]'2025-06-01T00:00:00Z'
        }
    }

    function Mock-EmptyPolicyCmdlets {
        Mock Get-Label -MockWith { @() }
        Mock Get-LabelPolicy -MockWith { @() }
        Mock Get-AutoSensitivityLabelPolicy -MockWith { @() }
        Mock Get-DlpCompliancePolicy -MockWith { @() }
        Mock Get-DlpComplianceRule -MockWith { @() }
        Mock Get-RetentionCompliancePolicy -MockWith { @() }
        Mock Get-ComplianceTag -MockWith { @() }
    }

    function New-MockActivityRecord {
        param(
            [string]$RecordIdentity = 'rec-1',
            [datetime]$Happened = [datetime]'2026-09-01T10:00:00Z',
            [string]$Activity = 'Label applied',
            [string]$ActivityId = 'LabelApplied',
            [string]$LabelEventType = 'LabelApplied'
        )
        [pscustomobject]@{
            RecordIdentity = $RecordIdentity
            Happened       = $Happened
            Activity       = $Activity
            ActivityId     = $ActivityId
            Workload       = 'SharePoint'
            User           = 'avery.abara@example.com'
            LabelEventType = $LabelEventType
        }
    }

    function New-ActivityExplorerResponse {
        param([object[]]$Records = @(), [bool]$LastPage = $true, [string]$Watermark = $null)
        [pscustomobject]@{
            LastPage   = $LastPage
            Watermark  = $Watermark
            ResultData = (@($Records) | ConvertTo-Json -Depth 6)
        }
    }

    function New-MockCopilotAuditRecord {
        param(
            [string]$RecordId = 'audit-1',
            [datetime]$CreationTime = [datetime]'2026-09-01T10:00:00Z',
            [object[]]$AccessedResources = @(
                @{ Id = 'res-1'; Name = 'Doc1.docx'; Type = 'File'; Action = 'Read'; SiteUrl = 'https://example.sharepoint.com/sites/a'; SensitivityLabelId = 'lbl-1'; Status = 'success' }
            )
        )
        $auditData = @{
            CreationTime    = $CreationTime.ToString('o')
            Id              = $RecordId
            Operation       = 'CopilotInteraction'
            Workload        = 'Teams'
            UserId          = 'avery.abara@example.com'
            UserKey         = 'avery.abara@example.com'
            UserType        = 'Regular'
            AppIdentity     = 'Microsoft 365 Copilot'
            AgentId         = 'agent-1'
            AgentName       = 'Sales Agent'
            CopilotEventData = @{
                AppHost           = 'Teams'
                ThreadId          = 'thread-1'
                AccessedResources = $AccessedResources
            }
        } | ConvertTo-Json -Depth 6 -Compress

        [pscustomobject]@{ AuditData = $auditData }
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

    It 'policies.csv' {
        Mock-EmptyPolicyCmdlets
        Mock Get-Label -MockWith { New-MockLabel }

        & (Join-Path $script:Collectors 'Get-Policies.ps1') -OutputPath $script:folder

        $produced = Join-Path $script:folder 'policies.csv'
        @(Import-Csv -LiteralPath $produced).Count | Should -BeGreaterThan 0
        Get-HeaderText -Path $produced | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples 'policies.csv'))
        Get-HeaderText -Path $produced | Should -Be ($script:Schema.Policies -join ',')
    }

    It 'activity-explorer-events.csv' {
        Mock Export-ActivityExplorerData -MockWith { New-ActivityExplorerResponse -Records @(New-MockActivityRecord) }

        & (Join-Path $script:Collectors 'Get-ActivityExplorerEvents.ps1') -OutputPath $script:folder -RetentionDays 1

        $produced = Join-Path $script:folder 'activity-explorer-events.csv'
        @(Import-Csv -LiteralPath $produced).Count | Should -BeGreaterThan 0
        Get-HeaderText -Path $produced | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples 'activity-explorer-events.csv'))
    }

    It 'content-explorer-snapshot.csv' {
        Mock Get-Label -MockWith { New-MockLabel }
        Mock Get-ComplianceTag -MockWith { @() }
        Mock Get-DlpSensitiveInformationType -MockWith { @() }
        Mock Export-ContentExplorerData -MockWith { [pscustomobject]@{ TotalCount = 5 } }

        & (Join-Path $script:Collectors 'Get-ContentExplorerSnapshot.ps1') -OutputPath $script:folder

        $produced = Join-Path $script:folder 'content-explorer-snapshot.csv'
        @(Import-Csv -LiteralPath $produced).Count | Should -BeGreaterThan 0
        Get-HeaderText -Path $produced | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples 'content-explorer-snapshot.csv'))
    }

    It 'copilot-accessed-resources.csv' {
        Mock Search-UnifiedAuditLog -MockWith { New-MockCopilotAuditRecord }

        & (Join-Path $script:Collectors 'Get-CopilotAccessedResources.ps1') -OutputPath $script:folder -LookbackDays 1

        $produced = Join-Path $script:folder 'copilot-accessed-resources.csv'
        @(Import-Csv -LiteralPath $produced).Count | Should -BeGreaterThan 0
        Get-HeaderText -Path $produced | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples 'copilot-accessed-resources.csv'))
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

Describe 'Get-Policies.ps1' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
        Mock-EmptyPolicyCmdlets
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'marks a DLP policy scoped to Copilot through the CopilotExperiences enforcement plane' {
        Mock Get-DlpCompliancePolicy -MockWith { New-MockDlpPolicy -EnforcementPlanes @('CopilotExperiences') }

        & (Join-Path $script:Collectors 'Get-Policies.ps1') -OutputPath $script:folder

        $row = Import-Csv -LiteralPath (Join-Path $script:folder 'policies.csv')
        $row.AppliesToCopilot | Should -Be 'True'
    }

    It 'marks a DLP policy scoped to Copilot through the location GUID' {
        Mock Get-DlpCompliancePolicy -MockWith { New-MockDlpPolicy -Locations @('470f2276-e011-4e9d-a6ec-20768be3a4b0') }

        & (Join-Path $script:Collectors 'Get-Policies.ps1') -OutputPath $script:folder

        $row = Import-Csv -LiteralPath (Join-Path $script:folder 'policies.csv')
        $row.AppliesToCopilot | Should -Be 'True'
    }

    It 'does not mark an ordinary DLP policy as scoped to Copilot' {
        Mock Get-DlpCompliancePolicy -MockWith { New-MockDlpPolicy }

        & (Join-Path $script:Collectors 'Get-Policies.ps1') -OutputPath $script:folder

        $row = Import-Csv -LiteralPath (Join-Path $script:folder 'policies.csv')
        $row.AppliesToCopilot | Should -Be 'False'
    }

    It 'resolves a sublabel''s ParentName from the parent''s Guid' {
        Mock Get-Label -MockWith {
            @(
                New-MockLabel -Guid 'parent-1' -Name 'Confidential' -Priority 2
                New-MockLabel -Guid 'child-1' -Name 'Confidential-Internal' -ParentId 'parent-1' -Priority 3
            )
        }

        & (Join-Path $script:Collectors 'Get-Policies.ps1') -OutputPath $script:folder

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:folder 'policies.csv'))
        ($rows | Where-Object ObjectId -eq 'child-1').ParentName | Should -Be 'Confidential'
    }

    It 'appends a second snapshot rather than rewriting the file' {
        Mock Get-Label -MockWith { New-MockLabel }

        & (Join-Path $script:Collectors 'Get-Policies.ps1') -OutputPath $script:folder
        & (Join-Path $script:Collectors 'Get-Policies.ps1') -OutputPath $script:folder

        # Same day, same policies: the composite RunDate+ObjectType+ObjectId key
        # keeps a re-run idempotent.
        $rows = @(Import-Csv -LiteralPath (Join-Path $script:folder 'policies.csv'))
        $rows.Count | Should -Be 1
    }

    It 'writes the header only when Security & Compliance PowerShell refuses the request' {
        Mock Get-Label -MockWith { throw 'Insufficient privileges to complete the operation.' }

        & (Join-Path $script:Collectors 'Get-Policies.ps1') -OutputPath $script:folder -WarningAction SilentlyContinue

        $csv = Join-Path $script:folder 'policies.csv'
        @(Get-Content -LiteralPath $csv).Count | Should -Be 1
        Get-Content -LiteralPath (Join-Path $script:folder 'run.log') -Raw | Should -Match 'Writing the header only'
    }
}

Describe 'Get-ActivityExplorerEvents.ps1' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'normalises a display-name activity to its filter-enum name' {
        Mock Export-ActivityExplorerData -MockWith {
            New-ActivityExplorerResponse -Records @(New-MockActivityRecord -Activity 'Label applied' -ActivityId '')
        }

        & (Join-Path $script:Collectors 'Get-ActivityExplorerEvents.ps1') -OutputPath $script:folder -RetentionDays 1

        $row = Import-Csv -LiteralPath (Join-Path $script:folder 'activity-explorer-events.csv')
        $row.Activity | Should -Be 'LabelApplied'
        $row.ActivityCategory | Should -Be 'Labeling'
    }

    It 'categorises an unmapped activity as Other and logs it' {
        Mock Export-ActivityExplorerData -MockWith {
            New-ActivityExplorerResponse -Records @(New-MockActivityRecord -Activity 'Something New' -ActivityId '' -LabelEventType '')
        }

        & (Join-Path $script:Collectors 'Get-ActivityExplorerEvents.ps1') -OutputPath $script:folder -RetentionDays 1

        $row = Import-Csv -LiteralPath (Join-Path $script:folder 'activity-explorer-events.csv')
        $row.ActivityCategory | Should -Be 'Other'
        Get-Content -LiteralPath (Join-Path $script:folder 'run.log') -Raw | Should -Match 'Could not map'
    }

    It 'marks a downgrade from LabelEventType' {
        Mock Export-ActivityExplorerData -MockWith {
            New-ActivityExplorerResponse -Records @(New-MockActivityRecord -LabelEventType 'LabelDowngraded')
        }

        & (Join-Path $script:Collectors 'Get-ActivityExplorerEvents.ps1') -OutputPath $script:folder -RetentionDays 1

        $row = Import-Csv -LiteralPath (Join-Path $script:folder 'activity-explorer-events.csv')
        $row.IsLabelDowngrade | Should -Be 'True'
    }

    It 'resumes from the watermark already in the file' {
        # Relative to real now, not a fixed date: a fixed past watermark would make the
        # window (and the number of slices below) grow with wall-clock time.
        $watermark = [datetime]::UtcNow.AddHours(-3).ToString('yyyy-MM-ddTHH:mm:ssZ')
        Export-AppendCsv -Path (Join-Path $script:folder 'activity-explorer-events.csv') -Column $script:Schema.ActivityExplorerEvents -Rows @(
            [pscustomobject]@{ RecordIdentity = 'rec-0'; Happened = $watermark }
        )
        Mock Export-ActivityExplorerData -MockWith { New-ActivityExplorerResponse }

        & (Join-Path $script:Collectors 'Get-ActivityExplorerEvents.ps1') -OutputPath $script:folder

        Should -Invoke Export-ActivityExplorerData -ParameterFilter {
            $StartTime -eq [datetime]::Parse($watermark, [cultureinfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AdjustToUniversal -bor [System.Globalization.DateTimeStyles]::AssumeUniversal)
        }
    }

    It 'skips an event it already holds when the window overlaps' {
        Mock Export-ActivityExplorerData -MockWith { New-ActivityExplorerResponse -Records @(New-MockActivityRecord -RecordIdentity 'rec-1') }

        & (Join-Path $script:Collectors 'Get-ActivityExplorerEvents.ps1') -OutputPath $script:folder -RetentionDays 1
        & (Join-Path $script:Collectors 'Get-ActivityExplorerEvents.ps1') -OutputPath $script:folder -RetentionDays 1

        @(Import-Csv -LiteralPath (Join-Path $script:folder 'activity-explorer-events.csv')).Count | Should -Be 1
    }

    It 'writes the header only when Activity Explorer is unavailable' {
        Mock Export-ActivityExplorerData -MockWith { throw 'Insufficient privileges to complete the operation.' }

        & (Join-Path $script:Collectors 'Get-ActivityExplorerEvents.ps1') -OutputPath $script:folder -RetentionDays 1 -WarningAction SilentlyContinue

        $csv = Join-Path $script:folder 'activity-explorer-events.csv'
        @(Get-Content -LiteralPath $csv).Count | Should -Be 1
        Get-HeaderText -Path $csv | Should -Be ($script:Schema.ActivityExplorerEvents -join ',')
    }
}

Describe 'Get-ContentExplorerSnapshot.ps1' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
        Mock Get-Label -MockWith { New-MockLabel }
        Mock Get-ComplianceTag -MockWith { @() }
        Mock Get-DlpSensitiveInformationType -MockWith { @() }
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'snapshots every workload in Commercial, including Teams' {
        Mock Export-ContentExplorerData -MockWith { [pscustomobject]@{ TotalCount = 12 } }

        & (Join-Path $script:Collectors 'Get-ContentExplorerSnapshot.ps1') -OutputPath $script:folder -Environment Commercial

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:folder 'content-explorer-snapshot.csv'))
        ($rows.Workload | Sort-Object -Unique) | Should -Contain 'Teams'
    }

    It 'leaves the Teams workload out in GCC High' {
        Mock Export-ContentExplorerData -MockWith { [pscustomobject]@{ TotalCount = 12 } }

        & (Join-Path $script:Collectors 'Get-ContentExplorerSnapshot.ps1') -OutputPath $script:folder -Environment GCCHigh

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:folder 'content-explorer-snapshot.csv'))
        ($rows.Workload | Sort-Object -Unique) | Should -Not -Contain 'Teams'
        Get-Content -LiteralPath (Join-Path $script:folder 'run.log') -Raw | Should -Match 'Leaving the Teams workload out'
    }

    It 'appends a new snapshot date rather than rewriting the file' {
        $prior = [datetime]::UtcNow.AddDays(-1).ToString('yyyy-MM-dd')
        Export-AppendCsv -Path (Join-Path $script:folder 'content-explorer-snapshot.csv') -Column $script:Schema.ContentExplorerSnapshot -Rows @(
            [pscustomobject]@{ RunDate = $prior; TagType = 'Sensitivity'; TagName = 'General'; Workload = 'EXO'; TotalCount = 1 }
        )
        Mock Export-ContentExplorerData -MockWith { [pscustomobject]@{ TotalCount = 12 } }

        & (Join-Path $script:Collectors 'Get-ContentExplorerSnapshot.ps1') -OutputPath $script:folder

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:folder 'content-explorer-snapshot.csv'))
        $rows.RunDate | Should -Contain $prior
        @($rows.RunDate | Sort-Object -Unique).Count | Should -Be 2
    }
}

Describe 'Get-CopilotAccessedResources.ps1' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'flattens each accessed resource into its own row, sharing the interaction''s fields' {
        Mock Search-UnifiedAuditLog -MockWith {
            New-MockCopilotAuditRecord -AccessedResources @(
                @{ Id = 'res-1'; Name = 'Doc1.docx'; Type = 'File'; Action = 'Read'; Status = 'success' }
                @{ Id = 'res-2'; Name = 'Doc2.docx'; Type = 'File'; Action = 'Read'; Status = 'success' }
            )
        }

        & (Join-Path $script:Collectors 'Get-CopilotAccessedResources.ps1') -OutputPath $script:folder -LookbackDays 1

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:folder 'copilot-accessed-resources.csv'))
        $rows.Count | Should -Be 2
        ($rows.RecordId | Sort-Object -Unique) | Should -Be @('audit-1')
        ($rows.ResourceId | Sort-Object) | Should -Be @('res-1', 'res-2')
    }

    It 'keeps one row with empty resource columns for an interaction that touched nothing' {
        Mock Search-UnifiedAuditLog -MockWith { New-MockCopilotAuditRecord -AccessedResources @() }

        & (Join-Path $script:Collectors 'Get-CopilotAccessedResources.ps1') -OutputPath $script:folder -LookbackDays 1

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:folder 'copilot-accessed-resources.csv'))
        $rows.Count | Should -Be 1
        $rows[0].ResourceId | Should -BeNullOrEmpty
    }

    It 'marks AccessBlocked when a policy is named on the resource' {
        Mock Search-UnifiedAuditLog -MockWith {
            New-MockCopilotAuditRecord -AccessedResources @(
                @{
                    Id            = 'res-1'; Name = 'Doc1.docx'; Type = 'File'; Action = 'Read'; Status = 'failure'
                    PolicyDetails = @{ PolicyId = 'dlp-1'; PolicyName = 'Copilot data protection'; Rules = @('Block Copilot on Restricted label') }
                }
            )
        }

        & (Join-Path $script:Collectors 'Get-CopilotAccessedResources.ps1') -OutputPath $script:folder -LookbackDays 1

        $row = Import-Csv -LiteralPath (Join-Path $script:folder 'copilot-accessed-resources.csv')
        $row.AccessBlocked | Should -Be 'True'
        $row.PolicyName | Should -Be 'Copilot data protection'
        $row.PolicyRules | Should -Be 'Block Copilot on Restricted label'
    }

    It 'resumes from the watermark already in the file' {
        # Relative to real now, not a fixed date: a fixed past watermark would make the
        # window (and the number of retried windows below) grow with wall-clock time.
        $watermark = [datetime]::UtcNow.AddHours(-2).ToString('yyyy-MM-ddTHH:mm:ssZ')
        Export-AppendCsv -Path (Join-Path $script:folder 'copilot-accessed-resources.csv') -Column $script:Schema.CopilotAccessedResources -Rows @(
            [pscustomobject]@{ RecordId = 'audit-0'; CreationTime = $watermark }
        )
        Mock Search-UnifiedAuditLog -MockWith { @() }

        & (Join-Path $script:Collectors 'Get-CopilotAccessedResources.ps1') -OutputPath $script:folder

        Should -Invoke Search-UnifiedAuditLog -ParameterFilter {
            $StartDate -eq [datetime]::Parse($watermark, [cultureinfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AdjustToUniversal -bor [System.Globalization.DateTimeStyles]::AssumeUniversal)
        }
    }

    It 'does not write a window whose ResultCount exceeds the session cap' {
        Mock Search-UnifiedAuditLog -MockWith {
            $record = New-MockCopilotAuditRecord -AccessedResources @()
            $record | Add-Member -NotePropertyName ResultCount -NotePropertyValue 60000
            $record
        }

        {
            & (Join-Path $script:Collectors 'Get-CopilotAccessedResources.ps1') -OutputPath $script:folder -LookbackDays 1
        } | Should -Throw '*50,000*'
    }

    It 'disconnects Exchange Online when it opened the session' {
        Mock Disconnect-ExchangeOnline -MockWith { }
        Mock Search-UnifiedAuditLog -MockWith { @() }

        & (Join-Path $script:Collectors 'Get-CopilotAccessedResources.ps1') -OutputPath $script:folder -LookbackDays 1

        Should -Invoke Disconnect-ExchangeOnline -Times 1 -Exactly
    }

    It 'leaves the session alone when -SkipConnect is given' {
        Mock Disconnect-ExchangeOnline -MockWith { }
        Mock Search-UnifiedAuditLog -MockWith { @() }

        & (Join-Path $script:Collectors 'Get-CopilotAccessedResources.ps1') -OutputPath $script:folder -LookbackDays 1 -SkipConnect

        Should -Invoke Connect-M365Service -Times 0 -Exactly
        Should -Invoke Disconnect-ExchangeOnline -Times 0 -Exactly
    }
}

Describe 'Collectors connect to the cloud they were asked for' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
        Mock-EmptyPolicyCmdlets
        Mock Get-DlpSensitiveInformationType -MockWith { @() }
        Mock Export-ActivityExplorerData -MockWith { New-ActivityExplorerResponse }
        Mock Export-ContentExplorerData -MockWith { [pscustomobject]@{ TotalCount = 0 } }
        Mock Search-UnifiedAuditLog -MockWith { @() }
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It '<CollectorScript> asks for the <ExpectedService> service in <Cloud>' -ForEach @(
        @{ CollectorScript = 'Get-Policies.ps1'; ExpectedService = 'SecurityCompliance'; Cloud = 'Commercial'; Extra = @{} }
        @{ CollectorScript = 'Get-Policies.ps1'; ExpectedService = 'SecurityCompliance'; Cloud = 'GCCHigh'; Extra = @{} }
        @{ CollectorScript = 'Get-ActivityExplorerEvents.ps1'; ExpectedService = 'SecurityCompliance'; Cloud = 'GCC'; Extra = @{ RetentionDays = 1 } }
        @{ CollectorScript = 'Get-ContentExplorerSnapshot.ps1'; ExpectedService = 'SecurityCompliance'; Cloud = 'GCCHigh'; Extra = @{} }
        @{ CollectorScript = 'Get-CopilotAccessedResources.ps1'; ExpectedService = 'ExchangeOnline'; Cloud = 'Commercial'; Extra = @{ LookbackDays = 1 } }
        @{ CollectorScript = 'Get-CopilotAccessedResources.ps1'; ExpectedService = 'ExchangeOnline'; Cloud = 'GCCHigh'; Extra = @{ LookbackDays = 1 } }
    ) {
        $params = @{
            OutputPath    = $script:folder
            Environment   = $Cloud
            WarningAction = 'SilentlyContinue'
        } + $Extra
        & (Join-Path $script:Collectors $CollectorScript) @params

        Should -Invoke Connect-M365Service -Times 1 -Exactly -ParameterFilter {
            $Service -eq $ExpectedService -and $Environment -eq $Cloud
        }
    }

    It 'reuses the session when -SkipConnect is given' {
        & (Join-Path $script:Collectors 'Get-Policies.ps1') -OutputPath $script:folder -SkipConnect

        Should -Invoke Connect-M365Service -Times 0 -Exactly
    }
}

Describe 'A source that is unavailable in this cloud leaves a header-only CSV and a log line' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'content-explorer-snapshot.csv still writes the other workloads when Teams is dropped in GCC' {
        Mock Get-Label -MockWith { @() }
        Mock Get-ComplianceTag -MockWith { @() }
        Mock Get-DlpSensitiveInformationType -MockWith { @(New-MockLabel -Guid 'sit-1' -Name 'Credit Card Number') }
        Mock Export-ContentExplorerData -MockWith { [pscustomobject]@{ TotalCount = 3 } }

        & (Join-Path $script:Collectors 'Get-ContentExplorerSnapshot.ps1') -OutputPath $script:folder -Environment GCC

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:folder 'content-explorer-snapshot.csv'))
        $rows.Count | Should -Be 3
        ($rows.Workload | Sort-Object -Unique) | Should -Not -Contain 'Teams'
    }
}

Describe 'Run-All.ps1' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
        Mock Get-MgUser -ModuleName M365ReportLibrary -MockWith { }
        Mock Connect-MgGraph -ModuleName M365ReportLibrary -MockWith { }
        Mock-EmptyPolicyCmdlets
        Mock Get-DlpSensitiveInformationType -MockWith { @() }
        Mock Export-ActivityExplorerData -MockWith { New-ActivityExplorerResponse }
        Mock Export-ContentExplorerData -MockWith { [pscustomobject]@{ TotalCount = 0 } }
        Mock Search-UnifiedAuditLog -MockWith { @() }
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'writes every CSV of the report plus run.log' {
        & (Join-Path $script:Collectors 'Run-All.ps1') -OutputPath $script:folder -WarningAction SilentlyContinue

        foreach ($name in @('users.csv', 'policies.csv', 'activity-explorer-events.csv', 'content-explorer-snapshot.csv', 'copilot-accessed-resources.csv', 'run.log')) {
            Test-Path -LiteralPath (Join-Path $script:folder $name) | Should -BeTrue -Because "Run-All.ps1 should produce $name"
        }

        Get-Content -LiteralPath (Join-Path $script:folder 'run.log') -Raw | Should -Match 'Finished\.'
    }

    It 'keeps going when one collector fails outright, then reports the failure' {
        # A window that exceeds the 50,000-record session cap is a terminating error in
        # Get-CopilotAccessedResources.ps1, unlike a source that is merely unavailable
        # (caught inside the collector, not a run failure).
        Mock Search-UnifiedAuditLog -MockWith {
            $record = [pscustomobject]@{ AuditData = (@{ Id = 'audit-x' } | ConvertTo-Json) }
            $record | Add-Member -NotePropertyName ResultCount -NotePropertyValue 60000
            $record
        }

        {
            & (Join-Path $script:Collectors 'Run-All.ps1') -OutputPath $script:folder -WarningAction SilentlyContinue
        } | Should -Throw '*stopped with an error*'

        Test-Path -LiteralPath (Join-Path $script:folder 'policies.csv') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $script:folder 'users.csv') | Should -BeTrue
        Get-Content -LiteralPath (Join-Path $script:folder 'run.log') -Raw | Should -Match 'Finished\.'
        Get-Content -LiteralPath (Join-Path $script:folder 'run.log') -Raw | Should -Match 'copilot-accessed-resources collector stopped'
    }
}
