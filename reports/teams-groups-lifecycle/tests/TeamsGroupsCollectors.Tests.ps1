#Requires -Version 7.0

BeforeAll {
    $script:Root = Join-Path $PSScriptRoot '../../..' | Resolve-Path | Select-Object -ExpandProperty Path
    $script:Collectors = Join-Path $script:Root 'reports/teams-groups-lifecycle/collectors'
    $script:Samples = Join-Path $script:Root 'reports/teams-groups-lifecycle/samples'

    . (Join-Path $script:Root 'shared/tests/TenantCmdletStubs.ps1')
    Import-Module (Join-Path $script:Root 'shared/M365ReportLibrary.psm1') -Force

    $script:Schema = Import-PowerShellDataFile -LiteralPath (Join-Path $script:Collectors 'TeamsGroupsSchema.psd1')

    function New-TestFolder {
        $path = Join-Path ([System.IO.Path]::GetTempPath()) ('teams-groups-' + [guid]::NewGuid().ToString('N'))
        New-Item -Path $path -ItemType Directory -Force | Out-Null
        return $path
    }

    function Get-HeaderText {
        param([Parameter(Mandatory)][string]$Path)
        return ((Get-CsvHeaderColumn -Path $Path) -join ',')
    }

    # Group resource: https://learn.microsoft.com/graph/api/resources/group
    # List groups:    https://learn.microsoft.com/graph/api/group-list
    # resourceProvisioningOptions is not a typed property in the SDK model, so the SDK
    # surfaces it in AdditionalProperties; expirationDateTime and deletedDateTime are
    # nullable date-times, always UTC.
    function New-MockGroup {
        param(
            [string]$Id = 'group-1',
            [string]$Name = 'Project Northwind',
            [string[]]$Provisioning = @('Team'),
            [object]$Expiration = [datetime]'2027-02-03T04:05:06Z',
            [object]$Synced = $null
        )

        [pscustomobject]@{
            Id                    = $Id
            DisplayName           = $Name
            Mail                  = 'northwind@example.com'
            GroupTypes            = @('Unified')
            SecurityEnabled       = $false
            MailEnabled           = $true
            Visibility            = 'Private'
            CreatedDateTime       = [datetime]'2026-02-03T04:05:06Z'
            RenewedDateTime       = [datetime]'2026-02-03T04:05:06Z'
            ExpirationDateTime    = $Expiration
            DeletedDateTime       = $null
            OnPremisesSyncEnabled = $Synced
            AdditionalProperties  = @{ resourceProvisioningOptions = $Provisioning }
        }
    }

    # Soft-deleted groups: https://learn.microsoft.com/graph/api/directory-deleteditems-list
    function New-MockDeletedGroup {
        param([string]$Id = 'deleted-1', [datetime]$DeletedAt = [datetime]'2026-08-20T10:00:00Z')

        [pscustomobject]@{
            Id                   = $Id
            DisplayName          = 'Old Project'
            GroupTypes           = @('Unified')
            SecurityEnabled      = $false
            MailEnabled          = $true
            CreatedDateTime      = [datetime]'2025-01-02T03:04:05Z'
            DeletedDateTime      = $DeletedAt
            AdditionalProperties = @{ resourceProvisioningOptions = @('Team') }
        }
    }

    # List group owners: https://learn.microsoft.com/graph/api/group-list-owners
    # The collection holds directoryObjects; a user carries @odata.type, displayName
    # and userPrincipalName.
    function New-MockOwner {
        param([string]$Id = 'user-1')

        [pscustomobject]@{
            Id                   = $Id
            AdditionalProperties = @{
                '@odata.type'     = '#microsoft.graph.user'
                displayName       = 'Avery Abara'
                userPrincipalName = 'avery.abara@example.com'
            }
        }
    }

    # groupLifecyclePolicy: https://learn.microsoft.com/graph/api/resources/grouplifecyclepolicy
    function New-MockLifecyclePolicy {
        param([string]$Types = 'All')

        [pscustomobject]@{
            Id                          = 'policy-1'
            GroupLifetimeInDays         = 365
            ManagedGroupTypes           = $Types
            AlternateNotificationEmails = 'groups-admin@example.com'
        }
    }

    # Get team: https://learn.microsoft.com/graph/api/team-get
    function New-MockTeam {
        param([string]$Id = 'group-1', [bool]$Archived = $false)
        [pscustomobject]@{ Id = $Id; DisplayName = 'Project Northwind'; IsArchived = $Archived }
    }

    # Usage report CSV headers, as the header lists on the API pages give them:
    # https://learn.microsoft.com/graph/api/reportroot-getteamsteamactivitydetail
    # https://learn.microsoft.com/graph/api/reportroot-getoffice365groupsactivitydetail
    function New-TeamReportCsv {
        param([string]$TypeHeader = 'Team type')

        @(
            "Report Refresh Date,Team Name,Team Id,$TypeHeader,Last Activity Date,Report Period,Active users,Active Channels,Guests,Reactions,Meetings Organized,Post Messages,Reply Messages,Channel Messages,Urgent Messages,Mentions,Active Shared Channels,Active External Users"
            '2026-08-29,Project Northwind,group-1,Public,2026-08-20,180,12,3,2,40,5,60,80,140,1,9,0,1'
        ) -join "`r`n"
    }

    function New-GroupReportCsv {
        param([string]$ExternalHeader = 'External Member Count')

        @(
            "Report Refresh Date,Group Display Name,Is Deleted,Owner Principal Name,Last Activity Date,Group Type,Member Count,$ExternalHeader,Exchange Received Email Count,SharePoint Active File Count,Yammer Posted Message Count,Yammer Read Message Count,Yammer Liked Message Count,Exchange Mailbox Total Item Count,Exchange Mailbox Storage Used (Byte),SharePoint Total File Count,SharePoint Site Storage Used (Byte),Group Id,Report Period"
            '2026-08-29,Project Northwind,False,avery.abara@example.com,2026-08-20,Private,14,3,120,40,0,0,0,900,1048576,300,52428800,group-1,180'
        ) -join "`r`n"
    }

    # Unified audit log: https://learn.microsoft.com/purview/audit-log-detailed-properties
    # (TeamName, TeamGuid) and https://learn.microsoft.com/purview/audit-log-activities
    # (Operation names). The detail comes back as a JSON string in AuditData.
    function New-MockAuditRecord {
        param([string]$Id = 'event-1', [string]$Operation = 'TeamCreated')

        $auditData = @{
            CreationTime = '2026-08-10T12:00:00'
            Id           = $Id
            Operation    = $Operation
            UserId       = 'avery.abara@example.com'
            Workload     = 'MicrosoftTeams'
            ObjectId     = 'group-1'
            TeamName     = 'Project Northwind'
            TeamGuid     = 'group-1'
        } | ConvertTo-Json -Compress

        [pscustomobject]@{ RecordType = 'MicrosoftTeams'; AuditData = $auditData }
    }

    function Set-TestGroupsCsv {
        param([Parameter(Mandatory)][string]$OutputPath, [string[]]$Id = @('group-1'), [string]$Synced = '')

        $rows = foreach ($value in $Id) {
            [pscustomobject]@{
                RunDate = '2026-08-01'; Id = $value; DisplayName = "Group $value"; Mail = "$value@example.com"
                GroupTypes = 'Unified'; SecurityEnabled = $false; MailEnabled = $true; Visibility = 'Private'
                CreatedDateTime = '2026-02-03T04:05:06Z'; RenewedDateTime = '2026-02-03T04:05:06Z'
                ExpirationDateTime = '2027-02-03T04:05:06Z'; DeletedDateTime = ''
                OnPremisesSyncEnabled = $Synced; ResourceProvisioningOptions = 'Team'; IsTeam = $true
            }
        }
        Export-AppendCsv -Path (Join-Path $OutputPath 'groups.csv') -Rows @($rows) -Column $script:Schema.Groups
    }

    function Invoke-CollectorScript {
        param([Parameter(Mandatory)][string]$Name, [hashtable]$Arguments = @{})
        & (Join-Path $script:Collectors $Name) @Arguments
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

    It 'groups.csv' {
        Mock Get-MgGroup -MockWith { New-MockGroup }

        Invoke-CollectorScript 'Get-Groups.ps1' @{ OutputPath = $script:folder }

        $produced = Join-Path $script:folder 'groups.csv'
        @(Import-Csv -LiteralPath $produced).Count | Should -BeGreaterThan 0
        Get-HeaderText -Path $produced | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples 'groups.csv'))
        Get-HeaderText -Path $produced | Should -Be ($script:Schema.Groups -join ',')
    }

    It 'group-owners.csv' {
        Set-TestGroupsCsv -OutputPath $script:folder
        Mock Get-MgGroupOwner -MockWith { New-MockOwner }

        Invoke-CollectorScript 'Get-GroupOwners.ps1' @{ OutputPath = $script:folder }

        $produced = Join-Path $script:folder 'group-owners.csv'
        @(Import-Csv -LiteralPath $produced).Count | Should -BeGreaterThan 0
        Get-HeaderText -Path $produced | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples 'group-owners.csv'))
    }

    It 'deleted-groups.csv' {
        Mock Get-MgDirectoryDeletedItemAsGroup -MockWith { New-MockDeletedGroup }

        Invoke-CollectorScript 'Get-DeletedGroups.ps1' @{ OutputPath = $script:folder }

        $produced = Join-Path $script:folder 'deleted-groups.csv'
        @(Import-Csv -LiteralPath $produced).Count | Should -BeGreaterThan 0
        Get-HeaderText -Path $produced | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples 'deleted-groups.csv'))
    }

    It 'group-lifecycle-policies.csv' {
        Mock Get-MgGroupLifecyclePolicy -MockWith { New-MockLifecyclePolicy }

        Invoke-CollectorScript 'Get-GroupLifecyclePolicies.ps1' @{ OutputPath = $script:folder }

        $produced = Join-Path $script:folder 'group-lifecycle-policies.csv'
        @(Import-Csv -LiteralPath $produced).Count | Should -BeGreaterThan 0
        Get-HeaderText -Path $produced | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples 'group-lifecycle-policies.csv'))
    }

    It 'group-lifecycle-coverage.csv' {
        Set-TestGroupsCsv -OutputPath $script:folder
        Export-AppendCsv -Path (Join-Path $script:folder 'group-lifecycle-policies.csv') -Column $script:Schema.GroupLifecyclePolicies -Rows @(
            [pscustomobject]@{ RunDate = '2026-08-01'; Id = 'policy-1'; GroupLifetimeInDays = 365; ManagedGroupTypes = 'All'; AlternateNotificationEmails = '' }
        )

        Invoke-CollectorScript 'Get-GroupLifecycleCoverage.ps1' @{ OutputPath = $script:folder }

        $produced = Join-Path $script:folder 'group-lifecycle-coverage.csv'
        @(Import-Csv -LiteralPath $produced).Count | Should -BeGreaterThan 0
        Get-HeaderText -Path $produced | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples 'group-lifecycle-coverage.csv'))
    }

    It 'team-activity.csv' {
        Mock Get-MgReportTeamActivityDetail -MockWith { Set-Content -LiteralPath $OutFile -Value (New-TeamReportCsv) }

        Invoke-CollectorScript 'Get-TeamActivity.ps1' @{ OutputPath = $script:folder }

        $produced = Join-Path $script:folder 'team-activity.csv'
        @(Import-Csv -LiteralPath $produced).Count | Should -BeGreaterThan 0
        Get-HeaderText -Path $produced | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples 'team-activity.csv'))
    }

    It 'group-activity.csv' {
        Mock Get-MgReportOffice365GroupActivityDetail -MockWith { Set-Content -LiteralPath $OutFile -Value (New-GroupReportCsv) }

        Invoke-CollectorScript 'Get-GroupActivity.ps1' @{ OutputPath = $script:folder }

        $produced = Join-Path $script:folder 'group-activity.csv'
        @(Import-Csv -LiteralPath $produced).Count | Should -BeGreaterThan 0
        Get-HeaderText -Path $produced | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples 'group-activity.csv'))
    }

    It 'team-archive-status.csv' {
        Set-TestGroupsCsv -OutputPath $script:folder
        Mock Get-MgTeam -MockWith { New-MockTeam }

        Invoke-CollectorScript 'Get-ArchivedTeams.ps1' @{ OutputPath = $script:folder }

        $produced = Join-Path $script:folder 'team-archive-status.csv'
        @(Import-Csv -LiteralPath $produced).Count | Should -BeGreaterThan 0
        Get-HeaderText -Path $produced | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples 'team-archive-status.csv'))
    }

    It 'group-creation-events.csv' {
        Mock Search-UnifiedAuditLog -MockWith { New-MockAuditRecord }

        Invoke-CollectorScript 'Get-GroupCreationEvents.ps1' @{
            OutputPath = $script:folder
            StartDate  = [datetime]'2026-08-10T00:00:00Z'
            EndDate    = [datetime]'2026-08-11T00:00:00Z'
        }

        $produced = Join-Path $script:folder 'group-creation-events.csv'
        @(Import-Csv -LiteralPath $produced).Count | Should -BeGreaterThan 0
        Get-HeaderText -Path $produced | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples 'group-creation-events.csv'))
    }
}

Describe 'Get-Groups.ps1' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'lists Microsoft 365 groups and asks for every page with -All' {
        Mock Get-MgGroup -MockWith { New-MockGroup }

        Invoke-CollectorScript 'Get-Groups.ps1' @{ OutputPath = $script:folder }

        # -All is the SDK switch that follows @odata.nextLink until it is absent.
        Should -Invoke Get-MgGroup -Times 1 -Exactly -ParameterFilter {
            $All -and $Filter -eq "groupTypes/any(c:c eq 'Unified')" -and
            $Property -contains 'resourceProvisioningOptions' -and $Property -contains 'expirationDateTime'
        }
    }

    It 'keeps every object when the service returns more than one page of them' {
        # 1,250 objects is more than the 999-per-page ceiling. Without -All the
        # mock returns one page of 100, which is what list groups does by default.
        # https://learn.microsoft.com/graph/api/group-list
        # https://learn.microsoft.com/graph/paging
        Mock Get-MgGroup -MockWith {
            $limit = if ($All) { 1250 } else { 100 }
            1..$limit | ForEach-Object { New-MockGroup -Id "group-$_" }
        }

        Invoke-CollectorScript 'Get-Groups.ps1' @{ OutputPath = $script:folder }

        @(Import-Csv -LiteralPath (Join-Path $script:folder 'groups.csv')).Count | Should -Be 1250
    }

    It 'reads expiration, team membership and UTC timestamps' {
        Mock Get-MgGroup -MockWith {
            @(
                New-MockGroup -Id 'team-1' -Provisioning @('Team')
                New-MockGroup -Id 'plain-1' -Provisioning @() -Expiration $null
            )
        }

        Invoke-CollectorScript 'Get-Groups.ps1' @{ OutputPath = $script:folder }

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:folder 'groups.csv'))
        $team = $rows | Where-Object Id -EQ 'team-1'
        $plain = $rows | Where-Object Id -EQ 'plain-1'
        $team.IsTeam | Should -Be 'True'
        $team.ResourceProvisioningOptions | Should -Be 'Team'
        $team.ExpirationDateTime | Should -Be '2027-02-03T04:05:06Z'
        $plain.IsTeam | Should -Be 'False'
        $plain.ExpirationDateTime | Should -BeNullOrEmpty
    }

    It 'appends a second snapshot without repeating the header' {
        Mock Get-MgGroup -MockWith { New-MockGroup }

        Invoke-CollectorScript 'Get-Groups.ps1' @{ OutputPath = $script:folder }
        Invoke-CollectorScript 'Get-Groups.ps1' @{ OutputPath = $script:folder }

        # Same day, same group: RunDate + Id keeps a re-run idempotent.
        @(Import-Csv -LiteralPath (Join-Path $script:folder 'groups.csv')).Count | Should -Be 1
        (Get-Content -LiteralPath (Join-Path $script:folder 'groups.csv') | Select-String -Pattern '^"RunDate"').Count | Should -Be 1
    }

    It 'writes the header only and says why when the sign-in cannot list groups' {
        Mock Get-MgGroup -MockWith { throw 'Insufficient privileges to complete the operation.' }

        Invoke-CollectorScript 'Get-Groups.ps1' @{ OutputPath = $script:folder; WarningAction = 'SilentlyContinue' }

        $csv = Join-Path $script:folder 'groups.csv'
        (Get-Content -LiteralPath $csv).Count | Should -Be 1
        Get-Content -LiteralPath (Join-Path $script:folder 'run.log') -Raw | Should -Match 'Group.Read.All'
    }
}

Describe 'Get-GroupOwners.ps1' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'reads owners with -All and writes one row per owner' {
        Set-TestGroupsCsv -OutputPath $script:folder
        Mock Get-MgGroupOwner -MockWith { @(New-MockOwner -Id 'user-1'; New-MockOwner -Id 'user-2') }

        Invoke-CollectorScript 'Get-GroupOwners.ps1' @{ OutputPath = $script:folder }

        Should -Invoke Get-MgGroupOwner -Times 1 -Exactly -ParameterFilter { $All -and $GroupId -eq 'group-1' }
        $rows = @(Import-Csv -LiteralPath (Join-Path $script:folder 'group-owners.csv'))
        $rows.Count | Should -Be 2
        $rows[0].OwnerListStatus | Should -Be 'Listed'
        $rows[0].OwnerType | Should -Be 'user'
        $rows[0].OwnerUserPrincipalName | Should -Be 'avery.abara@example.com'
    }

    It 'writes None for a group whose call succeeded and returned no owner' {
        Set-TestGroupsCsv -OutputPath $script:folder
        Mock Get-MgGroupOwner -MockWith { @() }

        Invoke-CollectorScript 'Get-GroupOwners.ps1' @{ OutputPath = $script:folder }

        $row = Import-Csv -LiteralPath (Join-Path $script:folder 'group-owners.csv')
        $row.OwnerListStatus | Should -Be 'None'
        $row.OwnerId | Should -BeNullOrEmpty
    }

    It 'writes Unknown, not None, for a group synchronized from on-premises without calling Graph' {
        Set-TestGroupsCsv -OutputPath $script:folder -Synced 'True'
        Mock Get-MgGroupOwner -MockWith { @() }

        Invoke-CollectorScript 'Get-GroupOwners.ps1' @{ OutputPath = $script:folder }

        (Import-Csv -LiteralPath (Join-Path $script:folder 'group-owners.csv')).OwnerListStatus | Should -Be 'Unknown'
        Should -Invoke Get-MgGroupOwner -Times 0 -Exactly
    }

    It 'writes Unknown for the one group that fails and keeps the rest' {
        Set-TestGroupsCsv -OutputPath $script:folder -Id @('group-1', 'group-2')
        Mock Get-MgGroupOwner -MockWith { throw 'Request_ResourceNotFound' } -ParameterFilter { $GroupId -eq 'group-1' }
        Mock Get-MgGroupOwner -MockWith { New-MockOwner } -ParameterFilter { $GroupId -eq 'group-2' }

        Invoke-CollectorScript 'Get-GroupOwners.ps1' @{ OutputPath = $script:folder; WarningAction = 'SilentlyContinue' }

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:folder 'group-owners.csv'))
        ($rows | Where-Object GroupId -EQ 'group-1').OwnerListStatus | Should -Be 'Unknown'
        ($rows | Where-Object GroupId -EQ 'group-2').OwnerListStatus | Should -Be 'Listed'
    }

    It 'writes the header only when groups.csv does not exist yet' {
        Invoke-CollectorScript 'Get-GroupOwners.ps1' @{ OutputPath = $script:folder; WarningAction = 'SilentlyContinue' }

        (Get-Content -LiteralPath (Join-Path $script:folder 'group-owners.csv')).Count | Should -Be 1
        Get-Content -LiteralPath (Join-Path $script:folder 'run.log') -Raw | Should -Match 'Run Get-Groups.ps1 first'
    }

    It 'writes the header only when no owner can be read for any group' {
        Set-TestGroupsCsv -OutputPath $script:folder
        Mock Get-MgGroupOwner -MockWith { throw 'Insufficient privileges to complete the operation.' }

        Invoke-CollectorScript 'Get-GroupOwners.ps1' @{ OutputPath = $script:folder; WarningAction = 'SilentlyContinue' }

        (Get-Content -LiteralPath (Join-Path $script:folder 'group-owners.csv')).Count | Should -Be 1
        Get-Content -LiteralPath (Join-Path $script:folder 'run.log') -Raw | Should -Match 'GroupMember.Read.All'
    }
}

Describe 'Get-DeletedGroups.ps1' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'asks for every page and works out the purge date as 30 days after deletion' {
        Mock Get-MgDirectoryDeletedItemAsGroup -MockWith { New-MockDeletedGroup -DeletedAt ([datetime]'2026-08-20T10:00:00Z') }

        Invoke-CollectorScript 'Get-DeletedGroups.ps1' @{ OutputPath = $script:folder }

        Should -Invoke Get-MgDirectoryDeletedItemAsGroup -Times 1 -Exactly -ParameterFilter {
            $All -and $Property -contains 'deletedDateTime'
        }
        $row = Import-Csv -LiteralPath (Join-Path $script:folder 'deleted-groups.csv')
        $row.DeletedDateTime | Should -Be '2026-08-20T10:00:00Z'
        $row.PurgeDateTime | Should -Be '2026-09-19T10:00:00Z'
        $row.GroupTypes | Should -Be 'Unified'
        $row.IsTeam | Should -Be 'True'
    }

    It 'keeps a soft-deleted security group that reports securityEnabled false and an empty groupTypes list' {
        # https://learn.microsoft.com/graph/api/directory-deleteditems-list
        Mock Get-MgDirectoryDeletedItemAsGroup -MockWith {
            [pscustomobject]@{
                Id                   = 'security-1'
                DisplayName          = 'Role assignable group'
                GroupTypes           = @()
                SecurityEnabled      = $false
                MailEnabled          = $false
                CreatedDateTime      = [datetime]'2025-01-02T03:04:05Z'
                DeletedDateTime      = [datetime]'2026-08-20T10:00:00Z'
                AdditionalProperties = @{ resourceProvisioningOptions = @() }
            }
        }

        Invoke-CollectorScript 'Get-DeletedGroups.ps1' @{ OutputPath = $script:folder }

        $row = Import-Csv -LiteralPath (Join-Path $script:folder 'deleted-groups.csv')
        $row.Id | Should -Be 'security-1'
        $row.GroupTypes | Should -BeNullOrEmpty
        $row.SecurityEnabled | Should -Be 'False'
        $row.MailEnabled | Should -Be 'False'
        $row.IsTeam | Should -Be 'False'
        $row.PurgeDateTime | Should -Be '2026-09-19T10:00:00Z'
    }

    It 'writes the header only when the container cannot be read' {
        Mock Get-MgDirectoryDeletedItemAsGroup -MockWith { throw 'Insufficient privileges to complete the operation.' }

        Invoke-CollectorScript 'Get-DeletedGroups.ps1' @{ OutputPath = $script:folder; WarningAction = 'SilentlyContinue' }

        (Get-Content -LiteralPath (Join-Path $script:folder 'deleted-groups.csv')).Count | Should -Be 1
    }
}

Describe 'The group expiration policy collectors' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'writes the policy lifetime, scope and notification address' {
        Mock Get-MgGroupLifecyclePolicy -MockWith { New-MockLifecyclePolicy -Types 'Selected' }

        Invoke-CollectorScript 'Get-GroupLifecyclePolicies.ps1' @{ OutputPath = $script:folder }

        Should -Invoke Get-MgGroupLifecyclePolicy -Times 1 -Exactly -ParameterFilter { $All }
        $row = Import-Csv -LiteralPath (Join-Path $script:folder 'group-lifecycle-policies.csv')
        $row.GroupLifetimeInDays | Should -Be '365'
        $row.ManagedGroupTypes | Should -Be 'Selected'
        $row.AlternateNotificationEmails | Should -Be 'groups-admin@example.com'
    }

    It 'writes the header only and says so when the tenant has no policy' {
        Mock Get-MgGroupLifecyclePolicy -MockWith { }

        Invoke-CollectorScript 'Get-GroupLifecyclePolicies.ps1' @{ OutputPath = $script:folder }

        (Get-Content -LiteralPath (Join-Path $script:folder 'group-lifecycle-policies.csv')).Count | Should -Be 1
        Get-Content -LiteralPath (Join-Path $script:folder 'run.log') -Raw | Should -Match 'no group expiration policy'
    }

    It 'writes the header only and names the licence when the policy cannot be read' {
        Mock Get-MgGroupLifecyclePolicy -MockWith { throw 'Insufficient privileges to complete the operation.' }

        Invoke-CollectorScript 'Get-GroupLifecyclePolicies.ps1' @{ OutputPath = $script:folder; WarningAction = 'SilentlyContinue' }

        (Get-Content -LiteralPath (Join-Path $script:folder 'group-lifecycle-policies.csv')).Count | Should -Be 1
        Get-Content -LiteralPath (Join-Path $script:folder 'run.log') -Raw | Should -Match 'Entra ID P1 or P2'
    }

    It 'covers every group without a per-group call when the policy manages All' {
        Set-TestGroupsCsv -OutputPath $script:folder -Id @('group-1', 'group-2')
        Export-AppendCsv -Path (Join-Path $script:folder 'group-lifecycle-policies.csv') -Column $script:Schema.GroupLifecyclePolicies -Rows @(
            [pscustomobject]@{ RunDate = '2026-08-01'; Id = 'policy-1'; GroupLifetimeInDays = 365; ManagedGroupTypes = 'All'; AlternateNotificationEmails = '' }
        )
        Mock Get-MgGroupLifecyclePolicyByGroup -MockWith { }

        Invoke-CollectorScript 'Get-GroupLifecycleCoverage.ps1' @{ OutputPath = $script:folder }

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:folder 'group-lifecycle-coverage.csv'))
        $rows.CoverageStatus | Should -Be @('Covered', 'Covered')
        Should -Invoke Get-MgGroupLifecyclePolicyByGroup -Times 0 -Exactly
    }

    It 'asks Graph per group when the policy manages Selected groups' {
        Set-TestGroupsCsv -OutputPath $script:folder -Id @('group-1', 'group-2')
        Export-AppendCsv -Path (Join-Path $script:folder 'group-lifecycle-policies.csv') -Column $script:Schema.GroupLifecyclePolicies -Rows @(
            [pscustomobject]@{ RunDate = '2026-08-01'; Id = 'policy-1'; GroupLifetimeInDays = 365; ManagedGroupTypes = 'Selected'; AlternateNotificationEmails = '' }
        )
        Mock Get-MgGroupLifecyclePolicyByGroup -MockWith { New-MockLifecyclePolicy -Types 'Selected' } -ParameterFilter { $GroupId -eq 'group-1' }
        Mock Get-MgGroupLifecyclePolicyByGroup -MockWith { } -ParameterFilter { $GroupId -eq 'group-2' }

        Invoke-CollectorScript 'Get-GroupLifecycleCoverage.ps1' @{ OutputPath = $script:folder }

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:folder 'group-lifecycle-coverage.csv'))
        ($rows | Where-Object GroupId -EQ 'group-1').CoverageStatus | Should -Be 'Covered'
        ($rows | Where-Object GroupId -EQ 'group-2').CoverageStatus | Should -Be 'NotCovered'
    }

    It 'covers no group when the tenant has no policy' {
        Set-TestGroupsCsv -OutputPath $script:folder
        Export-AppendCsv -Path (Join-Path $script:folder 'group-lifecycle-policies.csv') -Column $script:Schema.GroupLifecyclePolicies

        Invoke-CollectorScript 'Get-GroupLifecycleCoverage.ps1' @{ OutputPath = $script:folder }

        (Import-Csv -LiteralPath (Join-Path $script:folder 'group-lifecycle-coverage.csv')).CoverageStatus | Should -Be 'NotCovered'
    }
}

Describe 'The usage report collectors' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'asks for the D180 period and saves the redirect download to a file' {
        Mock Get-MgReportTeamActivityDetail -MockWith { Set-Content -LiteralPath $OutFile -Value (New-TeamReportCsv) }

        Invoke-CollectorScript 'Get-TeamActivity.ps1' @{ OutputPath = $script:folder }

        Should -Invoke Get-MgReportTeamActivityDetail -Times 1 -Exactly -ParameterFilter {
            $Period -eq 'D180' -and -not [string]::IsNullOrEmpty($OutFile)
        }
        $row = Import-Csv -LiteralPath (Join-Path $script:folder 'team-activity.csv')
        $row.TeamId | Should -Be 'group-1'
        $row.LastActivityDate | Should -Be '2026-08-20'
        $row.ActiveUsers | Should -Be '12'
    }

    It 'removes the temporary download' {
        Mock Get-MgReportTeamActivityDetail -MockWith {
            $global:TeamsGroupsDownloadPath = $OutFile
            Set-Content -LiteralPath $OutFile -Value (New-TeamReportCsv)
        }

        Invoke-CollectorScript 'Get-TeamActivity.ps1' @{ OutputPath = $script:folder }

        Test-Path -LiteralPath $global:TeamsGroupsDownloadPath | Should -BeFalse
    }

    It 'reads the team type under either spelling of its header' -ForEach @(
        @{ Header = 'Team type' }
        @{ Header = 'Team Type' }
    ) {
        $global:TeamsGroupsTypeHeader = $Header
        Mock Get-MgReportTeamActivityDetail -MockWith { Set-Content -LiteralPath $OutFile -Value (New-TeamReportCsv -TypeHeader $global:TeamsGroupsTypeHeader) }

        Invoke-CollectorScript 'Get-TeamActivity.ps1' @{ OutputPath = $script:folder }

        (Import-Csv -LiteralPath (Join-Path $script:folder 'team-activity.csv')).TeamType | Should -Be 'Public'
    }

    It 'reads the external member count under either name' -ForEach @(
        @{ Header = 'External Member Count' }
        @{ Header = 'Guest Count' }
    ) {
        $global:TeamsGroupsExternalHeader = $Header
        Mock Get-MgReportOffice365GroupActivityDetail -MockWith { Set-Content -LiteralPath $OutFile -Value (New-GroupReportCsv -ExternalHeader $global:TeamsGroupsExternalHeader) }

        Invoke-CollectorScript 'Get-GroupActivity.ps1' @{ OutputPath = $script:folder }

        $row = Import-Csv -LiteralPath (Join-Path $script:folder 'group-activity.csv')
        $row.ExternalMemberCount | Should -Be '3'
        $row.GroupId | Should -Be 'group-1'
        $row.ExchangeMailboxStorageUsedByte | Should -Be '1048576'
    }

    It '<Script> writes a header only in GCC High, without calling Graph' -ForEach @(
        @{ Script = 'Get-TeamActivity.ps1'; Csv = 'team-activity.csv' }
        @{ Script = 'Get-GroupActivity.ps1'; Csv = 'group-activity.csv' }
    ) {
        Mock Get-MgReportTeamActivityDetail -MockWith { }
        Mock Get-MgReportOffice365GroupActivityDetail -MockWith { }

        Invoke-CollectorScript $Script @{ OutputPath = $script:folder; Environment = 'GCCHigh'; WarningAction = 'SilentlyContinue' }

        $produced = Join-Path $script:folder $Csv
        (Get-Content -LiteralPath $produced).Count | Should -Be 1
        Get-HeaderText -Path $produced | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples "gcchigh/$Csv"))
        Should -Invoke Get-MgReportTeamActivityDetail -Times 0 -Exactly
        Should -Invoke Get-MgReportOffice365GroupActivityDetail -Times 0 -Exactly
        Should -Invoke Connect-M365Service -Times 0 -Exactly
        $log = Get-Content -LiteralPath (Join-Path $script:folder 'run.log') -Raw
        $log | Should -Match 'unavailable in GCCHigh'
    }

    It '<Script> attempts the report in GCC and logs the UNVERIFIED warning' -ForEach @(
        @{ Script = 'Get-TeamActivity.ps1'; Csv = 'team-activity.csv' }
        @{ Script = 'Get-GroupActivity.ps1'; Csv = 'group-activity.csv' }
    ) {
        Mock Get-MgReportTeamActivityDetail -MockWith { Set-Content -LiteralPath $OutFile -Value (New-TeamReportCsv) }
        Mock Get-MgReportOffice365GroupActivityDetail -MockWith { Set-Content -LiteralPath $OutFile -Value (New-GroupReportCsv) }

        Invoke-CollectorScript $Script @{ OutputPath = $script:folder; Environment = 'GCC'; WarningAction = 'SilentlyContinue' }

        @(Import-Csv -LiteralPath (Join-Path $script:folder $Csv)).Count | Should -Be 1
        Get-Content -LiteralPath (Join-Path $script:folder 'run.log') -Raw | Should -Match 'UNVERIFIED'
    }

    It 'records a refusal in run.log and writes the header only when GCC turns the report down' {
        Mock Get-MgReportTeamActivityDetail -MockWith { throw 'Forbidden: the API is not available in this cloud.' }

        Invoke-CollectorScript 'Get-TeamActivity.ps1' @{ OutputPath = $script:folder; Environment = 'GCC'; WarningAction = 'SilentlyContinue' }

        (Get-Content -LiteralPath (Join-Path $script:folder 'team-activity.csv')).Count | Should -Be 1
        $log = Get-Content -LiteralPath (Join-Path $script:folder 'run.log') -Raw
        $log | Should -Match 'UNVERIFIED'
        $log | Should -Match 'Reports.Read.All'
    }
}

Describe 'Get-ArchivedTeams.ps1' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'asks Graph about every group, including one with no Team provisioning option' {
        Set-TestGroupsCsv -OutputPath $script:folder -Id @('group-1', 'old-team-1')
        Mock Get-MgTeam -MockWith { New-MockTeam -Id $TeamId }

        Invoke-CollectorScript 'Get-ArchivedTeams.ps1' @{ OutputPath = $script:folder }

        Should -Invoke Get-MgTeam -Times 1 -Exactly -ParameterFilter { $TeamId -eq 'group-1' }
        Should -Invoke Get-MgTeam -Times 1 -Exactly -ParameterFilter { $TeamId -eq 'old-team-1' }
    }

    It 'records isArchived and skips a group that is not a team' {
        Set-TestGroupsCsv -OutputPath $script:folder -Id @('team-1', 'plain-1')
        Mock Get-MgTeam -MockWith { New-MockTeam -Id 'team-1' -Archived $true } -ParameterFilter { $TeamId -eq 'team-1' }
        Mock Get-MgTeam -MockWith { throw 'Response status code does not indicate success: NotFound (Not Found).' } -ParameterFilter { $TeamId -eq 'plain-1' }

        Invoke-CollectorScript 'Get-ArchivedTeams.ps1' @{ OutputPath = $script:folder; WarningAction = 'SilentlyContinue' }

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:folder 'team-archive-status.csv'))
        $rows.Count | Should -Be 1
        $rows[0].TeamId | Should -Be 'team-1'
        $rows[0].IsArchived | Should -Be 'True'
    }

    It 'keeps the teams it could read and fails when another team read errors' {
        Set-TestGroupsCsv -OutputPath $script:folder -Id @('team-1', 'team-2')
        Mock Get-MgTeam -MockWith { New-MockTeam -Id 'team-1' -Archived $true } -ParameterFilter { $TeamId -eq 'team-1' }
        Mock Get-MgTeam -MockWith { throw 'Forbidden: insufficient privileges.' } -ParameterFilter { $TeamId -eq 'team-2' }

        {
            Invoke-CollectorScript 'Get-ArchivedTeams.ps1' @{ OutputPath = $script:folder; WarningAction = 'SilentlyContinue' }
        } | Should -Throw '*could not be read*'

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:folder 'team-archive-status.csv'))
        $rows.Count | Should -Be 1
        $rows[0].TeamId | Should -Be 'team-1'
        Get-Content -LiteralPath (Join-Path $script:folder 'run.log') -Raw | Should -Match 'team-2'
    }

    It 'writes the header only when no team can be read' {
        Set-TestGroupsCsv -OutputPath $script:folder
        Mock Get-MgTeam -MockWith { throw 'Forbidden: insufficient privileges.' }

        Invoke-CollectorScript 'Get-ArchivedTeams.ps1' @{ OutputPath = $script:folder; WarningAction = 'SilentlyContinue' }

        (Get-Content -LiteralPath (Join-Path $script:folder 'team-archive-status.csv')).Count | Should -Be 1
        Get-Content -LiteralPath (Join-Path $script:folder 'run.log') -Raw | Should -Match 'Team.ReadBasic.All'
    }
}

Describe 'Get-GroupCreationEvents.ps1' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
        $script:watermark = [datetime]::UtcNow.AddHours(-30).ToString('yyyy-MM-ddTHH:mm:00Z')
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'searches for AddGroup and TeamCreated only, in a ReturnLargeSet session' {
        Mock Search-UnifiedAuditLog -MockWith { New-MockAuditRecord }

        Invoke-CollectorScript 'Get-GroupCreationEvents.ps1' @{
            OutputPath = $script:folder
            StartDate  = [datetime]'2026-08-10T00:00:00Z'
            EndDate    = [datetime]'2026-08-11T00:00:00Z'
        }

        Should -Invoke Search-UnifiedAuditLog -Times 1 -Exactly -ParameterFilter {
            $SessionCommand -eq 'ReturnLargeSet' -and -not [string]::IsNullOrEmpty($SessionId) -and
            $Operations.Count -eq 2 -and $Operations -ccontains 'AddGroup' -and $Operations -ccontains 'TeamCreated'
        }
    }

    It 'reads the creator and the team out of AuditData' {
        Mock Search-UnifiedAuditLog -MockWith { New-MockAuditRecord -Id 'event-9' }

        Invoke-CollectorScript 'Get-GroupCreationEvents.ps1' @{
            OutputPath = $script:folder
            StartDate  = [datetime]'2026-08-10T00:00:00Z'
            EndDate    = [datetime]'2026-08-11T00:00:00Z'
        }

        $row = Import-Csv -LiteralPath (Join-Path $script:folder 'group-creation-events.csv')
        $row.Id | Should -Be 'event-9'
        $row.Operation | Should -Be 'TeamCreated'
        $row.UserId | Should -Be 'avery.abara@example.com'
        $row.TargetDisplayName | Should -Be 'Project Northwind'
        $row.TargetGroupId | Should -Be 'group-1'
        $row.CreationTime | Should -Be '2026-08-10T12:00:00Z'
    }

    It 'starts at the latest CreationTime already collected' {
        Export-AppendCsv -Path (Join-Path $script:folder 'group-creation-events.csv') -Column $script:Schema.GroupCreationEvents -Rows @(
            [pscustomobject]@{
                CreationTime = $script:watermark; Id = 'event-0'; Operation = 'AddGroup'; UserId = 'avery.abara@example.com'
                Workload = 'AzureActiveDirectory'; ObjectId = 'group-0'; TargetDisplayName = ''; TargetGroupId = ''
            }
        )
        Mock Search-UnifiedAuditLog -MockWith { @() }

        Invoke-CollectorScript 'Get-GroupCreationEvents.ps1' @{ OutputPath = $script:folder }

        Should -Invoke Search-UnifiedAuditLog -Times 1 -ParameterFilter {
            $StartDate -eq [datetime]::Parse($script:watermark, [cultureinfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AdjustToUniversal -bor [System.Globalization.DateTimeStyles]::AssumeUniversal)
        }
    }

    It 'refuses a range that ends at or before it starts' {
        Mock Search-UnifiedAuditLog -MockWith { @() }

        {
            Invoke-CollectorScript 'Get-GroupCreationEvents.ps1' @{
                OutputPath = $script:folder
                StartDate  = [datetime]'2026-08-16T00:00:00Z'
                EndDate    = [datetime]'2026-08-15T00:00:00Z'
            }
        } | Should -Throw '*range is empty*'
    }

    It 'attempts GCC and GCC High with a warning, because the records there are UNVERIFIED' -ForEach @(
        @{ Cloud = 'GCC' }
        @{ Cloud = 'GCCHigh' }
    ) {
        Mock Search-UnifiedAuditLog -MockWith { New-MockAuditRecord }

        Invoke-CollectorScript 'Get-GroupCreationEvents.ps1' @{
            OutputPath    = $script:folder
            Environment   = $Cloud
            StartDate     = [datetime]'2026-08-10T00:00:00Z'
            EndDate       = [datetime]'2026-08-11T00:00:00Z'
            WarningAction = 'SilentlyContinue'
        }

        Should -Invoke Search-UnifiedAuditLog -Times 1 -Exactly
        Get-Content -LiteralPath (Join-Path $script:folder 'run.log') -Raw | Should -Match 'UNVERIFIED'
    }

    It 'fetches the next page when ResultCount is larger than the page just returned' {
        # A short first page used to end the session. ResultCount is the hit count
        # across iterations, so one record with ResultCount 2 still has a second page.
        # https://learn.microsoft.com/powershell/module/exchangepowershell/search-unifiedauditlog
        $global:TeamsGroupsAuditCalls = 0
        Mock Search-UnifiedAuditLog -MockWith {
            $global:TeamsGroupsAuditCalls++
            if ($global:TeamsGroupsAuditCalls -gt 2) { return @() }
            $recordId = if ($global:TeamsGroupsAuditCalls -eq 1) { 'event-1' } else { 'event-2' }
            $record = New-MockAuditRecord -Id $recordId
            $record | Add-Member -NotePropertyName ResultCount -NotePropertyValue 2
            return $record
        }

        Invoke-CollectorScript 'Get-GroupCreationEvents.ps1' @{
            OutputPath = $script:folder
            StartDate  = [datetime]'2026-08-10T00:00:00Z'
            EndDate    = [datetime]'2026-08-11T00:00:00Z'
        }

        $ids = @(Import-Csv -LiteralPath (Join-Path $script:folder 'group-creation-events.csv') | Select-Object -ExpandProperty Id | Sort-Object)
        $ids | Should -Be @('event-1', 'event-2')
        $global:TeamsGroupsAuditCalls | Should -BeGreaterOrEqual 2
    }

    It 'fetches the next page when moreRecordsAvailable is true even if ResultCount looks complete' {
        $global:TeamsGroupsAuditCalls = 0
        Mock Search-UnifiedAuditLog -MockWith {
            $global:TeamsGroupsAuditCalls++
            if ($global:TeamsGroupsAuditCalls -eq 1) {
                $record = New-MockAuditRecord -Id 'event-1'
                $record | Add-Member -NotePropertyName ResultCount -NotePropertyValue 1
                $record | Add-Member -NotePropertyName AuditSearchRequestMetadata -NotePropertyValue ([pscustomobject]@{ moreRecordsAvailable = $true })
                return $record
            }
            if ($global:TeamsGroupsAuditCalls -eq 2) {
                $record = New-MockAuditRecord -Id 'event-2'
                $record | Add-Member -NotePropertyName ResultCount -NotePropertyValue 2
                $record | Add-Member -NotePropertyName AuditSearchRequestMetadata -NotePropertyValue ([pscustomobject]@{ moreRecordsAvailable = $false })
                return $record
            }
            return @()
        }

        Invoke-CollectorScript 'Get-GroupCreationEvents.ps1' @{
            OutputPath = $script:folder
            StartDate  = [datetime]'2026-08-10T00:00:00Z'
            EndDate    = [datetime]'2026-08-11T00:00:00Z'
        }

        $ids = @(Import-Csv -LiteralPath (Join-Path $script:folder 'group-creation-events.csv') | Select-Object -ExpandProperty Id | Sort-Object)
        $ids | Should -Be @('event-1', 'event-2')
    }

    It 'writes the header only when the audit log cannot be searched' {
        Mock Search-UnifiedAuditLog -MockWith { throw 'The term ''Search-UnifiedAuditLog'' is not recognized.' }

        Invoke-CollectorScript 'Get-GroupCreationEvents.ps1' @{
            OutputPath    = $script:folder
            StartDate     = [datetime]'2026-08-10T00:00:00Z'
            EndDate       = [datetime]'2026-08-11T00:00:00Z'
            WarningAction = 'SilentlyContinue'
        }

        (Get-Content -LiteralPath (Join-Path $script:folder 'group-creation-events.csv')).Count | Should -Be 1
    }
}

Describe 'Collectors connect to the cloud they were asked for' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
        Mock Get-MgGroup -MockWith { }
        Mock Get-MgGroupOwner -MockWith { }
        Mock Get-MgDirectoryDeletedItemAsGroup -MockWith { }
        Mock Get-MgGroupLifecyclePolicy -MockWith { }
        Mock Get-MgGroupLifecyclePolicyByGroup -MockWith { }
        Mock Get-MgTeam -MockWith { }
        Mock Get-MgReportTeamActivityDetail -MockWith { }
        Mock Get-MgReportOffice365GroupActivityDetail -MockWith { }
        Mock Search-UnifiedAuditLog -MockWith { @() }
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    # The data keys are deliberately not named Service or Environment: inside a
    # -ParameterFilter those names are the mocked call's own bound parameters.
    It '<CollectorScript> asks for the <ExpectedService> service in <Cloud>' -ForEach @(
        @{ CollectorScript = 'Get-Groups.ps1'; ExpectedService = 'Graph'; Cloud = 'Commercial'; Extra = @{} }
        @{ CollectorScript = 'Get-Groups.ps1'; ExpectedService = 'Graph'; Cloud = 'GCC'; Extra = @{} }
        @{ CollectorScript = 'Get-Groups.ps1'; ExpectedService = 'Graph'; Cloud = 'GCCHigh'; Extra = @{} }
        @{ CollectorScript = 'Get-GroupOwners.ps1'; ExpectedService = 'Graph'; Cloud = 'GCCHigh'; Extra = @{} }
        @{ CollectorScript = 'Get-DeletedGroups.ps1'; ExpectedService = 'Graph'; Cloud = 'GCC'; Extra = @{} }
        @{ CollectorScript = 'Get-DeletedGroups.ps1'; ExpectedService = 'Graph'; Cloud = 'GCCHigh'; Extra = @{} }
        @{ CollectorScript = 'Get-GroupLifecyclePolicies.ps1'; ExpectedService = 'Graph'; Cloud = 'GCCHigh'; Extra = @{} }
        @{ CollectorScript = 'Get-ArchivedTeams.ps1'; ExpectedService = 'Graph'; Cloud = 'GCCHigh'; Extra = @{} }
        @{ CollectorScript = 'Get-TeamActivity.ps1'; ExpectedService = 'Graph'; Cloud = 'Commercial'; Extra = @{} }
        @{ CollectorScript = 'Get-GroupActivity.ps1'; ExpectedService = 'Graph'; Cloud = 'GCC'; Extra = @{} }
        @{ CollectorScript = 'Get-GroupCreationEvents.ps1'; ExpectedService = 'ExchangeOnline'; Cloud = 'Commercial'; Extra = @{ LookbackDays = 1 } }
        @{ CollectorScript = 'Get-GroupCreationEvents.ps1'; ExpectedService = 'ExchangeOnline'; Cloud = 'GCCHigh'; Extra = @{ LookbackDays = 1 } }
    ) {
        Set-TestGroupsCsv -OutputPath $script:folder

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

    It 'resolves the Graph endpoint for each cloud' -ForEach @(
        @{ Cloud = 'Commercial'; Endpoint = 'https://graph.microsoft.com'; GraphEnv = 'Global' }
        @{ Cloud = 'GCC'; Endpoint = 'https://graph.microsoft.com'; GraphEnv = 'Global' }
        @{ Cloud = 'GCCHigh'; Endpoint = 'https://graph.microsoft.us'; GraphEnv = 'USGov' }
    ) {
        $resolved = Get-M365ServiceEndpoint -Service Graph -Environment $Cloud
        $resolved.ResourceEndpoint | Should -Be $Endpoint
        $resolved.GraphEnvironment | Should -Be $GraphEnv
    }

    It 'reuses the session when -SkipConnect is given' {
        & (Join-Path $script:Collectors 'Get-Groups.ps1') -OutputPath $script:folder -SkipConnect -WarningAction SilentlyContinue

        Should -Invoke Connect-M365Service -Times 0 -Exactly
    }
}

Describe 'Run-All.ps1' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
        Mock Connect-M365Service -ModuleName M365ReportLibrary -MockWith { }
        Mock Get-MgUser -ModuleName M365ReportLibrary -MockWith { }
        Mock Get-MgGroup -MockWith { New-MockGroup }
        Mock Get-MgGroupOwner -MockWith { New-MockOwner }
        Mock Get-MgDirectoryDeletedItemAsGroup -MockWith { New-MockDeletedGroup }
        Mock Get-MgGroupLifecyclePolicy -MockWith { New-MockLifecyclePolicy }
        Mock Get-MgGroupLifecyclePolicyByGroup -MockWith { }
        Mock Get-MgTeam -MockWith { New-MockTeam }
        Mock Get-MgReportTeamActivityDetail -MockWith { Set-Content -LiteralPath $OutFile -Value (New-TeamReportCsv) }
        Mock Get-MgReportOffice365GroupActivityDetail -MockWith { Set-Content -LiteralPath $OutFile -Value (New-GroupReportCsv) }
        Mock Search-UnifiedAuditLog -MockWith { New-MockAuditRecord }
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'writes all ten CSVs in dependency order' {
        & (Join-Path $script:Collectors 'Run-All.ps1') -OutputPath $script:folder `
            -StartDate ([datetime]'2026-08-10T00:00:00Z') -EndDate ([datetime]'2026-08-11T00:00:00Z') -WarningAction SilentlyContinue

        $expected = @(
            'users', 'groups', 'group-owners', 'deleted-groups', 'group-lifecycle-policies'
            'group-lifecycle-coverage', 'team-archive-status', 'team-activity', 'group-activity', 'group-creation-events'
        )
        foreach ($name in $expected) {
            Test-Path -LiteralPath (Join-Path $script:folder "$name.csv") | Should -BeTrue -Because "$name.csv should exist"
        }
        @(Import-Csv -LiteralPath (Join-Path $script:folder 'group-lifecycle-coverage.csv')).Count | Should -BeGreaterThan 0
    }
}
