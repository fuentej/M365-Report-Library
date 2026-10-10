#Requires -Version 7.0

BeforeAll {
    $script:Root = Join-Path $PSScriptRoot '../../..' | Resolve-Path | Select-Object -ExpandProperty Path
    $script:Collectors = Join-Path $script:Root 'reports/oversharing/collectors'
    $script:Samples = Join-Path $script:Root 'reports/oversharing/samples'

    . (Join-Path $script:Root 'shared/tests/TenantCmdletStubs.ps1')
    . (Join-Path $PSScriptRoot 'OversharingStubs.ps1')
    Import-Module (Join-Path $script:Root 'shared/M365ReportLibrary.psm1') -Force

    $script:Schema = Import-PowerShellDataFile -LiteralPath (Join-Path $script:Collectors 'OversharingSchema.psd1')
    . (Join-Path $script:Collectors 'OversharingHelpers.ps1')

    $script:Today = [datetime]::UtcNow.ToString('yyyy-MM-dd')

    function New-TestFolder {
        $path = Join-Path ([System.IO.Path]::GetTempPath()) ('oversharing-' + [guid]::NewGuid().ToString('N'))
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

    # getAllSites: https://learn.microsoft.com/graph/api/site-getallsites (site resource, siteCollection)
    function New-MockSiteResponse {
        param([string]$Id = 'contoso.sharepoint.com,11111111-1111-1111-1111-111111111111,22222222-2222-2222-2222-222222222222', [string]$NextLink = '')
        $page = @{
            value = @(@{
                    id = $Id; name = 'Finance'; webUrl = 'https://contoso.sharepoint.com/sites/finance'
                    isPersonalSite = $false; siteCollection = @{ hostname = 'contoso.sharepoint.com'; dataLocationCode = 'NAM' }
                })
        }
        if ($NextLink) { $page['@odata.nextLink'] = $NextLink }
        return $page
    }

    # Search-UnifiedAuditLog record: RecordType, ResultCount, AuditData as a JSON string.
    # https://learn.microsoft.com/powershell/module/exchangepowershell/search-unifiedauditlog
    function New-MockAuditRecord {
        param(
            [string]$Id = 'event-1',
            [string]$Operation = 'AnonymousLinkCreated',
            [object]$ResultCount = $null,
            [string]$CreationTime = '2026-08-10T12:00:00',
            [object]$MoreRecords = $null
        )
        $auditData = [ordered]@{
            CreationTime = $CreationTime; Id = $Id; Operation = $Operation; UserId = 'avery.abara@example.com'
            Workload = 'SharePoint'; ObjectId = 'https://contoso.sharepoint.com/sites/finance/Shared Documents/plan.docx'
            ItemType = 'File'; SiteUrl = 'https://contoso.sharepoint.com/sites/finance'; ClientIP = '203.0.113.5'
        } | ConvertTo-Json -Compress
        $record = [ordered]@{ RecordType = 'SharePointSharingOperation'; AuditData = $auditData }
        if ($null -ne $ResultCount) { $record['ResultCount'] = $ResultCount }
        if ($null -ne $MoreRecords) { $record['AuditSearchRequestMetadata'] = [pscustomobject]@{ moreRecordsAvailable = $MoreRecords } }
        return [pscustomobject]$record
    }

    # The CSV a Data access governance export holds. Columns for the site permissions report
    # are those at https://learn.microsoft.com/sharepoint/data-access-governance-site-permissions-report
    $global:OversharingTestCsv = (
        '"Site ID","Site URL","Site Name","Number of users having access","Report Date","Recipient","RoleDefinition"',
        '"11111111-1111-1111-1111-111111111111","https://contoso.sharepoint.com/sites/finance","Finance","42","2026-08-10","Everyone except external users","Read"'
    ) -join "`n"

    function Set-DagMock {
        # Reports are listed as none, started, completed at once, and exported as one CSV.
        param([string]$Status = 'Completed')
        $global:OversharingTestStatus = $Status
        Mock Get-SPODataAccessGovernanceInsight -MockWith {
            if ($ReportID) {
                [pscustomobject]@{ ReportId = $ReportID; Status = $global:OversharingTestStatus; ReportStartTime = '2026-07-13T00:00:00Z'; ReportEndTime = '2026-08-10T00:00:00Z' }
            }
        }
        Mock Start-SPODataAccessGovernanceInsight -MockWith { [pscustomobject]@{ ReportId = "report-$ReportEntity-$Workload" } }
        Mock Export-SPODataAccessGovernanceInsight -MockWith { Set-Content -LiteralPath (Join-Path $DownloadPath 'report.csv') -Value $global:OversharingTestCsv }
        Mock Get-SPOAuditDataCollectionStatusForActivityInsights -MockWith { [pscustomobject]@{ Status = 'InProgress' } }
    }

    $script:DagCases = @(
        @{ Script = 'Get-SitePermissionBreadth.ps1'; Csv = 'site-permission-breadth.csv'; Extra = @{} }
        @{ Script = 'Get-EveryoneItemExposure.ps1'; Csv = 'everyone-item-exposure.csv'; Extra = @{} }
        @{ Script = 'Get-SharingLinkActivity.ps1'; Csv = 'sharing-link-activity.csv'; Extra = @{} }
        @{ Script = 'Get-EeeuActivity.ps1'; Csv = 'eeeu-activity.csv'; Extra = @{} }
        @{ Script = 'Get-LabeledFileSites.ps1'; Csv = 'labeled-file-sites.csv'; Extra = @{ LabelGuid = [guid]'33333333-3333-3333-3333-333333333333' } }
    )
}

AfterAll {
    Remove-Variable -Name OversharingTestCsv, OversharingTestStatus -Scope Global -ErrorAction SilentlyContinue
    Remove-Module M365ReportLibrary -Force -ErrorAction SilentlyContinue
}

Describe 'Each collector writes the columns of its sample' {
    BeforeEach {
        $script:folder = New-TestFolder
        Set-DagMock
        Mock Invoke-MgGraphRequest -MockWith {
            if ($Uri -like '*getAllSites*') { New-MockSiteResponse }
            elseif ($Uri -like '*/drives?*') { @{ value = @(@{ id = 'drive-1'; name = 'Documents'; driveType = 'documentLibrary' }) } }
            elseif ($Uri -like '*/children*') { @{ value = @(@{ id = 'item-1'; name = 'plan.docx'; webUrl = 'https://contoso.sharepoint.com/plan.docx' }) } }
            elseif ($Uri -like '*/permissions') {
                @{ value = @(@{ id = 'perm-1'; roles = @('read'); link = @{ scope = 'anonymous'; type = 'view' }; hasPassword = $false }) }
            }
        }
        Mock Get-SPOTenant -MockWith { [pscustomobject]@{ SharingCapability = 'ExternalUserAndGuestSharing'; DefaultSharingLinkType = 'Internal'; DisableCompanyWideSharingLinks = 'NotDisabled' } }
        Mock Get-SPOSite -MockWith {
            if ($Identity) { [pscustomobject]@{ Url = $Identity; Title = 'Finance'; Template = 'GROUP#0'; SharingCapability = 'ExternalUserSharingOnly'; DefaultSharingLinkType = 'Direct'; DisableCompanyWideSharingLinks = $false; SensitivityLabel = '' } }
            else { [pscustomobject]@{ Url = 'https://contoso.sharepoint.com/sites/finance' } }
        }
        Mock Search-UnifiedAuditLog -MockWith { New-MockAuditRecord }
        Mock Get-AdminAuditLogConfig -MockWith { [pscustomobject]@{ UnifiedAuditLogIngestionEnabled = $true } }
        Mock Get-Label -MockWith { @() }
    }
    AfterEach { Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue }

    It '<Script> writes <Csv> with the sample columns and the schema columns' -ForEach @(
        @{ Script = 'Get-Sites.ps1'; Csv = 'sites.csv'; Key = 'Sites' }
        @{ Script = 'Get-ItemSharingPermissions.ps1'; Csv = 'item-permissions.csv'; Key = 'ItemPermissions' }
        @{ Script = 'Get-SitePermissionBreadth.ps1'; Csv = 'site-permission-breadth.csv'; Key = 'SitePermissionBreadth' }
        @{ Script = 'Get-EveryoneItemExposure.ps1'; Csv = 'everyone-item-exposure.csv'; Key = 'EveryoneItemExposure' }
        @{ Script = 'Get-SharingLinkActivity.ps1'; Csv = 'sharing-link-activity.csv'; Key = 'SharingLinkActivity' }
        @{ Script = 'Get-EeeuActivity.ps1'; Csv = 'eeeu-activity.csv'; Key = 'EeeuActivity' }
        @{ Script = 'Get-LabeledFileSites.ps1'; Csv = 'labeled-file-sites.csv'; Key = 'LabeledFileSites' }
        @{ Script = 'Get-SiteSharingSettings.ps1'; Csv = 'site-sharing-settings.csv'; Key = 'SiteSharingSettings' }
        @{ Script = 'Get-AnonymousLinkEvents.ps1'; Csv = 'anonymous-link-events.csv'; Key = 'AuditEvents' }
        @{ Script = 'Get-SharingEvents.ps1'; Csv = 'sharing-events.csv'; Key = 'AuditEvents' }
        @{ Script = 'Get-AuditLogStatus.ps1'; Csv = 'audit-log-status.csv'; Key = 'AuditLogStatus' }
    ) {
        $arguments = @{ OutputPath = $script:folder; SkipConnect = $true }
        if ($Script -eq 'Get-LabeledFileSites.ps1') { $arguments['LabelGuid'] = [guid]'33333333-3333-3333-3333-333333333333' }
        if ($Script -eq 'Get-ItemSharingPermissions.ps1') { $arguments['SiteId'] = 'site-1' }

        Invoke-CollectorScript $Script $arguments

        $produced = Join-Path $script:folder $Csv
        @(Import-Csv -LiteralPath $produced).Count | Should -BeGreaterThan 0
        Get-HeaderText -Path $produced | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples $Csv))
        Get-HeaderText -Path $produced | Should -Be ($script:Schema[$Key] -join ',')
    }
}

Describe 'State sources stamp the run date and append' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Invoke-MgGraphRequest -MockWith { New-MockSiteResponse }
        Mock Get-AdminAuditLogConfig -MockWith { [pscustomobject]@{ UnifiedAuditLogIngestionEnabled = 'True' } }
    }
    AfterEach { Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue }

    It 'Get-Sites.ps1 tags rows with today''s UTC date' {
        Invoke-CollectorScript 'Get-Sites.ps1' @{ OutputPath = $script:folder; SkipConnect = $true }

        @(Import-Csv -LiteralPath (Join-Path $script:folder 'sites.csv'))[0].RunDate | Should -Be $script:Today
    }

    It 'Get-AuditLogStatus.ps1 tags the row with the run date and writes the flag' {
        Invoke-CollectorScript 'Get-AuditLogStatus.ps1' @{ OutputPath = $script:folder; SkipConnect = $true }

        $row = @(Import-Csv -LiteralPath (Join-Path $script:folder 'audit-log-status.csv'))[0]
        $row.RunDate | Should -Be $script:Today
        $row.UnifiedAuditLogIngestionEnabled | Should -Be 'True'
    }

    It 'a later run appends a new RunDate block instead of overwriting' {
        Invoke-CollectorScript 'Get-Sites.ps1' @{ OutputPath = $script:folder; SkipConnect = $true }
        $path = Join-Path $script:folder 'sites.csv'
        Set-Content -LiteralPath $path -Value ((Get-Content -LiteralPath $path -Raw) -replace [regex]::Escape($script:Today), '2026-01-01') -NoNewline
        Invoke-CollectorScript 'Get-Sites.ps1' @{ OutputPath = $script:folder; SkipConnect = $true }

        (@(Import-Csv -LiteralPath $path).RunDate | Sort-Object -Unique) -join ',' | Should -Be "2026-01-01,$script:Today"
    }
}

Describe 'Each connection targets the endpoints of its -Environment' {
    BeforeEach {
        $script:folder = New-TestFolder
        Set-DagMock
        Mock Connect-MgGraph -ModuleName M365ReportLibrary -MockWith { }
        Mock Connect-ExchangeOnline -ModuleName M365ReportLibrary -MockWith { }
        Mock Disconnect-ExchangeOnline -MockWith { }
        Mock Connect-SPOService -MockWith { }
        Mock Disconnect-SPOService -MockWith { }
        Mock Invoke-MgGraphRequest -MockWith { @{ value = @() } }
        Mock Search-UnifiedAuditLog -MockWith { @() }
        Mock Get-AdminAuditLogConfig -MockWith { [pscustomobject]@{ UnifiedAuditLogIngestionEnabled = $true } }
        Mock Get-SPOTenant -MockWith { [pscustomobject]@{ SharingCapability = 'Disabled' } }
        Mock Get-SPOSite -MockWith { @() }
        Mock Start-Sleep -MockWith { }
    }
    AfterEach { Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue }

    # https://learn.microsoft.com/graph/deployments: GCC calls the global service;
    # GCC High is the USGov environment.
    It '<Script> signs in to <Graph> Graph for <Environment>' -ForEach @(
        foreach ($script in 'Get-Sites.ps1', 'Get-ItemSharingPermissions.ps1') {
            @{ Script = $script; Environment = 'Commercial'; Graph = 'Global' }
            @{ Script = $script; Environment = 'GCC'; Graph = 'Global' }
            @{ Script = $script; Environment = 'GCCHigh'; Graph = 'USGov' }
        }
    ) {
        $arguments = @{ OutputPath = $script:folder; Environment = $Environment }
        if ($Script -eq 'Get-ItemSharingPermissions.ps1') { $arguments['SiteId'] = 'site-1' }
        Invoke-CollectorScript $Script $arguments

        $expected = $Graph
        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter { $Environment -eq $expected }
    }

    It '<Script> signs in to Exchange Online for <Environment> as <Name>' -ForEach @(
        foreach ($script in 'Get-AnonymousLinkEvents.ps1', 'Get-SharingEvents.ps1', 'Get-AuditLogStatus.ps1') {
            @{ Script = $script; Environment = 'Commercial'; Name = 'O365Default' }
            @{ Script = $script; Environment = 'GCC'; Name = 'O365Default' }
            @{ Script = $script; Environment = 'GCCHigh'; Name = 'O365USGovGCCHigh' }
        }
    ) {
        Invoke-CollectorScript $Script @{ OutputPath = $script:folder; Environment = $Environment; WarningAction = 'SilentlyContinue' }

        $expected = $Name
        Should -Invoke Connect-ExchangeOnline -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter { $ExchangeEnvironmentName -eq $expected }
    }

    It '<Script> signs in to SharePoint Online at the admin URL for <Environment>' -ForEach @(
        foreach ($script in 'Get-SitePermissionBreadth.ps1', 'Get-SiteSharingSettings.ps1', 'Get-SharingLinkActivity.ps1') {
            @{ Script = $script; Environment = 'Commercial'; Region = $null }
            @{ Script = $script; Environment = 'GCC'; Region = $null }
            @{ Script = $script; Environment = 'GCCHigh'; Region = 'ITAR' }
        }
    ) {
        $arguments = @{ OutputPath = $script:folder; Environment = $Environment; AdminUrl = 'https://contoso-admin.sharepoint.com'; WarningAction = 'SilentlyContinue' }
        if ($Script -ne 'Get-SiteSharingSettings.ps1') { $arguments['WaitMinutes'] = 0 }
        Invoke-CollectorScript $Script $arguments

        $expected = $Region
        Should -Invoke Connect-SPOService -Times 1 -Exactly -ParameterFilter {
            $Url -eq 'https://contoso-admin.sharepoint.com' -and $(if ($expected) { $Region -eq $expected } else { -not $Region })
        }
        Should -Invoke Disconnect-SPOService -Times 1 -Exactly
    }

    It 'asks for Sites.Read.All for the site list and Files.Read.All for item permissions' {
        Invoke-CollectorScript 'Get-Sites.ps1' @{ OutputPath = $script:folder; WarningAction = 'SilentlyContinue' }
        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter { $Scopes -contains 'Sites.Read.All' }
    }

    It 'requires -AdminUrl to sign in to SharePoint Online' {
        { Invoke-CollectorScript 'Get-SitePermissionBreadth.ps1' @{ OutputPath = $script:folder } } | Should -Throw '*-AdminUrl is required*'
    }

    It 'refuses a half-specified app-only SharePoint sign-in' {
        {
            Invoke-CollectorScript 'Get-SiteSharingSettings.ps1' @{ OutputPath = $script:folder; AdminUrl = 'https://contoso-admin.sharepoint.com'; AppId = 'app-1' }
        } | Should -Throw '*needs both -AppId and -CertificateThumbprint*'
    }
}

Describe 'Paging is followed' {
    BeforeEach {
        $script:folder = New-TestFolder
        $global:OversharingTestUris = [System.Collections.Generic.List[string]]::new()
    }
    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Variable -Name OversharingTestUris -Scope Global -ErrorAction SilentlyContinue
    }

    It 'follows the returned @odata.nextLink of getAllSites as it is (it changes the path)' {
        Mock Invoke-MgGraphRequest -MockWith {
            $global:OversharingTestUris.Add($Uri)
            if ($Uri -like '*skiptoken=page2') { New-MockSiteResponse -Id 'site-2' }
            else { New-MockSiteResponse -Id 'site-1' -NextLink 'https://graph.microsoft.com/v1.0/sites/oneDrive.getAllSites?$skiptoken=page2' }
        }

        Invoke-CollectorScript 'Get-Sites.ps1' @{ OutputPath = $script:folder; SkipConnect = $true }

        $global:OversharingTestUris.Count | Should -Be 2
        $global:OversharingTestUris[1] | Should -Be 'https://graph.microsoft.com/v1.0/sites/oneDrive.getAllSites?$skiptoken=page2'
        (@(Import-Csv -LiteralPath (Join-Path $script:folder 'sites.csv')).SiteId | Sort-Object) -join ',' | Should -Be 'site-1,site-2'
    }

    It 'stops on a nextLink it has already followed' {
        Mock Invoke-MgGraphRequest -MockWith { New-MockSiteResponse -NextLink 'https://graph.microsoft.com/v1.0/sites/getAllSites' }

        Invoke-CollectorScript 'Get-Sites.ps1' @{ OutputPath = $script:folder; SkipConnect = $true; WarningAction = 'SilentlyContinue' }

        Get-LogText -Folder $script:folder | Should -Match 'nextLink it had already returned'
        @(Import-Csv -LiteralPath (Join-Path $script:folder 'sites.csv')).Count | Should -Be 0
    }

    It 'walks every folder and pages the drives, children and permissions of item permissions' {
        Mock Invoke-MgGraphRequest -MockWith {
            $global:OversharingTestUris.Add($Uri)
            switch -Wildcard ($Uri) {
                '*/sites/site-1/drives?*' { @{ value = @(@{ id = 'drive-1' }) } }
                '*/drives/drive-1/root/children*' {
                    @{ value = @(@{ id = 'folder-1'; name = 'Plans'; folder = @{ childCount = 1 } }); '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/drives/drive-1/root/children?$skiptoken=2' }
                }
                '*/drives/drive-1/root/children?$skiptoken=2' { @{ value = @(@{ id = 'file-root'; name = 'root.docx' }) } }
                '*/drives/drive-1/items/folder-1/children*' { @{ value = @(@{ id = 'file-1'; name = 'plan.docx' }) } }
                '*/permissions' {
                    @{ value = @(@{ id = 'p-1'; roles = @('read'); link = @{ scope = 'organization'; type = 'view' } }); '@odata.nextLink' = "$Uri`?`$skiptoken=3" }
                }
                '*/permissions?$skiptoken=3' { @{ value = @(@{ id = 'p-2'; roles = @('write'); link = @{ scope = 'anonymous'; type = 'edit' } }) } }
                default { @{ value = @() } }
            }
        }

        Invoke-CollectorScript 'Get-ItemSharingPermissions.ps1' @{ OutputPath = $script:folder; SkipConnect = $true; SiteId = 'site-1'; WarningAction = 'SilentlyContinue' }

        ($global:OversharingTestUris | Where-Object { $_ -like '*root/children*skiptoken=2' }).Count | Should -Be 1
        ($global:OversharingTestUris | Where-Object { $_ -like '*items/folder-1/children*' }).Count | Should -Be 1
        $rows = @(Import-Csv -LiteralPath (Join-Path $script:folder 'item-permissions.csv'))
        ($rows | Where-Object ItemId -eq 'file-1' | Select-Object -ExpandProperty PermissionId | Sort-Object) -join ',' | Should -Be 'p-1,p-2'
        @($rows | Where-Object PermissionId -eq 'p-2').LinkScope | Select-Object -Unique | Should -Be 'anonymous'
    }

    It 'never writes a sharing link''s secret URL' {
        Mock Invoke-MgGraphRequest -MockWith {
            if ($Uri -like '*/permissions') { @{ value = @(@{ id = 'p-1'; link = @{ scope = 'anonymous'; type = 'view'; webUrl = 'https://contoso.sharepoint.com/:w:/s/secret'; } ; shareId = 'SECRETSHAREID' }) } }
            elseif ($Uri -like '*/drives?*') { @{ value = @(@{ id = 'drive-1' }) } }
            elseif ($Uri -like '*/children*') { @{ value = @(@{ id = 'file-1'; name = 'plan.docx' }) } }
        }

        Invoke-CollectorScript 'Get-ItemSharingPermissions.ps1' @{ OutputPath = $script:folder; SkipConnect = $true; SiteId = 'site-1'; WarningAction = 'SilentlyContinue' }

        (Get-Content -LiteralPath (Join-Path $script:folder 'item-permissions.csv') -Raw) | Should -Not -Match 'SECRETSHAREID|/s/secret'
    }

    It 'reads the next page of the audit log while moreRecordsAvailable is true, in one ReturnLargeSet session' {
        $global:OversharingTestSessions = [System.Collections.Generic.List[string]]::new()
        $global:OversharingTestPages = 0
        Mock Search-UnifiedAuditLog -MockWith {
            $global:OversharingTestSessions.Add("$SessionId|$SessionCommand|$ResultSize")
            $global:OversharingTestPages++
            if ($global:OversharingTestPages -eq 1) { New-MockAuditRecord -Id 'e-1' -MoreRecords $true }
            elseif ($global:OversharingTestPages -eq 2) { New-MockAuditRecord -Id 'e-2' -MoreRecords $false }
        }

        Invoke-CollectorScript 'Get-AnonymousLinkEvents.ps1' @{ OutputPath = $script:folder; SkipConnect = $true; StartDate = [datetime]'2026-08-10T00:00:00Z'; EndDate = [datetime]'2026-08-11T00:00:00Z' }

        $global:OversharingTestSessions.Count | Should -Be 2
        ($global:OversharingTestSessions | ForEach-Object { ($_ -split '\|')[0] } | Sort-Object -Unique).Count | Should -Be 1
        $global:OversharingTestSessions[0] | Should -Match '\|ReturnLargeSet\|5000$'
        (@(Import-Csv -LiteralPath (Join-Path $script:folder 'anonymous-link-events.csv')).Id | Sort-Object) -join ',' | Should -Be 'e-1,e-2'
        Remove-Variable -Name OversharingTestSessions, OversharingTestPages -Scope Global -ErrorAction SilentlyContinue
    }
}

Describe 'Event sources resume from the last exported timestamp' {
    BeforeEach {
        $script:folder = New-TestFolder
        $global:OversharingTestStarts = [System.Collections.Generic.List[datetime]]::new()
        Mock Start-Sleep -MockWith { }
    }
    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Variable -Name OversharingTestStarts -Scope Global -ErrorAction SilentlyContinue
    }

    It '<Script> searches from the newest CreationTime already in <Csv> on the second run' -ForEach @(
        @{ Script = 'Get-AnonymousLinkEvents.ps1'; Csv = 'anonymous-link-events.csv'; Operation = 'AnonymousLinkCreated' }
        @{ Script = 'Get-SharingEvents.ps1'; Csv = 'sharing-events.csv'; Operation = 'SharingSet' }
    ) {
        $op = $Operation
        Mock Search-UnifiedAuditLog -MockWith {
            $global:OversharingTestStarts.Add($StartDate.ToUniversalTime())
            New-MockAuditRecord -Id ('e-' + $global:OversharingTestStarts.Count) -Operation $op -CreationTime '2026-08-10T12:00:00'
        }
        $first = @{ OutputPath = $script:folder; SkipConnect = $true; StartDate = [datetime]'2026-08-10T00:00:00Z'; EndDate = [datetime]'2026-08-11T00:00:00Z' }

        Invoke-CollectorScript $Script $first
        Invoke-CollectorScript $Script @{ OutputPath = $script:folder; SkipConnect = $true; LookbackDays = 1; WindowHours = 24 }

        $global:OversharingTestStarts[0] | Should -Be ([datetime]'2026-08-10T00:00:00Z').ToUniversalTime()
        $global:OversharingTestStarts[1].ToString('yyyy-MM-ddTHH:mm:ss') | Should -Be '2026-08-10T12:00:00'
        @(Import-Csv -LiteralPath (Join-Path $script:folder $Csv)).Count | Should -BeGreaterThan 0
    }

    It 'searches the operations the contract lists' {
        Mock Search-UnifiedAuditLog -MockWith { @() }

        Invoke-CollectorScript 'Get-AnonymousLinkEvents.ps1' @{ OutputPath = $script:folder; SkipConnect = $true; StartDate = [datetime]'2026-08-10T00:00:00Z'; EndDate = [datetime]'2026-08-10T06:00:00Z' }

        Should -Invoke Search-UnifiedAuditLog -Times 1 -ParameterFilter {
            $Operations.Count -eq 4 -and $Operations -ccontains 'AnonymousLinkCreated' -and $Operations -ccontains 'AnonymousLinkRemoved' -and $Formatted
        }
    }

    It 'does not write a window that reaches the 50,000-record cap' {
        Mock Search-UnifiedAuditLog -MockWith { New-MockAuditRecord -ResultCount 60000 }

        {
            Invoke-CollectorScript 'Get-SharingEvents.ps1' @{
                OutputPath = $script:folder; SkipConnect = $true; WarningAction = 'SilentlyContinue'
                StartDate = [datetime]'2026-08-10T00:00:00Z'; EndDate = [datetime]'2026-08-11T00:00:00Z'
            }
        } | Should -Throw '*50,000*'

        @(Import-Csv -LiteralPath (Join-Path $script:folder 'sharing-events.csv')).Count | Should -Be 0
    }

    It 'does not write a window whose ResultCount is exactly 50,000 when the next call is empty' {
        $global:OversharingTestPages = 0
        Mock Search-UnifiedAuditLog -MockWith {
            $global:OversharingTestPages++
            if ($global:OversharingTestPages -eq 1) { New-MockAuditRecord -Id 'capped' -ResultCount 50000 -MoreRecords $false }
            else { @() }
        }

        {
            Invoke-CollectorScript 'Get-AnonymousLinkEvents.ps1' @{
                OutputPath = $script:folder; SkipConnect = $true; WarningAction = 'SilentlyContinue'
                StartDate = [datetime]'2026-08-10T00:00:00Z'; EndDate = [datetime]'2026-08-11T00:00:00Z'
            }
        } | Should -Throw '*50,000*'

        @(Import-Csv -LiteralPath (Join-Path $script:folder 'anonymous-link-events.csv')).Count | Should -Be 0
        Remove-Variable -Name OversharingTestPages -Scope Global -ErrorAction SilentlyContinue
    }
}

Describe 'Data access governance reports are read, reused and left running' {
    BeforeEach {
        $script:folder = New-TestFolder
        Set-DagMock
        Mock Get-SPOSite -MockWith { @() }
    }
    AfterEach { Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue }

    It 'starts the site permissions report for each workload with -Name and -CountOfUsersMoreThan, then exports it' {
        Invoke-CollectorScript 'Get-SitePermissionBreadth.ps1' @{ OutputPath = $script:folder; SkipConnect = $true }

        foreach ($workload in 'SharePoint', 'OneDriveForBusiness') {
            $expected = $workload
            Should -Invoke Start-SPODataAccessGovernanceInsight -Times 1 -Exactly -ParameterFilter {
                $ReportEntity -eq 'PermissionedUsers' -and $ReportType -eq 'Snapshot' -and $Workload -eq $expected -and $Name -and $CountOfUsersMoreThan -eq 0
            }
        }
        Should -Invoke Export-SPODataAccessGovernanceInsight -Times 2 -Exactly
        $row = @(Import-Csv -LiteralPath (Join-Path $script:folder 'site-permission-breadth.csv'))[0]
        $row.SiteUrl | Should -Be 'https://contoso.sharepoint.com/sites/finance'
        $row.UsersWithAccess | Should -Be '42'
    }

    It 'reuses a completed report that is newer than -MaxReportAgeHours instead of starting another' {
        Mock Get-SPODataAccessGovernanceInsight -MockWith {
            if ($ReportID) { [pscustomobject]@{ ReportId = $ReportID; Status = 'Completed' } }
            else { [pscustomobject]@{ ReportId = 'earlier'; Status = 'Completed'; CreatedDateTime = [datetime]::UtcNow.AddHours(-2); CountOfUsersMoreThan = 0 } }
        }

        Invoke-CollectorScript 'Get-SitePermissionBreadth.ps1' @{ OutputPath = $script:folder; SkipConnect = $true }

        Should -Invoke Start-SPODataAccessGovernanceInsight -Times 0 -Exactly
        Get-LogText -Folder $script:folder | Should -Match 'Reusing the completed'
    }

    It 'leaves a report that is still running, writes the header only and does not fail' {
        Set-DagMock -Status 'InProgress'

        { Invoke-CollectorScript 'Get-SitePermissionBreadth.ps1' @{ OutputPath = $script:folder; SkipConnect = $true; WaitMinutes = 0; WarningAction = 'SilentlyContinue' } } | Should -Not -Throw

        Should -Invoke Export-SPODataAccessGovernanceInsight -Times 0 -Exactly
        @(Import-Csv -LiteralPath (Join-Path $script:folder 'site-permission-breadth.csv')).Count | Should -Be 0
        Get-LogText -Folder $script:folder | Should -Match 'left running'
    }

    It 'starts one sensitivity label report per label GUID and keeps the whole row as JSON' {
        Invoke-CollectorScript 'Get-LabeledFileSites.ps1' @{ OutputPath = $script:folder; SkipConnect = $true; LabelGuid = [guid]'33333333-3333-3333-3333-333333333333' }

        Should -Invoke Start-SPODataAccessGovernanceInsight -Times 1 -Exactly -ParameterFilter {
            $ReportEntity -eq 'SensitivityLabelForFiles' -and $Workload -eq 'SharePoint' -and $FileSensitivityLabelGUID -eq [guid]'33333333-3333-3333-3333-333333333333'
        }
        $row = @(Import-Csv -LiteralPath (Join-Path $script:folder 'labeled-file-sites.csv'))[0]
        $row.LabelGuid | Should -Be '33333333-3333-3333-3333-333333333333'
        ($row.ReportRow | ConvertFrom-Json).'Site Name' | Should -Be 'Finance'
    }

    It 'starts the Everyone and Everyone-except-external-users item reports' {
        Invoke-CollectorScript 'Get-EveryoneItemExposure.ps1' @{ OutputPath = $script:folder; SkipConnect = $true }

        foreach ($entity in 'EveryoneExceptExternalUsers', 'Everyone') {
            $expected = $entity
            Should -Invoke Start-SPODataAccessGovernanceInsight -Times 1 -Exactly -ParameterFilter { $ReportEntity -eq $expected -and $ReportType -eq 'Snapshot' -and -not $Workload }
        }
    }

    It 'starts the sharing link and EEEU activity reports as RecentActivity for both workloads' {
        Invoke-CollectorScript 'Get-SharingLinkActivity.ps1' @{ OutputPath = $script:folder; SkipConnect = $true }
        Invoke-CollectorScript 'Get-EeeuActivity.ps1' @{ OutputPath = $script:folder; SkipConnect = $true }

        Should -Invoke Start-SPODataAccessGovernanceInsight -Times 6 -Exactly -ParameterFilter { $ReportType -eq 'RecentActivity' -and $ReportEntity -like 'SharingLinks_*' }
        # At site level SharePoint only; at item level both workloads (OneDrive supports the item level only).
        Should -Invoke Start-SPODataAccessGovernanceInsight -Times 3 -Exactly -ParameterFilter { $ReportType -eq 'RecentActivity' -and $ReportEntity -like 'EveryoneExceptExternalUsers*' -and $Name }
    }

    It 'only reads the data collection status and never starts data collection, which changes the tenant' {
        Mock Start-SPOAuditDataCollectionForActivityInsights -MockWith { }

        Invoke-CollectorScript 'Get-SharingLinkActivity.ps1' @{ OutputPath = $script:folder; SkipConnect = $true }

        Should -Invoke Get-SPOAuditDataCollectionStatusForActivityInsights -Times 1
        Should -Invoke Start-SPOAuditDataCollectionForActivityInsights -Times 0 -Exactly
    }

    It 'warns when data collection is not InProgress' {
        Mock Get-SPOAuditDataCollectionStatusForActivityInsights -MockWith { [pscustomobject]@{ Status = 'NotInitiated' } }

        Invoke-CollectorScript 'Get-SharingLinkActivity.ps1' @{ OutputPath = $script:folder; SkipConnect = $true; WarningAction = 'SilentlyContinue' }

        Get-LogText -Folder $script:folder | Should -Match "NotInitiated"
    }
}

Describe 'A source that cannot be read writes the header only' {
    BeforeEach { $script:folder = New-TestFolder }
    AfterEach { Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue }

    It 'Get-OversharingSourceAvailability skips only NotAvailable' {
        $custom = @{ SourceAvailability = @{ Demo = @{
                    Commercial = @{ Status = 'Available'; Reference = 'r' }
                    GCC        = @{ Status = 'NotAvailable'; Reference = 'r' }
                    GCCHigh    = @{ Status = 'Unverified'; Reference = 'r' }
                } } }

        (Get-OversharingSourceAvailability -Source 'Demo' -Environment 'Commercial' -Schema $custom).ShouldSkip | Should -BeFalse
        (Get-OversharingSourceAvailability -Source 'Demo' -Environment 'GCC' -Schema $custom).ShouldSkip | Should -BeTrue
        (Get-OversharingSourceAvailability -Source 'Demo' -Environment 'GCCHigh' -Schema $custom).ShouldSkip | Should -BeFalse
        { Get-OversharingSourceAvailability -Source 'Nope' -Environment 'GCC' -Schema $custom } | Should -Throw '*Unknown source*'
    }

    It 'the real schema marks no source NotAvailable, so every source is attempted in every cloud' {
        foreach ($entry in $script:Schema.SourceAvailability.GetEnumerator()) {
            foreach ($cloud in 'Commercial', 'GCC', 'GCCHigh') { $entry.Value[$cloud].Status | Should -BeIn @('Available', 'Unverified') }
        }
    }

    It 'Get-Sites.ps1 logs the refusal, leaves the header and does not throw' {
        Mock Invoke-MgGraphRequest -MockWith { throw 'Forbidden' }

        { Invoke-CollectorScript 'Get-Sites.ps1' @{ OutputPath = $script:folder; SkipConnect = $true; WarningAction = 'SilentlyContinue' } } | Should -Not -Throw

        Get-HeaderText -Path (Join-Path $script:folder 'sites.csv') | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples 'sites.csv'))
        @(Import-Csv -LiteralPath (Join-Path $script:folder 'sites.csv')).Count | Should -Be 0
        Get-LogText -Folder $script:folder | Should -Match 'Sites.Read.All'
    }

    It 'Get-AuditLogStatus.ps1 logs the refusal and leaves the header' {
        Mock Get-AdminAuditLogConfig -MockWith { throw 'denied' }

        Invoke-CollectorScript 'Get-AuditLogStatus.ps1' @{ OutputPath = $script:folder; SkipConnect = $true; WarningAction = 'SilentlyContinue' }

        @(Import-Csv -LiteralPath (Join-Path $script:folder 'audit-log-status.csv')).Count | Should -Be 0
    }

    It 'Get-SiteSharingSettings.ps1 leaves the header when Get-SPOSite is refused' {
        Mock Get-SPOTenant -MockWith { [pscustomobject]@{ SharingCapability = 'Disabled' } }
        Mock Get-SPOSite -MockWith { throw 'denied' }

        Invoke-CollectorScript 'Get-SiteSharingSettings.ps1' @{ OutputPath = $script:folder; SkipConnect = $true; WarningAction = 'SilentlyContinue' }

        @(Import-Csv -LiteralPath (Join-Path $script:folder 'site-sharing-settings.csv')).Count | Should -Be 0
    }
}

Describe 'An UNVERIFIED source is attempted and warned about' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Get-SPOTenant -MockWith { [pscustomobject]@{ SharingCapability = 'Disabled' } }
        Mock Get-SPOSite -MockWith { @() }
        Mock Get-AdminAuditLogConfig -MockWith { [pscustomobject]@{ UnifiedAuditLogIngestionEnabled = $true } }
    }
    AfterEach { Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue }

    It 'reads site sharing settings and logs UNVERIFIED in <Environment>' -ForEach @(
        @{ Environment = 'Commercial' }, @{ Environment = 'GCC' }, @{ Environment = 'GCCHigh' }
    ) {
        Invoke-CollectorScript 'Get-SiteSharingSettings.ps1' @{ OutputPath = $script:folder; SkipConnect = $true; Environment = $Environment; WarningAction = 'SilentlyContinue' }

        Should -Invoke Get-SPOSite -Times 1 -Exactly
        Get-LogText -Folder $script:folder | Should -Match 'UNVERIFIED'
    }

    It 'reads the audit log flag in GCC High and logs UNVERIFIED' {
        Invoke-CollectorScript 'Get-AuditLogStatus.ps1' @{ OutputPath = $script:folder; SkipConnect = $true; Environment = 'GCCHigh'; WarningAction = 'SilentlyContinue' }

        Should -Invoke Get-AdminAuditLogConfig -Times 1 -Exactly
        Get-LogText -Folder $script:folder | Should -Match 'UNVERIFIED'
    }

    It 'logs that an app-only Graph caller may not see every link (row 2)' {
        Mock Invoke-MgGraphRequest -MockWith { @{ value = @() } }
        Invoke-CollectorScript 'Get-ItemSharingPermissions.ps1' @{ OutputPath = $script:folder; SkipConnect = $true; SiteId = 'site-1'; WarningAction = 'SilentlyContinue' }

        Get-LogText -Folder $script:folder | Should -Match 'UNVERIFIED'
    }
}

Describe 'Run-All.ps1' {
    BeforeEach {
        $script:folder = New-TestFolder
        Set-DagMock
        Mock Connect-M365Service -MockWith { }
        Mock Connect-SPOService -MockWith { }
        Mock Disconnect-SPOService -MockWith { }
        Mock Disconnect-ExchangeOnline -MockWith { }
        Mock Invoke-MgGraphRequest -MockWith { if ($Uri -like '*getAllSites*') { New-MockSiteResponse } else { @{ value = @() } } }
        Mock Get-SPOTenant -MockWith { [pscustomobject]@{ SharingCapability = 'Disabled' } }
        Mock Get-SPOSite -MockWith { @() }
        Mock Search-UnifiedAuditLog -MockWith { @() }
        Mock Get-AdminAuditLogConfig -MockWith { [pscustomobject]@{ UnifiedAuditLogIngestionEnabled = $true } }
        Mock Get-Label -MockWith { @() }
        Mock Start-Sleep -MockWith { }
    }
    AfterEach { Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue }

    It 'writes every CSV of the report, with header-only files where nothing came back' {
        Invoke-CollectorScript 'Run-All.ps1' @{
            OutputPath = $script:folder; AdminUrl = 'https://contoso-admin.sharepoint.com'; LabelGuid = [guid]'33333333-3333-3333-3333-333333333333'
            WarningAction = 'SilentlyContinue'; WaitMinutes = 0
            StartDate = [datetime]'2026-08-10T00:00:00Z'; EndDate = [datetime]'2026-08-10T06:00:00Z'
        }

        foreach ($file in Get-ChildItem -LiteralPath $script:Samples -Filter '*.csv') {
            Test-Path -LiteralPath (Join-Path $script:folder $file.Name) | Should -BeTrue -Because $file.Name
        }
    }

    It 'keeps going when one collector fails, then throws' {
        # A 50,000-record window is the one failure a collector throws on.
        Mock Search-UnifiedAuditLog -MockWith { New-MockAuditRecord -ResultCount 60000 }

        {
            Invoke-CollectorScript 'Run-All.ps1' @{
                OutputPath = $script:folder; AdminUrl = 'https://contoso-admin.sharepoint.com'; WarningAction = 'SilentlyContinue'; WaitMinutes = 0
                StartDate = [datetime]'2026-08-10T00:00:00Z'; EndDate = [datetime]'2026-08-10T06:00:00Z'
            }
        } | Should -Throw '*2 collector(s) stopped*'

        Test-Path -LiteralPath (Join-Path $script:folder 'sites.csv') | Should -BeTrue
    }
}
