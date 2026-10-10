#Requires -Version 7.0

BeforeAll {
    $script:Root = Join-Path $PSScriptRoot '../../..' | Resolve-Path | Select-Object -ExpandProperty Path
    $script:Collectors = Join-Path $script:Root 'reports/sharepoint-onedrive-activity/collectors'
    $script:Samples = Join-Path $script:Root 'reports/sharepoint-onedrive-activity/samples'

    . (Join-Path $script:Root 'shared/tests/TenantCmdletStubs.ps1')
    . (Join-Path $PSScriptRoot 'SharePointOneDriveStubs.ps1')
    Import-Module (Join-Path $script:Root 'shared/M365ReportLibrary.psm1') -Force
    . (Join-Path $script:Collectors 'SharePointOneDriveHelpers.ps1')

    $script:Schema = Import-PowerShellDataFile -LiteralPath (Join-Path $script:Collectors 'SharePointOneDriveSchema.psd1')

    function New-TestFolder {
        $path = Join-Path ([System.IO.Path]::GetTempPath()) ('spo-activity-' + [guid]::NewGuid().ToString('N'))
        New-Item -Path $path -ItemType Directory -Force | Out-Null
        return $path
    }

    function Get-HeaderText { param([string]$Path) return ((Get-CsvHeaderColumn -Path $Path) -join ',') }

    function Get-LogText { param([string]$Folder) return (Get-Content -LiteralPath (Join-Path $Folder 'run.log') -Raw) }

    function Invoke-CollectorScript {
        param([Parameter(Mandatory)][string]$Name, [hashtable]$Arguments = @{})
        & (Join-Path $script:Collectors $Name) @Arguments
    }

    # Header lists of the Graph usage reports, in the order the contract gives them, which it
    # takes from each API page. A mock CSV is a header and one row.
    # https://learn.microsoft.com/graph/api/reportroot-getsharepointsiteusagedetail
    # https://learn.microsoft.com/graph/api/reportroot-getonedriveusageaccountdetail
    # https://learn.microsoft.com/graph/api/reportroot-getsharepointsiteusagestorage
    # https://learn.microsoft.com/graph/api/reportroot-getonedriveusagestorage
    # https://learn.microsoft.com/graph/api/reportroot-getsharepointactivityuserdetail
    # https://learn.microsoft.com/graph/api/reportroot-getonedriveactivityuserdetail
    $global:SpoTest = @{
        Headers = @{
            SiteDetail      = 'Report Refresh Date,Site Id,Site URL,Owner Display Name,Is Deleted,Last Activity Date,File Count,Active File Count,Page View Count,Visited Page Count,Storage Used (Byte),Storage Allocated (Byte),Root Web Template,Owner Principal Name,Report Period'
            AccountDetail   = 'Report Refresh Date,Site URL,Owner Display Name,Is Deleted,Last Activity Date,File Count,Active File Count,Storage Used (Byte),Storage Allocated (Byte),Owner Principal Name,Report Period'
            Storage         = 'Report Refresh Date,Site Type,Storage Used (Byte),Report Date,Report Period'
            SharePointUser  = 'Report Refresh Date,User Principal Name,Is Deleted,Deleted Date,Last Activity Date,Viewed Or Edited File Count,Synced File Count,Shared Internally File Count,Shared Externally File Count,Visited Page Count,Assigned Products,Report Period'
            OneDriveUser    = 'Report Refresh Date,User Principal Name,Is Deleted,Deleted Date,Last Activity Date,Viewed Or Edited File Count,Synced File Count,Shared Internally File Count,Shared Externally File Count,Assigned Products,Report Period'
        }
        Rows = @{
            SiteDetail      = '2026-10-06,site-1,https://contoso.sharepoint.example.com/sites/finance,Avery Abara,False,2026-10-05,100,10,40,12,2048,4096,GROUP#0,avery.abara@example.com,30'
            AccountDetail   = '2026-10-06,https://contoso-my.sharepoint.example.com/personal/avery,Avery Abara,False,2026-10-05,100,10,2048,4096,avery.abara@example.com,30'
            Storage         = '2026-10-06,All,2048,2026-10-05,30'
            SharePointUser  = '2026-10-06,avery.abara@example.com,False,,2026-10-05,5,4,3,2,1,MICROSOFT 365 E3,7'
            OneDriveUser    = '2026-10-06,avery.abara@example.com,False,,2026-10-05,5,4,3,2,MICROSOFT 365 E3,7'
        }
        Files = [System.Collections.Generic.List[string]]::new()
        Csv = @{}
    }

    function Set-UsageMock {
        # Writes the report to -OutFile, as the cmdlet does after Graph's 302 redirect.
        param([string]$Cmdlet, [string]$Kind)
        $global:SpoTest.Csv[$Kind] = $global:SpoTest.Headers[$Kind] + "`n" + $global:SpoTest.Rows[$Kind]
        # One script block per mock, so each mock keeps its own report. $OutFile is bound by the mock.
        Mock $Cmdlet -MockWith ([scriptblock]::Create("`$global:SpoTest.Files.Add(`$OutFile); Set-Content -LiteralPath `$OutFile -Encoding utf8 -Value `$global:SpoTest.Csv['$Kind']"))
    }

    # A site from GET /sites/getAllSites. https://learn.microsoft.com/graph/api/site-getallsites
    function New-MockSite {
        param([string]$Id, [string]$Url = 'https://contoso.sharepoint.example.com/sites/finance', [bool]$Personal = $false)
        @{ id = $Id; name = 'finance'; webUrl = $Url; isPersonalSite = $Personal; siteCollection = @{ hostname = 'contoso.sharepoint.example.com' } }
    }

    # A drive with its quota. https://learn.microsoft.com/graph/api/resources/quota
    function New-MockDrive {
        param([string]$Id, [int64]$Used = 2048)
        @{ id = $Id; name = 'Documents'; driveType = 'documentLibrary'; webUrl = 'https://contoso.sharepoint.example.com/Shared Documents'
            lastModifiedDateTime = '2026-10-07T14:03:11Z'
            quota = @{ deleted = 10; remaining = (4096 - $Used); state = 'normal'; total = 4096; used = $Used } }
    }

    # Search-UnifiedAuditLog record: RecordType, ResultCount, AuditData as a JSON string.
    # https://learn.microsoft.com/powershell/module/exchangepowershell/search-unifiedauditlog
    function New-MockAuditRecord {
        param(
            [string]$Id = 'event-1',
            [string]$Operation = 'FileAccessed',
            [string]$CreationTime = '2026-10-05T12:00:00',
            [string]$UserId = 'avery.abara@example.com',
            [string]$Workload = 'SharePoint',
            [object]$MoreRecords = $null
        )
        $auditData = [ordered]@{ CreationTime = $CreationTime; Id = $Id; Operation = $Operation; UserId = $UserId; Workload = $Workload } | ConvertTo-Json -Compress
        $record = [ordered]@{ RecordType = 'SharePointFileOperation'; AuditData = $auditData }
        if ($null -ne $MoreRecords) { $record['AuditSearchRequestMetadata'] = [pscustomobject]@{ moreRecordsAvailable = $MoreRecords } }
        return [pscustomobject]$record
    }
}

AfterAll {
    Remove-Variable -Name SpoTest -Scope Global -ErrorAction SilentlyContinue
    Remove-Module M365ReportLibrary -Force -ErrorAction SilentlyContinue
}

Describe 'Each collector writes the columns of its sample' {
    BeforeEach {
        $script:Out = New-TestFolder
        Mock Connect-MgGraph -ModuleName M365ReportLibrary -MockWith { }
        Mock Connect-ExchangeOnline -ModuleName M365ReportLibrary -MockWith { }
        Mock Disconnect-ExchangeOnline { }
        Mock Start-Sleep { }
        Set-UsageMock 'Get-MgReportSharePointSiteUsageDetail' 'SiteDetail'
        Set-UsageMock 'Get-MgReportOneDriveUsageAccountDetail' 'AccountDetail'
        Set-UsageMock 'Get-MgReportSharePointSiteUsageStorage' 'Storage'
        Set-UsageMock 'Get-MgReportOneDriveUsageStorage' 'Storage'
        Set-UsageMock 'Get-MgReportSharePointActivityUserDetail' 'SharePointUser'
        Set-UsageMock 'Get-MgReportOneDriveActivityUserDetail' 'OneDriveUser'
        Mock Get-MgAdminReportSetting { [pscustomobject]@{ DisplayConcealedNames = $false } }
        Mock Get-SPOTenant { [pscustomobject]@{ StorageQuota = 5242880; StorageQuotaAllocated = 1048576; ResourceQuota = 300; ResourceQuotaAllocated = 0; OneDriveStorageQuota = 1048576 } }
        Mock Get-SPOSite { [pscustomobject]@{ Url = 'https://contoso.sharepoint.example.com/sites/finance'; Title = 'Finance'; Template = 'GROUP#0'; StorageUsageCurrent = 1024; ResourceUsageCurrent = 0; WebsCount = 1 } }
        Mock Invoke-MgGraphRequest {
            if ($Uri -like '*getAllSites*') { @{ value = @(New-MockSite -Id 'site-1') } }
            elseif ($Uri -like '*/drives*') { @{ value = @(New-MockDrive -Id 'drive-1') } }
            elseif ($Uri -like '*getActivitiesByInterval*') {
                @{ value = @(@{ startDateTime = '2026-10-05T00:00:00Z'; endDateTime = '2026-10-06T00:00:00Z'; access = @{ actionCount = 5; actorCount = 3 }; edit = @{ actionCount = 2; actorCount = 1 } }) }
            }
        }
        Mock Search-UnifiedAuditLog { New-MockAuditRecord }
    }
    AfterEach { Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue }

    It '<Script> writes <Csv> with the schema columns, the sample columns and at least one row' -ForEach @(
        @{ Script = 'Get-SharePointSiteUsageDetail.ps1'; Csv = 'sharepoint-site-usage-detail.csv'; Key = 'SharePointSiteUsageDetail'; Extra = @{} }
        @{ Script = 'Get-OneDriveUsageAccountDetail.ps1'; Csv = 'onedrive-usage-account-detail.csv'; Key = 'OneDriveUsageAccountDetail'; Extra = @{} }
        @{ Script = 'Get-SharePointSiteUsageStorage.ps1'; Csv = 'sharepoint-site-usage-storage.csv'; Key = 'SharePointSiteUsageStorage'; Extra = @{} }
        @{ Script = 'Get-OneDriveUsageStorage.ps1'; Csv = 'onedrive-usage-storage.csv'; Key = 'OneDriveUsageStorage'; Extra = @{} }
        @{ Script = 'Get-SharePointActivityUserDetail.ps1'; Csv = 'sharepoint-activity-user-detail.csv'; Key = 'SharePointActivityUserDetail'; Extra = @{} }
        @{ Script = 'Get-OneDriveActivityUserDetail.ps1'; Csv = 'onedrive-activity-user-detail.csv'; Key = 'OneDriveActivityUserDetail'; Extra = @{} }
        @{ Script = 'Get-ReportSettings.ps1'; Csv = 'report-settings.csv'; Key = 'ReportSettings'; Extra = @{} }
        @{ Script = 'Get-TenantStorage.ps1'; Csv = 'tenant-storage.csv'; Key = 'TenantStorage'; Extra = @{} }
        @{ Script = 'Get-SpoSites.ps1'; Csv = 'spo-sites.csv'; Key = 'SpoSites'; Extra = @{} }
        @{ Script = 'Get-DriveQuota.ps1'; Csv = 'drive-quota.csv'; Key = 'DriveQuota'; Extra = @{} }
        @{ Script = 'Get-FileEvents.ps1'; Csv = 'file-events.csv'; Key = 'FileEvents'; Extra = @{ StartDate = [datetime]'2026-10-05T00:00:00Z'; EndDate = [datetime]'2026-10-06T00:00:00Z' } }
        @{ Script = 'Get-SiteActivity.ps1'; Csv = 'site-activity.csv'; Key = 'SiteActivity'; Extra = @{} }
    ) {
        $arguments = @{ OutputPath = $script:Out; SkipConnect = $true } + $Extra
        Invoke-CollectorScript $Script $arguments

        $path = Join-Path $script:Out $Csv
        Get-HeaderText $path | Should -Be ($script:Schema[$Key] -join ',')
        Get-HeaderText $path | Should -Be (Get-HeaderText (Join-Path $script:Samples $Csv))
        @(Import-Csv -LiteralPath $path).Count | Should -BeGreaterThan 0
    }

    It 'reads each Graph usage report column by its header name' {
        Invoke-CollectorScript 'Get-SharePointSiteUsageDetail.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        $row = Import-Csv -LiteralPath (Join-Path $script:Out 'sharepoint-site-usage-detail.csv')
        $row.SiteId | Should -Be 'site-1'
        $row.SiteUrl | Should -Be 'https://contoso.sharepoint.example.com/sites/finance'
        $row.StorageUsedByte | Should -Be '2048'
        $row.StorageAllocatedByte | Should -Be '4096'
        $row.RootWebTemplate | Should -Be 'GROUP#0'
        $row.ReportPeriod | Should -Be '30'
        foreach ($file in $global:SpoTest.Files) { Test-Path -LiteralPath $file | Should -BeFalse }
    }
}

Describe 'Cloud availability' {
    BeforeEach {
        $script:Out = New-TestFolder
        Mock Connect-MgGraph -ModuleName M365ReportLibrary -MockWith { }
        Mock Connect-ExchangeOnline -ModuleName M365ReportLibrary -MockWith { }
        Mock Connect-SPOService { }
        Mock Disconnect-SPOService { }
        Mock Disconnect-ExchangeOnline { }
        Mock Get-MgReportSharePointSiteUsageDetail { }
        Mock Get-MgReportOneDriveUsageAccountDetail { }
        Mock Get-MgReportSharePointSiteUsageStorage { }
        Mock Get-MgReportOneDriveUsageStorage { }
        Mock Get-MgReportSharePointActivityUserDetail { }
        Mock Get-MgReportOneDriveActivityUserDetail { }
        Mock Get-MgAdminReportSetting { [pscustomobject]@{ DisplayConcealedNames = $false } }
        Mock Get-SPOTenant { [pscustomobject]@{ StorageQuota = 1 } }
        Mock Get-SPOSite { [pscustomobject]@{ Url = 'https://contoso.sharepoint.example.com'; Title = 'Root'; Template = 'SITEPAGEPUBLISHING#0' } }
        Mock Search-UnifiedAuditLog { }
        Mock Invoke-MgGraphRequest { @{ value = @() } }
    }
    AfterEach { Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue }

    It 'writes a header only for <Script> in GCC High and never calls the report' -ForEach @(
        @{ Script = 'Get-SharePointSiteUsageDetail.ps1'; Csv = 'sharepoint-site-usage-detail.csv'; Key = 'SharePointSiteUsageDetail'; Cmdlet = 'Get-MgReportSharePointSiteUsageDetail' }
        @{ Script = 'Get-OneDriveUsageAccountDetail.ps1'; Csv = 'onedrive-usage-account-detail.csv'; Key = 'OneDriveUsageAccountDetail'; Cmdlet = 'Get-MgReportOneDriveUsageAccountDetail' }
        @{ Script = 'Get-SharePointSiteUsageStorage.ps1'; Csv = 'sharepoint-site-usage-storage.csv'; Key = 'SharePointSiteUsageStorage'; Cmdlet = 'Get-MgReportSharePointSiteUsageStorage' }
        @{ Script = 'Get-OneDriveUsageStorage.ps1'; Csv = 'onedrive-usage-storage.csv'; Key = 'OneDriveUsageStorage'; Cmdlet = 'Get-MgReportOneDriveUsageStorage' }
        @{ Script = 'Get-SharePointActivityUserDetail.ps1'; Csv = 'sharepoint-activity-user-detail.csv'; Key = 'SharePointActivityUserDetail'; Cmdlet = 'Get-MgReportSharePointActivityUserDetail' }
        @{ Script = 'Get-OneDriveActivityUserDetail.ps1'; Csv = 'onedrive-activity-user-detail.csv'; Key = 'OneDriveActivityUserDetail'; Cmdlet = 'Get-MgReportOneDriveActivityUserDetail' }
    ) {
        Invoke-CollectorScript $Script @{ OutputPath = $script:Out; Environment = 'GCCHigh' }

        $path = Join-Path $script:Out $Csv
        Get-HeaderText $path | Should -Be ($script:Schema[$Key] -join ',')
        @(Import-Csv -LiteralPath $path).Count | Should -Be 0
        Get-LogText $script:Out | Should -Match 'documented as unavailable in GCCHigh'
        Should -Invoke $Cmdlet -Times 0 -Exactly
        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 0 -Exactly
    }

    It 'attempts <Script> in <Cloud> and logs a warning that availability is UNVERIFIED' -ForEach @(
        @{ Script = 'Get-TenantStorage.ps1'; Cloud = 'Commercial'; Cmdlet = 'Get-SPOTenant' }
        @{ Script = 'Get-TenantStorage.ps1'; Cloud = 'GCCHigh'; Cmdlet = 'Get-SPOTenant' }
        @{ Script = 'Get-SpoSites.ps1'; Cloud = 'GCC'; Cmdlet = 'Get-SPOSite' }
        @{ Script = 'Get-SpoSites.ps1'; Cloud = 'GCCHigh'; Cmdlet = 'Get-SPOSite' }
        @{ Script = 'Get-ReportSettings.ps1'; Cloud = 'GCCHigh'; Cmdlet = 'Get-MgAdminReportSetting' }
        @{ Script = 'Get-SiteActivity.ps1'; Cloud = 'GCCHigh'; Cmdlet = 'Invoke-MgGraphRequest' }
    ) {
        Invoke-CollectorScript $Script @{ OutputPath = $script:Out; Environment = $Cloud; SkipConnect = $true }
        Get-LogText $script:Out | Should -Match 'UNVERIFIED'
        Should -Invoke $Cmdlet -Times 1 -Exactly -Scope It
    }

    It 'does not warn that <Script> is unverified in <Cloud>, where the contract marks it available' -ForEach @(
        @{ Script = 'Get-DriveQuota.ps1'; Cloud = 'GCCHigh' }
        @{ Script = 'Get-FileEvents.ps1'; Cloud = 'GCCHigh' }
        @{ Script = 'Get-ReportSettings.ps1'; Cloud = 'GCC' }
    ) {
        Invoke-CollectorScript $Script @{ OutputPath = $script:Out; Environment = $Cloud; SkipConnect = $true }
        Get-LogText $script:Out | Should -Not -Match 'UNVERIFIED'
    }
}

Describe 'Connection endpoint per -Environment' {
    BeforeEach {
        $script:Out = New-TestFolder
        Mock Connect-MgGraph -ModuleName M365ReportLibrary -MockWith { }
        Mock Connect-ExchangeOnline -ModuleName M365ReportLibrary -MockWith { }
        Mock Connect-SPOService { }
        Mock Disconnect-SPOService { }
        Mock Disconnect-ExchangeOnline { }
        Mock Get-MgReportSharePointSiteUsageDetail { }
        Mock Get-SPOTenant { [pscustomobject]@{ StorageQuota = 1 } }
        Mock Search-UnifiedAuditLog { }
        Mock Invoke-MgGraphRequest { @{ value = @() } }
    }
    AfterEach { Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue }

    It 'signs in to Graph with the <Graph> environment for <Cloud>' -ForEach @(
        @{ Cloud = 'Commercial'; Graph = 'Global' }
        @{ Cloud = 'GCC'; Graph = 'Global' }
    ) {
        Invoke-CollectorScript 'Get-SharePointSiteUsageDetail.ps1' @{ OutputPath = $script:Out; Environment = $Cloud }
        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter { $Environment -eq $Graph }
    }

    It 'signs in to Graph with the USGov environment for GCC High' {
        Invoke-CollectorScript 'Get-DriveQuota.ps1' @{ OutputPath = $script:Out; Environment = 'GCCHigh' }
        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter { $Environment -eq 'USGov' }
    }

    It 'requests the site and file read scopes for the site collectors' -ForEach @(
        @{ Script = 'Get-DriveQuota.ps1' }
        @{ Script = 'Get-SiteActivity.ps1' }
    ) {
        Invoke-CollectorScript $Script @{ OutputPath = $script:Out }
        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter { $Scopes -contains 'Sites.Read.All' -and $Scopes -contains 'Files.Read.All' }
    }

    It 'signs in to Exchange Online with <Name> for <Cloud>' -ForEach @(
        @{ Cloud = 'Commercial'; Name = 'O365Default' }
        @{ Cloud = 'GCC'; Name = 'O365Default' }
        @{ Cloud = 'GCCHigh'; Name = 'O365USGovGCCHigh' }
    ) {
        Invoke-CollectorScript 'Get-FileEvents.ps1' @{ OutputPath = $script:Out; Environment = $Cloud }
        Should -Invoke Connect-ExchangeOnline -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter { $ExchangeEnvironmentName -eq $Name }
        Should -Invoke Disconnect-ExchangeOnline -Times 1 -Exactly
    }

    It 'signs in to SharePoint Online with <Region> for <Cloud>' -ForEach @(
        @{ Cloud = 'Commercial'; Region = $null }
        @{ Cloud = 'GCC'; Region = $null }
        @{ Cloud = 'GCCHigh'; Region = 'ITAR' }
    ) {
        Invoke-CollectorScript 'Get-TenantStorage.ps1' @{ OutputPath = $script:Out; Environment = $Cloud; AdminUrl = 'https://contoso-admin.sharepoint.example.com' }
        if ($Region) {
            Should -Invoke Connect-SPOService -Times 1 -Exactly -ParameterFilter { $Region -eq 'ITAR' -and $Url -eq 'https://contoso-admin.sharepoint.example.com' }
        }
        else {
            Should -Invoke Connect-SPOService -Times 1 -Exactly -ParameterFilter { -not $Region -and $Url -eq 'https://contoso-admin.sharepoint.example.com' }
        }
        Should -Invoke Disconnect-SPOService -Times 1 -Exactly
    }

    It 'needs -AdminUrl to sign in to SharePoint Online' {
        { Invoke-CollectorScript 'Get-SpoSites.ps1' @{ OutputPath = $script:Out } } | Should -Throw '*-AdminUrl*'
        Should -Invoke Connect-SPOService -Times 0 -Exactly
    }

    It 'does not sign in when asked to reuse a session' {
        Invoke-CollectorScript 'Get-FileEvents.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        Invoke-CollectorScript 'Get-TenantStorage.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        Should -Invoke Connect-ExchangeOnline -ModuleName M365ReportLibrary -Times 0 -Exactly
        Should -Invoke Connect-SPOService -Times 0 -Exactly
        Should -Invoke Disconnect-ExchangeOnline -Times 0 -Exactly
        Should -Invoke Disconnect-SPOService -Times 0 -Exactly
    }

    It 'stops half an app-only credential' {
        { Invoke-CollectorScript 'Get-TenantStorage.ps1' @{ OutputPath = $script:Out; AdminUrl = 'https://contoso-admin.sharepoint.example.com'; AppId = 'app' } } | Should -Throw '*-AppId and -CertificateThumbprint*'
    }
}

Describe 'State sources stamp the run date and append' {
    BeforeEach {
        $script:Out = New-TestFolder
        $script:RunDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')
        $global:SpoTest.Files.Clear()
        Mock Get-SPOTenant { [pscustomobject]@{ StorageQuota = 5242880; StorageQuotaAllocated = 1048576; ResourceQuota = 300; ResourceQuotaAllocated = 0; OneDriveStorageQuota = 1048576 } }
        Mock Get-SPOSite { [pscustomobject]@{ Url = 'https://contoso.sharepoint.example.com/sites/finance'; Title = 'Finance'; Template = 'GROUP#0' } }
        Mock Get-MgAdminReportSetting { [pscustomobject]@{ AdditionalProperties = @{ displayConcealedNames = $true } } }
    }
    AfterEach { Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue }

    It 'stamps a usage report with the run date and does not duplicate a row on a second run' {
        Set-UsageMock 'Get-MgReportSharePointSiteUsageDetail' 'SiteDetail'
        Invoke-CollectorScript 'Get-SharePointSiteUsageDetail.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; Period = 'D7' }
        Invoke-CollectorScript 'Get-SharePointSiteUsageDetail.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; Period = 'D7' }

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:Out 'sharepoint-site-usage-detail.csv'))
        $rows.Count | Should -Be 1
        $rows[0].RunDate | Should -Be $script:RunDate
        Should -Invoke Get-MgReportSharePointSiteUsageDetail -Times 2 -Exactly -ParameterFilter { $Period -eq 'D7' }
    }

    It 'sends -Date and not -Period for a one-day activity report and records the day asked for' {
        Set-UsageMock 'Get-MgReportOneDriveActivityUserDetail' 'OneDriveUser'
        Invoke-CollectorScript 'Get-OneDriveActivityUserDetail.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; Date = [datetime]'2026-10-05' }

        $row = Import-Csv -LiteralPath (Join-Path $script:Out 'onedrive-activity-user-detail.csv')
        $row.RunDate | Should -Be $script:RunDate
        $row.QueryDate | Should -Be '2026-10-05'
        Should -Invoke Get-MgReportOneDriveActivityUserDetail -Times 1 -Exactly -ParameterFilter { $Date -eq [datetime]'2026-10-05' -and -not $Period }
    }

    It 'sends the calendar day as UTC midnight, not a local conversion of the clock time' {
        # https://learn.microsoft.com/graph/api/reportroot-getsharepointactivityuserdetail
        Set-UsageMock 'Get-MgReportSharePointActivityUserDetail' 'SharePointUser'
        $stamp = [datetime]::new(2026, 10, 5, 23, 30, 0, [DateTimeKind]::Unspecified)
        Invoke-CollectorScript 'Get-SharePointActivityUserDetail.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; Date = $stamp }

        (Import-Csv -LiteralPath (Join-Path $script:Out 'sharepoint-activity-user-detail.csv')).QueryDate | Should -Be '2026-10-05'
        Should -Invoke Get-MgReportSharePointActivityUserDetail -Times 1 -Exactly -ParameterFilter {
            $Date.Kind -eq [DateTimeKind]::Utc -and $Date.Hour -eq 0 -and $Date.ToString('yyyy-MM-dd') -eq '2026-10-05'
        }
    }

    It 'rejects an activity date outside the past 30 days instead of calling the report' {
        # The date form is only the past 30 days. A refusal log would be the wrong result.
        Set-UsageMock 'Get-MgReportOneDriveActivityUserDetail' 'OneDriveUser'
        $tooOld = [datetime]::UtcNow.Date.AddDays(-31)
        { Invoke-CollectorScript 'Get-OneDriveActivityUserDetail.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; Date = $tooOld } } | Should -Throw '*past 30 days*'
        Should -Invoke Get-MgReportOneDriveActivityUserDetail -Times 0 -Exactly

        Set-UsageMock 'Get-MgReportSharePointActivityUserDetail' 'SharePointUser'
        $oldest = [datetime]::UtcNow.Date.AddDays(-30)
        Invoke-CollectorScript 'Get-SharePointActivityUserDetail.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; Date = $oldest }
        Should -Invoke Get-MgReportSharePointActivityUserDetail -Times 1 -Exactly -ParameterFilter { $Date.ToString('yyyy-MM-dd') -eq $oldest.ToString('yyyy-MM-dd') }
    }

    It 'retries a usage report 429 and does not record it as a refused report' {
        # https://learn.microsoft.com/graph/throttling-limits
        Mock Start-Sleep { }
        $global:SpoTest.UsageTries = 0
        Mock Get-MgReportSharePointSiteUsageDetail {
            $global:SpoTest.UsageTries++
            if ($global:SpoTest.UsageTries -eq 1) { throw 'Response status code does not indicate success: 429 (Too Many Requests).' }
            Set-Content -LiteralPath $OutFile -Encoding utf8 -Value ($global:SpoTest.Headers.SiteDetail + "`n" + $global:SpoTest.Rows.SiteDetail)
        }
        Invoke-CollectorScript 'Get-SharePointSiteUsageDetail.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }

        (Import-Csv -LiteralPath (Join-Path $script:Out 'sharepoint-site-usage-detail.csv')).SiteId | Should -Be 'site-1'
        $global:SpoTest.UsageTries | Should -Be 2
        Get-LogText $script:Out | Should -Match 'HTTP 429'
        Get-LogText $script:Out | Should -Not -Match 'unavailable to this sign-in'
    }

    It 'sends -Period and records no query date for a period activity report' {
        Set-UsageMock 'Get-MgReportSharePointActivityUserDetail' 'SharePointUser'
        Invoke-CollectorScript 'Get-SharePointActivityUserDetail.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; Period = 'D30' }

        (Import-Csv -LiteralPath (Join-Path $script:Out 'sharepoint-activity-user-detail.csv')).QueryDate | Should -Be ''
        Should -Invoke Get-MgReportSharePointActivityUserDetail -Times 1 -Exactly -ParameterFilter { $Period -eq 'D30' -and -not $Date }
    }

    It 'writes a header only and logs the reason when a usage report is refused' {
        Mock Get-MgReportOneDriveUsageStorage { throw 'Forbidden' }
        Invoke-CollectorScript 'Get-OneDriveUsageStorage.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }

        Get-HeaderText (Join-Path $script:Out 'onedrive-usage-storage.csv') | Should -Be ($script:Schema.OneDriveUsageStorage -join ',')
        @(Import-Csv -LiteralPath (Join-Path $script:Out 'onedrive-usage-storage.csv')).Count | Should -Be 0
        Get-LogText $script:Out | Should -Match 'Reports.Read.All'
    }

    It 'stamps report settings and warns when names are concealed' {
        Invoke-CollectorScript 'Get-ReportSettings.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        $row = Import-Csv -LiteralPath (Join-Path $script:Out 'report-settings.csv')
        $row.RunDate | Should -Be $script:RunDate
        $row.DisplayConcealedNames | Should -Be 'True'
        Get-LogText $script:Out | Should -Match 'conceal names'
    }

    It 'retries report settings on 429 and does not record throttle as a missing permission' {
        # https://learn.microsoft.com/graph/throttling-limits
        Mock Start-Sleep { }
        $global:SpoTest.SettingsTries = 0
        Mock Get-MgAdminReportSetting {
            $global:SpoTest.SettingsTries++
            if ($global:SpoTest.SettingsTries -eq 1) { throw 'Response status code does not indicate success: 429 (Too Many Requests).' }
            [pscustomobject]@{ AdditionalProperties = @{ displayConcealedNames = $false } }
        }
        Invoke-CollectorScript 'Get-ReportSettings.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }

        (Import-Csv -LiteralPath (Join-Path $script:Out 'report-settings.csv')).DisplayConcealedNames | Should -Be 'False'
        $global:SpoTest.SettingsTries | Should -Be 2
        Get-LogText $script:Out | Should -Match 'HTTP 429'
        Get-LogText $script:Out | Should -Not -Match 'ReportSettings.Read.All'
    }

    It 'stamps the tenant storage and the site list with the run date' {
        Invoke-CollectorScript 'Get-TenantStorage.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        Invoke-CollectorScript 'Get-SpoSites.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }

        $tenant = Import-Csv -LiteralPath (Join-Path $script:Out 'tenant-storage.csv')
        $tenant.RunDate | Should -Be $script:RunDate
        $tenant.StorageQuota | Should -Be '5242880'
        $tenant.OneDriveStorageQuota | Should -Be '1048576'
        $sites = Import-Csv -LiteralPath (Join-Path $script:Out 'spo-sites.csv')
        $sites.RunDate | Should -Be $script:RunDate
        $sites.Template | Should -Be 'GROUP#0'
    }

    It 'asks Get-SPOSite for every site and the personal sites, and leaves unreturned storage empty with a warning' {
        Invoke-CollectorScript 'Get-SpoSites.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }

        Should -Invoke Get-SPOSite -Times 1 -Exactly -ParameterFilter { $Limit -eq 'ALL' -and $IncludePersonalSite -eq $true -and -not $Identity }
        (Import-Csv -LiteralPath (Join-Path $script:Out 'spo-sites.csv')).StorageUsageCurrent | Should -Be ''
        Get-LogText $script:Out | Should -Match 'without -Detailed'
    }

    It 'warns for the tenant storage properties Get-SPOTenant does not return' {
        Mock Get-SPOTenant { [pscustomobject]@{ StorageQuota = 100 } }
        Invoke-CollectorScript 'Get-TenantStorage.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        (Import-Csv -LiteralPath (Join-Path $script:Out 'tenant-storage.csv')).OneDriveStorageQuota | Should -Be ''
        Get-LogText $script:Out | Should -Match 'did not return: .*OneDriveStorageQuota'
    }

    It 'writes a header only when Get-SPOSite or Get-SPOTenant is refused' {
        Mock Get-SPOSite { throw 'Access denied' }
        Mock Get-SPOTenant { throw 'Access denied' }
        Invoke-CollectorScript 'Get-SpoSites.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        Invoke-CollectorScript 'Get-TenantStorage.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        @(Import-Csv -LiteralPath (Join-Path $script:Out 'spo-sites.csv')).Count | Should -Be 0
        @(Import-Csv -LiteralPath (Join-Path $script:Out 'tenant-storage.csv')).Count | Should -Be 0
        Get-HeaderText (Join-Path $script:Out 'spo-sites.csv') | Should -Be ($script:Schema.SpoSites -join ',')
    }
}

Describe 'Paged sources follow @odata.nextLink as returned' {
    BeforeEach {
        $script:Out = New-TestFolder
        $global:SpoTest.Requested = [System.Collections.Generic.List[string]]::new()
    }
    AfterEach { Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue }

    It 'follows the getAllSites nextLink, including the oneDrive.getAllSites path, and the drive list nextLink' {
        Mock Invoke-MgGraphRequest {
            $global:SpoTest.Requested.Add($Uri)
            switch -Wildcard ($Uri) {
                '/v1.0/sites/getAllSites' {
                    @{ value = @(New-MockSite -Id 'site-1'); '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/sites/microsoft.graph.oneDrive.getAllSites?$skiptoken=abc' }
                    break
                }
                '*oneDrive.getAllSites*' {
                    @{ value = @(New-MockSite -Id 'site-2' -Url 'https://contoso-my.sharepoint.example.com/personal/avery' -Personal $true) }
                    break
                }
                # The first page's address also contains /drives, so the skiptoken case has to win.
                '*sites/site-1/drives*skiptoken=d2*' {
                    @{ value = @(New-MockDrive -Id 'drive-1b') }
                    break
                }
                '*sites/site-1/drives*' {
                    @{ value = @(New-MockDrive -Id 'drive-1'); '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/sites/site-1/drives?$skiptoken=d2' }
                    break
                }
                '*sites/site-2/drives*' {
                    @{ value = @(New-MockDrive -Id 'drive-2' -Used 4096) }
                    break
                }
            }
        }
        Invoke-CollectorScript 'Get-DriveQuota.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:Out 'drive-quota.csv'))
        ($rows.DriveId | Sort-Object) -join ',' | Should -Be 'drive-1,drive-1b,drive-2'
        ($rows | Where-Object DriveId -EQ 'drive-2').IsPersonalSite | Should -Be 'True'
        $global:SpoTest.Requested | Should -Contain 'https://graph.microsoft.com/v1.0/sites/microsoft.graph.oneDrive.getAllSites?$skiptoken=abc'
        $global:SpoTest.Requested | Should -Not -Contain '/v1.0/sites/getAllSites?$skiptoken=abc'
        $global:SpoTest.Requested | Should -Contain 'https://graph.microsoft.com/v1.0/sites/site-1/drives?$skiptoken=d2'
        Should -Invoke Invoke-MgGraphRequest -Times 5 -Exactly
    }

    It 'selects system so drives with that facet are not hidden' {
        # https://learn.microsoft.com/graph/api/drive-list
        Mock Invoke-MgGraphRequest {
            $global:SpoTest.Requested.Add($Uri)
            if ($Uri -like '*getAllSites*') { return @{ value = @(New-MockSite -Id 'site-1') } }
            @{ value = @(New-MockDrive -Id 'drive-1' -Used 3900) }
        }
        Invoke-CollectorScript 'Get-DriveQuota.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }

        $driveCall = @($global:SpoTest.Requested | Where-Object { $_ -like '*/drives*' })
        $driveCall.Count | Should -Be 1
        $driveCall[0] | Should -Match '\$select=id,name,driveType,webUrl,quota,lastModifiedDateTime,system'
        (Import-Csv -LiteralPath (Join-Path $script:Out 'drive-quota.csv')).QuotaUsed | Should -Be '3900'
    }

    It 'retries a drive-list 429 and then writes the drive' {
        # https://learn.microsoft.com/sharepoint/dev/general-development/how-to-avoid-getting-throttled-or-blocked-in-sharepoint-online
        Mock Start-Sleep { }
        $global:SpoTest.DriveTries = 0
        Mock Invoke-MgGraphRequest {
            if ($Uri -like '*getAllSites*') { return @{ value = @(New-MockSite -Id 'site-1') } }
            $global:SpoTest.DriveTries++
            if ($global:SpoTest.DriveTries -eq 1) { throw 'Response status code does not indicate success: 429 (Too Many Requests).' }
            @{ value = @(New-MockDrive -Id 'drive-1') }
        }
        Invoke-CollectorScript 'Get-DriveQuota.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }

        (Import-Csv -LiteralPath (Join-Path $script:Out 'drive-quota.csv')).DriveId | Should -Be 'drive-1'
        $global:SpoTest.DriveTries | Should -Be 2
        Get-LogText $script:Out | Should -Match 'HTTP 429'
    }

    It 'does not write a partial snapshot when a drive list stays throttled' {
        Mock Start-Sleep { }
        Mock Invoke-MgGraphRequest {
            if ($Uri -like '*getAllSites*') { return @{ value = @((New-MockSite -Id 'site-1'), (New-MockSite -Id 'site-2')) } }
            throw 'Response status code does not indicate success: 429 (Too Many Requests).'
        }
        { Invoke-CollectorScript 'Get-DriveQuota.ps1' @{ OutputPath = $script:Out; SkipConnect = $true } } | Should -Throw '*429*'
        Test-Path -LiteralPath (Join-Path $script:Out 'drive-quota.csv') | Should -BeFalse
    }

    It 'writes the quota in bytes, its state and the drive-modified time' {
        Mock Invoke-MgGraphRequest {
            if ($Uri -like '*getAllSites*') { @{ value = @(New-MockSite -Id 'site-1') } } else { @{ value = @(New-MockDrive -Id 'drive-1' -Used 3900) } }
        }
        Invoke-CollectorScript 'Get-DriveQuota.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        $row = Import-Csv -LiteralPath (Join-Path $script:Out 'drive-quota.csv')
        $row.QuotaUsed | Should -Be '3900'
        $row.QuotaTotal | Should -Be '4096'
        $row.QuotaRemaining | Should -Be '196'
        $row.QuotaDeleted | Should -Be '10'
        $row.QuotaState | Should -Be 'normal'
        $row.LastModifiedDateTime | Should -Be '2026-10-07T14:03:11Z'
        $row.RunDate | Should -Be ([datetime]::UtcNow.ToString('yyyy-MM-dd'))
    }

    It 'refuses a nextLink that is not a Graph address, so the token is not sent there' {
        Mock Invoke-MgGraphRequest {
            $global:SpoTest.Requested.Add($Uri)
            @{ value = @(New-MockSite -Id 'site-1'); '@odata.nextLink' = 'https://evil.example.com/next' }
        }
        Invoke-CollectorScript 'Get-DriveQuota.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        $global:SpoTest.Requested | Should -Not -Contain 'https://evil.example.com/next'
        Get-LogText $script:Out | Should -Match 'Refusing to request a page outside Microsoft Graph'
    }

    It 'skips a site whose drives cannot be read and warns' {
        Mock Invoke-MgGraphRequest {
            if ($Uri -like '*getAllSites*') { @{ value = @((New-MockSite -Id 'site-1'), (New-MockSite -Id 'site-2')) } }
            elseif ($Uri -like '*site-1/drives*') { throw 'Forbidden' }
            else { @{ value = @(New-MockDrive -Id 'drive-2') } }
        }
        Invoke-CollectorScript 'Get-DriveQuota.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        (Import-Csv -LiteralPath (Join-Path $script:Out 'drive-quota.csv')).DriveId | Should -Be 'drive-2'
        Get-LogText $script:Out | Should -Match '1 site\(s\) could not have their drives read'
    }

    It 'follows the nextLink of the site activity for every site' {
        Mock Invoke-MgGraphRequest {
            $global:SpoTest.Requested.Add($Uri)
            if ($Uri -like '*getAllSites*') { return @{ value = @(New-MockSite -Id 'site-1') } }
            if ($Uri -like '*getActivitiesByInterval*') {
                return @{ value = @(@{ startDateTime = '2026-10-05T00:00:00Z'; endDateTime = '2026-10-06T00:00:00Z'; access = @{ actionCount = 5; actorCount = 3 } })
                    '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/sites/site-1/next-interval' }
            }
            @{ value = @(@{ startDateTime = '2026-10-06T00:00:00Z'; endDateTime = '2026-10-07T00:00:00Z'; edit = @{ actionCount = 1; actorCount = 1 } }) }
        }
        Invoke-CollectorScript 'Get-SiteActivity.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        @(Import-Csv -LiteralPath (Join-Path $script:Out 'site-activity.csv')).Count | Should -Be 2
        $global:SpoTest.Requested | Should -Contain 'https://graph.microsoft.com/v1.0/sites/site-1/next-interval'
    }
}

Describe 'Site activity (source 13)' {
    BeforeEach { $script:Out = New-TestFolder }
    AfterEach { Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue }

    It 'asks for daily intervals over fewer than 90 days and writes absent actions as empty, not zero' {
        $global:SpoTest.Requested = [System.Collections.Generic.List[string]]::new()
        Mock Invoke-MgGraphRequest {
            $global:SpoTest.Requested.Add($Uri)
            if ($Uri -like '*getAllSites*') { return @{ value = @(New-MockSite -Id 'site-1') } }
            @{ value = @(
                    @{ startDateTime = '2026-10-05T00:00:00Z'; endDateTime = '2026-10-06T00:00:00Z'; access = @{ actionCount = 5; actorCount = 3 }; edit = @{ actionCount = 2; actorCount = 1 } }
                    @{ startDateTime = '2026-10-06T00:00:00Z'; endDateTime = '2026-10-07T00:00:00Z'; incompleteData = @{ missingDataBeforeDateTime = '2026-10-06T06:00:00Z'; wasThrottled = $false } }
                ) }
        }
        Invoke-CollectorScript 'Get-SiteActivity.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; LookbackDays = 30 }

        $call = $global:SpoTest.Requested | Where-Object { $_ -like '*getActivitiesByInterval*' }
        $call | Should -Match "interval='day'"
        $matches = [regex]::Match($call, "startDateTime='([\d-]+)',endDateTime='([\d-]+)'")
        ([datetime]$matches.Groups[2].Value - [datetime]$matches.Groups[1].Value).TotalDays | Should -Be 30

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:Out 'site-activity.csv'))
        $rows[0].AccessActionCount | Should -Be '5'
        $rows[0].EditActorCount | Should -Be '1'
        $rows[0].CreateActionCount | Should -Be ''
        $rows[0].IncompleteData | Should -Be 'False'
        $rows[1].IncompleteData | Should -Be 'True'
        $rows[1].MissingDataBeforeDateTime | Should -Be '2026-10-06T06:00:00Z'
        $rows[1].WasThrottled | Should -Be 'False'
        Get-LogText $script:Out | Should -Match 'incomplete data'
    }

    It 'does not accept a lookback of 90 days or more' {
        { Invoke-CollectorScript 'Get-SiteActivity.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; LookbackDays = 90 } } | Should -Throw
    }

    It 'warns and skips the sites that return nothing, as in a cloud where itemAnalytics is not available' {
        Mock Invoke-MgGraphRequest {
            if ($Uri -like '*getAllSites*') { return @{ value = @(New-MockSite -Id 'site-1') } }
            throw 'NotImplemented'
        }
        Invoke-CollectorScript 'Get-SiteActivity.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; Environment = 'GCCHigh' }
        @(Import-Csv -LiteralPath (Join-Path $script:Out 'site-activity.csv')).Count | Should -Be 0
        Get-LogText $script:Out | Should -Match 'itemAnalytics is not yet available'
    }

    It 'does not write a partial snapshot when site activity stays throttled' {
        # https://learn.microsoft.com/sharepoint/dev/general-development/how-to-avoid-getting-throttled-or-blocked-in-sharepoint-online
        Mock Start-Sleep { }
        Mock Invoke-MgGraphRequest {
            if ($Uri -like '*getAllSites*') { return @{ value = @((New-MockSite -Id 'site-1'), (New-MockSite -Id 'site-2')) } }
            throw 'Response status code does not indicate success: 429 (Too Many Requests).'
        }
        { Invoke-CollectorScript 'Get-SiteActivity.ps1' @{ OutputPath = $script:Out; SkipConnect = $true } } | Should -Throw '*429*'
        Test-Path -LiteralPath (Join-Path $script:Out 'site-activity.csv') | Should -BeFalse
    }
}

Describe 'File events (source 12) resume from the last day collected' {
    BeforeEach {
        $script:Out = New-TestFolder
        $global:SpoTest.Windows = [System.Collections.Generic.List[object]]::new()
        Mock Start-Sleep { }
    }
    AfterEach { Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue }

    It 'counts records per UTC day, workload, user and operation' {
        Mock Search-UnifiedAuditLog {
            if ($StartDate.ToUniversalTime().Date -eq [datetime]'2026-10-05') {
                New-MockAuditRecord -Id 'a' -Operation 'FileAccessed' -CreationTime '2026-10-05T01:00:00'
                New-MockAuditRecord -Id 'b' -Operation 'FileAccessed' -CreationTime '2026-10-05T23:59:00'
                New-MockAuditRecord -Id 'c' -Operation 'FileModified' -CreationTime '2026-10-05T10:00:00'
                New-MockAuditRecord -Id 'd' -Operation 'FileSyncUploadedFull' -CreationTime '2026-10-05T11:00:00' -Workload 'OneDrive' -UserId 'devon.diaz@example.com'
            }
        }
        Invoke-CollectorScript 'Get-FileEvents.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; StartDate = [datetime]'2026-10-05T00:00:00Z'; EndDate = [datetime]'2026-10-06T00:00:00Z' }

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:Out 'file-events.csv'))
        $rows.Count | Should -Be 3
        ($rows | Where-Object Operation -EQ 'FileAccessed').EventCount | Should -Be '2'
        ($rows | Where-Object Operation -EQ 'FileModified').EventCount | Should -Be '1'
        ($rows | Where-Object Operation -EQ 'FileSyncUploadedFull').Workload | Should -Be 'OneDrive'
        ($rows.Date | Sort-Object -Unique) | Should -Be '2026-10-05'
        Should -Invoke Search-UnifiedAuditLog -ParameterFilter { $Operations -contains 'PageViewed' -and $Operations -contains 'SharingSet' -and $Operations.Count -eq 10 }
    }

    It 'starts the day after the latest Date already in the file and stops before today' {
        Mock Search-UnifiedAuditLog {
            $global:SpoTest.Windows.Add([pscustomobject]@{ Start = $StartDate.ToUniversalTime(); End = $EndDate.ToUniversalTime() })
            New-MockAuditRecord -Id ('e-' + $global:SpoTest.Windows.Count) -CreationTime ($StartDate.ToUniversalTime().AddHours(1).ToString('s'))
        }
        $yesterday = [datetime]::UtcNow.Date.AddDays(-1)
        $twoDaysAgo = [datetime]::UtcNow.Date.AddDays(-2)
        Export-AppendCsv -Path (Join-Path $script:Out 'file-events.csv') -Column $script:Schema.FileEvents -Rows @(
            [pscustomobject]@{ Date = $twoDaysAgo.ToString('yyyy-MM-dd'); Workload = 'SharePoint'; UserId = 'avery.abara@example.com'; Operation = 'FileAccessed'; EventCount = 7 })

        Invoke-CollectorScript 'Get-FileEvents.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }

        $global:SpoTest.Windows.Count | Should -Be 1
        $global:SpoTest.Windows[0].Start | Should -Be $yesterday
        $global:SpoTest.Windows[0].End | Should -Be ([datetime]::UtcNow.Date)
        @(Import-Csv -LiteralPath (Join-Path $script:Out 'file-events.csv')).Count | Should -Be 2
    }

    It 'reaches back -LookbackDays on the first run and splits the range into -WindowHours windows' {
        Mock Search-UnifiedAuditLog { $global:SpoTest.Windows.Add([pscustomobject]@{ Start = $StartDate.ToUniversalTime() }) }
        Invoke-CollectorScript 'Get-FileEvents.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; LookbackDays = 2; WindowHours = 12 }
        # An empty first page is retried, so count the distinct window starts.
        $starts = @($global:SpoTest.Windows.Start | Sort-Object -Unique)
        $starts.Count | Should -Be 4
        $starts[0] | Should -Be ([datetime]::UtcNow.Date.AddDays(-2))
        $starts[1] | Should -Be ([datetime]::UtcNow.Date.AddDays(-2).AddHours(12))
    }

    It 'writes nothing new and says so when the last whole day is already collected' {
        Export-AppendCsv -Path (Join-Path $script:Out 'file-events.csv') -Column $script:Schema.FileEvents -Rows @(
            [pscustomobject]@{ Date = [datetime]::UtcNow.Date.AddDays(-1).ToString('yyyy-MM-dd'); Workload = 'SharePoint'; UserId = 'a@example.com'; Operation = 'FileAccessed'; EventCount = 1 })
        Mock Search-UnifiedAuditLog { throw 'should not be called' }
        Invoke-CollectorScript 'Get-FileEvents.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        Should -Invoke Search-UnifiedAuditLog -Times 0 -Exactly
        Get-LogText $script:Out | Should -Match 'Nothing new to collect'
    }

    It 'reads the next page in one ReturnLargeSet session while moreRecordsAvailable is true' {
        $global:SpoTest.Sessions = [System.Collections.Generic.List[string]]::new()
        Mock Search-UnifiedAuditLog {
            $global:SpoTest.Sessions.Add("$SessionId|$SessionCommand|$ResultSize")
            if ($global:SpoTest.Sessions.Count -eq 1) { New-MockAuditRecord -Id 'p1' -MoreRecords $true }
            elseif ($global:SpoTest.Sessions.Count -eq 2) { New-MockAuditRecord -Id 'p2' -MoreRecords $false }
        }
        Invoke-CollectorScript 'Get-FileEvents.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; StartDate = [datetime]'2026-10-05T00:00:00Z'; EndDate = [datetime]'2026-10-06T00:00:00Z' }

        $global:SpoTest.Sessions.Count | Should -Be 2
        ($global:SpoTest.Sessions | ForEach-Object { ($_ -split '\|')[0] } | Sort-Object -Unique).Count | Should -Be 1
        $global:SpoTest.Sessions[0] | Should -Match '\|ReturnLargeSet\|5000$'
        (Import-Csv -LiteralPath (Join-Path $script:Out 'file-events.csv')).EventCount | Should -Be '2'
    }

    It 'writes no day from a window that reaches the 50,000-record cap, logs an error and stops' {
        Mock Search-UnifiedAuditLog {
            if ($StartDate.ToUniversalTime().Date -eq [datetime]'2026-10-05') { New-MockAuditRecord -Id 'ok' -CreationTime '2026-10-05T05:00:00' }
            else {
                $record = New-MockAuditRecord -Id 'big' -CreationTime '2026-10-06T05:00:00'
                $record | Add-Member -NotePropertyName ResultCount -NotePropertyValue 50000
                $record
            }
        }
        { Invoke-CollectorScript 'Get-FileEvents.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; StartDate = [datetime]'2026-10-05T00:00:00Z'; EndDate = [datetime]'2026-10-07T00:00:00Z' } } | Should -Throw '*50,000-record session cap*'

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:Out 'file-events.csv'))
        ($rows.Date | Sort-Object -Unique) | Should -Be '2026-10-05'
        Get-LogText $script:Out | Should -Match 'matches 50,000 or more records'
    }

    It 'writes a header only and logs the reason when the audit log is refused' {
        Mock Search-UnifiedAuditLog { throw 'The role is missing' }
        Invoke-CollectorScript 'Get-FileEvents.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        Get-HeaderText (Join-Path $script:Out 'file-events.csv') | Should -Be ($script:Schema.FileEvents -join ',')
        @(Import-Csv -LiteralPath (Join-Path $script:Out 'file-events.csv')).Count | Should -Be 0
        Get-LogText $script:Out | Should -Match 'View-Only Audit Logs'
    }

    It 'rejects an empty explicit range' {
        { Invoke-CollectorScript 'Get-FileEvents.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; StartDate = [datetime]'2026-10-06T00:00:00Z'; EndDate = [datetime]'2026-10-05T00:00:00Z' } } | Should -Throw '*requested range is empty*'
    }

    It 'accepts a 365-day lookback so the E5 year is not cut off' {
        # https://learn.microsoft.com/purview/audit-log-retention-policies
        Mock Search-UnifiedAuditLog { $global:SpoTest.Windows.Add($StartDate.ToUniversalTime()); throw 'stop' }
        Invoke-CollectorScript 'Get-FileEvents.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; LookbackDays = 365; WindowHours = 24 }
        $global:SpoTest.Windows[0].Date | Should -Be ([datetime]::UtcNow.Date.AddDays(-365))
        { Invoke-CollectorScript 'Get-FileEvents.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; LookbackDays = 366 } } | Should -Throw
    }
}

Describe 'Run-All' {
    BeforeEach {
        $script:Out = New-TestFolder
        Mock Connect-MgGraph -ModuleName M365ReportLibrary -MockWith { }
        Mock Connect-ExchangeOnline -ModuleName M365ReportLibrary -MockWith { }
        Mock Connect-SPOService { }
        Mock Disconnect-SPOService { }
        Mock Disconnect-ExchangeOnline { }
        Mock Get-MgReportSharePointSiteUsageDetail { }
        Mock Get-MgReportOneDriveUsageAccountDetail { }
        Mock Get-MgReportSharePointSiteUsageStorage { }
        Mock Get-MgReportOneDriveUsageStorage { }
        Mock Get-MgReportSharePointActivityUserDetail { }
        Mock Get-MgReportOneDriveActivityUserDetail { }
        Mock Get-MgAdminReportSetting { [pscustomobject]@{ DisplayConcealedNames = $false } }
        Mock Get-SPOTenant { [pscustomobject]@{ StorageQuota = 1 } }
        Mock Get-SPOSite { [pscustomobject]@{ Url = 'https://contoso.sharepoint.example.com'; Title = 'Root'; Template = 'SITEPAGEPUBLISHING#0' } }
        Mock Search-UnifiedAuditLog { }
        Mock Invoke-MgGraphRequest { @{ value = @() } }
        Mock Start-Sleep { }
    }
    AfterEach { Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue }

    It 'leaves all twelve CSVs, signs in once to each service and signs out' {
        & (Join-Path $script:Collectors 'Run-All.ps1') -OutputPath $script:Out -AdminUrl 'https://contoso-admin.sharepoint.example.com'

        @(Get-ChildItem -LiteralPath $script:Out -Filter '*.csv').Count | Should -Be 12
        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 1 -Exactly
        Should -Invoke Connect-ExchangeOnline -ModuleName M365ReportLibrary -Times 1 -Exactly
        Should -Invoke Connect-SPOService -Times 1 -Exactly
        Should -Invoke Disconnect-SPOService -Times 1 -Exactly
        Should -Invoke Disconnect-ExchangeOnline -Times 1 -Exactly
    }

    It 'writes header-only files for the six Graph usage reports in GCC High and still runs the others' {
        & (Join-Path $script:Collectors 'Run-All.ps1') -OutputPath $script:Out -Environment GCCHigh -AdminUrl 'https://contoso-admin.sharepoint.us'

        foreach ($name in 'sharepoint-site-usage-detail', 'onedrive-usage-account-detail', 'sharepoint-site-usage-storage', 'onedrive-usage-storage', 'sharepoint-activity-user-detail', 'onedrive-activity-user-detail') {
            Test-Path -LiteralPath (Join-Path $script:Out "$name.csv") | Should -BeTrue
            @(Import-Csv -LiteralPath (Join-Path $script:Out "$name.csv")).Count | Should -Be 0
        }
        Should -Invoke Get-MgReportSharePointSiteUsageDetail -Times 0 -Exactly
        Should -Invoke Connect-SPOService -Times 1 -Exactly -ParameterFilter { $Region -eq 'ITAR' }
        Should -Invoke Get-SPOSite -Times 1 -Exactly
        Should -Invoke Search-UnifiedAuditLog -Scope It -Times 1
    }

    It 'writes the two SharePoint Online headers and a warning when no -AdminUrl is given' {
        & (Join-Path $script:Collectors 'Run-All.ps1') -OutputPath $script:Out

        Should -Invoke Connect-SPOService -Times 0 -Exactly
        Should -Invoke Get-SPOSite -Times 0 -Exactly
        Get-HeaderText (Join-Path $script:Out 'spo-sites.csv') | Should -Be ($script:Schema.SpoSites -join ',')
        Get-HeaderText (Join-Path $script:Out 'tenant-storage.csv') | Should -Be ($script:Schema.TenantStorage -join ',')
        Get-LogText $script:Out | Should -Match 'No -AdminUrl was given'
    }

    It 'keeps running the other collectors when one stops' {
        Mock Get-MgAdminReportSetting { throw 'Forbidden' }
        & (Join-Path $script:Collectors 'Run-All.ps1') -OutputPath $script:Out -AdminUrl 'https://contoso-admin.sharepoint.example.com'
        @(Get-ChildItem -LiteralPath $script:Out -Filter '*.csv').Count | Should -Be 12
    }

    It 'passes a 365-day lookback through to the audit search' {
        $global:SpoTest.Windows = [System.Collections.Generic.List[datetime]]::new()
        Mock Search-UnifiedAuditLog { $global:SpoTest.Windows.Add($StartDate.ToUniversalTime()); throw 'stop' }
        & (Join-Path $script:Collectors 'Run-All.ps1') -OutputPath $script:Out -LookbackDays 365
        $global:SpoTest.Windows[0].Date | Should -Be ([datetime]::UtcNow.Date.AddDays(-365))
        { & (Join-Path $script:Collectors 'Run-All.ps1') -OutputPath $script:Out -LookbackDays 366 } | Should -Throw
    }
}
