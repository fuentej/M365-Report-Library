#Requires -Version 7.0

BeforeAll {
    $global:TaTest = @{}
    $script:Root = Join-Path $PSScriptRoot '../../..' | Resolve-Path | Select-Object -ExpandProperty Path
    $script:Collectors = Join-Path $script:Root 'reports/teams-activity/collectors'

    . (Join-Path $script:Root 'shared/tests/TenantCmdletStubs.ps1')
    . (Join-Path $PSScriptRoot 'TeamsStubs.ps1')
    Import-Module (Join-Path $script:Root 'shared/M365ReportLibrary.psm1') -Force
    . (Join-Path $script:Collectors 'TeamsActivityHelpers.ps1')

    $script:Schema = Import-PowerShellDataFile -LiteralPath (Join-Path $script:Collectors 'TeamsActivitySchema.psd1')

    function New-TestFolder {
        $path = Join-Path ([System.IO.Path]::GetTempPath()) ('teams-' + [guid]::NewGuid().ToString('N'))
        New-Item -Path $path -ItemType Directory -Force | Out-Null
        return $path
    }

    function Get-HeaderText { param([string]$Path) return ((Get-CsvHeaderColumn -Path $Path) -join ',') }

    function Get-LogText { param([string]$Folder) return (Get-Content -LiteralPath (Join-Path $Folder 'run.log') -Raw) }

    function Invoke-CollectorScript {
        param([Parameter(Mandatory)][string]$Name, [hashtable]$Arguments = @{})
        & (Join-Path $script:Collectors $Name) @Arguments
    }

    # Header lists, in the order the API pages give them.
    # https://learn.microsoft.com/graph/api/reportroot-getteamsuseractivityuserdetail
    # https://learn.microsoft.com/graph/api/reportroot-getteamsuseractivitycounts
    # https://learn.microsoft.com/graph/api/reportroot-getteamsdeviceusageuserdetail
    $global:TaTest.UserHeader = 'Report Refresh Date,Tenant Display Name,Shared Channel Tenant Display Names,User Id,User Principal Name,Last Activity Date,Is Deleted,Deleted Date,Assigned Products,Team Chat Message Count,Private Chat Message Count,Call Count,Meeting Count,Post Messages,Reply Messages,Urgent Messages,Meetings Organized Count,Meetings Attended Count,Ad Hoc Meetings Organized Count,Ad Hoc Meetings Attended Count,Scheduled One-time Meetings Organized Count,Scheduled One-time Meetings Attended Count,Scheduled Recurring Meetings Organized Count,Scheduled Recurring Meetings Attended Count,Audio Duration,Video Duration,Screen Share Duration,Audio Duration In Seconds,Video Duration In Seconds,Screen Share Duration In Seconds,Has Other Action,Is Licensed,Report Period'
    $global:TaTest.CountsHeader = 'Report Refresh Date,Report Date,Team Chat Messages,Post Messages,Reply Messages,Private Chat Messages,Calls,Meetings,Audio Duration,Video Duration,Screen Share Duration,Meetings Organized,Meetings Attended,Report Period'
    $global:TaTest.DeviceHeader = 'Report Refresh Date,User Id,User Principal Name,Last Activity Date,Is Deleted,Deleted Date,Used Web,Used Windows Phone,Used iOS,Used Mac,Used Android Phone,Used Windows,Used Chrome OS,Used Linux,Is Licensed,Report Period'

    function New-UserReportCsv {
        param([string[]]$Upn = @('avery.abara@example.com'))
        $lines = @($global:TaTest.UserHeader)
        foreach ($u in $Upn) {
            $lines += "2026-10-06,Contoso,,id-$u,$u,2026-10-05,False,,MICROSOFT 365 E3,12,34,5,6,7,8,1,2,3,0,1,1,1,1,1,PT1H,PT20M,PT5M,3600,1200,300,Yes,Yes,30"
        }
        $lines -join "`n"
    }

    # A call record from GET /communications/callRecords/{id}?$expand=sessions.
    # https://learn.microsoft.com/graph/api/callrecords-callrecord-get
    function New-MockCallRecord {
        param([string]$Id, [string]$Start = '2026-10-07T14:00:00Z', [int]$Version = 1, [object[]]$Session = @(), [string]$SessionsNext = '')
        $record = @{
            id = $Id; version = $Version; type = 'groupCall'; modalities = @('audio', 'video')
            startDateTime = $Start; endDateTime = '2026-10-07T14:45:00Z'; lastModifiedDateTime = '2026-10-07T14:47:00Z'
            sessions = $Session
        }
        if ($SessionsNext) { $record['sessions@odata.nextLink'] = $SessionsNext }
        $record
    }

    function New-MockSession {
        param([string]$Id, [string]$Caller = 'user-1', [string]$CallerPlatform = 'windows', [string]$Callee = '', [string]$CalleePlatform = '')
        $session = @{
            id = $Id; startDateTime = '2026-10-07T14:00:00Z'; endDateTime = '2026-10-07T14:45:00Z'
            caller = @{ associatedIdentity = @{ id = $Caller }; userAgent = @{ platform = $CallerPlatform } }
        }
        if ($Callee) { $session['callee'] = @{ associatedIdentity = @{ id = $Callee }; userAgent = @{ platform = $CalleePlatform } } }
        $session
    }

    # Search-UnifiedAuditLog record: RecordType, ResultCount, AuditData as a JSON string.
    # https://learn.microsoft.com/powershell/module/exchangepowershell/search-unifiedauditlog
    function New-MockAuditRecord {
        param(
            [string]$Id = 'event-1',
            [string]$Operation = 'MeetingDetail',
            [string]$CreationTime = '2026-10-05T12:00:00',
            [string]$UserId = 'avery.abara@example.com',
            [string]$Workload = 'MicrosoftTeams',
            [object]$MoreRecords = $null
        )
        $auditData = [ordered]@{ CreationTime = $CreationTime; Id = $Id; Operation = $Operation; UserId = $UserId; Workload = $Workload } | ConvertTo-Json -Compress
        $record = [ordered]@{ RecordType = 'MicrosoftTeams'; AuditData = $auditData }
        if ($null -ne $MoreRecords) { $record['AuditSearchRequestMetadata'] = [pscustomobject]@{ moreRecordsAvailable = $MoreRecords } }
        return [pscustomobject]$record
    }
}

AfterAll {
    Remove-Variable -Name TaTest -Scope Global -ErrorAction SilentlyContinue
    Remove-Module M365ReportLibrary -Force -ErrorAction SilentlyContinue
}

Describe 'Cloud availability' {
    BeforeEach {
        $script:Out = New-TestFolder
        Mock Connect-MgGraph -ModuleName M365ReportLibrary -MockWith { }
        Mock Get-MgReportTeamUserActivityUserDetail { }
        Mock Get-MgReportTeamUserActivityCount { }
        Mock Get-MgReportTeamDeviceUsageUserDetail { }
        Mock Get-MgAdminReportSetting { [pscustomobject]@{ DisplayConcealedNames = $false } }
        Mock Invoke-MgGraphRequest { @{ value = @() } }
        Mock Start-Sleep { }
    }
    AfterEach { Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue }

    It 'writes a header only for <Script> in GCC High and never calls the report' -ForEach @(
        @{ Script = 'Get-TeamsUserActivityUserDetail.ps1'; Csv = 'teams-user-activity-user-detail.csv'; Key = 'TeamsUserActivityUserDetail'; Cmdlet = 'Get-MgReportTeamUserActivityUserDetail' }
        @{ Script = 'Get-TeamsUserActivityCounts.ps1'; Csv = 'teams-user-activity-counts.csv'; Key = 'TeamsUserActivityCounts'; Cmdlet = 'Get-MgReportTeamUserActivityCount' }
        @{ Script = 'Get-TeamsDeviceUsageUserDetail.ps1'; Csv = 'teams-device-usage-user-detail.csv'; Key = 'TeamsDeviceUsageUserDetail'; Cmdlet = 'Get-MgReportTeamDeviceUsageUserDetail' }
    ) {
        Invoke-CollectorScript $Script @{ OutputPath = $script:Out; Environment = 'GCCHigh' }

        $path = Join-Path $script:Out $Csv
        Get-HeaderText $path | Should -Be ($script:Schema[$Key] -join ',')
        @(Import-Csv -LiteralPath $path).Count | Should -Be 0
        Get-LogText $script:Out | Should -Match 'documented as unavailable in GCCHigh'
        Should -Invoke $Cmdlet -Times 0 -Exactly
        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 0 -Exactly
    }

    It 'attempts <Script> in GCC and logs a warning that availability is UNVERIFIED' -ForEach @(
        @{ Script = 'Get-TeamsUserActivityUserDetail.ps1'; Cmdlet = 'Get-MgReportTeamUserActivityUserDetail' }
        @{ Script = 'Get-TeamsUserActivityCounts.ps1'; Cmdlet = 'Get-MgReportTeamUserActivityCount' }
        @{ Script = 'Get-TeamsDeviceUsageUserDetail.ps1'; Cmdlet = 'Get-MgReportTeamDeviceUsageUserDetail' }
    ) {
        Invoke-CollectorScript $Script @{ OutputPath = $script:Out; Environment = 'GCC'; SkipConnect = $true }
        Get-LogText $script:Out | Should -Match 'UNVERIFIED'
        Should -Invoke $Cmdlet -Times 1 -Exactly
    }

    It 'attempts the report settings in GCC and GCC High with a warning' -ForEach @(@{ Cloud = 'GCC' }, @{ Cloud = 'GCCHigh' }) {
        Invoke-CollectorScript 'Get-ReportSettings.ps1' @{ OutputPath = $script:Out; Environment = $Cloud; SkipConnect = $true }
        Should -Invoke Get-MgAdminReportSetting -Times 1 -Exactly
        Get-LogText $script:Out | Should -Match 'UNVERIFIED'
    }

    It 'reads call records in GCC High without a warning, since Learn marks them available there' {
        Invoke-CollectorScript 'Get-CallRecords.ps1' @{ OutputPath = $script:Out; Environment = 'GCCHigh'; SkipConnect = $true }
        Should -Invoke Invoke-MgGraphRequest -Times 1 -Exactly
        Get-LogText $script:Out | Should -Not -Match 'UNVERIFIED'
    }
}

Describe 'Connection endpoint per -Environment' {
    BeforeEach {
        $script:Out = New-TestFolder
        Mock Connect-MgGraph -ModuleName M365ReportLibrary -MockWith { }
        Mock Connect-ExchangeOnline -ModuleName M365ReportLibrary -MockWith { }
        Mock Disconnect-ExchangeOnline { }
        Mock Get-MgReportTeamUserActivityCount { Set-Content -LiteralPath $OutFile -Value $global:TaTest.CountsHeader }
        Mock Invoke-MgGraphRequest { @{ value = @() } }
        Mock Search-UnifiedAuditLog { }
        Mock Start-Sleep { }
    }
    AfterEach { Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue }

    It 'signs in to Graph with the <Graph> environment for <Cloud>' -ForEach @(
        @{ Cloud = 'Commercial'; Graph = 'Global' }
        @{ Cloud = 'GCC'; Graph = 'Global' }
    ) {
        Invoke-CollectorScript 'Get-TeamsUserActivityCounts.ps1' @{ OutputPath = $script:Out; Environment = $Cloud }
        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter { $Environment -eq $Graph }
    }

    It 'signs in to Graph with the USGov environment for GCC High, without the application-only call-records scope' {
        Invoke-CollectorScript 'Get-CallRecords.ps1' @{ OutputPath = $script:Out; Environment = 'GCCHigh' }
        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter {
            $Environment -eq 'USGov' -and $Scopes -notcontains 'CallRecords.Read.All'
        }
    }

    It 'signs in to Exchange Online with <Name> for <Cloud>' -ForEach @(
        @{ Cloud = 'Commercial'; Name = 'O365Default' }
        @{ Cloud = 'GCC'; Name = 'O365Default' }
        @{ Cloud = 'GCCHigh'; Name = 'O365USGovGCCHigh' }
    ) {
        Invoke-CollectorScript 'Get-TeamsAuditEvents.ps1' @{ OutputPath = $script:Out; Environment = $Cloud }
        Should -Invoke Connect-ExchangeOnline -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter { $ExchangeEnvironmentName -eq $Name }
        Should -Invoke Disconnect-ExchangeOnline -Times 1 -Exactly
    }

    It 'does not sign in when asked to reuse a session' {
        Invoke-CollectorScript 'Get-TeamsAuditEvents.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        Invoke-CollectorScript 'Get-CallRecords.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        Should -Invoke Connect-ExchangeOnline -ModuleName M365ReportLibrary -Times 0 -Exactly
        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 0 -Exactly
        Should -Invoke Disconnect-ExchangeOnline -Times 0 -Exactly
    }
}

Describe 'Graph usage reports (sources 1 to 3) are state snapshots stamped with the run date' {
    BeforeEach {
        $script:Out = New-TestFolder
        $script:RunDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')
        $global:TaTest.Files = [System.Collections.Generic.List[string]]::new()
    }
    AfterEach { Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue }

    It 'stamps user activity with the run date, keeps the count columns, and removes the download' {
        Mock Get-MgReportTeamUserActivityUserDetail {
            $global:TaTest.Files.Add($OutFile)
            Set-Content -LiteralPath $OutFile -Encoding utf8 -Value (New-UserReportCsv -Upn 'avery.abara@example.com', 'blake.bishop@example.com')
        }
        Invoke-CollectorScript 'Get-TeamsUserActivityUserDetail.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; Period = 'D7' }

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:Out 'teams-user-activity-user-detail.csv'))
        $rows.Count | Should -Be 2
        ($rows.RunDate | Sort-Object -Unique) | Should -Be $script:RunDate
        $rows[0].QueryDate | Should -Be ''
        $rows[0].TeamChatMessageCount | Should -Be '12'
        $rows[0].PrivateChatMessageCount | Should -Be '34'
        $rows[0].CallCount | Should -Be '5'
        $rows[0].AudioDurationInSeconds | Should -Be '3600'
        $rows[0].IsLicensed | Should -Be 'Yes'
        $rows[0].ReportPeriod | Should -Be '30'
        Should -Invoke Get-MgReportTeamUserActivityUserDetail -Times 1 -Exactly -ParameterFilter { $Period -eq 'D7' -and -not $PSBoundParameters.ContainsKey('Date') }
        foreach ($file in $global:TaTest.Files) { Test-Path -LiteralPath $file | Should -BeFalse }
    }

    It 'sends the date alone when one is given, as UTC midnight of that calendar day' -ForEach @(
        @{ Script = 'Get-TeamsUserActivityUserDetail.ps1'; Cmdlet = 'Get-MgReportTeamUserActivityUserDetail'; Header = 'UserHeader' }
        @{ Script = 'Get-TeamsDeviceUsageUserDetail.ps1'; Cmdlet = 'Get-MgReportTeamDeviceUsageUserDetail'; Header = 'DeviceHeader' }
    ) {
        $script:ActivityDay = [datetime]::UtcNow.Date.AddDays(-2)
        Mock $Cmdlet -MockWith ([scriptblock]::Create("Set-Content -LiteralPath `$OutFile -Value `$global:TaTest.$Header"))
        Invoke-CollectorScript $Script @{ OutputPath = $script:Out; SkipConnect = $true; Date = $script:ActivityDay }
        Should -Invoke $Cmdlet -Times 1 -Exactly -ParameterFilter {
            $Date.Kind -eq [DateTimeKind]::Utc -and $Date -eq $script:ActivityDay -and $Date.TimeOfDay -eq [timespan]::Zero -and [string]::IsNullOrEmpty($Period)
        }
    }

    It 'keeps an Unspecified activity time on that calendar day' {
        $yesterday = [datetime]::UtcNow.Date.AddDays(-1)
        $script:ActivityDay = [datetime]::new($yesterday.Year, $yesterday.Month, $yesterday.Day, 23, 30, 0, [DateTimeKind]::Unspecified)
        Mock Get-MgReportTeamUserActivityUserDetail { Set-Content -LiteralPath $OutFile -Value $global:TaTest.UserHeader }
        Invoke-CollectorScript 'Get-TeamsUserActivityUserDetail.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; Date = $script:ActivityDay }
        Should -Invoke Get-MgReportTeamUserActivityUserDetail -Times 1 -Exactly -ParameterFilter {
            $Date.Kind -eq [DateTimeKind]::Utc -and $Date.Year -eq $script:ActivityDay.Year -and $Date.Month -eq $script:ActivityDay.Month -and $Date.Day -eq $script:ActivityDay.Day -and $Date.Hour -eq 0
        }
    }

    It 'rejects a user-activity date outside the past 30 days instead of calling the report' {
        $tooOld = [datetime]::UtcNow.Date.AddDays(-31)
        Mock Get-MgReportTeamUserActivityUserDetail { throw 'should not be called' }
        { Invoke-CollectorScript 'Get-TeamsUserActivityUserDetail.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; Date = $tooOld } } | Should -Throw '*past 30 days*'
        Should -Invoke Get-MgReportTeamUserActivityUserDetail -Times 0 -Exactly
    }

    It 'rejects a device date outside the past 28 days, including a day user activity still accepts' {
        $day = [datetime]::UtcNow.Date.AddDays(-29)
        Mock Get-MgReportTeamDeviceUsageUserDetail { throw 'should not be called' }
        { Invoke-CollectorScript 'Get-TeamsDeviceUsageUserDetail.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; Date = $day } } | Should -Throw '*past 28 days*'
        Should -Invoke Get-MgReportTeamDeviceUsageUserDetail -Times 0 -Exactly

        Mock Get-MgReportTeamUserActivityUserDetail { Set-Content -LiteralPath $OutFile -Value $global:TaTest.UserHeader }
        Invoke-CollectorScript 'Get-TeamsUserActivityUserDetail.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; Date = $day }
        Should -Invoke Get-MgReportTeamUserActivityUserDetail -Times 1 -Exactly
    }

    It 'reads the device usage columns as yes/no per platform' {
        Mock Get-MgReportTeamDeviceUsageUserDetail {
            Set-Content -LiteralPath $OutFile -Encoding utf8 -Value ($global:TaTest.DeviceHeader + "`n2026-10-06,id-1,avery.abara@example.com,2026-10-05,False,,Yes,No,Yes,No,No,Yes,No,No,Yes,30")
        }
        Invoke-CollectorScript 'Get-TeamsDeviceUsageUserDetail.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }

        $row = @(Import-Csv -LiteralPath (Join-Path $script:Out 'teams-device-usage-user-detail.csv'))[0]
        $row.RunDate | Should -Be $script:RunDate
        $row.UsedWeb | Should -Be 'Yes'
        $row.UsediOS | Should -Be 'Yes'
        $row.UsedWindows | Should -Be 'Yes'
        $row.UsedMac | Should -Be 'No'
        Should -Invoke Get-MgReportTeamDeviceUsageUserDetail -Times 1 -Exactly -ParameterFilter { $Period -eq 'D30' }
    }

    It 'writes one row per report date for tenant totals' {
        Mock Get-MgReportTeamUserActivityCount {
            Set-Content -LiteralPath $OutFile -Encoding utf8 -Value ($global:TaTest.CountsHeader + "`n2026-10-06,2026-10-05,200,120,80,500,30,70,PT40H,PT12H,PT3H,25,70,30`n2026-10-06,2026-10-04,190,110,80,480,28,65,PT38H,PT11H,PT3H,22,65,30")
        }
        Invoke-CollectorScript 'Get-TeamsUserActivityCounts.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:Out 'teams-user-activity-counts.csv'))
        $rows.Count | Should -Be 2
        $rows[0].ReportDate | Should -Be '2026-10-05'
        $rows[0].TeamChatMessages | Should -Be '200'
        $rows[0].RunDate | Should -Be $script:RunDate
        Should -Invoke Get-MgReportTeamUserActivityCount -Times 1 -Exactly -ParameterFilter { $Period -eq 'D30' }
    }

    It 'does not append a second copy of the same snapshot, but a new run date appends again' {
        Mock Get-MgReportTeamUserActivityUserDetail { Set-Content -LiteralPath $OutFile -Value (New-UserReportCsv) -Encoding utf8 }
        Invoke-CollectorScript 'Get-TeamsUserActivityUserDetail.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        Invoke-CollectorScript 'Get-TeamsUserActivityUserDetail.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        $path = Join-Path $script:Out 'teams-user-activity-user-detail.csv'
        @(Import-Csv -LiteralPath $path).Count | Should -Be 1

        Export-AppendCsv -Path $path -Column $script:Schema.TeamsUserActivityUserDetail -Rows @(
            [pscustomobject]@{ RunDate = '2026-09-01'; QueryDate = ''; UserPrincipalName = 'avery.abara@example.com'; ReportPeriod = '30' }) -KeyColumn @('RunDate', 'QueryDate', 'UserPrincipalName', 'ReportPeriod')
        @(Import-Csv -LiteralPath $path).Count | Should -Be 2
    }

    It 'retries a usage report 429 and does not record it as a refused report' {
        $global:TaTest.UsageTries = 0
        Mock Get-MgReportTeamUserActivityUserDetail {
            $global:TaTest.UsageTries++
            if ($global:TaTest.UsageTries -eq 1) { throw 'Response status code does not indicate success: 429 (Too Many Requests).' }
            Set-Content -LiteralPath $OutFile -Value (New-UserReportCsv) -Encoding utf8
        }
        Mock Start-Sleep { }
        Invoke-CollectorScript 'Get-TeamsUserActivityUserDetail.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        $global:TaTest.UsageTries | Should -Be 2
        @(Import-Csv -LiteralPath (Join-Path $script:Out 'teams-user-activity-user-detail.csv')).Count | Should -Be 1
        $log = Get-LogText $script:Out
        $log | Should -Match 'HTTP 429'
        $log | Should -Not -Match 'Reports\.Read\.All'
    }

    It 'logs a usage report that stays throttled as throttle, not a missing permission' {
        Mock Get-MgReportTeamUserActivityCount { throw 'Response status code does not indicate success: 429 (Too Many Requests).' }
        Mock Start-Sleep { }
        Invoke-CollectorScript 'Get-TeamsUserActivityCounts.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        @(Import-Csv -LiteralPath (Join-Path $script:Out 'teams-user-activity-counts.csv')).Count | Should -Be 0
        $log = Get-LogText $script:Out
        $log | Should -Match 'still throttled'
        $log | Should -Not -Match 'Reports\.Read\.All'
    }

    It 'writes the header only and logs the refusal when <Script> is refused' -ForEach @(
        @{ Script = 'Get-TeamsUserActivityUserDetail.ps1'; Cmdlet = 'Get-MgReportTeamUserActivityUserDetail'; Csv = 'teams-user-activity-user-detail.csv'; Key = 'TeamsUserActivityUserDetail' }
        @{ Script = 'Get-TeamsUserActivityCounts.ps1'; Cmdlet = 'Get-MgReportTeamUserActivityCount'; Csv = 'teams-user-activity-counts.csv'; Key = 'TeamsUserActivityCounts' }
        @{ Script = 'Get-TeamsDeviceUsageUserDetail.ps1'; Cmdlet = 'Get-MgReportTeamDeviceUsageUserDetail'; Csv = 'teams-device-usage-user-detail.csv'; Key = 'TeamsDeviceUsageUserDetail' }
    ) {
        Mock $Cmdlet -MockWith { throw 'Forbidden' }
        Invoke-CollectorScript $Script @{ OutputPath = $script:Out; SkipConnect = $true }
        $path = Join-Path $script:Out $Csv
        Get-HeaderText $path | Should -Be ($script:Schema[$Key] -join ',')
        @(Import-Csv -LiteralPath $path).Count | Should -Be 0
        Get-LogText $script:Out | Should -Match 'Reports.Read.All'
    }
}

Describe 'Report settings (source 5)' {
    BeforeEach { $script:Out = New-TestFolder }
    AfterEach { Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue }

    It 'records <Expected> when displayConcealedNames is <Value>' -ForEach @(
        @{ Value = $true; Expected = 'True' }
        @{ Value = $false; Expected = 'False' }
    ) {
        Mock Get-MgAdminReportSetting { [pscustomobject]@{ DisplayConcealedNames = $Value } }.GetNewClosure()
        Invoke-CollectorScript 'Get-ReportSettings.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        $row = @(Import-Csv -LiteralPath (Join-Path $script:Out 'report-settings.csv'))[0]
        $row.DisplayConcealedNames | Should -Be $Expected
        $row.RunDate | Should -Be ([datetime]::UtcNow.ToString('yyyy-MM-dd'))
        if ($Expected -eq 'True') { Get-LogText $script:Out | Should -Match 'conceal names' }
    }

    It 'retries report settings on 429 and does not record throttle as a missing permission' {
        $global:TaTest.SettingsTries = 0
        Mock Get-MgAdminReportSetting {
            $global:TaTest.SettingsTries++
            if ($global:TaTest.SettingsTries -eq 1) { throw 'Response status code does not indicate success: 429 (Too Many Requests).' }
            [pscustomobject]@{ DisplayConcealedNames = $false }
        }
        Mock Start-Sleep { }
        Invoke-CollectorScript 'Get-ReportSettings.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        $global:TaTest.SettingsTries | Should -Be 2
        (Import-Csv -LiteralPath (Join-Path $script:Out 'report-settings.csv')).DisplayConcealedNames | Should -Be 'False'
        $log = Get-LogText $script:Out
        $log | Should -Match 'HTTP 429'
        $log | Should -Not -Match 'ReportSettings\.Read\.All'
    }

    It 'writes the header only when the setting cannot be read' {
        Mock Get-MgAdminReportSetting { throw 'Forbidden' }
        Invoke-CollectorScript 'Get-ReportSettings.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        Get-HeaderText (Join-Path $script:Out 'report-settings.csv') | Should -Be ($script:Schema.ReportSettings -join ',')
        @(Import-Csv -LiteralPath (Join-Path $script:Out 'report-settings.csv')).Count | Should -Be 0
        Get-LogText $script:Out | Should -Match 'ReportSettings.Read.All'
    }
}

Describe 'Call records (source 6) are an event source that pages and resumes' {
    BeforeEach {
        $script:Out = New-TestFolder
        $global:TaTest.Requested = [System.Collections.Generic.List[string]]::new()
        Mock Start-Sleep { }
    }
    AfterEach { Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue }

    It 'follows the list nextLink and the sessions nextLink as returned, with GET only, and writes one row per session' {
        Mock Invoke-MgGraphRequest {
            $global:TaTest.Requested.Add("$Method $Uri")
            switch -Wildcard ($Uri) {
                '*communications/callRecords[?]*skiptoken=p2*' { @{ value = @(@{ id = 'rec-2' }) }; break }
                '*communications/callRecords[?]*' { @{ value = @(@{ id = 'rec-1' }); '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/communications/callRecords?$skiptoken=p2' }; break }
                '*callRecords/rec-1/sessions*skiptoken=s2*' { @{ value = @(New-MockSession -Id 's-1b' -CallerPlatform 'iOS') }; break }
                '*callRecords/rec-1*' {
                    New-MockCallRecord -Id 'rec-1' -Session @(New-MockSession -Id 's-1a') -SessionsNext 'https://graph.microsoft.com/v1.0/communications/callRecords/rec-1/sessions?$skiptoken=s2'
                    break
                }
                '*callRecords/rec-2*' { New-MockCallRecord -Id 'rec-2' -Start '2026-10-07T15:00:00Z' -Session @(New-MockSession -Id 's-2a' -Caller 'user-2' -CallerPlatform 'macOS' -Callee 'user-3' -CalleePlatform 'web'); break }
            }
        }
        Invoke-CollectorScript 'Get-CallRecords.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:Out 'call-records.csv'))
        ($rows.SessionId | Sort-Object) -join ',' | Should -Be 's-1a,s-1b,s-2a'
        ($rows | Where-Object SessionId -EQ 's-1b').CallerPlatform | Should -Be 'iOS'
        $two = $rows | Where-Object SessionId -EQ 's-2a'
        $two.CallerUserId | Should -Be 'user-2'
        $two.CalleePlatform | Should -Be 'web'
        $rows[0].Modalities | Should -Be 'audio;video'
        $rows[0].Type | Should -Be 'groupCall'
        $global:TaTest.Requested | Where-Object { $_ -notlike 'GET *' } | Should -BeNullOrEmpty
        @($global:TaTest.Requested | Where-Object { $_ -like '*skiptoken=p2*' }).Count | Should -Be 1
        @($global:TaTest.Requested | Where-Object { $_ -like '*skiptoken=s2*' }).Count | Should -Be 1
    }

    It 'filters on startDateTime with ge and lt in ISO 8601 UTC, ending before now' {
        Mock Invoke-MgGraphRequest { $global:TaTest.Requested.Add($Uri); @{ value = @() } }
        Invoke-CollectorScript 'Get-CallRecords.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; LookbackDays = 5; DelayMinutes = 180 }

        $uri = [uri]::UnescapeDataString($global:TaTest.Requested[0])
        $uri | Should -Match 'startDateTime ge \d{4}-\d\d-\d\dT\d\d:\d\d:\d\dZ and startDateTime lt \d{4}-\d\d-\d\dT\d\d:\d\d:\d\dZ'
        $m = [regex]::Match($uri, 'ge (\S+) and startDateTime lt (\S+)')
        $from = [datetime]::Parse($m.Groups[1].Value, [cultureinfo]::InvariantCulture, 'AdjustToUniversal,AssumeUniversal')
        $to = [datetime]::Parse($m.Groups[2].Value, [cultureinfo]::InvariantCulture, 'AdjustToUniversal,AssumeUniversal')
        ([datetime]::UtcNow - $from).TotalDays | Should -BeGreaterThan 4.9
        ([datetime]::UtcNow - $from).TotalDays | Should -BeLessThan 5.1
        ([datetime]::UtcNow - $to).TotalMinutes | Should -BeGreaterThan 179
        ([datetime]::UtcNow - $to).TotalMinutes | Should -BeLessThan 181
    }

    It 'still requests a call that started before the latest stored start' {
        # A later start already in the file must not become the filter. The record can appear
        # up to 150 minutes after the call ends, and the list filters on startDateTime.
        # https://learn.microsoft.com/graph/callrecords-api-faq
        $stored = [datetime]::UtcNow.AddDays(-1).ToString('yyyy-MM-ddTHH:mm:ssZ')
        Export-AppendCsv -Path (Join-Path $script:Out 'call-records.csv') -Column $script:Schema.CallRecords -Rows @(
            [pscustomobject]@{ CallRecordId = 'newer'; Version = '1'; StartDateTime = $stored; SessionId = 'already' })
        $global:TaTest.RecordStart = [datetime]::UtcNow.AddDays(-10)
        Mock Invoke-MgGraphRequest {
            $text = [uri]::UnescapeDataString([string]$Uri)
            if ($text -like '*communications/callRecords[?]*') {
                $matched = [regex]::Match($text, 'ge (\S+) and')
                $from = [datetime]::Parse($matched.Groups[1].Value, [cultureinfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AdjustToUniversal -bor [System.Globalization.DateTimeStyles]::AssumeUniversal)
                if ($from -le $global:TaTest.RecordStart) { return @{ value = @(@{ id = 'late' }) } }
                return @{ value = @() }
            }
            New-MockCallRecord -Id 'late' -Start ($global:TaTest.RecordStart.ToString('yyyy-MM-ddTHH:mm:ssZ')) -Session @(New-MockSession -Id 'late-s')
        }
        Invoke-CollectorScript 'Get-CallRecords.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; LookbackDays = 30; DelayMinutes = 180 }

        (Import-Csv -LiteralPath (Join-Path $script:Out 'call-records.csv')).CallRecordId | Should -Contain 'late'
    }

    It 'keeps both rows when one session id is repeated for two service identities' {
        Mock Invoke-MgGraphRequest {
            if ($Uri -like '*communications/callRecords[?]*') { return @{ value = @(@{ id = 'rec-1' }) } }
            $first = New-MockSession -Id 'same' -Caller 'user-1' -Callee 'svc-1' -CalleePlatform 'unknown'
            $second = New-MockSession -Id 'same' -Caller 'user-1' -Callee 'svc-2' -CalleePlatform 'unknown'
            $second.endDateTime = '2026-10-07T14:20:00Z'
            New-MockCallRecord -Id 'rec-1' -Session @($first, $second)
        }
        Invoke-CollectorScript 'Get-CallRecords.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        $rows = @(Import-Csv -LiteralPath (Join-Path $script:Out 'call-records.csv'))
        $rows.Count | Should -Be 2
        ($rows.CalleeUserId | Sort-Object) -join ',' | Should -Be 'svc-1,svc-2'
    }

    It 'sends Prefer include-unknown-enum-members on the list and the record' {
        $global:TaTest.Prefer = [System.Collections.Generic.List[string]]::new()
        Mock Invoke-MgGraphRequest {
            $global:TaTest.Prefer.Add([string]$Headers['Prefer'])
            if ($Uri -like '*communications/callRecords[?]*') { return @{ value = @(@{ id = 'rec-1' }) } }
            New-MockCallRecord -Id 'rec-1' -Session @(New-MockSession -Id 's-1')
        }
        Invoke-CollectorScript 'Get-CallRecords.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        $global:TaTest.Prefer.Count | Should -BeGreaterThan 1
        @($global:TaTest.Prefer | Where-Object { $_ -ne 'include-unknown-enum-members' }).Count | Should -Be 0
    }

    It 'does not repeat a row it already has, and keeps a later version of the same record as its own rows' {
        $global:TaTest.Version = 1
        Mock Invoke-MgGraphRequest {
            if ($Uri -like '*callRecords[?]*') { return @{ value = @(@{ id = 'rec-1' }) } }
            New-MockCallRecord -Id 'rec-1' -Version $global:TaTest.Version -Session @(New-MockSession -Id 's-1')
        }
        Invoke-CollectorScript 'Get-CallRecords.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        Invoke-CollectorScript 'Get-CallRecords.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        @(Import-Csv -LiteralPath (Join-Path $script:Out 'call-records.csv')).Count | Should -Be 1

        $global:TaTest.Version = 2
        Invoke-CollectorScript 'Get-CallRecords.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        $rows = @(Import-Csv -LiteralPath (Join-Path $script:Out 'call-records.csv'))
        ($rows.Version | Sort-Object) -join ',' | Should -Be '1,2'
    }

    It 'writes one row with empty session columns when a record has no sessions' {
        Mock Invoke-MgGraphRequest {
            if ($Uri -like '*callRecords[?]*') { return @{ value = @(@{ id = 'rec-1' }) } }
            New-MockCallRecord -Id 'rec-1'
        }
        Invoke-CollectorScript 'Get-CallRecords.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        $rows = @(Import-Csv -LiteralPath (Join-Path $script:Out 'call-records.csv'))
        $rows.Count | Should -Be 1
        $rows[0].SessionId | Should -Be ''
        $rows[0].CallerPlatform | Should -Be ''
    }

    It 'skips a record that answers 404 and keeps the others' {
        Mock Invoke-MgGraphRequest {
            if ($Uri -like '*callRecords[?]*') { return @{ value = @(@{ id = 'gone' }, @{ id = 'rec-2' }) } }
            if ($Uri -like '*callRecords/gone*') { throw 'Response status code does not indicate success: 404 (Not Found).' }
            New-MockCallRecord -Id 'rec-2' -Session @(New-MockSession -Id 's-2')
        }
        Invoke-CollectorScript 'Get-CallRecords.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        (Import-Csv -LiteralPath (Join-Path $script:Out 'call-records.csv')).CallRecordId | Should -Be 'rec-2'
        Get-LogText $script:Out | Should -Match 'gone answered 404'
    }

    It 'refuses a nextLink on a host that is not Microsoft Graph' {
        Mock Invoke-MgGraphRequest { @{ value = @(); '@odata.nextLink' = 'https://evil.example.net/v1.0/communications/callRecords?$skiptoken=x' } }
        Invoke-CollectorScript 'Get-CallRecords.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        Get-LogText $script:Out | Should -Match 'outside Microsoft Graph'
        Should -Invoke Invoke-MgGraphRequest -Times 1 -Exactly
        @(Import-Csv -LiteralPath (Join-Path $script:Out 'call-records.csv')).Count | Should -Be 0
    }

    It 'writes the header only and names the application permission when the list is refused' {
        Mock Invoke-MgGraphRequest { throw 'Forbidden' }
        Invoke-CollectorScript 'Get-CallRecords.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        Get-HeaderText (Join-Path $script:Out 'call-records.csv') | Should -Be ($script:Schema.CallRecords -join ',')
        Get-LogText $script:Out | Should -Match 'application permission CallRecords.Read.All'
    }

    It 'throws when call records stay throttled instead of recording a missing permission' {
        Mock Invoke-MgGraphRequest { throw 'Response status code does not indicate success: 429 (Too Many Requests).' }
        { Invoke-CollectorScript 'Get-CallRecords.ps1' @{ OutputPath = $script:Out; SkipConnect = $true } } | Should -Throw '*429*'
        if (Test-Path -LiteralPath (Join-Path $script:Out 'run.log')) {
            Get-LogText $script:Out | Should -Not -Match 'application permission CallRecords\.Read\.All'
        }
    }

    It 'retries a throttled page instead of keeping an empty file' {
        $global:TaTest.Calls = 0
        Mock Invoke-MgGraphRequest {
            $global:TaTest.Calls++
            if ($global:TaTest.Calls -eq 1) { throw 'Response status code does not indicate success: 429 (Too Many Requests).' }
            @{ value = @() }
        }
        Invoke-CollectorScript 'Get-CallRecords.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        $global:TaTest.Calls | Should -Be 2
        Get-LogText $script:Out | Should -Match 'HTTP 429'
    }
}

Describe 'Teams audit events (source 8) resume from the last day collected' {
    BeforeEach {
        $script:Out = New-TestFolder
        $global:TaTest.Windows = [System.Collections.Generic.List[object]]::new()
        Mock Start-Sleep { }
    }
    AfterEach { Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue }

    It 'counts records per UTC day, workload, user and operation, searching the five Teams operations' {
        Mock Search-UnifiedAuditLog {
            if ($StartDate.ToUniversalTime().Date -eq [datetime]'2026-10-05') {
                New-MockAuditRecord -Id 'a' -Operation 'MeetingDetail' -CreationTime '2026-10-05T01:00:00'
                New-MockAuditRecord -Id 'b' -Operation 'MeetingDetail' -CreationTime '2026-10-05T23:59:00'
                New-MockAuditRecord -Id 'c' -Operation 'MessageSent' -CreationTime '2026-10-05T10:00:00'
                New-MockAuditRecord -Id 'd' -Operation 'ChatCreated' -CreationTime '2026-10-05T11:00:00' -UserId 'devon.diaz@example.com'
            }
        }
        Invoke-CollectorScript 'Get-TeamsAuditEvents.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; StartDate = [datetime]'2026-10-05T00:00:00Z'; EndDate = [datetime]'2026-10-06T00:00:00Z' }

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:Out 'teams-audit-events.csv'))
        $rows.Count | Should -Be 3
        ($rows | Where-Object Operation -EQ 'MeetingDetail').EventCount | Should -Be '2'
        ($rows | Where-Object Operation -EQ 'ChatCreated').UserId | Should -Be 'devon.diaz@example.com'
        ($rows.Date | Sort-Object -Unique) | Should -Be '2026-10-05'
        Should -Invoke Search-UnifiedAuditLog -ParameterFilter { $Operations.Count -eq 5 -and $Operations -contains 'MeetingParticipantDetail' -and $Operations -contains 'CallParticipantDetail' }
    }

    It 'starts the day after the latest Date already in the file and stops before today' {
        Mock Search-UnifiedAuditLog {
            $global:TaTest.Windows.Add([pscustomobject]@{ Start = $StartDate.ToUniversalTime(); End = $EndDate.ToUniversalTime() })
            New-MockAuditRecord -Id ('e-' + $global:TaTest.Windows.Count) -CreationTime ($StartDate.ToUniversalTime().AddHours(1).ToString('s'))
        }
        $yesterday = [datetime]::UtcNow.Date.AddDays(-1)
        $twoDaysAgo = [datetime]::UtcNow.Date.AddDays(-2)
        Export-AppendCsv -Path (Join-Path $script:Out 'teams-audit-events.csv') -Column $script:Schema.TeamsAuditEvents -Rows @(
            [pscustomobject]@{ Date = $twoDaysAgo.ToString('yyyy-MM-dd'); Workload = 'MicrosoftTeams'; UserId = 'avery.abara@example.com'; Operation = 'MeetingDetail'; EventCount = 7 })

        Invoke-CollectorScript 'Get-TeamsAuditEvents.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }

        $global:TaTest.Windows.Count | Should -Be 1
        $global:TaTest.Windows[0].Start | Should -Be $yesterday
        $global:TaTest.Windows[0].End | Should -Be ([datetime]::UtcNow.Date)
        @(Import-Csv -LiteralPath (Join-Path $script:Out 'teams-audit-events.csv')).Count | Should -Be 2
    }

    It 'reaches back -LookbackDays on the first run and splits the range into -WindowHours windows' {
        Mock Search-UnifiedAuditLog { $global:TaTest.Windows.Add([pscustomobject]@{ Start = $StartDate.ToUniversalTime() }) }
        Invoke-CollectorScript 'Get-TeamsAuditEvents.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; LookbackDays = 2; WindowHours = 12 }
        $starts = @($global:TaTest.Windows.Start | Sort-Object -Unique)
        $starts.Count | Should -Be 4
        $starts[0] | Should -Be ([datetime]::UtcNow.Date.AddDays(-2))
        $starts[1] | Should -Be ([datetime]::UtcNow.Date.AddDays(-2).AddHours(12))
    }

    It 'writes nothing new and says so when the last whole day is already collected' {
        Export-AppendCsv -Path (Join-Path $script:Out 'teams-audit-events.csv') -Column $script:Schema.TeamsAuditEvents -Rows @(
            [pscustomobject]@{ Date = [datetime]::UtcNow.Date.AddDays(-1).ToString('yyyy-MM-dd'); Workload = 'MicrosoftTeams'; UserId = 'a@example.com'; Operation = 'MeetingDetail'; EventCount = 1 })
        Mock Search-UnifiedAuditLog { throw 'should not be called' }
        Invoke-CollectorScript 'Get-TeamsAuditEvents.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        Should -Invoke Search-UnifiedAuditLog -Times 0 -Exactly
        Get-LogText $script:Out | Should -Match 'Nothing new to collect'
    }

    It 'reads the next page in one ReturnLargeSet session while moreRecordsAvailable is true' {
        $global:TaTest.Sessions = [System.Collections.Generic.List[string]]::new()
        Mock Search-UnifiedAuditLog {
            $global:TaTest.Sessions.Add("$SessionId|$SessionCommand|$ResultSize")
            if ($global:TaTest.Sessions.Count -eq 1) { New-MockAuditRecord -Id 'p1' -MoreRecords $true }
            elseif ($global:TaTest.Sessions.Count -eq 2) { New-MockAuditRecord -Id 'p2' -MoreRecords $false }
        }
        Invoke-CollectorScript 'Get-TeamsAuditEvents.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; StartDate = [datetime]'2026-10-05T00:00:00Z'; EndDate = [datetime]'2026-10-06T00:00:00Z' }

        $global:TaTest.Sessions.Count | Should -Be 2
        ($global:TaTest.Sessions | ForEach-Object { ($_ -split '\|')[0] } | Sort-Object -Unique).Count | Should -Be 1
        $global:TaTest.Sessions[0] | Should -Match '\|ReturnLargeSet\|5000$'
        (Import-Csv -LiteralPath (Join-Path $script:Out 'teams-audit-events.csv')).EventCount | Should -Be '2'
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
        { Invoke-CollectorScript 'Get-TeamsAuditEvents.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; StartDate = [datetime]'2026-10-05T00:00:00Z'; EndDate = [datetime]'2026-10-07T00:00:00Z' } } | Should -Throw '*50,000-record session cap*'

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:Out 'teams-audit-events.csv'))
        ($rows.Date | Sort-Object -Unique) | Should -Be '2026-10-05'
        Get-LogText $script:Out | Should -Match 'matches 50,000 or more records'
    }

    It 'writes a header only and logs the reason when the audit log is refused' {
        Mock Search-UnifiedAuditLog { throw 'The role is missing' }
        Invoke-CollectorScript 'Get-TeamsAuditEvents.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        Get-HeaderText (Join-Path $script:Out 'teams-audit-events.csv') | Should -Be ($script:Schema.TeamsAuditEvents -join ',')
        @(Import-Csv -LiteralPath (Join-Path $script:Out 'teams-audit-events.csv')).Count | Should -Be 0
        Get-LogText $script:Out | Should -Match 'View-Only Audit Logs'
    }

    It 'warns in <Cloud> that three operations are UNVERIFIED, and not in Commercial' -ForEach @(
        @{ Cloud = 'GCC'; Warn = $true }, @{ Cloud = 'GCCHigh'; Warn = $true }, @{ Cloud = 'Commercial'; Warn = $false }
    ) {
        Mock Search-UnifiedAuditLog { }
        Invoke-CollectorScript 'Get-TeamsAuditEvents.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; Environment = $Cloud }
        $log = Get-LogText $script:Out
        if ($Warn) { $log | Should -Match 'CallParticipantDetail, MessageSent, ChatCreated are logged in \w+ is UNVERIFIED' } else { $log | Should -Not -Match 'UNVERIFIED' }
    }

    It 'rejects an empty explicit range and a lookback past the 180-day retention' {
        { Invoke-CollectorScript 'Get-TeamsAuditEvents.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; StartDate = [datetime]'2026-10-06T00:00:00Z'; EndDate = [datetime]'2026-10-05T00:00:00Z' } } | Should -Throw '*requested range is empty*'
        { Invoke-CollectorScript 'Get-TeamsAuditEvents.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; LookbackDays = 181 } } | Should -Throw
    }
}

Describe 'Run-All' {
    BeforeEach {
        $script:Out = New-TestFolder
        Mock Connect-MgGraph -ModuleName M365ReportLibrary -MockWith { }
        Mock Connect-ExchangeOnline -ModuleName M365ReportLibrary -MockWith { }
        Mock Disconnect-ExchangeOnline { }
        Mock Get-MgReportTeamUserActivityUserDetail { }
        Mock Get-MgReportTeamUserActivityCount { }
        Mock Get-MgReportTeamDeviceUsageUserDetail { }
        Mock Get-MgReportTeamActivityDetail { }
        Mock Get-MgAdminReportSetting { [pscustomobject]@{ DisplayConcealedNames = $false } }
        Mock Invoke-MgGraphRequest { @{ value = @() } }
        Mock Search-UnifiedAuditLog { }
        Mock Start-Sleep { }
    }
    AfterEach { Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue }

    It 'leaves seven CSVs including the lifecycle report''s team-activity.csv, after one sign-in to each service' {
        & (Join-Path $script:Collectors 'Run-All.ps1') -OutputPath $script:Out

        @(Get-ChildItem -LiteralPath $script:Out -Filter '*.csv').Count | Should -Be 7
        Test-Path -LiteralPath (Join-Path $script:Out 'team-activity.csv') | Should -BeTrue
        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter { $Scopes -contains 'Reports.Read.All' -and $Scopes -contains 'ReportSettings.Read.All' -and $Scopes -notcontains 'CallRecords.Read.All' }
        Should -Invoke Connect-ExchangeOnline -ModuleName M365ReportLibrary -Times 1 -Exactly
        Should -Invoke Disconnect-ExchangeOnline -Times 1 -Exactly
    }

    It 'leaves a header-only file for each usage report in GCC High and still reads call records and the audit log' {
        & (Join-Path $script:Collectors 'Run-All.ps1') -OutputPath $script:Out -Environment GCCHigh

        foreach ($name in 'teams-user-activity-user-detail', 'teams-user-activity-counts', 'teams-device-usage-user-detail', 'team-activity') {
            Test-Path -LiteralPath (Join-Path $script:Out "$name.csv") | Should -BeTrue
            @(Import-Csv -LiteralPath (Join-Path $script:Out "$name.csv")).Count | Should -Be 0
        }
        Should -Invoke Get-MgReportTeamUserActivityUserDetail -Times 0 -Exactly
        Should -Invoke Get-MgReportTeamActivityDetail -Times 0 -Exactly
        Should -Invoke Invoke-MgGraphRequest -Times 1
        Should -Invoke Search-UnifiedAuditLog -Scope It -Times 1
    }

    It 'keeps running the other collectors when one stops' {
        Mock Get-MgAdminReportSetting { throw 'Forbidden' }
        & (Join-Path $script:Collectors 'Run-All.ps1') -OutputPath $script:Out
        @(Get-ChildItem -LiteralPath $script:Out -Filter '*.csv').Count | Should -Be 7
    }

    It 'passes the period to the usage reports and the lifecycle team report' {
        & (Join-Path $script:Collectors 'Run-All.ps1') -OutputPath $script:Out -Period D90
        Should -Invoke Get-MgReportTeamUserActivityCount -Times 1 -Exactly -ParameterFilter { $Period -eq 'D90' }
        Should -Invoke Get-MgReportTeamActivityDetail -Times 1 -Exactly -ParameterFilter { $Period -eq 'D90' }
    }
}
