#Requires -Version 7.0

BeforeAll {
    $script:Root = Join-Path $PSScriptRoot '../../..' | Resolve-Path | Select-Object -ExpandProperty Path
    $script:Collectors = Join-Path $script:Root 'reports/copilot-usage/collectors'
    $script:Samples = Join-Path $script:Root 'reports/copilot-usage/samples'

    . (Join-Path $script:Root 'shared/tests/TenantCmdletStubs.ps1')
    . (Join-Path $PSScriptRoot 'CopilotUsageStubs.ps1')
    Import-Module (Join-Path $script:Root 'shared/M365ReportLibrary.psm1') -Force
    . (Join-Path $script:Collectors 'CopilotUsageHelpers.ps1')

    $script:Schema = Import-PowerShellDataFile -LiteralPath (Join-Path $script:Collectors 'CopilotUsageSchema.psd1')
    $global:CuTest = @{}

    $script:ReportCases = @(
        @{ Script = 'Get-CopilotUsageUserDetail.ps1'; Csv = 'copilot-usage-user-detail.csv'; MapKey = 'UsageUserDetailMap'; Function = 'getMicrosoft365CopilotUsageUserDetail' }
        @{ Script = 'Get-CopilotUserCountSummary.ps1'; Csv = 'copilot-user-count-summary.csv'; MapKey = 'UserCountSummaryMap'; Function = 'getMicrosoft365CopilotUserCountSummary' }
        @{ Script = 'Get-CopilotUserCountTrend.ps1'; Csv = 'copilot-user-count-trend.csv'; MapKey = 'UserCountTrendMap'; Function = 'getMicrosoft365CopilotUserCountTrend' }
    )

    function New-TestFolder {
        $path = Join-Path ([System.IO.Path]::GetTempPath()) ('copilot-usage-' + [guid]::NewGuid().ToString('N'))
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

    function Get-ExpectedColumn {
        param([string]$MapKey)
        return (Get-CopilotColumn -Schema $script:Schema -MapKey $MapKey) -join ','
    }

    # A usage report CSV stream: the headers are the ones on the Learn page, so a row is
    # built from the schema's own header map. Values are fake.
    # https://learn.microsoft.com/microsoft-365-copilot/extensibility/api/admin-settings/reports/copilotreportroot-getmicrosoft365copilotusageuserdetail
    function New-ReportCsv {
        param([string]$MapKey, [int]$Rows = 1, [string]$Prefix = 'user')

        $map = @($script:Schema[$MapKey])
        $header = ($map | ForEach-Object { $_.Header }) -join ','
        $lines = foreach ($n in 1..$Rows) {
            ($map | ForEach-Object {
                    switch -Wildcard ($_.Header) {
                        'User Principal Name' { "$Prefix$n@example.com" }
                        'Display Name' { "Fake User $n" }
                        'Report Date' { ([datetime]'2026-09-01').AddDays($n % 28).ToString('yyyy-MM-dd') }
                        'Report Refresh Date' { '2026-09-29' }
                        'Report Period' { '28' }
                        '*Date' { '2026-09-20' }
                        default { '7' }
                    }
                }) -join ','
        }
        return (@($header) + @($lines)) -join "`n"
    }

    # aiInteraction: https://learn.microsoft.com/microsoft-365-copilot/extensibility/api/ai-services/interaction-export/aiinteractionhistory-getallenterpriseinteractions
    function New-MockInteraction {
        param(
            [string]$Id = '1731701801008',
            [string]$At = '2026-10-05T20:16:41.008Z',
            [string]$Type = 'aiResponse',
            [string]$Text = 'SECRET-PROMPT-TEXT: what is on my radar?'
        )

        @{
            id = $Id
            sessionId = '19:icg2t_AWPYJyJ2oDLB_CZyh29QXpZvbdpljKf7qKotk1@thread.v2'
            requestId = '7336770c-fb25-48ac-8303-4493ad11ed71'
            appClass = 'IPM.SkypeTeams.Message.Copilot.Teams'
            interactionType = $Type
            conversationType = 'appchat'
            etag = $Id
            createdDateTime = $At
            locale = 'en-us'
            contexts = @(@{ contextReference = 'https://microsoft.teams.com/threads/19:meeting_x@thread.v2'; displayName = 'Teams Meeting Copilot'; contextType = 'TeamsMeeting' })
            from = @{ user = $null; application = @{ id = 'fb8d773d-7ef8-4ec0-a117-179f88add510'; displayName = 'Copilot in Teams' } }
            body = @{ contentType = 'text'; content = $Text }
            attachments = @(@{ attachmentId = 'a1'; contentType = 'reference'; content = 'SECRET-ATTACHMENT-TEXT'; name = 'Teams Meeting Copilot' })
            links = @()
            mentions = @()
        }
    }

    # Search-UnifiedAuditLog record: RecordType, ResultCount, AuditData as a JSON string.
    # CopilotInteraction properties: https://learn.microsoft.com/purview/audit-copilot
    function New-MockAuditRecord {
        param(
            [string]$Id = 'event-1',
            [string]$CreationTime = '2026-10-05T12:00:00',
            [string]$UserId = 'avery.abara@example.com',
            [object]$MoreRecords = $null,
            [switch]$AgentAtTop
        )

        $eventData = [ordered]@{
            AppHost           = 'Teams'
            AppIdentity       = 'Copilot.MicrosoftCopilot.BizChat'
            Contexts          = @([ordered]@{ ID = 'chat-1'; Type = 'TeamsChat' })
            Messages          = @([ordered]@{ ID = '1715186983849'; isPrompt = $true }, [ordered]@{ ID = '1715186984291'; isPrompt = $false }, [ordered]@{ ID = '1715186984292'; isPrompt = $false })
            AccessedResources = @([ordered]@{ Action = 'Read'; ID = 'f1'; Name = 'Document1.docx'; Type = 'docx' })
            AISystemPlugin    = @([ordered]@{ Name = 'Bing'; ID = 'BingWebSearch'; Version = '1' })
        }
        $audit = [ordered]@{ CreationTime = $CreationTime; Id = $Id; Operation = 'CopilotInteraction'; RecordType = 261; UserId = $UserId; Workload = 'Copilot' }
        if ($AgentAtTop) { $audit['AgentId'] = 'CopilotStudio.Declarative.8ad83f3e'; $audit['AgentName'] = 'SalesAgent' }
        else { $eventData['AgentId'] = 'CopilotStudio.Declarative.11fd28b5'; $eventData['AgentName'] = 'ReminderBot' }
        $audit['CopilotEventData'] = $eventData

        $record = [ordered]@{ RecordType = 'CopilotInteraction'; AuditData = ($audit | ConvertTo-Json -Compress -Depth 6) }
        if ($null -ne $MoreRecords) { $record['AuditSearchRequestMetadata'] = [pscustomobject]@{ moreRecordsAvailable = $MoreRecords } }
        return [pscustomobject]$record
    }

    function New-UsersCsv {
        param([string]$Folder, [string[]]$Id = @('user-1'))
        $rows = foreach ($i in $Id) { [pscustomobject]@{ RunDate = '2026-10-01'; Id = $i } }
        Export-AppendCsv -Path (Join-Path $Folder 'users.csv') -Rows @($rows) -Column @('RunDate', 'Id')
    }
}

AfterAll {
    Remove-Variable -Name CuTest -Scope Global -ErrorAction SilentlyContinue
    Remove-Module M365ReportLibrary -Force -ErrorAction SilentlyContinue
}

Describe 'Each collector writes the columns of its sample file' {
    BeforeEach {
        $script:Out = New-TestFolder
        Mock Connect-M365Service -MockWith { }
        Mock Disconnect-ExchangeOnline { }
        Mock Start-Sleep { }
    }
    AfterEach { Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue }

    It '<Csv>' -ForEach $script:ReportCases {
        $global:CuTest.Csv = New-ReportCsv -MapKey $MapKey
        Mock Invoke-MgGraphRequest { Set-Content -LiteralPath $OutputFilePath -Value $global:CuTest.Csv }

        Invoke-CollectorScript $Script @{ OutputPath = $script:Out }

        $produced = Join-Path $script:Out $Csv
        @(Import-Csv -LiteralPath $produced).Count | Should -Be 1
        Get-HeaderText $produced | Should -Be (Get-HeaderText (Join-Path $script:Samples $Csv))
        Get-HeaderText $produced | Should -Be (Get-ExpectedColumn $MapKey)
    }

    It 'copilot-audit-events.csv' {
        Mock Search-UnifiedAuditLog { New-MockAuditRecord }
        Invoke-CollectorScript 'Get-CopilotAuditEvents.ps1' @{ OutputPath = $script:Out; LookbackDays = 1 }
        $produced = Join-Path $script:Out 'copilot-audit-events.csv'
        @(Import-Csv -LiteralPath $produced).Count | Should -BeGreaterThan 0
        Get-HeaderText $produced | Should -Be (Get-HeaderText (Join-Path $script:Samples 'copilot-audit-events.csv'))
        Get-HeaderText $produced | Should -Be ($script:Schema.CopilotAuditEvents -join ',')
    }

    It 'copilot-interactions.csv' {
        New-UsersCsv -Folder $script:Out
        Mock Invoke-MgGraphRequest { @{ value = @(New-MockInteraction) } }
        Invoke-CollectorScript 'Get-CopilotInteractions.ps1' @{ OutputPath = $script:Out }
        $produced = Join-Path $script:Out 'copilot-interactions.csv'
        @(Import-Csv -LiteralPath $produced).Count | Should -Be 1
        Get-HeaderText $produced | Should -Be (Get-HeaderText (Join-Path $script:Samples 'copilot-interactions.csv'))
        Get-HeaderText $produced | Should -Be ($script:Schema.CopilotInteractions -join ',')
    }

    It 'copilot-feature-availability.csv' {
        Invoke-CollectorScript 'Get-CopilotFeatureAvailability.ps1' @{ OutputPath = $script:Out }
        $produced = Join-Path $script:Out 'copilot-feature-availability.csv'
        @(Import-Csv -LiteralPath $produced).Count | Should -Be $script:Schema.FeatureRows.Count
        Get-HeaderText $produced | Should -Be (Get-HeaderText (Join-Path $script:Samples 'copilot-feature-availability.csv'))
        Get-HeaderText $produced | Should -Be ($script:Schema.CopilotFeatureAvailability -join ',')
    }

    It 'reads each report column by the header Learn gives it' {
        $header = (Get-Content -LiteralPath (Join-Path $script:Samples 'copilot-usage-user-detail.csv') -TotalCount 1)
        $header | Should -Match 'LastActivityDate'
        $names = $script:Schema.UsageUserDetailMap.Header
        foreach ($expected in 'Report Refresh Date', 'User Principal Name', 'Display Name', 'Last Activity Date', 'Copilot Chat Last Activity Date',
            'Microsoft Teams Copilot Last Activity Date', 'Word Copilot Last Activity Date', 'Excel Copilot Last Activity Date',
            'PowerPoint Copilot Last Activity Date', 'Outlook Copilot Last Activity Date', 'OneNote Copilot Last Activity Date',
            'Loop Copilot Last Activity Date', 'Report Period', 'Prompts submitted for all apps', 'Active Usage Days for all apps',
            'Copilot Agent Last Activity Date') { $names | Should -Contain $expected }
        $script:Schema.UserCountSummaryMap.Header | Should -Contain 'Microsoft Teams Enabled Users'
        $script:Schema.UserCountSummaryMap.Header | Should -Contain 'Total prompts submitted'
        $script:Schema.UserCountTrendMap.Header | Should -Contain 'Prompts submitted'
        $script:Schema.UserCountTrendMap.Header | Should -Not -Contain 'Total prompts submitted'
    }
}

Describe 'The samples' {
    It 'has every sample with rows, the exact columns, and the GCC High header-only case' {
        foreach ($case in $script:ReportCases) {
            $columns = Get-ExpectedColumn $case.MapKey
            Get-HeaderText (Join-Path $script:Samples $case.Csv) | Should -Be $columns
            @(Import-Csv -LiteralPath (Join-Path $script:Samples $case.Csv)).Count | Should -BeGreaterThan 0
            $gcch = Join-Path $script:Samples "gcchigh/$($case.Csv)"
            Get-HeaderText $gcch | Should -Be $columns
            (Get-Content -LiteralPath $gcch).Count | Should -Be 1
        }
        Get-HeaderText (Join-Path $script:Samples 'copilot-audit-events.csv') | Should -Be ($script:Schema.CopilotAuditEvents -join ',')
        Get-HeaderText (Join-Path $script:Samples 'copilot-interactions.csv') | Should -Be ($script:Schema.CopilotInteractions -join ',')
    }

    It 'uses example.com addresses only and carries no prompt text' {
        $text = (Get-ChildItem -LiteralPath $script:Samples -Recurse -Filter '*.csv' | Get-Content -Raw) -join ''
        $text | Should -Not -Match '@(?!example\.com)[a-z0-9.-]+\.[a-z]{2,}'
        $script:Schema.CopilotInteractions | Should -Not -Contain 'Body'
        $script:Schema.CopilotInteractions | Should -Not -Contain 'Content'
    }
}

Describe 'Each connection targets the endpoints of its -Environment' {
    BeforeEach {
        $script:Out = New-TestFolder
        Mock Connect-MgGraph -ModuleName M365ReportLibrary -MockWith { }
        Mock Connect-ExchangeOnline -ModuleName M365ReportLibrary -MockWith { }
        Mock Disconnect-ExchangeOnline { }
        Mock Search-UnifiedAuditLog { }
        Mock Invoke-MgGraphRequest { if ($OutputFilePath) { Set-Content -LiteralPath $OutputFilePath -Value '' } else { @{ value = @() } } }
        Mock Start-Sleep { }
        New-UsersCsv -Folder $script:Out
    }
    AfterEach { Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue }

    # https://learn.microsoft.com/graph/deployments: GCC calls the global service;
    # GCC High is the USGov environment.
    It '<Script> signs in to <Graph> for <Environment>' -ForEach @(
        foreach ($script in 'Get-CopilotUsageUserDetail.ps1', 'Get-CopilotUserCountSummary.ps1', 'Get-CopilotUserCountTrend.ps1') {
            @{ Script = $script; Environment = 'Commercial'; Graph = 'Global' }
            @{ Script = $script; Environment = 'GCC'; Graph = 'Global' }
        }
        foreach ($environment in 'Commercial', 'GCC', 'GCCHigh') {
            @{ Script = 'Get-CopilotInteractions.ps1'; Environment = $environment; Graph = $(if ($environment -eq 'GCCHigh') { 'USGov' } else { 'Global' }) }
        }
    ) {
        Invoke-CollectorScript $Script @{ OutputPath = $script:Out; Environment = $Environment }

        $expected = $Graph
        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter { $Environment -eq $expected }
    }

    It 'resolves the Graph resource endpoint for each cloud' -ForEach @(
        @{ Environment = 'Commercial'; Endpoint = 'https://graph.microsoft.com' }
        @{ Environment = 'GCC'; Endpoint = 'https://graph.microsoft.com' }
        @{ Environment = 'GCCHigh'; Endpoint = 'https://graph.microsoft.us' }
    ) {
        (Get-M365ServiceEndpoint -Service Graph -Environment $Environment).ResourceEndpoint | Should -Be $Endpoint
    }

    It 'signs the audit collector in to Exchange Online for <Environment> and signs out' -ForEach @(
        @{ Environment = 'Commercial'; Name = 'O365Default' }
        @{ Environment = 'GCC'; Name = 'O365Default' }
        @{ Environment = 'GCCHigh'; Name = 'O365USGovGCCHigh' }
    ) {
        Invoke-CollectorScript 'Get-CopilotAuditEvents.ps1' @{ OutputPath = $script:Out; Environment = $Environment; LookbackDays = 1 } 3>$null

        $expected = $Name
        Should -Invoke Connect-ExchangeOnline -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter { $ExchangeEnvironmentName -eq $expected }
        Should -Invoke Disconnect-ExchangeOnline -Times 1 -Exactly
    }

    It 'asks for Reports.Read.All for the usage reports' -ForEach @(
        @{ Script = 'Get-CopilotUsageUserDetail.ps1'; Scope = 'Reports.Read.All' }
        @{ Script = 'Get-CopilotUserCountSummary.ps1'; Scope = 'Reports.Read.All' }
        @{ Script = 'Get-CopilotUserCountTrend.ps1'; Scope = 'Reports.Read.All' }
    ) {
        Invoke-CollectorScript $Script @{ OutputPath = $script:Out }

        $wanted = $Scope
        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter { $Scopes -contains $wanted }
    }

    It 'does not request AiEnterpriseInteraction.Read.All on a delegated sign-in' {
        # Delegated is not supported. -Scopes is the delegated list, so the application
        # permission must not be placed on an interactive sign-in.
        # https://learn.microsoft.com/microsoft-365-copilot/extensibility/api/ai-services/interaction-export/aiinteractionhistory-getallenterpriseinteractions
        Invoke-CollectorScript 'Get-CopilotInteractions.ps1' @{ OutputPath = $script:Out }

        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter {
            $Scopes -notcontains 'AiEnterpriseInteraction.Read.All'
        }
        Get-LogText $script:Out | Should -Match 'not a delegated scope'
    }

    It 'signs in app-only without scopes when given a certificate' {
        Invoke-CollectorScript 'Get-CopilotInteractions.ps1' @{
            OutputPath = $script:Out; AppId = 'app-1'; CertificateThumbprint = 'AB12'; TenantId = 'tenant-1'
        }

        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter {
            $ClientId -eq 'app-1' -and $CertificateThumbprint -eq 'AB12' -and $TenantId -eq 'tenant-1' -and -not $Scopes
        }
    }

    It 'the feature reference makes no sign-in at all' {
        Invoke-CollectorScript 'Get-CopilotFeatureAvailability.ps1' @{ OutputPath = $script:Out; Environment = 'GCCHigh' }
        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 0 -Exactly
        Should -Invoke Connect-ExchangeOnline -ModuleName M365ReportLibrary -Times 0 -Exactly
    }
}

Describe 'The usage reports' {
    BeforeEach {
        $script:Out = New-TestFolder
        Mock Connect-M365Service -MockWith { }
        Mock Start-Sleep { }
        $global:CuTest.Uris = [System.Collections.Generic.List[string]]::new()
    }
    AfterEach { Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue }

    It '<Function> is a GET to the v1.0 copilot reports path with the period' -ForEach $script:ReportCases {
        $global:CuTest.Csv = New-ReportCsv -MapKey $MapKey
        Mock Invoke-MgGraphRequest { $global:CuTest.Uris.Add($Uri); Set-Content -LiteralPath $OutputFilePath -Value $global:CuTest.Csv }

        Invoke-CollectorScript $Script @{ OutputPath = $script:Out; Period = 'D7' }

        $global:CuTest.Uris.Count | Should -Be 1
        $global:CuTest.Uris[0] | Should -Be "v1.0/copilot/reports/$Function(period='D7')"
        Should -Invoke Invoke-MgGraphRequest -Times 1 -Exactly -ParameterFilter { $Method -eq 'GET' -and $OutputFilePath }
    }

    It 'defaults to D28, accepts ALL, and refuses D30, which is a v1 value' {
        $global:CuTest.Csv = New-ReportCsv -MapKey 'UserCountSummaryMap'
        Mock Invoke-MgGraphRequest { $global:CuTest.Uris.Add($Uri); Set-Content -LiteralPath $OutputFilePath -Value $global:CuTest.Csv }

        Invoke-CollectorScript 'Get-CopilotUserCountSummary.ps1' @{ OutputPath = $script:Out }
        Invoke-CollectorScript 'Get-CopilotUserCountSummary.ps1' @{ OutputPath = $script:Out; Period = 'ALL' }
        { Invoke-CollectorScript 'Get-CopilotUserCountSummary.ps1' @{ OutputPath = $script:Out; Period = 'D30' } } | Should -Throw

        $global:CuTest.Uris[0] | Should -Match "period='D28'"
        $global:CuTest.Uris[1] | Should -Match "period='ALL'"
    }

    It '<Csv> keeps every row of a long CSV stream (the reports are one stream, not pages)' -ForEach $script:ReportCases {
        $global:CuTest.Csv = New-ReportCsv -MapKey $MapKey -Rows 1250
        Mock Invoke-MgGraphRequest { Set-Content -LiteralPath $OutputFilePath -Value $global:CuTest.Csv }

        Invoke-CollectorScript $Script @{ OutputPath = $script:Out }

        # The summary is keyed by RunDate and period and the trend by date, so only the
        # per-user report is guaranteed to keep all 1,250 rows.
        if ($MapKey -eq 'UsageUserDetailMap') {
            @(Import-Csv -LiteralPath (Join-Path $script:Out $Csv)).Count | Should -Be 1250
        }
        else {
            @(Import-Csv -LiteralPath (Join-Path $script:Out $Csv)).Count | Should -BeGreaterThan 0
        }
    }

    It 'stamps every row with the UTC run date and does not repeat a row on a second run the same day' -ForEach $script:ReportCases {
        $global:CuTest.Csv = New-ReportCsv -MapKey $MapKey -Rows 2
        Mock Invoke-MgGraphRequest { Set-Content -LiteralPath $OutputFilePath -Value $global:CuTest.Csv }

        Invoke-CollectorScript $Script @{ OutputPath = $script:Out }
        $first = @(Import-Csv -LiteralPath (Join-Path $script:Out $Csv)).Count
        Invoke-CollectorScript $Script @{ OutputPath = $script:Out }
        $rows = @(Import-Csv -LiteralPath (Join-Path $script:Out $Csv))

        $rows.Count | Should -Be $first
        ($rows.RunDate | Sort-Object -Unique) | Should -Be ([datetime]::UtcNow.ToString('yyyy-MM-dd'))
    }

    It 'a v1 report, which has fewer columns, leaves the v2 columns empty' {
        $global:CuTest.Csv = "Report Refresh Date,User Principal Name,Display Name,Last Activity Date,Report Period`n2026-09-29,a@example.com,A,2026-09-20,28"
        Mock Invoke-MgGraphRequest { Set-Content -LiteralPath $OutputFilePath -Value $global:CuTest.Csv }

        Invoke-CollectorScript 'Get-CopilotUsageUserDetail.ps1' @{ OutputPath = $script:Out }

        $row = Import-Csv -LiteralPath (Join-Path $script:Out 'copilot-usage-user-detail.csv')
        $row.UserPrincipalName | Should -Be 'a@example.com'
        $row.PromptsSubmittedAllApps | Should -BeNullOrEmpty
    }

    It 'keeps hashed names exactly as returned when names are concealed' {
        $global:CuTest.Csv = "Report Refresh Date,User Principal Name,Display Name,Report Period`n2026-09-29,E58EEF6A6BBD3D0293CB306D24A42057,6A4630397AAA1A9C8EED362168DC88A0,7"
        Mock Invoke-MgGraphRequest { Set-Content -LiteralPath $OutputFilePath -Value $global:CuTest.Csv }

        Invoke-CollectorScript 'Get-CopilotUsageUserDetail.ps1' @{ OutputPath = $script:Out }

        (Import-Csv -LiteralPath (Join-Path $script:Out 'copilot-usage-user-detail.csv')).UserPrincipalName | Should -Be 'E58EEF6A6BBD3D0293CB306D24A42057'
    }

    It 'retries a 429, then reads the report' {
        $global:CuTest.Calls = 0
        $global:CuTest.Csv = New-ReportCsv -MapKey 'UserCountTrendMap'
        Mock Invoke-MgGraphRequest {
            $global:CuTest.Calls++
            if ($global:CuTest.Calls -eq 1) { throw 'Response status code does not indicate success: 429 (Too Many Requests).' }
            Set-Content -LiteralPath $OutputFilePath -Value $global:CuTest.Csv
        }

        Invoke-CollectorScript 'Get-CopilotUserCountTrend.ps1' @{ OutputPath = $script:Out } 3>$null

        $global:CuTest.Calls | Should -Be 2
        Should -Invoke Start-Sleep -Times 1
        @(Import-Csv -LiteralPath (Join-Path $script:Out 'copilot-user-count-trend.csv')).Count | Should -BeGreaterThan 0
    }

    It 'logs a persistent 429 as throttling, not as an empty report' {
        Mock Invoke-MgGraphRequest { throw 'Response status code does not indicate success: 429 (Too Many Requests).' }

        Invoke-CollectorScript 'Get-CopilotUserCountTrend.ps1' @{ OutputPath = $script:Out } 3>$null

        Get-LogText $script:Out | Should -Match 'still throttled'
        (Get-Content -LiteralPath (Join-Path $script:Out 'copilot-user-count-trend.csv')).Count | Should -Be 1
    }

    It 'writes the header only and logs the permission when the report is refused' {
        Mock Invoke-MgGraphRequest { throw 'Response status code does not indicate success: 403 (Forbidden).' }

        Invoke-CollectorScript 'Get-CopilotUsageUserDetail.ps1' @{ OutputPath = $script:Out } 3>$null

        (Get-Content -LiteralPath (Join-Path $script:Out 'copilot-usage-user-detail.csv')).Count | Should -Be 1
        Get-LogText $script:Out | Should -Match 'Reports.Read.All'
    }
}

Describe 'The interaction export' {
    BeforeEach {
        $script:Out = New-TestFolder
        Mock Connect-M365Service -MockWith { }
        Mock Start-Sleep { }
        $global:CuTest.Uris = [System.Collections.Generic.List[string]]::new()
    }
    AfterEach { Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue }

    It 'asks for one user per call, 100 per page, with both bounds on createdDateTime' {
        New-UsersCsv -Folder $script:Out -Id 'user-1', 'user-2'
        Mock Invoke-MgGraphRequest { $global:CuTest.Uris.Add([uri]::UnescapeDataString($Uri)); @{ value = @() } }

        Invoke-CollectorScript 'Get-CopilotInteractions.ps1' @{
            OutputPath = $script:Out; StartDate = [datetime]'2026-10-01T00:00:00Z'; EndDate = [datetime]'2026-10-02T00:00:00Z'
        }

        $global:CuTest.Uris.Count | Should -Be 2
        $global:CuTest.Uris[0] | Should -Be 'v1.0/copilot/users/user-1/interactionHistory/getAllEnterpriseInteractions?$top=100&$filter=createdDateTime gt 2026-10-01T00:00:00Z and createdDateTime lt 2026-10-02T00:00:00Z'
        $global:CuTest.Uris[1] | Should -Match '/users/user-2/'
        Should -Invoke Invoke-MgGraphRequest -Times 2 -Exactly -ParameterFilter { $Method -eq 'GET' }
    }

    It 'follows @odata.nextLink to the last page' {
        New-UsersCsv -Folder $script:Out
        Mock Invoke-MgGraphRequest {
            $global:CuTest.Uris.Add($Uri)
            if ($global:CuTest.Uris.Count -eq 1) { @{ value = @(1..100 | ForEach-Object { New-MockInteraction -Id "a-$_" }); '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/copilot/users/user-1/interactionHistory/getAllEnterpriseInteractions?$skiptoken=page2' } }
            else { @{ value = @(1..50 | ForEach-Object { New-MockInteraction -Id "b-$_" }) } }
        }

        Invoke-CollectorScript 'Get-CopilotInteractions.ps1' @{ OutputPath = $script:Out }

        @(Import-Csv -LiteralPath (Join-Path $script:Out 'copilot-interactions.csv')).Count | Should -Be 150
        $global:CuTest.Uris[1] | Should -Match 'skiptoken=page2'
    }

    It 'does not send the session token to a nextLink on another host' {
        New-UsersCsv -Folder $script:Out
        Mock Invoke-MgGraphRequest { @{ value = @(New-MockInteraction); '@odata.nextLink' = 'https://evil.example.com/steal?token=1' } }

        { Invoke-CollectorScript 'Get-CopilotInteractions.ps1' @{ OutputPath = $script:Out } 3>$null } | Should -Throw '*every one of the 1 users*'

        Should -Invoke Invoke-MgGraphRequest -Times 1 -Exactly
        Get-LogText $script:Out | Should -Match 'outside Microsoft Graph'
    }

    It 'stores metadata only: no prompt, response or attachment text reaches the file' {
        New-UsersCsv -Folder $script:Out
        Mock Invoke-MgGraphRequest { @{ value = @(New-MockInteraction -Type 'userPrompt'; New-MockInteraction -Id '2') } }

        Invoke-CollectorScript 'Get-CopilotInteractions.ps1' @{ OutputPath = $script:Out }

        $raw = Get-Content -LiteralPath (Join-Path $script:Out 'copilot-interactions.csv') -Raw
        $raw | Should -Not -Match 'SECRET'
        $row = (Import-Csv -LiteralPath (Join-Path $script:Out 'copilot-interactions.csv'))[0]
        $row.InteractionType | Should -Be 'userPrompt'
        $row.AppClass | Should -Be 'IPM.SkypeTeams.Message.Copilot.Teams'
        $row.ContextTypes | Should -Be 'TeamsMeeting'
        $row.UserId | Should -Be 'user-1'
    }

    It 'resumes per user from that user''s latest CreatedDateTime' {
        New-UsersCsv -Folder $script:Out -Id 'user-1', 'user-2'
        Export-AppendCsv -Path (Join-Path $script:Out 'copilot-interactions.csv') -Column $script:Schema.CopilotInteractions -Rows @(
            [pscustomobject]@{ CreatedDateTime = '2026-10-05T10:00:00Z'; Id = 'x1'; UserId = 'user-1'; SessionId = 's'; RequestId = 'r'; AppClass = 'c'; InteractionType = 'userPrompt'; ConversationType = 'bizchat'; Locale = 'en-us'; ContextCount = 0; ContextTypes = '' })
        Mock Invoke-MgGraphRequest { $global:CuTest.Uris.Add([uri]::UnescapeDataString($Uri)); @{ value = @() } }

        Invoke-CollectorScript 'Get-CopilotInteractions.ps1' @{ OutputPath = $script:Out; EndDate = [datetime]'2026-10-08T00:00:00Z'; LookbackDays = 30 }

        # The stored stamp is that second. ge keeps a sibling interaction on the same
        # second; gt would drop it. The user with no stamp still uses gt from the lookback.
        # https://learn.microsoft.com/microsoft-365-copilot/extensibility/api/ai-services/interaction-export/aiinteractionhistory-getallenterpriseinteractions
        ($global:CuTest.Uris | Where-Object { $_ -match '/users/user-1/' }) | Should -Match 'createdDateTime ge 2026-10-05T10:00:00Z and'
        ($global:CuTest.Uris | Where-Object { $_ -match '/users/user-2/' }) | Should -Match 'createdDateTime gt 2026-09-08T00:00:00Z and'
    }

    It 'does not repeat an interaction already collected' {
        New-UsersCsv -Folder $script:Out
        Mock Invoke-MgGraphRequest { @{ value = @(New-MockInteraction -Id 'same') } }
        Invoke-CollectorScript 'Get-CopilotInteractions.ps1' @{ OutputPath = $script:Out; StartDate = [datetime]'2026-10-01T00:00:00Z'; EndDate = [datetime]'2026-10-07T00:00:00Z' }
        Invoke-CollectorScript 'Get-CopilotInteractions.ps1' @{ OutputPath = $script:Out; StartDate = [datetime]'2026-10-01T00:00:00Z'; EndDate = [datetime]'2026-10-07T00:00:00Z' }
        @(Import-Csv -LiteralPath (Join-Path $script:Out 'copilot-interactions.csv')).Count | Should -Be 1
    }

    It 'skips a user the API refuses, keeps the others, and says so' {
        New-UsersCsv -Folder $script:Out -Id 'user-1', 'user-2'
        Mock Invoke-MgGraphRequest {
            if ($Uri -match '/users/user-1/') { throw 'Response status code does not indicate success: 403 (Forbidden).' }
            @{ value = @(New-MockInteraction) }
        }

        Invoke-CollectorScript 'Get-CopilotInteractions.ps1' @{ OutputPath = $script:Out } 3>$null

        (Import-Csv -LiteralPath (Join-Path $script:Out 'copilot-interactions.csv')).UserId | Should -Be 'user-2'
        Get-LogText $script:Out | Should -Match 'Skipping user user-1'
    }

    It 'does not report success when every user is refused' {
        New-UsersCsv -Folder $script:Out -Id 'user-1', 'user-2'
        Mock Invoke-MgGraphRequest { throw 'Response status code does not indicate success: 403 (Forbidden).' }

        { Invoke-CollectorScript 'Get-CopilotInteractions.ps1' @{ OutputPath = $script:Out } 3>$null } | Should -Throw '*every one of the 2 users*'
    }

    It 'stops on a persistent 429, keeps what it read and says a 429 is not an empty history' {
        New-UsersCsv -Folder $script:Out -Id 'user-1', 'user-2'
        Mock Invoke-MgGraphRequest {
            if ($Uri -match '/users/user-2/') { throw 'Response status code does not indicate success: 429 (Too Many Requests).' }
            @{ value = @(New-MockInteraction) }
        }

        { Invoke-CollectorScript 'Get-CopilotInteractions.ps1' @{ OutputPath = $script:Out } 3>$null } | Should -Throw '*throttling*'

        (Import-Csv -LiteralPath (Join-Path $script:Out 'copilot-interactions.csv')).UserId | Should -Be 'user-1'
        Get-LogText $script:Out | Should -Match 'not an empty history'
    }

    It 'writes the header only and says so when there are no users' {
        Mock Invoke-MgGraphRequest { throw 'must not be called' }

        Invoke-CollectorScript 'Get-CopilotInteractions.ps1' @{ OutputPath = $script:Out } 3>$null

        (Get-Content -LiteralPath (Join-Path $script:Out 'copilot-interactions.csv')).Count | Should -Be 1
        Get-LogText $script:Out | Should -Match 'users.csv holds no users'
    }

    It 'reads the users named by -UserId instead of users.csv' {
        Mock Invoke-MgGraphRequest { $global:CuTest.Uris.Add($Uri); @{ value = @() } }
        Invoke-CollectorScript 'Get-CopilotInteractions.ps1' @{ OutputPath = $script:Out; UserId = @('only-user') }
        $global:CuTest.Uris.Count | Should -Be 1
        $global:CuTest.Uris[0] | Should -Match '/users/only-user/'
    }

    It 'keeps an Unspecified createdDateTime bound on that UTC instant when the machine zone is not UTC' {
        # createdDateTime filters are UTC. Kind Unspecified is not the local zone.
        # https://learn.microsoft.com/microsoft-365-copilot/extensibility/api/ai-services/interaction-export/aiinteractionhistory-getallenterpriseinteractions
        $repo = $script:Root.Replace("'", "''")
        $probe = @"
Set-StrictMode -Version Latest
`$ErrorActionPreference = 'Stop'
. '$repo/reports/copilot-usage/tests/CopilotUsageStubs.ps1'
. '$repo/shared/tests/TenantCmdletStubs.ps1'
function global:Connect-MgGraph { }
function global:Invoke-MgGraphRequest {
    param([string]`$Method, [string]`$Uri, [hashtable]`$Headers, [string]`$OutputFilePath)
    `$global:SeenUri = [uri]::UnescapeDataString(`$Uri)
    @{ value = @() }
}
`$out = Join-Path ([System.IO.Path]::GetTempPath()) ('copilot-tz-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path `$out | Out-Null
Import-Module '$repo/shared/M365ReportLibrary.psm1' -Force
[pscustomobject]@{ RunDate = '2026-10-01'; Id = 'user-1' } | Export-Csv -LiteralPath (Join-Path `$out 'users.csv') -NoTypeInformation
`$start = [datetime]::SpecifyKind([datetime]::new(2026, 10, 1, 0, 0, 0), [DateTimeKind]::Unspecified)
`$end = [datetime]::SpecifyKind([datetime]::new(2026, 10, 2, 0, 0, 0), [DateTimeKind]::Unspecified)
& '$repo/reports/copilot-usage/collectors/Get-CopilotInteractions.ps1' -OutputPath `$out -UserId 'user-1' -StartDate `$start -EndDate `$end
if (`$global:SeenUri -notmatch 'createdDateTime gt 2026-10-01T00:00:00Z and createdDateTime lt 2026-10-02T00:00:00Z') {
    Write-Output `$global:SeenUri
    exit 1
}
exit 0
"@
        $previous = $env:TZ
        try {
            $env:TZ = 'America/New_York'
            & pwsh -NoProfile -Command $probe
            $LASTEXITCODE | Should -Be 0
        }
        finally {
            if ($null -eq $previous) { Remove-Item Env:TZ -ErrorAction SilentlyContinue }
            else { $env:TZ = $previous }
        }
    }

    It 'rejects an empty explicit range' {
        New-UsersCsv -Folder $script:Out
        { Invoke-CollectorScript 'Get-CopilotInteractions.ps1' @{ OutputPath = $script:Out; StartDate = [datetime]'2026-10-08T00:00:00Z'; EndDate = [datetime]'2026-10-01T00:00:00Z' } } | Should -Throw '*range is empty*'
    }
}

Describe 'The audit events' {
    BeforeEach {
        $script:Out = New-TestFolder
        Mock Connect-M365Service -MockWith { }
        Mock Disconnect-ExchangeOnline { }
        Mock Start-Sleep { }
        $global:CuTest.Windows = [System.Collections.Generic.List[object]]::new()
        $script:Range = @{ StartDate = [datetime]'2026-10-05T00:00:00Z'; EndDate = [datetime]'2026-10-06T00:00:00Z' }
    }
    AfterEach { Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue }

    It 'searches the CopilotInteraction operation and counts messages, prompts, resources and plug-ins per record' {
        Mock Search-UnifiedAuditLog { New-MockAuditRecord }
        Invoke-CollectorScript 'Get-CopilotAuditEvents.ps1' ($script:Range + @{ OutputPath = $script:Out })

        Should -Invoke Search-UnifiedAuditLog -ParameterFilter { $Operations.Count -eq 1 -and $Operations[0] -eq 'CopilotInteraction' -and $SessionCommand -eq 'ReturnLargeSet' }
        $row = Import-Csv -LiteralPath (Join-Path $script:Out 'copilot-audit-events.csv')
        $row.MessageCount | Should -Be '3'
        $row.PromptMessageCount | Should -Be '1'
        $row.AccessedResourceCount | Should -Be '1'
        $row.PluginIds | Should -Be 'BingWebSearch'
        $row.AppHost | Should -Be 'Teams'
        $row.AppIdentity | Should -Be 'Copilot.MicrosoftCopilot.BizChat'
        $row.Operation | Should -Be 'CopilotInteraction'
        $row.CreationTime | Should -Be '2026-10-05T12:00:00Z'
    }

    It 'keeps AgentId and AgentName wherever the record places them' {
        Mock Search-UnifiedAuditLog {
            New-MockAuditRecord -Id 'inside'
            New-MockAuditRecord -Id 'top' -AgentAtTop
        }
        Invoke-CollectorScript 'Get-CopilotAuditEvents.ps1' ($script:Range + @{ OutputPath = $script:Out })

        $rows = Import-Csv -LiteralPath (Join-Path $script:Out 'copilot-audit-events.csv')
        ($rows | Where-Object Id -eq 'inside').AgentName | Should -Be 'ReminderBot'
        ($rows | Where-Object Id -eq 'top').AgentName | Should -Be 'SalesAgent'
    }

    It 'writes one row per record and no message text' {
        Mock Search-UnifiedAuditLog { New-MockAuditRecord -Id 'a'; New-MockAuditRecord -Id 'b' }
        Invoke-CollectorScript 'Get-CopilotAuditEvents.ps1' ($script:Range + @{ OutputPath = $script:Out })
        @(Import-Csv -LiteralPath (Join-Path $script:Out 'copilot-audit-events.csv')).Count | Should -Be 2
        $script:Schema.CopilotAuditEvents | Should -Not -Contain 'Prompt'
    }

    It 'starts at the latest CreationTime already exported and does not repeat that record' {
        Mock Search-UnifiedAuditLog { New-MockAuditRecord -Id 'e1' -CreationTime '2026-10-05T12:00:00' }
        Invoke-CollectorScript 'Get-CopilotAuditEvents.ps1' ($script:Range + @{ OutputPath = $script:Out })

        Mock Search-UnifiedAuditLog {
            $global:CuTest.Windows.Add($StartDate.ToUniversalTime())
            New-MockAuditRecord -Id 'e1' -CreationTime '2026-10-05T12:00:00'
            New-MockAuditRecord -Id 'e2' -CreationTime '2026-10-05T13:00:00'
        }
        Invoke-CollectorScript 'Get-CopilotAuditEvents.ps1' @{ OutputPath = $script:Out; EndDate = [datetime]'2026-10-06T00:00:00Z' }

        $global:CuTest.Windows[0] | Should -Be ([datetime]'2026-10-05T12:00:00Z')
        (Import-Csv -LiteralPath (Join-Path $script:Out 'copilot-audit-events.csv')).Id | Sort-Object | Should -Be @('e1', 'e2')
    }

    It 'reaches back -LookbackDays on the first run and splits the range into -WindowHours windows' {
        Mock Search-UnifiedAuditLog { $global:CuTest.Windows.Add($StartDate.ToUniversalTime()) }
        Invoke-CollectorScript 'Get-CopilotAuditEvents.ps1' @{ OutputPath = $script:Out; LookbackDays = 2; WindowHours = 12 }

        $starts = @($global:CuTest.Windows | Sort-Object -Unique)
        $starts.Count | Should -BeGreaterOrEqual 4
        ($starts[1] - $starts[0]).TotalHours | Should -Be 12
    }

    It 'reads the next page in one ReturnLargeSet session while moreRecordsAvailable is true' {
        $global:CuTest.Sessions = [System.Collections.Generic.List[string]]::new()
        Mock Search-UnifiedAuditLog {
            $global:CuTest.Sessions.Add("$SessionId|$SessionCommand|$ResultSize")
            if ($global:CuTest.Sessions.Count -eq 1) { New-MockAuditRecord -Id 'p1' -MoreRecords $true }
            elseif ($global:CuTest.Sessions.Count -eq 2) { New-MockAuditRecord -Id 'p2' -MoreRecords $false }
        }
        Invoke-CollectorScript 'Get-CopilotAuditEvents.ps1' ($script:Range + @{ OutputPath = $script:Out })

        $global:CuTest.Sessions.Count | Should -Be 2
        ($global:CuTest.Sessions | ForEach-Object { ($_ -split '\|')[0] } | Sort-Object -Unique).Count | Should -Be 1
        $global:CuTest.Sessions[0] | Should -Match '\|ReturnLargeSet\|5000$'
        @(Import-Csv -LiteralPath (Join-Path $script:Out 'copilot-audit-events.csv')).Count | Should -Be 2
    }

    It 'writes nothing from a window that reaches the 50,000-record cap, logs an error and stops' {
        Mock Search-UnifiedAuditLog {
            if ($StartDate.ToUniversalTime().Date -eq [datetime]'2026-10-05') { New-MockAuditRecord -Id 'ok' -CreationTime '2026-10-05T05:00:00' }
            else {
                $record = New-MockAuditRecord -Id 'big' -CreationTime '2026-10-06T05:00:00'
                $record | Add-Member -NotePropertyName ResultCount -NotePropertyValue 50000
                $record
            }
        }
        { Invoke-CollectorScript 'Get-CopilotAuditEvents.ps1' @{ OutputPath = $script:Out; StartDate = [datetime]'2026-10-05T00:00:00Z'; EndDate = [datetime]'2026-10-07T00:00:00Z' } 3>$null } |
            Should -Throw '*50,000-record session cap*'

        (Import-Csv -LiteralPath (Join-Path $script:Out 'copilot-audit-events.csv')).Id | Should -Be 'ok'
        Get-LogText $script:Out | Should -Match 'matches 50,000 or more records'
    }

    It 'writes the header only and logs the reason when the audit log is refused' {
        Mock Search-UnifiedAuditLog { throw 'The role is missing' }
        Invoke-CollectorScript 'Get-CopilotAuditEvents.ps1' ($script:Range + @{ OutputPath = $script:Out }) 3>$null
        (Get-Content -LiteralPath (Join-Path $script:Out 'copilot-audit-events.csv')).Count | Should -Be 1
        Get-LogText $script:Out | Should -Match 'View-Only Audit Logs'
    }

    It 'rejects an empty explicit range and a lookback past the 180-day retention' {
        { Invoke-CollectorScript 'Get-CopilotAuditEvents.ps1' @{ OutputPath = $script:Out; StartDate = [datetime]'2026-10-06T00:00:00Z'; EndDate = [datetime]'2026-10-05T00:00:00Z' } } | Should -Throw '*requested range is empty*'
        { Invoke-CollectorScript 'Get-CopilotAuditEvents.ps1' @{ OutputPath = $script:Out; LookbackDays = 181 } } | Should -Throw
    }
}

Describe 'A source Microsoft documents as NotAvailable writes the header only' {
    BeforeEach {
        $script:Out = New-TestFolder
        Mock Connect-MgGraph -ModuleName M365ReportLibrary -MockWith { }
        Mock Connect-ExchangeOnline -ModuleName M365ReportLibrary -MockWith { }
        Mock Invoke-MgGraphRequest { throw 'Invoke-MgGraphRequest must not be called' }
        Mock Search-UnifiedAuditLog { throw 'Search-UnifiedAuditLog must not be called' }
    }
    AfterEach { Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue }

    It '<Csv> in GCC High, as the contract says, signs in to nothing and logs why' -ForEach $script:ReportCases {
        Invoke-CollectorScript $Script @{ OutputPath = $script:Out; Environment = 'GCCHigh' } 3>$null

        $path = Join-Path $script:Out $Csv
        (Get-Content -LiteralPath $path).Count | Should -Be 1
        Get-HeaderText $path | Should -Be (Get-ExpectedColumn $MapKey)
        Get-HeaderText $path | Should -Be (Get-HeaderText (Join-Path $script:Samples "gcchigh/$Csv"))
        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 0 -Exactly
        Should -Invoke Invoke-MgGraphRequest -Times 0 -Exactly
        Get-LogText $script:Out | Should -Match 'documented as unavailable in GCCHigh'
    }

    It '<Script> skips in every cloud when the schema says NotAvailable' -ForEach @(
        @{ Script = 'Get-CopilotUsageUserDetail.ps1'; Csv = 'copilot-usage-user-detail.csv' }
        @{ Script = 'Get-CopilotUserCountSummary.ps1'; Csv = 'copilot-user-count-summary.csv' }
        @{ Script = 'Get-CopilotUserCountTrend.ps1'; Csv = 'copilot-user-count-trend.csv' }
        @{ Script = 'Get-CopilotAuditEvents.ps1'; Csv = 'copilot-audit-events.csv' }
        @{ Script = 'Get-CopilotInteractions.ps1'; Csv = 'copilot-interactions.csv' }
        @{ Script = 'Get-CopilotFeatureAvailability.ps1'; Csv = 'copilot-feature-availability.csv' }
    ) {
        # Run a copy of the collectors against a schema that marks every source NotAvailable.
        $copy = Join-Path $script:Out 'collectors'
        Copy-Item -LiteralPath $script:Collectors -Destination $copy -Recurse
        $schemaPath = Join-Path $copy 'CopilotUsageSchema.psd1'
        $text = Get-Content -LiteralPath $schemaPath -Raw
        Set-Content -LiteralPath $schemaPath -Value ($text -replace "Status = '(Available|Unverified)'", "Status = 'NotAvailable'") -Encoding utf8
        New-UsersCsv -Folder $script:Out

        # The copy sits three levels from the shared module, so point its import at the repo.
        $scriptPath = Join-Path $copy $Script
        (Get-Content -LiteralPath $scriptPath -Raw) -replace [regex]::Escape("(Join-Path `$PSScriptRoot '../../../shared/M365ReportLibrary.psm1')"), ("'" + (Join-Path $script:Root 'shared/M365ReportLibrary.psm1') + "'") |
            Set-Content -LiteralPath $scriptPath -Encoding utf8

        & $scriptPath -OutputPath $script:Out 3>$null

        $path = Join-Path $script:Out $Csv
        (Get-Content -LiteralPath $path).Count | Should -Be 1
        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 0 -Exactly
        Should -Invoke Connect-ExchangeOnline -ModuleName M365ReportLibrary -Times 0 -Exactly
        Get-LogText $script:Out | Should -Match 'documented as unavailable'
    }
}

Describe 'An UNVERIFIED source is attempted with a warning' {
    BeforeEach {
        $script:Out = New-TestFolder
        Mock Connect-M365Service -MockWith { }
        Mock Disconnect-ExchangeOnline { }
    }
    AfterEach { Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue }

    It 'asks <Environment> for the audit records, warns, and still writes what comes back' -ForEach @(
        @{ Environment = 'GCC' }, @{ Environment = 'GCCHigh' }
    ) {
        Mock Search-UnifiedAuditLog { New-MockAuditRecord }

        Invoke-CollectorScript 'Get-CopilotAuditEvents.ps1' @{ OutputPath = $script:Out; Environment = $Environment; LookbackDays = 1 } 3>$null

        Should -Invoke Search-UnifiedAuditLog -Times 1
        @(Import-Csv -LiteralPath (Join-Path $script:Out 'copilot-audit-events.csv')).Count | Should -BeGreaterThan 0
        Get-LogText $script:Out | Should -Match 'UNVERIFIED'
    }

    It 'records a refusal in GCC High without reporting success of the data' {
        Mock Search-UnifiedAuditLog { throw 'The term CopilotInteraction is not a recognised operation in this cloud.' }

        Invoke-CollectorScript 'Get-CopilotAuditEvents.ps1' @{ OutputPath = $script:Out; Environment = 'GCCHigh'; LookbackDays = 1 } 3>$null

        (Get-Content -LiteralPath (Join-Path $script:Out 'copilot-audit-events.csv')).Count | Should -Be 1
        Get-LogText $script:Out | Should -Match 'not a recognised operation'
    }

    It 'does not warn about an Available source in Commercial' {
        Mock Search-UnifiedAuditLog { New-MockAuditRecord }
        Invoke-CollectorScript 'Get-CopilotAuditEvents.ps1' @{ OutputPath = $script:Out; LookbackDays = 1 }
        Get-LogText $script:Out | Should -Not -Match 'UNVERIFIED'
    }
}

Describe 'The feature reference is a dated state snapshot' {
    BeforeEach { $script:Out = New-TestFolder }
    AfterEach { Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue }

    It 'stamps the UTC run date and the date the page was read, and appends one block per run date' {
        Invoke-CollectorScript 'Get-CopilotFeatureAvailability.ps1' @{ OutputPath = $script:Out }
        Invoke-CollectorScript 'Get-CopilotFeatureAvailability.ps1' @{ OutputPath = $script:Out }

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:Out 'copilot-feature-availability.csv'))
        $rows.Count | Should -Be $script:Schema.FeatureRows.Count
        ($rows.RunDate | Sort-Object -Unique) | Should -Be ([datetime]::UtcNow.ToString('yyyy-MM-dd'))
        ($rows.PageReadDate | Sort-Object -Unique) | Should -Be '2026-10-10'
    }

    It 'records what the contract says about GCC High: Teams and SharePoint Copilot are not currently available' {
        Invoke-CollectorScript 'Get-CopilotFeatureAvailability.ps1' @{ OutputPath = $script:Out }
        $rows = Import-Csv -LiteralPath (Join-Path $script:Out 'copilot-feature-availability.csv')
        ($rows | Where-Object Feature -like 'Copilot in Teams*').GCCHigh | Should -Be 'NotCurrentlyAvailable'
        ($rows | Where-Object Feature -eq 'Copilot in SharePoint').GCCHigh | Should -Be 'NotCurrentlyAvailable'
        ($rows | Where-Object Feature -like 'Copilot in Teams*').GCC | Should -Be 'Yes'
        ($rows | Where-Object Feature -eq 'Copilot in Outlook').GCCHigh | Should -Be 'Yes'
    }

    It 'uses only the four statuses' {
        $script:Schema.FeatureRows | ForEach-Object { $_.Commercial, $_.GCC, $_.GCCHigh } | Sort-Object -Unique |
            ForEach-Object { $_ | Should -BeIn @('Yes', 'NotCurrentlyAvailable', 'Limited', 'NotStated') }
    }
}

Describe 'Run-All.ps1' {
    BeforeEach {
        $script:Out = New-TestFolder
        Mock Connect-MgGraph -ModuleName M365ReportLibrary -MockWith { }
        Mock Connect-ExchangeOnline -ModuleName M365ReportLibrary -MockWith { }
        Mock Disconnect-ExchangeOnline { }
        Mock Get-MgUser -ModuleName M365ReportLibrary -MockWith {
            [pscustomobject]@{ Id = 'user-1'; DisplayName = 'Avery Abara'; UserPrincipalName = 'avery.abara@example.com'; Mail = $null; UserType = 'Member'; AccountEnabled = $true; CreatedDateTime = [datetime]'2025-01-01'; Department = ''; JobTitle = ''; City = ''; Country = ''; Manager = $null }
        }
        $global:CuTest.Csv = New-ReportCsv -MapKey 'UserCountSummaryMap'
        Mock Invoke-MgGraphRequest {
            if ($OutputFilePath) { Set-Content -LiteralPath $OutputFilePath -Value $global:CuTest.Csv } else { @{ value = @(New-MockInteraction) } }
        }
        Mock Search-UnifiedAuditLog { New-MockAuditRecord }
        Mock Start-Sleep { }
    }
    AfterEach { Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue }

    It 'runs every collector into one folder, users first so the interaction export has someone to read' {
        Invoke-CollectorScript 'Run-All.ps1' @{ OutputPath = $script:Out; StartDate = [datetime]'2026-10-05T00:00:00Z'; EndDate = [datetime]'2026-10-06T00:00:00Z' } 3>$null

        foreach ($csv in 'users.csv', 'copilot-usage-user-detail.csv', 'copilot-user-count-summary.csv', 'copilot-user-count-trend.csv',
            'copilot-audit-events.csv', 'copilot-feature-availability.csv', 'copilot-interactions.csv') {
            Test-Path -LiteralPath (Join-Path $script:Out $csv) | Should -BeTrue -Because $csv
        }
        @(Import-Csv -LiteralPath (Join-Path $script:Out 'copilot-interactions.csv')).Count | Should -BeGreaterThan 0
    }

    It '-SkipInteractions leaves out the most sensitive source' {
        Invoke-CollectorScript 'Run-All.ps1' @{ OutputPath = $script:Out; SkipInteractions = $true; StartDate = [datetime]'2026-10-05T00:00:00Z'; EndDate = [datetime]'2026-10-06T00:00:00Z' } 3>$null
        Test-Path -LiteralPath (Join-Path $script:Out 'copilot-interactions.csv') | Should -BeFalse
    }

    It 'continues past a failing collector and then fails the run' {
        Mock Search-UnifiedAuditLog { throw 'boom' }
        Mock Invoke-MgGraphRequest {
            if ($OutputFilePath) { Set-Content -LiteralPath $OutputFilePath -Value $global:CuTest.Csv } else { throw 'Response status code does not indicate success: 403 (Forbidden).' }
        }

        { Invoke-CollectorScript 'Run-All.ps1' @{ OutputPath = $script:Out; StartDate = [datetime]'2026-10-05T00:00:00Z'; EndDate = [datetime]'2026-10-06T00:00:00Z' } 3>$null } |
            Should -Throw '*collector(s) stopped with an error*'

        Test-Path -LiteralPath (Join-Path $script:Out 'copilot-feature-availability.csv') | Should -BeTrue
    }
}
