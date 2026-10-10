#Requires -Version 7.0

BeforeAll {
    $global:AuTest = @{}
    $script:Root = Join-Path $PSScriptRoot '../../..' | Resolve-Path | Select-Object -ExpandProperty Path
    $script:Collectors = Join-Path $script:Root 'reports/unified-audit-log/collectors'

    . (Join-Path $script:Root 'shared/tests/TenantCmdletStubs.ps1')
    . (Join-Path $PSScriptRoot 'AuditStubs.ps1')
    Import-Module (Join-Path $script:Root 'shared/M365ReportLibrary.psm1') -Force
    . (Join-Path $script:Collectors 'UnifiedAuditLogHelpers.ps1')

    $script:Schema = Import-PowerShellDataFile -LiteralPath (Join-Path $script:Collectors 'UnifiedAuditLogSchema.psd1')
    $global:AuTest.Org = '11111111-1111-1111-1111-111111111111'
    $script:Org = $global:AuTest.Org
    $script:TenantGuid = '33333333-3333-3333-3333-333333333333'
    $script:PublisherGuid = '44444444-4444-4444-4444-444444444444'
    $script:TokenText = 'tok-SECRET-0123456789'

    function New-TestFolder {
        $path = Join-Path ([System.IO.Path]::GetTempPath()) ('audit-' + [guid]::NewGuid().ToString('N'))
        New-Item -Path $path -ItemType Directory -Force | Out-Null
        return $path
    }

    function Get-HeaderText { param([string]$Path) return ((Get-CsvHeaderColumn -Path $Path) -join ',') }

    function Get-LogText { param([string]$Folder) return (Get-Content -LiteralPath (Join-Path $Folder 'run.log') -Raw) }

    function Invoke-CollectorScript {
        param([Parameter(Mandatory)][string]$Name, [hashtable]$Arguments = @{})
        & (Join-Path $script:Collectors $Name) @Arguments
    }

    function Merge-Arguments {
        param([hashtable]$Base, [hashtable]$Extra)
        $merged = $Base.Clone()
        foreach ($key in $Extra.Keys) { $merged[$key] = $Extra[$key] }
        return $merged
    }

    function New-Token { ConvertTo-SecureString $script:TokenText -AsPlainText -Force }

    function Get-Stamp { param([datetime]$At) $At.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ') }

    # The common schema's property names, from
    # https://learn.microsoft.com/office/office-365-management-api/office-365-management-activity-api-schema
    function New-AuditJson {
        param([string]$Id, [datetime]$At, [string]$Operation = 'FileAccessed', [string]$User = 'avery.abara@example.com')
        '{{"Id":"{0}","CreationTime":"{1:yyyy-MM-ddTHH:mm:ss}","Operation":"{2}","OrganizationId":"{3}","RecordType":6,"ResultStatus":"","UserType":0,"Workload":"SharePoint","UserId":"{4}","ClientIP":"203.0.113.1","ObjectId":"https://example.com/doc"}}' -f $Id, $At, $Operation, $global:AuTest.Org, $User
    }

    # Search-UnifiedAuditLog output: RecordType, CreationDate, UserIds, Operations, AuditData, Identity.
    # https://learn.microsoft.com/powershell/module/exchangepowershell/search-unifiedauditlog
    function New-SearchRecord {
        param([string]$Id = ([guid]::NewGuid().ToString()), [datetime]$At = [datetime]::UtcNow, [string]$Operation = 'FileAccessed')
        [pscustomobject]@{
            RecordType   = 'SharePointFileOperation'
            CreationDate = $At
            UserIds      = 'avery.abara@example.com'
            Operations   = $Operation
            AuditData    = New-AuditJson -Id $Id -At $At -Operation $Operation
            Identity     = $Id
        }
    }

    function New-ExistingCsv {
        param([string]$Folder, [string]$Name, [string]$Key, [hashtable]$Values)
        $row = [ordered]@{}
        foreach ($column in $script:Schema[$Key]) { $row[$column] = if ($Values.ContainsKey($column)) { $Values[$column] } else { '' } }
        Export-AppendCsv -Path (Join-Path $Folder $Name) -Rows @([pscustomobject]$row) -Column $script:Schema[$Key]
    }

    # auditLogRecord: https://learn.microsoft.com/graph/api/security-auditlogquery-list-records
    function New-GraphRecord {
        param([string]$Id, [datetime]$At = [datetime]::UtcNow)
        @{
            id = $Id; createdDateTime = (Get-Stamp $At); auditLogRecordType = 'sharePointFileOperation'; operation = 'FileAccessed'
            organizationId = $global:AuTest.Org; userType = 'regular'; userId = 'avery.abara@example.com'; service = 'SharePoint'
            objectId = 'https://example.com/doc'; userPrincipalName = 'avery.abara@example.com'; clientIp = '203.0.113.1'
            auditData = @{ Id = $Id; ResultStatus = 'Succeeded' }
        }
    }

    # Management Activity API shapes: https://learn.microsoft.com/office/office-365-management-api/office-365-management-activity-api-reference
    function New-WebResponse {
        param($Body, [hashtable]$Headers = @{})
        [pscustomobject]@{ Content = (ConvertTo-Json -InputObject $Body -Depth 10 -Compress); Headers = $Headers }
    }
}

AfterAll {
    Remove-Variable -Name AuTest -Scope Global -ErrorAction SilentlyContinue
    Remove-Module M365ReportLibrary -Force -ErrorAction SilentlyContinue
}

Describe 'Cloud availability' {
    BeforeEach {
        $script:Out = New-TestFolder
        $global:AuTest.Posts = [System.Collections.Generic.List[string]]::new()
        Mock Connect-MgGraph -ModuleName M365ReportLibrary -MockWith { }
        Mock Connect-ExchangeOnline -ModuleName M365ReportLibrary -MockWith { }
        Mock Connect-IPPSSession -ModuleName M365ReportLibrary -MockWith { }
        Mock Disconnect-ExchangeOnline { }
        Mock Search-UnifiedAuditLog { }
        Mock Get-AdminAuditLogConfig { [pscustomobject]@{ UnifiedAuditLogIngestionEnabled = $true } }
        Mock Get-UnifiedAuditLogRetentionPolicy { }
        Mock Invoke-MgGraphRequest { }
        Mock Invoke-WebRequest { New-WebResponse @() }
        Mock Start-Sleep { }
    }
    AfterEach { Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue }

    It 'writes a header only for the Graph audit records in GCC High and sends no request' {
        Invoke-CollectorScript 'Get-AuditGraphRecords.ps1' @{ OutputPath = $script:Out; Environment = 'GCCHigh' }

        $path = Join-Path $script:Out 'audit-graph-records.csv'
        Get-HeaderText $path | Should -Be ($script:Schema.AuditGraphRecords -join ',')
        @(Import-Csv -LiteralPath $path).Count | Should -Be 0
        Get-LogText $script:Out | Should -Match 'documented as unavailable in GCCHigh'
        Should -Invoke Invoke-MgGraphRequest -Times 0 -Exactly
        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 0 -Exactly
    }

    It 'attempts <Script> in <Cloud> and logs a warning that availability is UNVERIFIED' -ForEach @(
        @{ Script = 'Get-AuditSearchCmdlet.ps1'; Cloud = 'GCC'; Cmdlet = 'Search-UnifiedAuditLog'; Extra = @{ LookbackDays = 1; SliceMinutes = 1440 }; Calls = 4 }
        @{ Script = 'Get-AuditSearchCmdlet.ps1'; Cloud = 'GCCHigh'; Cmdlet = 'Search-UnifiedAuditLog'; Extra = @{ LookbackDays = 1; SliceMinutes = 1440 }; Calls = 4 }
        @{ Script = 'Get-AuditIngestion.ps1'; Cloud = 'GCC'; Cmdlet = 'Get-AdminAuditLogConfig'; Extra = @{}; Calls = 1 }
        @{ Script = 'Get-AuditIngestion.ps1'; Cloud = 'GCCHigh'; Cmdlet = 'Get-AdminAuditLogConfig'; Extra = @{}; Calls = 1 }
        @{ Script = 'Get-AuditRetentionPolicies.ps1'; Cloud = 'GCC'; Cmdlet = 'Get-UnifiedAuditLogRetentionPolicy'; Extra = @{}; Calls = 1 }
        @{ Script = 'Get-AuditRetentionPolicies.ps1'; Cloud = 'GCCHigh'; Cmdlet = 'Get-UnifiedAuditLogRetentionPolicy'; Extra = @{}; Calls = 1 }
    ) {
        Invoke-CollectorScript $Script (@{ OutputPath = $script:Out; Environment = $Cloud; SkipConnect = $true } + $Extra)
        Get-LogText $script:Out | Should -Match 'UNVERIFIED'
        Should -Invoke $Cmdlet -Times $Calls -Exactly
    }

    It 'attempts the activity feed in <Cloud> without a warning about availability' -ForEach @(
        @{ Cloud = 'Commercial' }, @{ Cloud = 'GCC' }, @{ Cloud = 'GCCHigh' }
    ) {
        Invoke-CollectorScript 'Get-AuditActivityFeed.ps1' @{
            OutputPath = $script:Out; Environment = $Cloud; TenantId = $script:TenantGuid
            PublisherIdentifier = $script:PublisherGuid; AccessToken = (New-Token); ContentType = 'Audit.General'; LookbackDays = 1
        }
        Should -Invoke Invoke-WebRequest -Times 1 -ParameterFilter { $Uri -like '*subscriptions/list*' }
        (Get-Content -LiteralPath (Join-Path $script:Out 'run.log') -Raw) | Should -Not -Match 'UNVERIFIED'
    }
}

Describe 'Connection endpoint per -Environment' {
    BeforeEach {
        $script:Out = New-TestFolder
        Mock Connect-MgGraph -ModuleName M365ReportLibrary -MockWith { }
        Mock Connect-ExchangeOnline -ModuleName M365ReportLibrary -MockWith { }
        Mock Connect-IPPSSession -ModuleName M365ReportLibrary -MockWith { }
        Mock Disconnect-ExchangeOnline { }
        Mock Search-UnifiedAuditLog { }
        Mock Get-AdminAuditLogConfig { [pscustomobject]@{ UnifiedAuditLogIngestionEnabled = $true } }
        Mock Get-UnifiedAuditLogRetentionPolicy { }
        Mock Invoke-MgGraphRequest { }
        Mock Start-Sleep { }
    }
    AfterEach { Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue }

    It 'signs in to Exchange Online with <Name> for <Cloud> (<Script>)' -ForEach @(
        @{ Script = 'Get-AuditSearchCmdlet.ps1'; Cloud = 'Commercial'; Name = 'O365Default' }
        @{ Script = 'Get-AuditSearchCmdlet.ps1'; Cloud = 'GCC'; Name = 'O365Default' }
        @{ Script = 'Get-AuditSearchCmdlet.ps1'; Cloud = 'GCCHigh'; Name = 'O365USGovGCCHigh' }
        @{ Script = 'Get-AuditIngestion.ps1'; Cloud = 'Commercial'; Name = 'O365Default' }
        @{ Script = 'Get-AuditIngestion.ps1'; Cloud = 'GCCHigh'; Name = 'O365USGovGCCHigh' }
    ) {
        Invoke-CollectorScript $Script @{ OutputPath = $script:Out; Environment = $Cloud }
        Should -Invoke Connect-ExchangeOnline -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter { $ExchangeEnvironmentName -eq $Name }
        Should -Invoke Disconnect-ExchangeOnline -Times 1 -Exactly
    }

    It 'signs in to Security & Compliance PowerShell with the default endpoint for <Cloud>' -ForEach @(
        @{ Cloud = 'Commercial' }, @{ Cloud = 'GCC' }
    ) {
        Invoke-CollectorScript 'Get-AuditRetentionPolicies.ps1' @{ OutputPath = $script:Out; Environment = $Cloud }
        Should -Invoke Connect-IPPSSession -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter { -not $ConnectionUri -and -not $AzureADAuthorizationEndpointUri }
    }

    It 'signs in to Security & Compliance PowerShell with the GCC High URIs' {
        Invoke-CollectorScript 'Get-AuditRetentionPolicies.ps1' @{ OutputPath = $script:Out; Environment = 'GCCHigh' }
        Should -Invoke Connect-IPPSSession -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter {
            $ConnectionUri -eq 'https://ps.compliance.protection.office365.us/powershell-liveid/' -and
            $AzureADAuthorizationEndpointUri -eq 'https://login.microsoftonline.us/organizations'
        }
    }

    It 'signs in to Graph with the Global environment and the audit query scopes for <Cloud>' -ForEach @(
        @{ Cloud = 'Commercial' }, @{ Cloud = 'GCC' }
    ) {
        # A finished query, so the assertion is the sign-in and not an empty response.
        Mock Invoke-MgGraphRequest {
            if ($Method -eq 'POST') { return @{ id = 'q1'; status = 'notStarted' } }
            if ($Uri -like '*/records') { return @{ value = @() } }
            return @{ id = 'q1'; status = 'succeeded'; isRecordCountLimitExceeded = $false }
        }
        Invoke-CollectorScript 'Get-AuditGraphRecords.ps1' @{ OutputPath = $script:Out; Environment = $Cloud; LookbackDays = 1 }
        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter {
            $Environment -eq 'Global' -and $Scopes -contains 'AuditLogsQuery.Read.All'
        }
    }

    It 'does not sign in when asked to reuse a session' {
        Invoke-CollectorScript 'Get-AuditIngestion.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        Should -Invoke Connect-ExchangeOnline -ModuleName M365ReportLibrary -Times 0 -Exactly
        Should -Invoke Disconnect-ExchangeOnline -Times 0 -Exactly
    }
}

Describe 'Search-UnifiedAuditLog (source 1)' {
    BeforeEach {
        $script:Out = New-TestFolder
        $global:AuTest.Calls = [System.Collections.Generic.List[object]]::new()
        Mock Disconnect-ExchangeOnline { }
        Mock Start-Sleep { }
    }
    AfterEach { Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue }

    It 'pages one slice with one SessionId and ReturnLargeSet until a call returns nothing' {
        Mock Search-UnifiedAuditLog {
            $global:AuTest.Calls.Add([pscustomobject]@{ Session = $SessionId; Command = $SessionCommand; Size = $ResultSize; Start = $StartDate; End = $EndDate })
            switch ($global:AuTest.Calls.Count) {
                1 { New-SearchRecord -Id 'r1'; New-SearchRecord -Id 'r2' }
                2 { New-SearchRecord -Id 'r3' }
                default { }
            }
        }
        Invoke-CollectorScript 'Get-AuditSearchCmdlet.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; LookbackDays = 1; SliceMinutes = 1440 }

        $global:AuTest.Calls.Count | Should -Be 3
        @($global:AuTest.Calls.Session | Sort-Object -Unique).Count | Should -Be 1
        @($global:AuTest.Calls.Command | Sort-Object -Unique) | Should -Be @('ReturnLargeSet')
        @($global:AuTest.Calls.Size | Sort-Object -Unique) | Should -Be @(5000)
        $rows = @(Import-Csv -LiteralPath (Join-Path $script:Out 'audit-search-cmdlet.csv'))
        ($rows.RecordId | Sort-Object) | Should -Be @('r1', 'r2', 'r3')
        $rows[0].RecordType | Should -Be 'SharePointFileOperation'
        $rows[0].Workload | Should -Be 'SharePoint'
        ($rows[0].AuditData | ConvertFrom-Json).Id | Should -Be $rows[0].RecordId
        Get-HeaderText (Join-Path $script:Out 'audit-search-cmdlet.csv') | Should -Be ($script:Schema.AuditSearchCmdlet -join ',')
    }

    It 'stops a session when the record says no more are available' {
        Mock Search-UnifiedAuditLog {
            $global:AuTest.Calls.Add($SessionId)
            $record = New-SearchRecord -Id 'only'
            $record | Add-Member -NotePropertyName AuditSearchRequestMetadata -NotePropertyValue ([pscustomobject]@{ moreRecordsAvailable = $false })
            $record
        }
        Invoke-CollectorScript 'Get-AuditSearchCmdlet.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; LookbackDays = 1; SliceMinutes = 1440 }
        $global:AuTest.Calls.Count | Should -Be 1
    }

    It 'continues a session while moreRecordsAvailable is true even if a page is short' {
        Mock Search-UnifiedAuditLog {
            $global:AuTest.Calls.Add($SessionId)
            if ($global:AuTest.Calls.Count -gt 2) { return }
            $record = New-SearchRecord -Id ('p' + $global:AuTest.Calls.Count)
            $record | Add-Member -NotePropertyName AuditSearchRequestMetadata -NotePropertyValue ([pscustomobject]@{ moreRecordsAvailable = ($global:AuTest.Calls.Count -lt 2) })
            $record
        }
        Invoke-CollectorScript 'Get-AuditSearchCmdlet.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; LookbackDays = 1; SliceMinutes = 1440 }
        $global:AuTest.Calls.Count | Should -Be 2
    }

    It 'reads a range that reaches 50,000 records as two halves' {
        $one = New-SearchRecord -Id 'bulk'
        $global:AuTest.Bulk = @($one) * 5000
        Mock Search-UnifiedAuditLog {
            $span = ($EndDate - $StartDate).TotalMinutes
            $served = if ($global:AuTest.Served.ContainsKey($SessionId)) { $global:AuTest.Served[$SessionId] } else { 0 }
            if ($span -gt 30) {
                if ($served -ge 50000) { return }
                $global:AuTest.Served[$SessionId] = $served + 5000
                return $global:AuTest.Bulk
            }
            if ($served -ge 1) { return }
            $global:AuTest.Served[$SessionId] = 1
            New-SearchRecord -Id ('half-{0:HHmm}' -f $StartDate) -At $StartDate
        }
        $global:AuTest.Served = @{}
        $start = [datetime]::new(2026, 10, 8, 9, 0, 0, [DateTimeKind]::Utc)

        $records = @(Get-AuditSearchSlice -Start $start -End $start.AddMinutes(60) -OutputPath $script:Out)

        $records.Count | Should -Be 2
        ($records.Identity | Sort-Object) | Should -Be @('half-0900', 'half-0930')
        Get-LogText $script:Out | Should -Match 'reached the 50,000-record limit'
    }

    It 'returns an incomplete range with a warning when it still reaches 50,000 at the smallest size' {
        $global:AuTest.Bulk = @(New-SearchRecord -Id 'bulk') * 5000
        $global:AuTest.Served = @{}
        Mock Search-UnifiedAuditLog {
            $served = if ($global:AuTest.Served.ContainsKey($SessionId)) { $global:AuTest.Served[$SessionId] } else { 0 }
            if ($served -ge 50000) { return }
            $global:AuTest.Served[$SessionId] = $served + 5000
            $global:AuTest.Bulk
        }
        $start = [datetime]::new(2026, 10, 8, 9, 0, 0, [DateTimeKind]::Utc)
        { Get-AuditSearchSlice -Start $start -End $start.AddMinutes(2) -OutputPath $script:Out } | Should -Throw '*50,000*'
        Get-LogText $script:Out | Should -Match 'Its rows are incomplete'
    }

    It 'does not keep a short page whose ResultCount is 50,000 when moreRecordsAvailable is false' {
        $global:AuTest.Served = @{}
        Mock Search-UnifiedAuditLog {
            $span = ($EndDate - $StartDate).TotalMinutes
            if ($span -gt 30) {
                $record = New-SearchRecord -Id 'capped'
                $record | Add-Member -NotePropertyName ResultCount -NotePropertyValue 50000
                $record | Add-Member -NotePropertyName AuditSearchRequestMetadata -NotePropertyValue ([pscustomobject]@{ moreRecordsAvailable = $false })
                return $record
            }
            if ($global:AuTest.Served.ContainsKey($SessionId)) { return }
            $global:AuTest.Served[$SessionId] = $true
            New-SearchRecord -Id ('half-{0:HHmm}' -f $StartDate) -At $StartDate
        }
        $start = [datetime]::new(2026, 10, 8, 9, 0, 0, [DateTimeKind]::Utc)
        $records = @(Get-AuditSearchSlice -Start $start -End $start.AddMinutes(60) -OutputPath $script:Out)
        $records.Identity | Should -Not -Contain 'capped'
        ($records.Identity | Sort-Object) | Should -Be @('half-0900', 'half-0930')
    }

    It 'does not write a two-minute slice that is still at the 50,000-record cap' {
        Mock Search-UnifiedAuditLog {
            $record = New-SearchRecord -Id 'capped'
            $record | Add-Member -NotePropertyName ResultCount -NotePropertyValue 50000
            $record | Add-Member -NotePropertyName AuditSearchRequestMetadata -NotePropertyValue ([pscustomobject]@{ moreRecordsAvailable = $false })
            $record
        }
        { Invoke-CollectorScript 'Get-AuditSearchCmdlet.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; LookbackDays = 1; SliceMinutes = 2 } } |
            Should -Throw '*50,000*'
        @(Import-Csv -LiteralPath (Join-Path $script:Out 'audit-search-cmdlet.csv')).Count | Should -Be 0
    }

    It 'retries the first empty page of a session and then reads the records' {
        $global:AuTest.Served = @{}
        Mock Search-UnifiedAuditLog {
            $n = if ($global:AuTest.Served.ContainsKey($SessionId)) { $global:AuTest.Served[$SessionId] } else { 0 }
            $n++
            $global:AuTest.Served[$SessionId] = $n
            if ($n -eq 1) { return }
            if ($n -eq 2) { return New-SearchRecord -Id 'late' }
        }
        Invoke-CollectorScript 'Get-AuditSearchCmdlet.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; LookbackDays = 1; SliceMinutes = 1440 }
        @(Import-Csv -LiteralPath (Join-Path $script:Out 'audit-search-cmdlet.csv')).RecordId | Should -Contain 'late'
    }

    It 'logs that results can be missing without -HighCompleteness and does not pass the switch' {
        Mock Search-UnifiedAuditLog { $global:AuTest.Calls.Add($PSBoundParameters.ContainsKey('HighCompleteness')) }
        Invoke-CollectorScript 'Get-AuditSearchCmdlet.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; LookbackDays = 1; SliceMinutes = 1440 }
        $global:AuTest.Calls | Should -Not -Contain $true
        Get-LogText $script:Out | Should -Match 'without -HighCompleteness'
        Get-LogText $script:Out | Should -Match 'results can be missing'
    }

    It 'accepts a lookback of 365 days so a first run can cover one-year retention' {
        Mock Search-UnifiedAuditLog { $global:AuTest.Calls.Add($StartDate); throw 'stop after the first window' }
        { Invoke-CollectorScript 'Get-AuditSearchCmdlet.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; LookbackDays = 365; SliceMinutes = 1440 } } |
            Should -Throw '*stop after the first window*'
        ([datetime]::UtcNow - $global:AuTest.Calls[0]).TotalDays | Should -BeGreaterThan 364
        ([datetime]::UtcNow - $global:AuTest.Calls[0]).TotalDays | Should -BeLessThan 366
    }

    It 'resumes from the latest CreationTime already in the file' {
        $watermark = [datetime]::UtcNow.AddHours(-3)
        New-ExistingCsv $script:Out 'audit-search-cmdlet.csv' 'AuditSearchCmdlet' @{ CreationTime = (Get-Stamp $watermark); RecordId = 'old' }
        $global:AuTest.Seen = [System.Collections.Generic.HashSet[string]]::new()
        Mock Search-UnifiedAuditLog {
            if ($global:AuTest.Seen.Add([string]$SessionId)) { $global:AuTest.Calls.Add($StartDate) }
        }

        Invoke-CollectorScript 'Get-AuditSearchCmdlet.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; SliceMinutes = 60 }

        (Get-Stamp $global:AuTest.Calls[0]) | Should -Be (Get-Stamp $watermark)
        $global:AuTest.Calls[0].Kind | Should -Be 'Utc'
        @($global:AuTest.Calls).Count | Should -BeGreaterThan 2
        # Slices run oldest first and do not overlap.
        $starts = @($global:AuTest.Calls)
        for ($i = 1; $i -lt $starts.Count; $i++) { $starts[$i] | Should -BeGreaterThan $starts[$i - 1] }
    }

    It 'starts <Days> days back on a first run' -ForEach @(@{ Days = 7 }) {
        Mock Search-UnifiedAuditLog { $global:AuTest.Calls.Add($StartDate) }
        Invoke-CollectorScript 'Get-AuditSearchCmdlet.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; SliceMinutes = 1440 }
        ([datetime]::UtcNow - $global:AuTest.Calls[0]).TotalDays | Should -BeGreaterThan ($Days - 0.01)
        ([datetime]::UtcNow - $global:AuTest.Calls[0]).TotalDays | Should -BeLessThan ($Days + 0.01)
    }

    It 'skips a record it already holds on a second run' {
        $stamp = [datetime]::UtcNow.AddMinutes(-30)
        Mock Search-UnifiedAuditLog {
            $global:AuTest.Calls.Add(1)
            if ($global:AuTest.Calls.Count % 2 -eq 1) { New-SearchRecord -Id 'same' -At $stamp }
        }
        Invoke-CollectorScript 'Get-AuditSearchCmdlet.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; LookbackDays = 1; SliceMinutes = 1440 }
        $global:AuTest.Calls.Clear()
        Invoke-CollectorScript 'Get-AuditSearchCmdlet.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; SliceMinutes = 1440 }
        @(Import-Csv -LiteralPath (Join-Path $script:Out 'audit-search-cmdlet.csv')).Count | Should -Be 1
        Get-LogText $script:Out | Should -Match '1 skipped'
    }

    It 'keeps the slices it finished and does not move past one that failed' {
        # The file holds whole seconds, so the slice boundaries are whole seconds too.
        $watermark = [datetime]::ParseExact((Get-Stamp ([datetime]::UtcNow.AddHours(-4))), 'yyyy-MM-ddTHH:mm:ssZ', [cultureinfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AdjustToUniversal)
        New-ExistingCsv $script:Out 'audit-search-cmdlet.csv' 'AuditSearchCmdlet' @{ CreationTime = (Get-Stamp $watermark); RecordId = 'old' }
        $global:AuTest.Watermark = $watermark
        $global:AuTest.Served = @{}
        Mock Search-UnifiedAuditLog {
            if ($StartDate -ge $global:AuTest.Watermark.AddHours(2)) { throw 'The operation timed out.' }
            if ($global:AuTest.Served.ContainsKey($SessionId)) { return }
            $global:AuTest.Served[$SessionId] = $true
            New-SearchRecord -Id ('s-{0:yyyyMMddHHmm}' -f $StartDate) -At $StartDate.AddMinutes(5)
        }
        { Invoke-CollectorScript 'Get-AuditSearchCmdlet.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; SliceMinutes = 60 } } |
            Should -Throw '*timed out*'

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:Out 'audit-search-cmdlet.csv'))
        $rows.Count | Should -Be 3
        $newest = ($rows.CreationTime | Sort-Object | Select-Object -Last 1)
        $newest | Should -BeLessThan (Get-Stamp $watermark.AddHours(2))
        Get-LogText $script:Out | Should -Match 'Search-UnifiedAuditLog failed'
    }

    It 'logs the failure and keeps the header when the cmdlet is refused' {
        Mock Search-UnifiedAuditLog { throw 'The term Search-UnifiedAuditLog is not recognized' }
        { Invoke-CollectorScript 'Get-AuditSearchCmdlet.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; Environment = 'GCC'; LookbackDays = 1 } } |
            Should -Throw '*not recognized*'
        Get-HeaderText (Join-Path $script:Out 'audit-search-cmdlet.csv') | Should -Be ($script:Schema.AuditSearchCmdlet -join ',')
        Get-LogText $script:Out | Should -Match 'Search-UnifiedAuditLog failed'
    }
}

Describe 'Graph Audit Search API (source 2)' {
    BeforeEach {
        $script:Out = New-TestFolder
        $global:AuTest.Requests = [System.Collections.Generic.List[object]]::new()
        $global:AuTest.Sleeps = [System.Collections.Generic.List[object]]::new()
        $global:AuTest.Polls = 0
        $global:AuTest.QueryCount = 0
        Mock Start-Sleep { $global:AuTest.Sleeps.Add($Seconds) }
    }
    AfterEach { Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue }

    It 'creates one query, polls it to succeeded, and follows @odata.nextLink' {
        Mock Invoke-MgGraphRequest {
            $global:AuTest.Requests.Add([pscustomobject]@{ Method = $Method; Uri = $Uri; Body = $Body })
            if ($Method -eq 'POST') { return @{ id = 'q1'; status = 'notStarted' } }
            if ($Uri -eq '/v1.0/security/auditLog/queries/q1') {
                $global:AuTest.Polls++
                return @{ id = 'q1'; status = $(if ($global:AuTest.Polls -lt 3) { 'running' } else { 'succeeded' }); isRecordCountLimitExceeded = $false }
            }
            if ($Uri -eq '/v1.0/security/auditLog/queries/q1/records') {
                return @{ value = @((New-GraphRecord 'g1'), (New-GraphRecord 'g2')); '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/security/auditLog/queries/q1/records?$skiptoken=abc' }
            }
            if ($Uri -like '*skiptoken=abc') { return @{ value = @((New-GraphRecord 'g3')) } }
            throw "Unexpected request $Method $Uri"
        }
        Invoke-CollectorScript 'Get-AuditGraphRecords.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; LookbackDays = 1; SliceMinutes = 1440; PollSeconds = 7 }

        @($global:AuTest.Requests | Where-Object Method -EQ 'POST').Count | Should -Be 1
        $post = $global:AuTest.Requests | Where-Object Method -EQ 'POST'
        $post.Uri | Should -Be '/v1.0/security/auditLog/queries'
        $post.Body | Should -Match '"@odata.type":\s*"#microsoft.graph.security.auditLogQuery"'
        $post.Body | Should -Match '"filterStartDateTime":\s*"\d{4}-\d\d-\d\dT\d\d:\d\d:\d\dZ"'
        $post.Body | Should -Match '"filterEndDateTime":\s*"\d{4}-\d\d-\d\dT\d\d:\d\d:\d\dZ"'
        $global:AuTest.Sleeps | Should -Be @(7, 7)

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:Out 'audit-graph-records.csv'))
        ($rows.RecordId | Sort-Object) | Should -Be @('g1', 'g2', 'g3')
        ($rows.QueryId | Sort-Object -Unique) | Should -Be 'q1'
        $rows[0].Workload | Should -Be 'SharePoint'
        $rows[0].RecordType | Should -Be 'sharePointFileOperation'
        $rows[0].ResultStatus | Should -Be 'Succeeded'
        Get-HeaderText (Join-Path $script:Out 'audit-graph-records.csv') | Should -Be ($script:Schema.AuditGraphRecords -join ',')
    }

    It 'sends one query per slice, oldest first' {
        $global:AuTest.Starts = [System.Collections.Generic.List[string]]::new()
        Mock Invoke-MgGraphRequest {
            if ($Method -eq 'POST') {
                $global:AuTest.QueryCount++
                $global:AuTest.Starts.Add([regex]::Match($Body, '"filterStartDateTime":\s*"([^"]+)"').Groups[1].Value)
                return @{ id = "q$($global:AuTest.QueryCount)" }
            }
            if ($Uri -like '*/records') { return @{ value = @() } }
            return @{ status = 'succeeded' }
        }
        Invoke-CollectorScript 'Get-AuditGraphRecords.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; LookbackDays = 3; SliceMinutes = 1440 }
        $global:AuTest.Starts.Count | Should -BeGreaterOrEqual 3
        @($global:AuTest.Starts | Sort-Object) | Should -Be @($global:AuTest.Starts)
    }

    It 'resumes from the latest CreationTime already in the file' {
        $watermark = [datetime]::UtcNow.AddHours(-5)
        New-ExistingCsv $script:Out 'audit-graph-records.csv' 'AuditGraphRecords' @{ CreationTime = (Get-Stamp $watermark); RecordId = 'old' }
        $global:AuTest.Starts = [System.Collections.Generic.List[string]]::new()
        Mock Invoke-MgGraphRequest {
            if ($Method -eq 'POST') { $global:AuTest.Starts.Add([regex]::Match($Body, '"filterStartDateTime":\s*"([^"]+)"').Groups[1].Value); return @{ id = 'q' } }
            if ($Uri -like '*/records') { return @{ value = @() } }
            return @{ status = 'succeeded' }
        }
        Invoke-CollectorScript 'Get-AuditGraphRecords.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; SliceMinutes = 1440 }
        $global:AuTest.Starts[0] | Should -Be (Get-Stamp $watermark)
    }

    It 'reads a range that went over the record limit as two halves' {
        Mock Invoke-MgGraphRequest {
            if ($Method -eq 'POST') { $global:AuTest.QueryCount++; return @{ id = "q$($global:AuTest.QueryCount)" } }
            if ($Uri -like '*/records') { return @{ value = @() } }
            return @{ status = 'succeeded'; isRecordCountLimitExceeded = ($Uri -like '*/q1') }
        }
        $start = [datetime]::new(2026, 10, 8, 0, 0, 0, [DateTimeKind]::Utc)
        $rows = @(Get-AuditGraphSlice -Start $start -End $start.AddHours(8) -OutputPath $script:Out -MinimumMinutes 60)
        $global:AuTest.QueryCount | Should -Be 3
        Get-LogText $script:Out | Should -Match 'went over the record limit'
    }

    It 'reads a range just over the minimum as two halves when the query exceeded its record limit' {
        Mock Invoke-MgGraphRequest {
            if ($Method -eq 'POST') {
                $global:AuTest.QueryCount++
                return @{ id = "q$($global:AuTest.QueryCount)" }
            }
            if ($Uri -like '*/records') { return @{ value = @((New-GraphRecord 'kept')) } }
            return @{ status = 'succeeded'; isRecordCountLimitExceeded = ($Uri -like '*/queries/q1') }
        }
        $start = [datetime]::new(2026, 10, 8, 0, 0, 0, [DateTimeKind]::Utc)
        $rows = @(Get-AuditGraphSlice -Start $start -End $start.AddMinutes(90) -OutputPath $script:Out -MinimumMinutes 60)
        $global:AuTest.QueryCount | Should -Be 3
        @($rows).Count | Should -Be 2
    }

    It 'does not return the records of a query that is still over the record limit at the minimum' {
        Mock Invoke-MgGraphRequest {
            if ($Method -eq 'POST') { return @{ id = 'q-floor' } }
            if ($Uri -like '*/records') { return @{ value = @((New-GraphRecord 'dropped')) } }
            return @{ status = 'succeeded'; isRecordCountLimitExceeded = $true }
        }
        $start = [datetime]::new(2026, 10, 8, 0, 0, 0, [DateTimeKind]::Utc)
        { Get-AuditGraphSlice -Start $start -End $start.AddMinutes(60) -OutputPath $script:Out -MinimumMinutes 60 } |
            Should -Throw '*record limit*'
        Get-LogText $script:Out | Should -Match 'were not written'
    }

    It 'throws on a query that failed instead of treating it as empty' {
        Mock Invoke-MgGraphRequest {
            if ($Method -eq 'POST') { return @{ id = 'q1' } }
            return @{ status = 'failed' }
        }
        { Invoke-CollectorScript 'Get-AuditGraphRecords.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; LookbackDays = 1 } } |
            Should -Throw '*ended with status failed*'
        Get-LogText $script:Out | Should -Match 'ended with status failed'
        @(Import-Csv -LiteralPath (Join-Path $script:Out 'audit-graph-records.csv')).Count | Should -Be 0
    }

    It 'gives up on a query that never succeeds' {
        Mock Invoke-MgGraphRequest {
            if ($Method -eq 'POST') { return @{ id = 'q1' } }
            return @{ status = 'running' }
        }
        { Invoke-CollectorScript 'Get-AuditGraphRecords.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; LookbackDays = 1; MaxPolls = 3; PollSeconds = 1 } } |
            Should -Throw '*did not succeed after 3 polls*'
        Get-LogText $script:Out | Should -Match 'did not succeed after 3 polls'
    }

    It 'refuses a nextLink outside Microsoft Graph and does not request it' {
        Mock Invoke-MgGraphRequest {
            $global:AuTest.Requests.Add($Uri)
            if ($Method -eq 'POST') { return @{ id = 'q1' } }
            if ($Uri -like '*/records') { return @{ value = @((New-GraphRecord 'g1')); '@odata.nextLink' = 'https://evil.example.net/steal' } }
            return @{ status = 'succeeded' }
        }
        { Invoke-CollectorScript 'Get-AuditGraphRecords.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; LookbackDays = 1 } } |
            Should -Throw '*outside Microsoft Graph*'
        $global:AuTest.Requests | Should -Not -Contain 'https://evil.example.net/steal'
        Get-LogText $script:Out | Should -Match 'Refusing to request a page outside Microsoft Graph'
    }

    It 'waits 30 seconds, then longer, on a 429 with no Retry-After and never retries at once' {
        Mock Invoke-MgGraphRequest {
            if ($Method -eq 'POST') {
                $global:AuTest.QueryCount++
                if ($global:AuTest.QueryCount -le 2) { throw 'Response status code does not indicate success: 429 (Too Many Requests).' }
                return @{ id = 'q1' }
            }
            if ($Uri -like '*/records') { return @{ value = @() } }
            return @{ status = 'succeeded' }
        }
        $start = [datetime]::new(2026, 10, 8, 0, 0, 0, [DateTimeKind]::Utc)
        @(Get-AuditGraphSlice -Start $start -End $start.AddHours(1) -OutputPath $script:Out)
        $global:AuTest.Sleeps | Should -Be @(30, 60)
    }

    It 'waits for Retry-After when the 429 carries one' {
        $response = [System.Net.Http.HttpResponseMessage]::new([System.Net.HttpStatusCode]429)
        $response.Headers.RetryAfter = [System.Net.Http.Headers.RetryConditionHeaderValue]::new([timespan]::FromSeconds(7))
        $global:AuTest.Throttle = [Microsoft.PowerShell.Commands.HttpResponseException]::new('Response status code does not indicate success: 429 (Too Many Requests).', $response)
        Mock Invoke-MgGraphRequest {
            if ($Method -eq 'POST') {
                $global:AuTest.QueryCount++
                if ($global:AuTest.QueryCount -eq 1) { throw $global:AuTest.Throttle }
                return @{ id = 'q1' }
            }
            if ($Uri -like '*/records') { return @{ value = @() } }
            return @{ status = 'succeeded' }
        }
        $start = [datetime]::new(2026, 10, 8, 0, 0, 0, [DateTimeKind]::Utc)
        @(Get-AuditGraphSlice -Start $start -End $start.AddHours(1) -OutputPath $script:Out)
        $global:AuTest.Sleeps | Should -Be @(7)
    }

    It 'rethrows an error that is not a 429 without waiting' {
        Mock Invoke-MgGraphRequest { throw 'Forbidden: Insufficient privileges' }
        { Invoke-CollectorScript 'Get-AuditGraphRecords.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; LookbackDays = 1 } } |
            Should -Throw '*Insufficient privileges*'
        $global:AuTest.Sleeps.Count | Should -Be 0
        Get-LogText $script:Out | Should -Match 'Insufficient privileges'
    }
}

Describe 'Office 365 Management Activity API (source 3)' {
    BeforeEach {
        $script:Out = New-TestFolder
        $global:AuTest.Requests = [System.Collections.Generic.List[object]]::new()
        $global:AuTest.Sleeps = [System.Collections.Generic.List[object]]::new()
        $global:AuTest.Subscribed = @('Audit.Exchange', 'Audit.General')
        $global:AuTest.Fail = $null
        Mock Start-Sleep { $global:AuTest.Sleeps.Add($Seconds) }

        $exchangeEvents = @(
            @{ Id = 'e1'; CreationTime = '2026-10-08T09:15:00'; Operation = 'MailItemsAccessed'; OrganizationId = $script:Org; RecordType = 50; ResultStatus = 'Succeeded'; UserId = 'avery.abara@example.com'; Workload = 'Exchange'; ClientIP = '203.0.113.1'; ObjectId = '' }
            @{ Id = 'e2'; CreationTime = '2026-10-08T09:16:00'; Operation = 'FolderBind'; OrganizationId = $script:Org; RecordType = 2; ResultStatus = 'Succeeded'; UserId = 'blake.bishop@example.com'; Workload = 'Exchange'; ClientIP = '203.0.113.2'; ObjectId = '' }
        )
        $generalEvents = @(
            @{ Id = 'g1'; CreationTime = '2026-10-08T10:00:00'; Operation = 'SendAs'; OrganizationId = $script:Org; RecordType = 2; ResultStatus = 'Succeeded'; UserId = 'casey.cho@example.com'; Workload = 'Exchange'; ClientIP = '203.0.113.3'; ObjectId = '' }
        )
        $global:AuTest.Blobs = @{
            'https://manage.office.com/api/v1.0/33333333-3333-3333-3333-333333333333/activity/feed/audit/blob-ex-1' = $exchangeEvents
            'https://manage.office.com/api/v1.0/33333333-3333-3333-3333-333333333333/activity/feed/audit/blob-gen-1' = $generalEvents
        }
        $base = 'https://manage.office.com/api/v1.0/33333333-3333-3333-3333-333333333333/activity/feed'
        $global:AuTest.Base = $base
        $created = (Get-Stamp ([datetime]::UtcNow.AddHours(-2)))

        # Both blobs are listed for their content type; Audit.Exchange takes two pages.
        Mock Invoke-WebRequest {
            $global:AuTest.Requests.Add([pscustomobject]@{ Method = $Method; Uri = $Uri; Headers = @($Headers.Keys) })
            if ($global:AuTest.Fail -and $Uri -like $global:AuTest.Fail) { throw 'Response status code does not indicate success: 500 (Internal Server Error).' }
            $base = $global:AuTest.Base
            if ($Method -eq 'Post') { return New-WebResponse @{ contentType = 'Audit.SharePoint'; status = 'enabled' } }
            if ($Uri -like "$base/subscriptions/list*") {
                return New-WebResponse @($global:AuTest.Subscribed | ForEach-Object { @{ contentType = $_; status = 'enabled'; webhook = $null } })
            }
            if ($Uri -like "$base/subscriptions/content*contentType=Audit.Exchange*") {
                if ($Uri -notlike '*nextPage=*') {
                    return New-WebResponse @(@{ contentType = 'Audit.Exchange'; contentId = 'blob-ex-1'; contentUri = "$base/audit/blob-ex-1"; contentCreated = $created; contentExpiration = '2099-01-01T00:00:00.000Z' }) `
                        -Headers @{ NextPageUri = @("$base/subscriptions/content?contentType=Audit.Exchange&startTime=2026-10-01&endTime=2026-10-02&nextPage=2") }
                }
                return New-WebResponse @()
            }
            if ($Uri -like "$base/subscriptions/content*contentType=Audit.General*") {
                return New-WebResponse @(@{ contentType = 'Audit.General'; contentId = 'blob-gen-1'; contentUri = "$base/audit/blob-gen-1"; contentCreated = $created; contentExpiration = '2099-01-01T00:00:00.000Z' })
            }
            if ($global:AuTest.Blobs.ContainsKey(($Uri -replace '[?&]PublisherIdentifier=.*$', ''))) {
                return New-WebResponse $global:AuTest.Blobs[($Uri -replace '[?&]PublisherIdentifier=.*$', '')]
            }
            throw "Unexpected request $Method $Uri"
        }

        $script:FeedArgs = @{
            OutputPath = $script:Out; TenantId = $script:TenantGuid; PublisherIdentifier = $script:PublisherGuid
            AccessToken = (New-Token); ContentType = @('Audit.Exchange', 'Audit.General'); LookbackDays = 1
        }
    }
    AfterEach { Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue }

    It 'lists the subscriptions, then pages the content listing and reads every blob' {
        Invoke-CollectorScript 'Get-AuditActivityFeed.ps1' $script:FeedArgs

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:Out 'audit-activity-feed.csv'))
        ($rows.RecordId | Sort-Object) | Should -Be @('e1', 'e2', 'g1')
        ($rows | Where-Object RecordId -EQ 'e1').ContentType | Should -Be 'Audit.Exchange'
        ($rows | Where-Object RecordId -EQ 'e1').ContentId | Should -Be 'blob-ex-1'
        ($rows | Where-Object RecordId -EQ 'e1').RecordType | Should -Be '50'
        ($rows | Where-Object RecordId -EQ 'e1').CreationTime | Should -Be '2026-10-08T09:15:00Z'
        Get-HeaderText (Join-Path $script:Out 'audit-activity-feed.csv') | Should -Be ($script:Schema.AuditActivityFeed -join ',')

        $uris = $global:AuTest.Requests.Uri
        @($uris | Where-Object { $_ -like '*subscriptions/content*contentType=Audit.Exchange*' }).Count | Should -Be 2
        @($uris | Where-Object { $_ -like '*nextPage=2*' }).Count | Should -Be 1
        $global:AuTest.Requests | Where-Object { $_.Method -eq 'Post' } | Should -BeNullOrEmpty
    }

    It 'keeps each event as it arrived in AuditData' {
        Invoke-CollectorScript 'Get-AuditActivityFeed.ps1' $script:FeedArgs
        $row = Import-Csv -LiteralPath (Join-Path $script:Out 'audit-activity-feed.csv') | Where-Object RecordId -EQ 'e1'
        $row.AuditData | Should -Match '"CreationTime":"2026-10-08T09:15:00"'
        ($row.AuditData | ConvertFrom-Json).Operation | Should -Be 'MailItemsAccessed'
    }

    It 'sends PublisherIdentifier on every request, including the next page and each blob' {
        Invoke-CollectorScript 'Get-AuditActivityFeed.ps1' $script:FeedArgs
        $global:AuTest.Requests.Count | Should -BeGreaterThan 5
        foreach ($request in $global:AuTest.Requests) { $request.Uri | Should -Match "[?&]PublisherIdentifier=$($script:PublisherGuid)" }
    }

    It 'lists a window of at most 24 hours, with startTime and endTime together, in UTC' {
        Invoke-CollectorScript 'Get-AuditActivityFeed.ps1' $script:FeedArgs
        $lists = @($global:AuTest.Requests.Uri | Where-Object { $_ -like '*subscriptions/content*' -and $_ -notlike '*nextPage*' })
        $lists.Count | Should -BeGreaterThan 0
        foreach ($uri in $lists) {
            $m = [regex]::Match($uri, 'startTime=(\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d)&endTime=(\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d)')
            $m.Success | Should -BeTrue
            $start = [datetime]::ParseExact($m.Groups[1].Value, 'yyyy-MM-ddTHH:mm:ss', $null)
            $end = [datetime]::ParseExact($m.Groups[2].Value, 'yyyy-MM-ddTHH:mm:ss', $null)
            ($end - $start).TotalHours | Should -BeLessOrEqual 24
            ($end - $start).TotalSeconds | Should -BeGreaterThan 0
        }
    }

    It 'sends the subscription start only for a content type that is not enabled' {
        $global:AuTest.Subscribed = @('Audit.Exchange')
        Invoke-CollectorScript 'Get-AuditActivityFeed.ps1' $script:FeedArgs

        $posts = @($global:AuTest.Requests | Where-Object { $_.Method -eq 'Post' })
        $posts.Count | Should -Be 1
        $posts[0].Uri | Should -Be "$($global:AuTest.Base)/subscriptions/start?contentType=Audit.General&PublisherIdentifier=$($script:PublisherGuid)"
        Get-LogText $script:Out | Should -Match 'Started the Audit.General subscription'
    }

    It 'sends no start with -NoStartSubscription and skips the type with a warning' {
        $global:AuTest.Subscribed = @('Audit.Exchange')
        Invoke-CollectorScript 'Get-AuditActivityFeed.ps1' (Merge-Arguments $script:FeedArgs @{ NoStartSubscription = $true })

        @($global:AuTest.Requests | Where-Object { $_.Method -eq 'Post' }).Count | Should -Be 0
        Get-LogText $script:Out | Should -Match 'Audit.General has no enabled subscription'
        @($global:AuTest.Requests.Uri | Where-Object { $_ -like '*contentType=Audit.General*' -and $_ -like '*subscriptions/content*' }).Count | Should -Be 0
        $rows = @(Import-Csv -LiteralPath (Join-Path $script:Out 'audit-activity-feed.csv'))
        ($rows.RecordId | Sort-Object) | Should -Be @('e1', 'e2')
    }

    It 'never sends a stop' {
        $global:AuTest.Subscribed = @()
        Invoke-CollectorScript 'Get-AuditActivityFeed.ps1' $script:FeedArgs
        @($global:AuTest.Requests.Uri | Where-Object { $_ -like '*subscriptions/stop*' }).Count | Should -Be 0
    }

    It 'uses the <Root> root for <Cloud>' -ForEach @(
        @{ Cloud = 'Commercial'; Root = 'https://manage.office.com' }
        @{ Cloud = 'GCC'; Root = 'https://manage-gcc.office.com' }
        @{ Cloud = 'GCCHigh'; Root = 'https://manage.office365.us' }
    ) {
        Mock Invoke-WebRequest {
            $global:AuTest.Requests.Add([pscustomobject]@{ Method = $Method; Uri = $Uri })
            New-WebResponse @()
        }
        Invoke-CollectorScript 'Get-AuditActivityFeed.ps1' (Merge-Arguments $script:FeedArgs @{ Environment = $Cloud })
        $expected = "$Root/api/v1.0/$($script:TenantGuid)/activity/feed/"
        foreach ($request in $global:AuTest.Requests) { $request.Uri | Should -BeLike "$expected*" }
    }

    It 'resumes from the latest ContentCreated already in the file' {
        $watermark = [datetime]::UtcNow.AddHours(-5)
        New-ExistingCsv $script:Out 'audit-activity-feed.csv' 'AuditActivityFeed' @{ ContentCreated = (Get-Stamp $watermark); RecordId = 'old' }
        Invoke-CollectorScript 'Get-AuditActivityFeed.ps1' ($script:FeedArgs.Clone())
        $first = $global:AuTest.Requests.Uri | Where-Object { $_ -like '*subscriptions/content*' } | Select-Object -First 1
        $first | Should -Match ('startTime=' + [regex]::Escape($watermark.ToString('yyyy-MM-ddTHH:mm:ss')))
    }

    It 'starts no further back than the feed holds and logs the gap' {
        $watermark = [datetime]::UtcNow.AddDays(-9)
        New-ExistingCsv $script:Out 'audit-activity-feed.csv' 'AuditActivityFeed' @{ ContentCreated = (Get-Stamp $watermark); RecordId = 'old' }
        Invoke-CollectorScript 'Get-AuditActivityFeed.ps1' ($script:FeedArgs.Clone())
        $first = $global:AuTest.Requests.Uri | Where-Object { $_ -like '*subscriptions/content*' } | Select-Object -First 1
        $m = [regex]::Match($first, 'startTime=(\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d)')
        $m.Success | Should -BeTrue
        ([datetime]::UtcNow - [datetime]::ParseExact($m.Groups[1].Value, 'yyyy-MM-ddTHH:mm:ss', $null)).TotalDays | Should -BeLessThan 7
        Get-LogText $script:Out | Should -Match 'older than the feed holds'
    }

    It 'skips an event it already holds on a second run' {
        Invoke-CollectorScript 'Get-AuditActivityFeed.ps1' $script:FeedArgs
        Invoke-CollectorScript 'Get-AuditActivityFeed.ps1' $script:FeedArgs
        @(Import-Csv -LiteralPath (Join-Path $script:Out 'audit-activity-feed.csv')).Count | Should -Be 3
        Get-LogText $script:Out | Should -Match '3 skipped'
    }

    It 'writes nothing for a window when one content type in it fails' {
        $global:AuTest.Fail = '*contentType=Audit.General*'
        { Invoke-CollectorScript 'Get-AuditActivityFeed.ps1' $script:FeedArgs } | Should -Throw '*500*'
        @(Import-Csv -LiteralPath (Join-Path $script:Out 'audit-activity-feed.csv')).Count | Should -Be 0
        Get-LogText $script:Out | Should -Match 'Management Activity API failed'
    }

    It 'skips a blob whose contentExpiration has passed and still writes the others' {
        Mock Invoke-WebRequest {
            $global:AuTest.Requests.Add([pscustomobject]@{ Method = $Method; Uri = $Uri })
            $base = $global:AuTest.Base
            if ($Uri -like "$base/subscriptions/list*") { return New-WebResponse @(@{ contentType = 'Audit.General'; status = 'enabled' }) }
            if ($Uri -like "$base/subscriptions/content*") {
                return New-WebResponse @(
                    @{ contentType = 'Audit.General'; contentId = 'old-blob'; contentUri = "$base/audit/old"; contentCreated = (Get-Stamp ([datetime]::UtcNow.AddHours(-3))); contentExpiration = '2000-01-01T00:00:00.000Z' }
                    @{ contentType = 'Audit.General'; contentId = 'new-blob'; contentUri = "$base/audit/new"; contentCreated = (Get-Stamp ([datetime]::UtcNow.AddHours(-1))); contentExpiration = '2099-01-01T00:00:00.000Z' }
                )
            }
            if ($Uri -like '*/audit/old*') { throw 'expired blob was requested' }
            if ($Uri -like '*/audit/new*') {
                return New-WebResponse @(@{ Id = 'kept'; CreationTime = '2026-10-08T09:15:00'; Operation = 'FileAccessed'; OrganizationId = $global:AuTest.Org; RecordType = 6; UserId = 'avery.abara@example.com'; Workload = 'SharePoint' })
            }
            throw "Unexpected request $Uri"
        }
        Invoke-CollectorScript 'Get-AuditActivityFeed.ps1' (Merge-Arguments $script:FeedArgs @{ ContentType = 'Audit.General' })
        @(Import-Csv -LiteralPath (Join-Path $script:Out 'audit-activity-feed.csv')).RecordId | Should -Be 'kept'
        @($global:AuTest.Requests.Uri | Where-Object { $_ -like '*/audit/old*' }).Count | Should -Be 0
        Get-LogText $script:Out | Should -Match 'contentExpiration'
    }

    It 'refuses a contentUri outside the feed host and never sends the token there' {
        $global:AuTest.Blobs = @{}
        Mock Invoke-WebRequest {
            $global:AuTest.Requests.Add([pscustomobject]@{ Method = $Method; Uri = $Uri })
            $base = $global:AuTest.Base
            if ($Uri -like "$base/subscriptions/list*") { return New-WebResponse @(@{ contentType = 'Audit.General'; status = 'enabled' }) }
            if ($Uri -like "$base/subscriptions/content*") {
                return New-WebResponse @(@{ contentType = 'Audit.General'; contentId = 'x'; contentUri = 'https://evil.example.net/blob'; contentCreated = (Get-Stamp ([datetime]::UtcNow.AddHours(-1))) })
            }
            throw "Unexpected request $Uri"
        }
        { Invoke-CollectorScript 'Get-AuditActivityFeed.ps1' (Merge-Arguments $script:FeedArgs @{ ContentType = 'Audit.General' }) } |
            Should -Throw '*outside the Management Activity API*'
        @($global:AuTest.Requests.Uri | Where-Object { $_ -like '*evil.example.net*' }).Count | Should -Be 0
        Get-LogText $script:Out | Should -Match 'outside the Management Activity API'
    }

    It 'refuses a NextPageUri outside the feed host' {
        Mock Invoke-WebRequest {
            $global:AuTest.Requests.Add([pscustomobject]@{ Method = $Method; Uri = $Uri })
            $base = $global:AuTest.Base
            if ($Uri -like "$base/subscriptions/list*") { return New-WebResponse @(@{ contentType = 'Audit.General'; status = 'enabled' }) }
            New-WebResponse @() -Headers @{ NextPageUri = @('http://manage.office.com/api/v1.0/x/activity/feed/subscriptions/content?nextPage=1') }
        }
        { Invoke-CollectorScript 'Get-AuditActivityFeed.ps1' (Merge-Arguments $script:FeedArgs @{ ContentType = 'Audit.General' }) } |
            Should -Throw '*outside the Management Activity API*'
        Get-LogText $script:Out | Should -Match 'outside the Management Activity API'
    }

    It 'waits 30 seconds on a 429 and then continues' {
        $global:AuTest.Throttled = 0
        Mock Invoke-WebRequest {
            $base = $global:AuTest.Base
            if ($Uri -like "$base/subscriptions/list*") {
                if ($global:AuTest.Throttled -eq 0) { $global:AuTest.Throttled = 1; throw 'Response status code does not indicate success: 429 (AF429: Too many requests).' }
                return New-WebResponse @(@{ contentType = 'Audit.General'; status = 'enabled' })
            }
            New-WebResponse @()
        }
        Invoke-CollectorScript 'Get-AuditActivityFeed.ps1' (Merge-Arguments $script:FeedArgs @{ ContentType = 'Audit.General' })
        $global:AuTest.Sleeps | Should -Be @(30)
    }

    It 'never writes the token to run.log, a URL or the CSV, including when a request fails' {
        Mock Invoke-WebRequest {
            $global:AuTest.Requests.Add([pscustomobject]@{ Method = $Method; Uri = $Uri })
            throw 'Response status code does not indicate success: 401 (Unauthorized).'
        }
        { Invoke-CollectorScript 'Get-AuditActivityFeed.ps1' $script:FeedArgs } | Should -Throw '*401*'
        (Get-LogText $script:Out) | Should -Not -Match ([regex]::Escape($script:TokenText))
        (Get-LogText $script:Out) | Should -Not -Match 'Bearer'
        (Get-Content -LiteralPath (Join-Path $script:Out 'audit-activity-feed.csv') -Raw) | Should -Not -Match ([regex]::Escape($script:TokenText))
        ($global:AuTest.Requests.Uri -join ' ') | Should -Not -Match ([regex]::Escape($script:TokenText))
    }

    It 'sends the token only in the Authorization header' {
        $global:AuTest.Seen = $null
        Mock Invoke-WebRequest {
            if ($null -eq $global:AuTest.Seen) { $global:AuTest.Seen = $Headers['Authorization'] }
            New-WebResponse @()
        }
        Invoke-CollectorScript 'Get-AuditActivityFeed.ps1' $script:FeedArgs
        $global:AuTest.Seen | Should -Be ('Bearer ' + $script:TokenText)
    }

    It 'rejects a TenantId that is not a GUID before any request' {
        { Invoke-CollectorScript 'Get-AuditActivityFeed.ps1' (Merge-Arguments $script:FeedArgs @{ TenantId = 'contoso.onmicrosoft.com' }) } | Should -Throw '*is not a GUID*'
        $global:AuTest.Requests.Count | Should -Be 0
    }
}

Describe 'Ingestion status and retention policies (sources 4 and 5)' {
    BeforeEach {
        $script:Out = New-TestFolder
        $script:RunDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')
        Mock Disconnect-ExchangeOnline { }
    }
    AfterEach { Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue }

    It 'stamps the ingestion setting with the run date and writes it once per day' {
        Mock Get-AdminAuditLogConfig { [pscustomobject]@{ UnifiedAuditLogIngestionEnabled = $false } }
        Invoke-CollectorScript 'Get-AuditIngestion.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        Invoke-CollectorScript 'Get-AuditIngestion.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        $rows = @(Import-Csv -LiteralPath (Join-Path $script:Out 'audit-ingestion.csv'))
        $rows.Count | Should -Be 1
        $rows[0].RunDate | Should -Be $script:RunDate
        $rows[0].UnifiedAuditLogIngestionEnabled | Should -Be 'False'
        Get-HeaderText (Join-Path $script:Out 'audit-ingestion.csv') | Should -Be ($script:Schema.AuditIngestion -join ',')
    }

    It 'writes the header only and logs why when Get-AdminAuditLogConfig is refused' {
        Mock Get-AdminAuditLogConfig { throw 'The term Get-AdminAuditLogConfig is not recognized' }
        Invoke-CollectorScript 'Get-AuditIngestion.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; Environment = 'GCCHigh' }
        @(Import-Csv -LiteralPath (Join-Path $script:Out 'audit-ingestion.csv')).Count | Should -Be 0
        Get-HeaderText (Join-Path $script:Out 'audit-ingestion.csv') | Should -Be ($script:Schema.AuditIngestion -join ',')
        Get-LogText $script:Out | Should -Match 'Get-AdminAuditLogConfig is unavailable'
    }

    It 'stamps every policy with the run date, joins lists, and keeps a duration the cmdlet list does not name' {
        Mock Get-UnifiedAuditLogRetentionPolicy {
            [pscustomobject]@{ Priority = 100; Name = 'Admin changes'; RecordTypes = @('AzureActiveDirectory', 'ExchangeAdmin'); Operations = @('Add member to role.'); UserIds = @(); RetentionDuration = 'TenYears' }
            [pscustomobject]@{ Priority = 200; Name = 'Short'; RecordTypes = @(); Operations = @(); UserIds = @('devon.diaz@example.com'); RetentionDuration = 'SevenYears' }
        }
        Invoke-CollectorScript 'Get-AuditRetentionPolicies.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        $rows = @(Import-Csv -LiteralPath (Join-Path $script:Out 'audit-retention-policies.csv'))
        $rows.Count | Should -Be 2
        ($rows.RunDate | Sort-Object -Unique) | Should -Be $script:RunDate
        $rows[0].RecordTypes | Should -Be 'AzureActiveDirectory;ExchangeAdmin'
        $rows[1].RetentionDuration | Should -Be 'SevenYears'
        $rows[1].UserIds | Should -Be 'devon.diaz@example.com'
        Get-HeaderText (Join-Path $script:Out 'audit-retention-policies.csv') | Should -Be ($script:Schema.AuditRetentionPolicies -join ',')
    }

    It 'says in run.log that no rows is not proof of no one-year retention' {
        Mock Get-UnifiedAuditLogRetentionPolicy { }
        Invoke-CollectorScript 'Get-AuditRetentionPolicies.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        @(Import-Csv -LiteralPath (Join-Path $script:Out 'audit-retention-policies.csv')).Count | Should -Be 0
        Get-LogText $script:Out | Should -Match 'does not return the default policy'
    }

    It 'writes the header only when the retention cmdlet is refused' {
        Mock Get-UnifiedAuditLogRetentionPolicy { throw 'You do not have permission' }
        Invoke-CollectorScript 'Get-AuditRetentionPolicies.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; Environment = 'GCC' }
        Get-HeaderText (Join-Path $script:Out 'audit-retention-policies.csv') | Should -Be ($script:Schema.AuditRetentionPolicies -join ',')
        Get-LogText $script:Out | Should -Match 'Get-UnifiedAuditLogRetentionPolicy is unavailable'
    }
}

Describe 'Run-All' {
    BeforeEach {
        $script:Out = New-TestFolder
        Mock Connect-MgGraph -ModuleName M365ReportLibrary -MockWith { }
        Mock Connect-ExchangeOnline -ModuleName M365ReportLibrary -MockWith { }
        Mock Connect-IPPSSession -ModuleName M365ReportLibrary -MockWith { }
        Mock Disconnect-ExchangeOnline { }
        Mock Search-UnifiedAuditLog { }
        Mock Get-AdminAuditLogConfig { [pscustomobject]@{ UnifiedAuditLogIngestionEnabled = $true } }
        Mock Get-UnifiedAuditLogRetentionPolicy { }
        Mock Invoke-MgGraphRequest { if ($Method -eq 'POST') { return @{ id = 'q1' } }; if ($Uri -like '*/records') { return @{ value = @() } }; @{ status = 'succeeded' } }
        Mock Invoke-WebRequest { throw 'Run-All must not call the feed without a token.' }
        Mock Start-Sleep { }
    }
    AfterEach { Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue }

    It 'runs every collector it can in GCC High, skips the Graph one and the feed without a token, and signs out of each Exchange session' {
        & (Join-Path $script:Collectors 'Run-All.ps1') -OutputPath $script:Out -Environment GCCHigh -LookbackDays 1

        foreach ($name in 'audit-search-cmdlet.csv', 'audit-graph-records.csv', 'audit-ingestion.csv', 'audit-retention-policies.csv') {
            Test-Path -LiteralPath (Join-Path $script:Out $name) | Should -BeTrue -Because $name
        }
        Test-Path -LiteralPath (Join-Path $script:Out 'audit-activity-feed.csv') | Should -BeFalse
        Get-LogText $script:Out | Should -Match 'documented as unavailable in GCCHigh'
        Get-LogText $script:Out | Should -Match 'Skipped: the Management Activity API needs'
        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 0 -Exactly
        Should -Invoke Connect-ExchangeOnline -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter { $ExchangeEnvironmentName -eq 'O365USGovGCCHigh' }
        Should -Invoke Connect-IPPSSession -ModuleName M365ReportLibrary -Times 1 -Exactly
        Should -Invoke Disconnect-ExchangeOnline -Times 2 -Exactly
        Should -Invoke Invoke-MgGraphRequest -Times 0 -Exactly
    }

    It 'keeps going when one sign-in fails' {
        Mock Connect-IPPSSession -ModuleName M365ReportLibrary -MockWith { throw 'AADSTS50076' }
        & (Join-Path $script:Collectors 'Run-All.ps1') -OutputPath $script:Out -LookbackDays 1
        Get-LogText $script:Out | Should -Match 'SecurityCompliance sign-in failed'
        Test-Path -LiteralPath (Join-Path $script:Out 'audit-ingestion.csv') | Should -BeTrue
        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 1 -Exactly
    }

    It 'passes a 365-day lookback through to Search-UnifiedAuditLog' {
        $global:AuTest.Starts = [System.Collections.Generic.List[datetime]]::new()
        Mock Search-UnifiedAuditLog { $global:AuTest.Starts.Add($StartDate); throw 'stop' }
        Mock Invoke-MgGraphRequest { throw 'stop graph' }
        & (Join-Path $script:Collectors 'Run-All.ps1') -OutputPath $script:Out -LookbackDays 365
        ([datetime]::UtcNow - $global:AuTest.Starts[0]).TotalDays | Should -BeGreaterThan 364
        ([datetime]::UtcNow - $global:AuTest.Starts[0]).TotalDays | Should -BeLessThan 366
    }
}
