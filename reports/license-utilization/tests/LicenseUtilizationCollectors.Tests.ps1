#Requires -Version 7.0

BeforeAll {
    $script:Root = Join-Path $PSScriptRoot '../../..' | Resolve-Path | Select-Object -ExpandProperty Path
    $script:Collectors = Join-Path $script:Root 'reports/license-utilization/collectors'
    $script:Samples = Join-Path $script:Root 'reports/license-utilization/samples'

    . (Join-Path $script:Root 'shared/tests/TenantCmdletStubs.ps1')
    . (Join-Path $PSScriptRoot 'LicenseUtilizationStubs.ps1')
    Import-Module (Join-Path $script:Root 'shared/M365ReportLibrary.psm1') -Force

    $script:Schema = Import-PowerShellDataFile -LiteralPath (Join-Path $script:Collectors 'LicenseUtilizationSchema.psd1')
    . (Join-Path $script:Collectors 'LicenseUtilizationHelpers.ps1')

    $script:Today = [datetime]::UtcNow.ToString('yyyy-MM-dd')
    # Mock bodies and helper functions resolve $script: against a different scope, so
    # what they read is kept in the global scope under a test-only name.
    $global:LicenseUtilizationTestSchema = $script:Schema

    function New-TestFolder {
        $path = Join-Path ([System.IO.Path]::GetTempPath()) ('license-utilization-' + [guid]::NewGuid().ToString('N'))
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

    # subscribedSku: https://learn.microsoft.com/graph/api/resources/subscribedsku
    # servicePlanInfo: https://learn.microsoft.com/graph/api/resources/serviceplaninfo
    function New-MockSku {
        [pscustomobject]@{
            SkuId         = '11111111-1111-1111-1111-111111111111'
            SkuPartNumber = 'SPE_E3'
            AdditionalProperties = @{
                appliesTo        = 'User'
                capabilityStatus = 'Enabled'
                consumedUnits    = 5
                prepaidUnits     = @{ enabled = 12; suspended = 0; warning = 0; lockedOut = 0 }
                servicePlans     = @(@{ servicePlanId = '22222222-2222-2222-2222-222222222222'; servicePlanName = 'TEAMS1'; provisioningStatus = 'Success'; appliesTo = 'User' })
            }
        }
    }

    # user with licenseAssignmentStates: https://learn.microsoft.com/graph/api/resources/licenseassignmentstate
    function New-MockLicensedUser {
        param([string]$Id = 'user-1')
        [pscustomobject]@{
            Id                = $Id
            UserPrincipalName = "$Id@example.com"
            AdditionalProperties = @{
                accountEnabled = $true
                usageLocation  = 'US'
                officeLocation = 'Tampa HQ'
                licenseAssignmentStates = @(@{ skuId = '11111111-1111-1111-1111-111111111111'; assignedByGroup = ''; state = 'Active'; error = 'None'; disabledPlans = @() })
                assignedLicenses = @(@{ skuId = '11111111-1111-1111-1111-111111111111'; disabledPlans = @() })
            }
        }
    }

    # signInActivity: https://learn.microsoft.com/graph/api/resources/signinactivity
    function New-MockSignInUser {
        param([string]$Id = 'user-1')
        [pscustomobject]@{
            Id = $Id; UserPrincipalName = "$Id@example.com"
            AdditionalProperties = @{ signInActivity = @{
                lastSignInDateTime = '2026-08-01T10:00:00Z'; lastNonInteractiveSignInDateTime = '0001-01-01T00:00:00Z'; lastSuccessfulSignInDateTime = ''
            } }
        }
    }

    # licenseDetails: https://learn.microsoft.com/graph/api/resources/licensedetails
    function New-MockLicenseDetail {
        [pscustomobject]@{
            SkuId = '11111111-1111-1111-1111-111111111111'; SkuPartNumber = 'SPE_E3'
            AdditionalProperties = @{ servicePlans = @(@{ servicePlanId = '22222222-2222-2222-2222-222222222222'; servicePlanName = 'TEAMS1'; provisioningStatus = 'Success' }) }
        }
    }

    # adminReportSettings: https://learn.microsoft.com/graph/api/resources/adminreportsettings
    function New-MockReportSettings {
        param([bool]$Concealed = $false)
        [pscustomobject]@{ AdditionalProperties = @{ displayConcealedNames = $Concealed } }
    }

    # A usage report CSV: the Learn page header, one user. https://learn.microsoft.com/graph/api/reportroot-getoffice365activeuserdetail
    function New-UsageCsvText {
        param([Parameter(Mandatory)][string]$Key, [string]$User = 'avery.abara@example.com')
        $headers = @($global:LicenseUtilizationTestSchema.UsageReports[$Key])
        $values = foreach ($header in $headers) {
            if ($header -eq 'User Principal Name') { $User } elseif ($header -eq 'Report Period') { '30' } else { '2026-08-01' }
        }
        return ((($headers | ForEach-Object { '"' + $_ + '"' }) -join ',') + "`n" + (($values | ForEach-Object { '"' + $_ + '"' }) -join ','))
    }

    $script:UsageCases = @(
        @{ Script = 'Get-ActiveUserUsage.ps1'; Key = 'ActiveUserUsage'; Csv = 'usage-active-users.csv' }
        @{ Script = 'Get-EmailActivityUsage.ps1'; Key = 'EmailActivityUsage'; Csv = 'usage-email-activity.csv' }
        @{ Script = 'Get-TeamsActivityUsage.ps1'; Key = 'TeamsActivityUsage'; Csv = 'usage-teams-activity.csv' }
        @{ Script = 'Get-SharePointActivityUsage.ps1'; Key = 'SharePointActivityUsage'; Csv = 'usage-sharepoint-activity.csv' }
        @{ Script = 'Get-OneDriveActivityUsage.ps1'; Key = 'OneDriveActivityUsage'; Csv = 'usage-onedrive-activity.csv' }
        @{ Script = 'Get-CopilotUsage.ps1'; Key = 'CopilotUsage'; Csv = 'usage-copilot.csv' }
    )
}

Describe 'Each state collector writes the columns of its sample' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
        Mock Get-MgSubscribedSku -MockWith { New-MockSku }
        Mock Get-MgUser -MockWith { New-MockLicensedUser }
        Mock Get-MgUserLicenseDetail -MockWith { New-MockLicenseDetail }
        Mock Get-MgAdminReportSetting -MockWith { New-MockReportSettings }
    }
    AfterEach { Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue }

    It 'Get-SubscribedSkus.ps1 writes subscribed-skus.csv and sku-service-plans.csv with a RunDate and a derived FreeUnits' {
        Invoke-CollectorScript 'Get-SubscribedSkus.ps1' @{ OutputPath = $script:folder }

        foreach ($pair in @(@('subscribed-skus.csv', 'SubscribedSkus'), @('sku-service-plans.csv', 'SkuServicePlans'))) {
            $produced = Join-Path $script:folder $pair[0]
            Get-HeaderText -Path $produced | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples $pair[0]))
            Get-HeaderText -Path $produced | Should -Be ($script:Schema[$pair[1]] -join ',')
        }
        $row = @(Import-Csv -LiteralPath (Join-Path $script:folder 'subscribed-skus.csv'))[0]
        $row.RunDate | Should -Be $script:Today
        $row.FreeUnits | Should -Be '7'
        @(Import-Csv -LiteralPath (Join-Path $script:folder 'sku-service-plans.csv'))[0].ServicePlanName | Should -Be 'TEAMS1'
    }

    It 'Get-UserLicenses.ps1 writes user-licenses.csv, one row per assignment state, stamped with the run date' {
        Invoke-CollectorScript 'Get-UserLicenses.ps1' @{ OutputPath = $script:folder }

        $produced = Join-Path $script:folder 'user-licenses.csv'
        Get-HeaderText -Path $produced | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples 'user-licenses.csv'))
        Get-HeaderText -Path $produced | Should -Be ($script:Schema.UserLicenses -join ',')
        $rows = @(Import-Csv -LiteralPath $produced)
        $rows.Count | Should -Be 1
        $rows[0].RunDate | Should -Be $script:Today
        $rows[0].AssignmentType | Should -Be 'Direct'
    }

    It 'Get-UserLicenses.ps1 marks a group-assigned licence' {
        Mock Get-MgUser -MockWith {
            $user = New-MockLicensedUser
            $user.AdditionalProperties.licenseAssignmentStates = @(@{ skuId = 'sku-a'; assignedByGroup = 'group-1'; state = 'Error'; error = 'CountViolation'; disabledPlans = @('p1', 'p2') })
            $user
        }
        Invoke-CollectorScript 'Get-UserLicenses.ps1' @{ OutputPath = $script:folder }

        $row = @(Import-Csv -LiteralPath (Join-Path $script:folder 'user-licenses.csv'))[0]
        $row.AssignmentType | Should -Be 'Group'
        $row.AssignedByGroup | Should -Be 'group-1'
        $row.Error | Should -Be 'CountViolation'
        $row.DisabledPlans | Should -Be 'p1;p2'
    }

    It 'Get-UserSignInActivity.ps1 writes user-signin-activity.csv and treats 0001-01-01 and blank as no sign-in' {
        Mock Get-MgUser -MockWith { New-MockSignInUser }
        Invoke-CollectorScript 'Get-UserSignInActivity.ps1' @{ OutputPath = $script:folder }

        $produced = Join-Path $script:folder 'user-signin-activity.csv'
        Get-HeaderText -Path $produced | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples 'user-signin-activity.csv'))
        Get-HeaderText -Path $produced | Should -Be ($script:Schema.UserSignInActivity -join ',')
        $row = @(Import-Csv -LiteralPath $produced)[0]
        $row.RunDate | Should -Be $script:Today
        $row.LastSignInDateTime | Should -Not -BeNullOrEmpty
        $row.LastNonInteractiveSignInDateTime | Should -BeNullOrEmpty
        $row.LastSuccessfulSignInDateTime | Should -BeNullOrEmpty
    }

    It 'Get-ReportSettings.ps1 writes report-settings.csv with the concealed-names flag' {
        Mock Get-MgAdminReportSetting -MockWith { New-MockReportSettings -Concealed $true }
        Invoke-CollectorScript 'Get-ReportSettings.ps1' @{ OutputPath = $script:folder }

        $produced = Join-Path $script:folder 'report-settings.csv'
        Get-HeaderText -Path $produced | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples 'report-settings.csv'))
        $row = @(Import-Csv -LiteralPath $produced)[0]
        $row.RunDate | Should -Be $script:Today
        $row.DisplayConcealedNames | Should -Be 'True'
    }

    It 'Get-LicenseDetails.ps1 reads the users from the latest user-licenses.csv snapshot' {
        Invoke-CollectorScript 'Get-UserLicenses.ps1' @{ OutputPath = $script:folder }
        Invoke-CollectorScript 'Get-LicenseDetails.ps1' @{ OutputPath = $script:folder }

        $produced = Join-Path $script:folder 'license-details.csv'
        Get-HeaderText -Path $produced | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples 'license-details.csv'))
        Get-HeaderText -Path $produced | Should -Be ($script:Schema.LicenseDetails -join ',')
        @(Import-Csv -LiteralPath $produced)[0].UserId | Should -Be 'user-1'
        Should -Invoke Get-MgUserLicenseDetail -Times 1 -Exactly -ParameterFilter { $UserId -eq 'user-1' }
    }

    It 'Get-LicenseDetails.ps1 writes the header only and warns when user-licenses.csv has not been collected' {
        Invoke-CollectorScript 'Get-LicenseDetails.ps1' @{ OutputPath = $script:folder }

        @(Import-Csv -LiteralPath (Join-Path $script:folder 'license-details.csv')).Count | Should -Be 0
        Get-LogText -Folder $script:folder | Should -Match 'Run Get-UserLicenses.ps1 first'
        Should -Invoke Get-MgUserLicenseDetail -Times 0 -Exactly
    }

    It 'a second run appends a new RunDate block instead of overwriting' {
        Invoke-CollectorScript 'Get-SubscribedSkus.ps1' @{ OutputPath = $script:folder }
        $path = Join-Path $script:folder 'subscribed-skus.csv'
        $old = (Get-Content -LiteralPath $path -Raw) -replace [regex]::Escape($script:Today), '2026-01-01'
        Set-Content -LiteralPath $path -Value $old -NoNewline
        Invoke-CollectorScript 'Get-SubscribedSkus.ps1' @{ OutputPath = $script:folder }

        (@(Import-Csv -LiteralPath $path).RunDate | Sort-Object -Unique) -join ',' | Should -Be "2026-01-01,$script:Today"
    }
}

Describe 'Each usage report writes the columns of its sample' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
    }
    AfterEach { Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue }

    It '<Script> downloads the CSV with a GET and writes <Csv>' -ForEach $script:UsageCases {
        $key = $Key
        Mock Invoke-MgGraphRequest -MockWith { Set-Content -LiteralPath $OutputFilePath -Value (New-UsageCsvText -Key $key) }

        Invoke-CollectorScript $Script @{ OutputPath = $script:folder }

        $produced = Join-Path $script:folder $Csv
        Get-HeaderText -Path $produced | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples $Csv))
        $row = @(Import-Csv -LiteralPath $produced)[0]
        $row.RunDate | Should -Be $script:Today
        $row.UserPrincipalName | Should -Be 'avery.abara@example.com'
        Should -Invoke Invoke-MgGraphRequest -Times 1 -Exactly -ParameterFilter { $Method -eq 'GET' -and $Uri -like 'https://graph.microsoft.com/v1.0/*' }
    }

    It 'Get-CopilotUsage.ps1 asks the v1.0 /copilot function with version v2 and a v2 period' {
        Mock Invoke-MgGraphRequest -MockWith { Set-Content -LiteralPath $OutputFilePath -Value (New-UsageCsvText -Key 'CopilotUsage') }
        Invoke-CollectorScript 'Get-CopilotUsage.ps1' @{ OutputPath = $script:folder }

        Should -Invoke Invoke-MgGraphRequest -Times 1 -Exactly -ParameterFilter {
            $Uri -eq "https://graph.microsoft.com/v1.0/copilot/reports/getMicrosoft365CopilotUsageUserDetail(period='D28',version='v2')"
        }
    }

    It 'a row with no user is not appended' {
        Mock Invoke-MgGraphRequest -MockWith { Set-Content -LiteralPath $OutputFilePath -Value (New-UsageCsvText -Key 'ActiveUserUsage' -User '') }
        Invoke-CollectorScript 'Get-ActiveUserUsage.ps1' @{ OutputPath = $script:folder }

        @(Import-Csv -LiteralPath (Join-Path $script:folder 'usage-active-users.csv')).Count | Should -Be 0
    }

    It 'fails when the usage-report download is not the CSV' {
        Mock Invoke-MgGraphRequest -MockWith { Set-Content -LiteralPath $OutputFilePath -Value 'Found' }

        { Invoke-CollectorScript 'Get-ActiveUserUsage.ps1' @{ OutputPath = $script:folder } } | Should -Throw '*not the CSV*'
    }

    It 'fails when the usage-report download writes no file' {
        Mock Invoke-MgGraphRequest -MockWith { }

        { Invoke-CollectorScript 'Get-ActiveUserUsage.ps1' @{ OutputPath = $script:folder } } | Should -Throw '*did not download*'
    }

    It 'warns when report-settings.csv says names are concealed' {
        Mock Get-MgAdminReportSetting -MockWith { New-MockReportSettings -Concealed $true }
        Mock Invoke-MgGraphRequest -MockWith { Set-Content -LiteralPath $OutputFilePath -Value (New-UsageCsvText -Key 'ActiveUserUsage') }
        Invoke-CollectorScript 'Get-ReportSettings.ps1' @{ OutputPath = $script:folder }
        Invoke-CollectorScript 'Get-ActiveUserUsage.ps1' @{ OutputPath = $script:folder }

        Get-LogText -Folder $script:folder | Should -Match 'conceal user names'
    }
}

Describe 'Microsoft 365 apps usage follows its JSON paging' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
    }
    AfterEach { Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue }

    # https://learn.microsoft.com/graph/api/reportroot-getm365appuserdetail (JSON form, @odata.nextLink)
    It 'reads every page until @odata.nextLink is absent' {
        $global:LicenseUtilizationTestCalls = [System.Collections.Generic.List[string]]::new()
        Mock Invoke-MgGraphRequest -MockWith {
            $global:LicenseUtilizationTestCalls.Add($Uri)
            if ($Uri -like '*skiptoken=2*') {
                @{ value = @(@{ userPrincipalName = 'two@example.com'; reportPeriod = '30'; details = @(@{ windows = $true }) }) }
            }
            else {
                @{ '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/reports/getM365AppUserDetail(period=''D30'')?$format=application/json&$skiptoken=2'
                    value = @(@{ userPrincipalName = 'one@example.com'; reportPeriod = '30'; details = @(@{ windows = $true }) }) }
            }
        }

        Invoke-CollectorScript 'Get-M365AppUsage.ps1' @{ OutputPath = $script:folder }

        $global:LicenseUtilizationTestCalls.Count | Should -Be 2
        $global:LicenseUtilizationTestCalls[0] | Should -BeLike '*?$format=application/json'
        (@(Import-Csv -LiteralPath (Join-Path $script:folder 'usage-m365-apps.csv')).UserPrincipalName | Sort-Object) -join ',' | Should -Be 'one@example.com,two@example.com'
        Get-HeaderText -Path (Join-Path $script:folder 'usage-m365-apps.csv') | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples 'usage-m365-apps.csv'))
    }

    It 'refuses a next link it has already followed' {
        Mock Invoke-MgGraphRequest -MockWith {
            @{ '@odata.nextLink' = $Uri
                value = @(@{ userPrincipalName = 'one@example.com'; reportPeriod = '30'; details = @(@{ windows = $true }) }) }
        }

        { Invoke-CollectorScript 'Get-M365AppUsage.ps1' @{ OutputPath = $script:folder } } | Should -Throw '*already followed*'
    }

    It 'refuses a next link on another host' {
        Mock Invoke-MgGraphRequest -MockWith {
            @{ '@odata.nextLink' = 'https://evil.example.net/page2'; value = @(@{ userPrincipalName = 'one@example.com' }) }
        }

        { Invoke-CollectorScript 'Get-M365AppUsage.ps1' @{ OutputPath = $script:folder } } | Should -Throw '*Refusing to follow*'
    }
}

Describe 'Paging is followed for the paged Graph reads' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
        Mock Get-MgUser -MockWith { @(New-MockLicensedUser -Id 'a'; New-MockLicensedUser -Id 'b') }
        Mock Get-MgSubscribedSku -MockWith { New-MockSku }
    }
    AfterEach { Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue }

    It 'asks Get-MgUser for every page (-All) and writes every user' {
        Invoke-CollectorScript 'Get-UserLicenses.ps1' @{ OutputPath = $script:folder }

        Should -Invoke Get-MgUser -Times 1 -Exactly -ParameterFilter { $All }
        @(Import-Csv -LiteralPath (Join-Path $script:folder 'user-licenses.csv')).Count | Should -Be 2
    }

    It 'asks for sign-in activity with -All' {
        Mock Get-MgUser -MockWith { New-MockSignInUser }
        Invoke-CollectorScript 'Get-UserSignInActivity.ps1' @{ OutputPath = $script:folder }

        Should -Invoke Get-MgUser -Times 1 -Exactly -ParameterFilter { $All -and $Property -contains 'signInActivity' }
    }

    It 'asks for subscribed SKUs with -All' {
        Invoke-CollectorScript 'Get-SubscribedSkus.ps1' @{ OutputPath = $script:folder }

        Should -Invoke Get-MgSubscribedSku -Times 1 -Exactly -ParameterFilter { $All }
    }
}

Describe 'Each connection targets the endpoints of its -Environment' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-MgGraph -ModuleName M365ReportLibrary -MockWith { }
        Mock Get-MgSubscribedSku -MockWith { @() }
        Mock Get-MgUser -MockWith { @() }
        Mock Get-MgAdminReportSetting -MockWith { New-MockReportSettings }
    }
    AfterEach { Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue }

    # https://learn.microsoft.com/graph/deployments: GCC calls the global service;
    # GCC High is the USGov environment.
    It '<Script> signs in to <Graph> for <Environment>' -ForEach @(
        foreach ($script in 'Get-SubscribedSkus.ps1', 'Get-UserLicenses.ps1', 'Get-UserSignInActivity.ps1', 'Get-ReportSettings.ps1') {
            @{ Script = $script; Environment = 'Commercial'; Graph = 'Global' }
            @{ Script = $script; Environment = 'GCC'; Graph = 'Global' }
            @{ Script = $script; Environment = 'GCCHigh'; Graph = 'USGov' }
        }
    ) {
        Invoke-CollectorScript $Script @{ OutputPath = $script:folder; Environment = $Environment }

        $expected = $Graph
        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter { $Environment -eq $expected }
    }

    It 'requests the permission each source adds to the shared scopes' -ForEach @(
        @{ Script = 'Get-SubscribedSkus.ps1'; Scope = 'LicenseAssignment.Read.All' }
        @{ Script = 'Get-UserSignInActivity.ps1'; Scope = 'AuditLog.Read.All' }
        @{ Script = 'Get-ReportSettings.ps1'; Scope = 'ReportSettings.Read.All' }
    ) {
        Invoke-CollectorScript $Script @{ OutputPath = $script:folder }

        $wanted = $Scope
        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter { $Scopes -contains $wanted }
    }

    It 'builds the usage report URI on the Graph host of the cloud' -ForEach @(
        @{ Environment = 'Commercial'; Base = 'https://graph.microsoft.com' }
        @{ Environment = 'GCC'; Base = 'https://graph.microsoft.com' }
        @{ Environment = 'GCCHigh'; Base = 'https://graph.microsoft.us' }
    ) {
        # GCC High is NotAvailable in the real schema, so a copy marks it Available to
        # prove which host a usage report would call.
        $text = Get-Content -LiteralPath (Join-Path $script:Collectors 'LicenseUtilizationSchema.psd1') -Raw
        $altered = [regex]::Replace($text, "(?s)(ActiveUserUsage\s*=\s*@\{.*?GCCHigh\s*=\s*@\{\s*Status\s*=\s*')NotAvailable", '${1}Available')
        $altered | Should -Not -Be $text
        $schemaCopy = Join-Path $script:folder 'schema.psd1'
        Set-Content -LiteralPath $schemaCopy -Value $altered
        Mock Invoke-MgGraphRequest -MockWith { Set-Content -LiteralPath $OutputFilePath -Value (New-UsageCsvText -Key 'ActiveUserUsage') }

        Invoke-CollectorScript 'Get-ActiveUserUsage.ps1' @{ OutputPath = $script:folder; Environment = $Environment; SchemaPath = $schemaCopy }

        $expected = $Base
        Should -Invoke Invoke-MgGraphRequest -Times 1 -Exactly -ParameterFilter { $Uri.StartsWith("$expected/v1.0/reports/getOffice365ActiveUserDetail") }
    }
}

Describe 'A source the doc marks NotAvailable writes the header only' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-MgGraph -ModuleName M365ReportLibrary -MockWith { }
        Mock Invoke-MgGraphRequest -MockWith { throw 'must not be called' }
    }
    AfterEach { Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue }

    It '<Script> writes <Csv> with its header and no rows in GCC High, without signing in' -ForEach @(
        $script:UsageCases + @(@{ Script = 'Get-M365AppUsage.ps1'; Key = 'M365AppUsage'; Csv = 'usage-m365-apps.csv' })
    ) {
        Invoke-CollectorScript $Script @{ OutputPath = $script:folder; Environment = 'GCCHigh' }

        $produced = Join-Path $script:folder $Csv
        Get-HeaderText -Path $produced | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples "gcchigh/$Csv"))
        @(Import-Csv -LiteralPath $produced).Count | Should -Be 0
        Get-LogText -Folder $script:folder | Should -Match 'unavailable in GCCHigh'
        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 0 -Exactly
        Should -Invoke Invoke-MgGraphRequest -Times 0 -Exactly
    }
}

Describe 'An UNVERIFIED source is attempted and warned about' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
    }
    AfterEach { Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue }

    It 'reads signInActivity in GCC High and logs UNVERIFIED (BRO-243)' {
        Mock Get-MgUser -MockWith { New-MockSignInUser }
        Invoke-CollectorScript 'Get-UserSignInActivity.ps1' @{ OutputPath = $script:folder; Environment = 'GCCHigh' }

        Should -Invoke Get-MgUser -Times 1 -Exactly
        Get-LogText -Folder $script:folder | Should -Match 'UNVERIFIED'
        @(Import-Csv -LiteralPath (Join-Path $script:folder 'user-signin-activity.csv')).Count | Should -Be 1
    }

    It 'reads report settings in GCC High and logs UNVERIFIED' {
        Mock Get-MgAdminReportSetting -MockWith { New-MockReportSettings }
        Invoke-CollectorScript 'Get-ReportSettings.ps1' @{ OutputPath = $script:folder; Environment = 'GCCHigh' }

        Should -Invoke Get-MgAdminReportSetting -Times 1 -Exactly
        Get-LogText -Folder $script:folder | Should -Match 'UNVERIFIED'
    }
}

Describe 'A refused read is logged, not hidden' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
    }
    AfterEach { Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue }

    It 'skips with a warning and the header only when the tenant lacks the Entra ID licence' {
        Mock Get-MgUser -MockWith { throw 'Request requires an Microsoft Entra ID P1 license (RequiresPremiumLicense)' }

        { Invoke-CollectorScript 'Get-UserSignInActivity.ps1' @{ OutputPath = $script:folder } } | Should -Not -Throw

        @(Import-Csv -LiteralPath (Join-Path $script:folder 'user-signin-activity.csv')).Count | Should -Be 0
        Get-LogText -Folder $script:folder | Should -Match 'not licensed'
    }

    It 'throws when the refusal names a permission, instead of skipping as unlicensed' {
        Mock Get-MgSubscribedSku -MockWith { throw 'Application must have one of the following permissions: LicenseAssignment.Read.All' }

        { Invoke-CollectorScript 'Get-SubscribedSkus.ps1' @{ OutputPath = $script:folder } } | Should -Throw '*unavailable*'
    }

    It 'logs an error, leaves the header and throws on any other failure' {
        Mock Get-MgSubscribedSku -MockWith { throw 'Insufficient privileges to complete the operation.' }

        { Invoke-CollectorScript 'Get-SubscribedSkus.ps1' @{ OutputPath = $script:folder } } | Should -Throw '*unavailable*'

        Test-Path -LiteralPath (Join-Path $script:folder 'subscribed-skus.csv') | Should -BeTrue
        Get-LogText -Folder $script:folder | Should -Match 'Insufficient privileges'
    }
}

Describe 'Run-All.ps1' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
        Mock Invoke-EntraUserCollector -MockWith { }
        Mock Get-MgSubscribedSku -MockWith { New-MockSku }
        Mock Get-MgUser -MockWith { New-MockLicensedUser }
        Mock Get-MgAdminReportSetting -MockWith { New-MockReportSettings }
        Mock Get-MgUserLicenseDetail -MockWith { New-MockLicenseDetail }
        Mock Invoke-MgGraphRequest -MockWith {
            if ($OutputFilePath) { Set-Content -LiteralPath $OutputFilePath -Value (New-UsageCsvText -Key 'ActiveUserUsage') }
            else { @{ value = @() } }
        }
    }
    AfterEach { Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue }

    It 'runs the shared users collector first and leaves license-details.csv out unless asked' {
        Invoke-CollectorScript 'Run-All.ps1' @{ OutputPath = $script:folder }

        Should -Invoke Invoke-EntraUserCollector -Times 1 -Exactly
        Test-Path -LiteralPath (Join-Path $script:folder 'license-details.csv') | Should -BeFalse
        foreach ($file in 'subscribed-skus.csv', 'sku-service-plans.csv', 'user-licenses.csv', 'user-signin-activity.csv', 'report-settings.csv', 'usage-active-users.csv', 'usage-copilot.csv') {
            Test-Path -LiteralPath (Join-Path $script:folder $file) | Should -BeTrue
        }
    }

    It 'adds license-details.csv with -IncludeLicenseDetails' {
        Invoke-CollectorScript 'Run-All.ps1' @{ OutputPath = $script:folder; IncludeLicenseDetails = $true }

        Test-Path -LiteralPath (Join-Path $script:folder 'license-details.csv') | Should -BeTrue
    }

    It 'in GCC High skips the usage reports with header-only files' {
        Invoke-CollectorScript 'Run-All.ps1' @{ OutputPath = $script:folder; Environment = 'GCCHigh' }

        @(Import-Csv -LiteralPath (Join-Path $script:folder 'usage-active-users.csv')).Count | Should -Be 0
        @(Import-Csv -LiteralPath (Join-Path $script:folder 'subscribed-skus.csv')).Count | Should -BeGreaterThan 0
    }

    It 'keeps going when one collector fails, then throws' {
        Mock Get-MgSubscribedSku -MockWith { throw 'Insufficient privileges' }

        { Invoke-CollectorScript 'Run-All.ps1' @{ OutputPath = $script:folder } } | Should -Throw '*collector(s) stopped*'

        Test-Path -LiteralPath (Join-Path $script:folder 'user-licenses.csv') | Should -BeTrue
    }
}
