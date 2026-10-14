#Requires -Version 7.0

# Sources 9 to 11: the collectors behind the E5 step-up features on the License value page
# that were not collected before (Access reviews, Defender for Office P2, Teams Phone).

BeforeDiscovery {
    $script:Cases = @(
        @{ Script = 'Get-AccessReviews.ps1'; Key = 'AccessReviewDefinitions'; Csv = 'access-reviews.csv'; Path = '/v1.0/identityGovernance/accessReviews/definitions'; Scope = 'AccessReview.Read.All'; Item = { New-MockAccessReview } }
        @{ Script = 'Get-AttackSimulations.ps1'; Key = 'AttackSimulations'; Csv = 'attack-simulations.csv'; Path = '/v1.0/security/attackSimulation/simulations'; Scope = 'AttackSimulation.Read.All'; Item = { New-MockSimulation } }
        @{ Script = 'Get-TeamsPstnCalls.ps1'; Key = 'TeamsPstnCalls'; Csv = 'teams-pstn-calls.csv'; Path = '/v1.0/communications/callRecords/getPstnCalls('; Scope = 'CallRecords.Read.All'; Item = { New-MockPstnCall } }
    )
}

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

    function New-TestFolder {
        $path = Join-Path ([System.IO.Path]::GetTempPath()) ('license-value-' + [guid]::NewGuid().ToString('N'))
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

    # accessReviewScheduleDefinition: https://learn.microsoft.com/graph/api/accessreviewset-list-definitions
    function New-MockAccessReview {
        param([string]$Id = 'review-1')
        @{
            id = $Id; displayName = 'Review guest access'; status = 'InProgress'
            scope = @{ query = './members/microsoft.graph.user/?$count=true'; queryType = 'MicrosoftGraph' }
            reviewers = @(@{ query = './manager'; queryType = 'MicrosoftGraph'; queryRoot = 'decisions' })
            settings = @{ autoApplyDecisionsEnabled = $true; recommendationsEnabled = $false
                recurrence = @{ pattern = @{ type = 'absoluteMonthly'; interval = 3 } } }
        }
    }

    # simulation: https://learn.microsoft.com/graph/api/attacksimulationroot-list-simulations
    function New-MockSimulation {
        param([string]$Id = 'sim-1')
        @{
            id = $Id; displayName = 'Sample Simulation'; attackType = 'social'; attackTechnique = 'credentialHarvesting'
            status = 'scheduled'; isAutomated = $false; payloadDeliveryPlatform = 'email'
            launchDateTime = '2026-08-01T02:01:01Z'; completionDateTime = '2026-08-07T01:01:01Z'
        }
    }

    # pstnCallLogRow: https://learn.microsoft.com/graph/api/callrecords-callrecord-getpstncalls
    function New-MockPstnCall {
        param([string]$Id = 'call-1')
        @{
            id = $Id; userId = 'user-1'; userPrincipalName = 'user-1@example.com'
            startDateTime = '2026-08-01T00:00:08Z'; duration = 219; callType = 'user_in'
            usageCountryCode = 'US'; licenseCapability = 'MCOPSTNU'; inventoryType = 'Subscriber'
            calleeNumber = '+15555550100'; callerNumber = '+15555550101'
        }
    }
}

Describe 'Each License value collector writes the columns of its sample' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
    }
    AfterEach { Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue }

    It '<Script> writes <Csv> matching the schema and its sample, stamped with the run date' -ForEach $script:Cases {
        $mock = & $Item
        Mock Invoke-MgGraphRequest -MockWith { @{ value = @($mock) } }.GetNewClosure()

        Invoke-CollectorScript $Script @{ OutputPath = $script:folder }

        $produced = Join-Path $script:folder $Csv
        Get-HeaderText -Path $produced | Should -Be ($script:Schema[$Key] -join ',')
        Get-HeaderText -Path $produced | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples $Csv))
        $rows = @(Import-Csv -LiteralPath $produced)
        $rows.Count | Should -Be 1
        $rows[0].RunDate | Should -Be $script:Today
    }

    It 'Get-AccessReviews.ps1 flattens scope, reviewers and recurrence into columns' {
        Mock Invoke-MgGraphRequest -MockWith { @{ value = @((New-MockAccessReview)) } }

        Invoke-CollectorScript 'Get-AccessReviews.ps1' @{ OutputPath = $script:folder }

        $row = @(Import-Csv -LiteralPath (Join-Path $script:folder 'access-reviews.csv'))[0]
        $row.DefinitionId | Should -Be 'review-1'
        $row.Status | Should -Be 'InProgress'
        $row.ScopeQuery | Should -BeLike './members/*'
        $row.ReviewerCount | Should -Be '1'
        $row.RecurrencePattern | Should -Be 'absoluteMonthly'
        $row.AutoApplyDecisionsEnabled | Should -Be 'True'
        $row.RecommendationsEnabled | Should -Be 'False'
    }

    It 'Get-TeamsPstnCalls.ps1 does not write the caller or callee number' {
        Mock Invoke-MgGraphRequest -MockWith { @{ value = @((New-MockPstnCall)) } }

        Invoke-CollectorScript 'Get-TeamsPstnCalls.ps1' @{ OutputPath = $script:folder }

        (Get-Content -LiteralPath (Join-Path $script:folder 'teams-pstn-calls.csv') -Raw) | Should -Not -Match '5555550'
    }

    It 'a second run appends a new RunDate block instead of overwriting' {
        Mock Invoke-MgGraphRequest -MockWith { @{ value = @((New-MockSimulation)) } }
        $path = Join-Path $script:folder 'attack-simulations.csv'
        Export-AppendCsv -Path $path -Rows @([pscustomobject]@{ RunDate = '2026-01-01'; SimulationId = 'sim-1' }) `
            -Column $script:Schema.AttackSimulations -KeyColumn @('RunDate', 'SimulationId')

        Invoke-CollectorScript 'Get-AttackSimulations.ps1' @{ OutputPath = $script:folder }

        @(Import-Csv -LiteralPath $path).RunDate | Should -Be @('2026-01-01', $script:Today)
    }
}

Describe 'Paging is followed for the License value reads' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
    }
    AfterEach { Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue }

    It '<Script> reads every page until @odata.nextLink is absent' -ForEach $script:Cases {
        $global:LicenseValueTestCalls = [System.Collections.Generic.List[string]]::new()
        $global:LicenseValueTestItem = $Item
        Mock Invoke-MgGraphRequest -MockWith {
            $global:LicenseValueTestCalls.Add($Uri)
            $id = if ($Uri -like '*page=2*') { 'two' } else { 'one' }
            $row = & $global:LicenseValueTestItem
            $row.id = $id
            if ($id -eq 'one') { @{ '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/next?page=2'; value = @($row) } }
            else { @{ value = @($row) } }
        }

        Invoke-CollectorScript $Script @{ OutputPath = $script:folder }

        $global:LicenseValueTestCalls.Count | Should -Be 2
        $rows = @(Import-Csv -LiteralPath (Join-Path $script:folder $Csv))
        $rows.Count | Should -Be 2
    }

    # simulation.attackTechnique: oAuthConsentGrant is returned as unknownFutureValue unless
    # the request carries Prefer: include-unknown-enum-members.
    # https://learn.microsoft.com/graph/api/resources/simulation#properties
    It 'Get-AttackSimulations.ps1 sends Prefer include-unknown-enum-members on every page' {
        Mock Invoke-MgGraphRequest -MockWith {
            if ($Uri -like '*page=2*') { @{ value = @((New-MockSimulation -Id 'two')) } }
            else { @{ '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/next?page=2'; value = @((New-MockSimulation -Id 'one')) } }
        }

        Invoke-CollectorScript 'Get-AttackSimulations.ps1' @{ OutputPath = $script:folder }

        Should -Invoke Invoke-MgGraphRequest -Times 2 -Exactly -ParameterFilter {
            $Headers -and $Headers['Prefer'] -eq 'include-unknown-enum-members'
        }
    }

    It 'refuses a next link on another host' {
        Mock Invoke-MgGraphRequest -MockWith {
            @{ '@odata.nextLink' = 'https://example.com/v1.0/next'; value = @((New-MockAccessReview)) }
        }

        { Invoke-CollectorScript 'Get-AccessReviews.ps1' @{ OutputPath = $script:folder } } | Should -Throw '*Refusing to follow*'
    }
}

Describe 'Each License value connection targets the endpoints of its -Environment' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-MgGraph -ModuleName M365ReportLibrary -MockWith { }
        Mock Invoke-MgGraphRequest -MockWith { @{ value = @() } }
    }
    AfterEach { Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue }

    # https://learn.microsoft.com/graph/deployments: GCC calls the global service;
    # GCC High is the USGov environment.
    It '<Script> signs in to <Graph> and reads from <Base> for <Environment>' -ForEach @(
        foreach ($case in $script:Cases) {
            foreach ($cloud in @(
                    @{ Environment = 'Commercial'; Graph = 'Global'; Base = 'https://graph.microsoft.com' }
                    @{ Environment = 'GCC'; Graph = 'Global'; Base = 'https://graph.microsoft.com' }
                    @{ Environment = 'GCCHigh'; Graph = 'USGov'; Base = 'https://graph.microsoft.us' })) {
                # Attack simulations are NotAvailable in GCC High; that is tested below.
                if ($case.Script -eq 'Get-AttackSimulations.ps1' -and $cloud.Environment -eq 'GCCHigh') { continue }
                @{ Script = $case.Script; Path = $case.Path } + $cloud
            }
        }
    ) {
        Invoke-CollectorScript $Script @{ OutputPath = $script:folder; Environment = $Environment }

        $expected = $Graph
        $uriStart = $Base + $Path
        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter { $Environment -eq $expected }
        Should -Invoke Invoke-MgGraphRequest -Times 1 -Exactly -ParameterFilter { $Uri.StartsWith($uriStart) }
    }

    It '<Script> requests <Scope> on top of the shared scopes' -ForEach $script:Cases {
        Invoke-CollectorScript $Script @{ OutputPath = $script:folder }

        $wanted = $Scope
        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter { $Scopes -contains $wanted }
    }

    It 'Get-TeamsPstnCalls.ps1 asks for the last -Days days and no more than 90' {
        Invoke-CollectorScript 'Get-TeamsPstnCalls.ps1' @{ OutputPath = $script:folder; Days = 10 }

        Should -Invoke Invoke-MgGraphRequest -Times 1 -Exactly -ParameterFilter {
            $match = [regex]::Match($Uri, 'fromDateTime=([^,]+),toDateTime=([^)]+)\)')
            $match.Success -and ([datetime]$match.Groups[2].Value - [datetime]$match.Groups[1].Value).TotalDays -ge 9.99 -and
            ([datetime]$match.Groups[2].Value - [datetime]$match.Groups[1].Value).TotalDays -le 10.01
        }
        { Invoke-CollectorScript 'Get-TeamsPstnCalls.ps1' @{ OutputPath = $script:folder; Days = 91 } } | Should -Throw
    }
}

Describe 'A License value source the doc marks NotAvailable writes the header only' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-MgGraph -ModuleName M365ReportLibrary -MockWith { }
        Mock Invoke-MgGraphRequest -MockWith { throw 'must not be called' }
    }
    AfterEach { Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue }

    # The attack simulation page marks US Government L4 and L5 unavailable.
    It 'Get-AttackSimulations.ps1 writes the header and no rows in GCC High, without signing in' {
        Invoke-CollectorScript 'Get-AttackSimulations.ps1' @{ OutputPath = $script:folder; Environment = 'GCCHigh' }

        $produced = Join-Path $script:folder 'attack-simulations.csv'
        Get-HeaderText -Path $produced | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples 'gcchigh/attack-simulations.csv'))
        @(Import-Csv -LiteralPath $produced).Count | Should -Be 0
        Get-LogText -Folder $script:folder | Should -Match 'unavailable in GCCHigh'
        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 0 -Exactly
        Should -Invoke Invoke-MgGraphRequest -Times 0 -Exactly
    }
}

Describe 'An UNVERIFIED License value source is attempted and warned about' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
        Mock Invoke-MgGraphRequest -MockWith { @{ value = @((New-MockSimulation)) } }
    }
    AfterEach { Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue }

    It 'reads attack simulations in GCC and logs UNVERIFIED' {
        Invoke-CollectorScript 'Get-AttackSimulations.ps1' @{ OutputPath = $script:folder; Environment = 'GCC' }

        Should -Invoke Invoke-MgGraphRequest -Times 1 -Exactly
        Get-LogText -Folder $script:folder | Should -Match 'UNVERIFIED'
        @(Import-Csv -LiteralPath (Join-Path $script:folder 'attack-simulations.csv')).Count | Should -Be 1
    }
}

Describe 'A refused License value read is logged, not hidden' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
    }
    AfterEach { Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue }

    It 'Get-TeamsPstnCalls.ps1 logs an error, leaves the header and throws when the token is refused' {
        Mock Invoke-MgGraphRequest -MockWith { throw 'Forbidden: delegated access is not supported' }

        { Invoke-CollectorScript 'Get-TeamsPstnCalls.ps1' @{ OutputPath = $script:folder } } | Should -Throw '*unavailable*'

        Get-HeaderText -Path (Join-Path $script:folder 'teams-pstn-calls.csv') | Should -Be ($script:Schema.TeamsPstnCalls -join ',')
        Get-LogText -Folder $script:folder | Should -Match 'application permission'
    }
}

Describe 'The License value sources are recorded in the schema and Run-All' {
    It '<Key> has all three clouds, and attack simulations are NotAvailable in GCC High' -ForEach $script:Cases {
        foreach ($cloud in 'Commercial', 'GCC', 'GCCHigh') {
            $script:Schema.SourceAvailability[$Key][$cloud].Status | Should -BeIn @('Available', 'NotAvailable', 'Unverified')
            $script:Schema.SourceAvailability[$Key][$cloud].Reference | Should -Match '^https://learn\.microsoft\.com/'
        }
        if ($Key -eq 'AttackSimulations') { $script:Schema.SourceAvailability[$Key].GCCHigh.Status | Should -Be 'NotAvailable' }
    }

    It 'Run-All runs the PSTN collector only with -AppId' {
        $folder = New-TestFolder
        try {
            Mock Connect-M365Service -MockWith { }
            Mock Invoke-EntraUserCollector -MockWith { }
            Mock Get-MgSubscribedSku -MockWith { @() }
            Mock Get-MgUser -MockWith { @() }
            Mock Get-MgAdminReportSetting -MockWith { New-Object psobject -Property @{ AdditionalProperties = @{ displayConcealedNames = $false } } }
            Mock Invoke-MgGraphRequest -MockWith {
                if ($OutputFilePath) { Set-Content -LiteralPath $OutputFilePath -Value '"User Principal Name"' } else { @{ value = @() } }
            }

            Invoke-CollectorScript 'Run-All.ps1' @{ OutputPath = $folder }
            Test-Path -LiteralPath (Join-Path $folder 'access-reviews.csv') | Should -BeTrue
            Test-Path -LiteralPath (Join-Path $folder 'attack-simulations.csv') | Should -BeTrue
            Test-Path -LiteralPath (Join-Path $folder 'teams-pstn-calls.csv') | Should -BeFalse
            Get-LogText -Folder $folder | Should -Match 'Skipping teams-pstn-calls.csv'

            Invoke-CollectorScript 'Run-All.ps1' @{ OutputPath = $folder; AppId = '00000000-0000-0000-0000-000000000000'; CertificateThumbprint = 'AB' }
            Test-Path -LiteralPath (Join-Path $folder 'teams-pstn-calls.csv') | Should -BeTrue
        }
        finally { Remove-Item -LiteralPath $folder -Recurse -Force -ErrorAction SilentlyContinue }
    }
}
