#Requires -Version 7.0

BeforeAll {
    $script:ModulePath = Join-Path $PSScriptRoot '../M365ReportLibrary.psm1'
    . (Join-Path $PSScriptRoot 'TenantCmdletStubs.ps1')
    Import-Module $script:ModulePath -Force

    function New-TestFolder {
        $path = Join-Path ([System.IO.Path]::GetTempPath()) ('m365rl-' + [guid]::NewGuid().ToString('N'))
        New-Item -Path $path -ItemType Directory -Force | Out-Null
        return $path
    }
}

AfterAll {
    Remove-Module M365ReportLibrary -Force -ErrorAction SilentlyContinue
}

Describe 'Get-M365ServiceEndpoint' {
    It 'sends Graph in <Environment> to <Expected>' -ForEach @(
        @{ Environment = 'Commercial'; Expected = 'https://graph.microsoft.com'; GraphEnvironment = 'Global' }
        @{ Environment = 'GCC'; Expected = 'https://graph.microsoft.com'; GraphEnvironment = 'Global' }
        @{ Environment = 'GCCHigh'; Expected = 'https://graph.microsoft.us'; GraphEnvironment = 'USGov' }
    ) {
        $endpoint = Get-M365ServiceEndpoint -Service Graph -Environment $Environment
        $endpoint.ResourceEndpoint | Should -Be $Expected
        $endpoint.GraphEnvironment | Should -Be $GraphEnvironment
    }

    It 'names the Exchange Online environment <Expected> in <Environment>' -ForEach @(
        @{ Environment = 'Commercial'; Expected = 'O365Default' }
        @{ Environment = 'GCC'; Expected = 'O365Default' }
        @{ Environment = 'GCCHigh'; Expected = 'O365USGovGCCHigh' }
    ) {
        (Get-M365ServiceEndpoint -Service ExchangeOnline -Environment $Environment).ExchangeEnvironmentName |
            Should -Be $Expected
    }

    It 'leaves the Security & Compliance connection URI unset in <Environment>' -ForEach @(
        @{ Environment = 'Commercial' }
        @{ Environment = 'GCC' }
    ) {
        (Get-M365ServiceEndpoint -Service SecurityCompliance -Environment $Environment).ConnectionUri |
            Should -BeNullOrEmpty
    }

    It 'points Security & Compliance in GCC High at the .us endpoints' {
        $endpoint = Get-M365ServiceEndpoint -Service SecurityCompliance -Environment GCCHigh
        $endpoint.ConnectionUri | Should -Be 'https://ps.compliance.protection.office365.us/powershell-liveid/'
        $endpoint.AzureADAuthorizationEndpointUri | Should -Be 'https://login.microsoftonline.us/organizations'
    }

    It 'defaults to the Commercial cloud' {
        (Get-M365ServiceEndpoint -Service Graph).Environment | Should -Be 'Commercial'
    }
}

Describe 'Connect-M365Service' {
    BeforeEach {
        Mock Connect-MgGraph -ModuleName M365ReportLibrary -MockWith { }
        Mock Connect-ExchangeOnline -ModuleName M365ReportLibrary -MockWith { }
        Mock Connect-IPPSSession -ModuleName M365ReportLibrary -MockWith { }
    }

    It 'connects Graph in <Environment> with -Environment <GraphEnvironment>' -ForEach @(
        @{ Environment = 'Commercial'; GraphEnvironment = 'Global' }
        @{ Environment = 'GCC'; GraphEnvironment = 'Global' }
        @{ Environment = 'GCCHigh'; GraphEnvironment = 'USGov' }
    ) {
        Connect-M365Service -Service Graph -Environment $Environment

        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter {
            $Environment -eq $GraphEnvironment
        }
    }

    It 'connects Exchange Online in <Environment> with -ExchangeEnvironmentName <Expected>' -ForEach @(
        @{ Environment = 'Commercial'; Expected = 'O365Default' }
        @{ Environment = 'GCC'; Expected = 'O365Default' }
        @{ Environment = 'GCCHigh'; Expected = 'O365USGovGCCHigh' }
    ) {
        Connect-M365Service -Service ExchangeOnline -Environment $Environment

        Should -Invoke Connect-ExchangeOnline -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter {
            $ExchangeEnvironmentName -eq $Expected
        }
    }

    It 'connects Security & Compliance in <Environment> without a connection URI' -ForEach @(
        @{ Environment = 'Commercial' }
        @{ Environment = 'GCC' }
    ) {
        Connect-M365Service -Service SecurityCompliance -Environment $Environment

        Should -Invoke Connect-IPPSSession -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter {
            [string]::IsNullOrEmpty($ConnectionUri)
        }
    }

    It 'connects Security & Compliance in GCC High through the .us endpoints' {
        Connect-M365Service -Service SecurityCompliance -Environment GCCHigh

        Should -Invoke Connect-IPPSSession -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter {
            $ConnectionUri -eq 'https://ps.compliance.protection.office365.us/powershell-liveid/' -and
            $AzureADAuthorizationEndpointUri -eq 'https://login.microsoftonline.us/organizations'
        }
    }

    It 'asks for read-only Graph scopes when signing in interactively' {
        Connect-M365Service -Service Graph

        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter {
            $Scopes -contains 'AuditLog.Read.All' -and ($Scopes | Where-Object { $_ -notlike '*.Read.All' }).Count -eq 0
        }
    }

    It 'signs in to Graph app-only when given a certificate' {
        Connect-M365Service -Service Graph -AppId '11111111-1111-1111-1111-111111111111' `
            -CertificateThumbprint 'ABC123' -TenantId 'contoso.onmicrosoft.com'

        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter {
            $ClientId -eq '11111111-1111-1111-1111-111111111111' -and
            $CertificateThumbprint -eq 'ABC123' -and
            $null -eq $Scopes
        }
    }

    It 'refuses app-only sign-in to Exchange Online without -Organization' {
        { Connect-M365Service -Service ExchangeOnline -AppId 'app' -CertificateThumbprint 'ABC123' } |
            Should -Throw '*requires -Organization*'
    }

    It 'refuses half a credential: <Given> without <Missing>' -ForEach @(
        @{ Given = '-AppId'; Missing = '-CertificateThumbprint'; Splat = @{ AppId = 'app' } }
        @{ Given = '-CertificateThumbprint'; Missing = '-AppId'; Splat = @{ CertificateThumbprint = 'ABC123' } }
    ) {
        # Falling back to interactive here would prompt, or sign in as whoever is at the
        # keyboard, on a run meant to be unattended.
        { Connect-M365Service -Service Graph @Splat } | Should -Throw '*needs both -AppId and -CertificateThumbprint*'
        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 0 -Exactly
    }

    It 'still signs in interactively when neither half is given' {
        { Connect-M365Service -Service Graph -AppId '' -CertificateThumbprint '' } | Should -Not -Throw
        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 1 -Exactly
    }
}

Describe 'Export-AppendCsv' {
    BeforeEach {
        $script:folder = New-TestFolder
        $script:csv = Join-Path $script:folder 'events.csv'
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'writes the header once and appends afterwards' {
        Export-AppendCsv -Path $script:csv -Rows @([pscustomobject]@{ Id = '1'; Name = 'one' }) -KeyColumn 'Id'
        Export-AppendCsv -Path $script:csv -Rows @([pscustomobject]@{ Id = '2'; Name = 'two' }) -KeyColumn 'Id'

        $lines = Get-Content -LiteralPath $script:csv
        $lines.Count | Should -Be 3
        $lines[0] | Should -Be '"Id","Name"'
        (Import-Csv -LiteralPath $script:csv).Id | Should -Be @('1', '2')
    }

    It 'skips rows whose key is already in the file' {
        Export-AppendCsv -Path $script:csv -Rows @(
            [pscustomobject]@{ Id = '1'; Name = 'one' }
            [pscustomobject]@{ Id = '2'; Name = 'two' }
        ) -KeyColumn 'Id'

        $result = Export-AppendCsv -Path $script:csv -KeyColumn 'Id' -PassThru -Rows @(
            [pscustomobject]@{ Id = '2'; Name = 'two again' }
            [pscustomobject]@{ Id = '3'; Name = 'three' }
        )

        $result.Written | Should -Be 1
        $result.Skipped | Should -Be 1
        (Import-Csv -LiteralPath $script:csv).Id | Should -Be @('1', '2', '3')
    }

    It 'skips a key repeated inside one batch' {
        $result = Export-AppendCsv -Path $script:csv -KeyColumn 'Id' -PassThru -Rows @(
            [pscustomobject]@{ Id = '1'; Name = 'one' }
            [pscustomobject]@{ Id = '1'; Name = 'one again' }
        )

        $result.Written | Should -Be 1
        $result.Skipped | Should -Be 1
    }

    It 'treats the combination of several key columns as the key' {
        Export-AppendCsv -Path $script:csv -Column @('Id', 'Name') -KeyColumn @('Id', 'Name') -Rows @(
            [pscustomobject]@{ Id = '1'; Name = 'one' }
        )

        $result = Export-AppendCsv -Path $script:csv -Column @('Id', 'Name') -KeyColumn @('Id', 'Name') -PassThru -Rows @(
            [pscustomobject]@{ Id = '1'; Name = 'one' }
            [pscustomobject]@{ Id = '1'; Name = 'two' }
        )

        $result.Written | Should -Be 1
        $result.Skipped | Should -Be 1
    }

    It 'throws when the incoming columns differ from the file header' {
        Export-AppendCsv -Path $script:csv -Rows @([pscustomobject]@{ Id = '1'; Name = 'one' })

        { Export-AppendCsv -Path $script:csv -Rows @([pscustomobject]@{ Id = '2'; Title = 'two' }) } |
            Should -Throw '*Column mismatch*'
    }

    It 'throws when the incoming columns are in a different order' {
        Export-AppendCsv -Path $script:csv -Rows @([pscustomobject]@{ Id = '1'; Name = 'one' })

        { Export-AppendCsv -Path $script:csv -Rows @([pscustomobject]@{ Name = 'two'; Id = '2' }) } |
            Should -Throw '*Column mismatch*'
    }

    It 'writes a header-only file when there are no rows' {
        Export-AppendCsv -Path $script:csv -Column @('Id', 'Name')

        Get-Content -LiteralPath $script:csv | Should -Be '"Id","Name"'
        @(Import-Csv -LiteralPath $script:csv).Count | Should -Be 0
    }

    It 'appends to a header-only file without repeating the header' {
        Export-AppendCsv -Path $script:csv -Column @('Id', 'Name')
        Export-AppendCsv -Path $script:csv -Column @('Id', 'Name') -Rows @([pscustomobject]@{ Id = '1'; Name = 'one' })

        (Get-Content -LiteralPath $script:csv).Count | Should -Be 2
    }

    It 'keeps the column order given by -Column, whatever order the row properties are in' {
        Export-AppendCsv -Path $script:csv -Column @('Id', 'Name') -Rows @([pscustomobject]@{ Name = 'one'; Id = '1' })

        Get-Content -LiteralPath $script:csv -TotalCount 1 | Should -Be '"Id","Name"'
    }

    It 'throws when it has neither -Column nor a row to take the columns from' {
        { Export-AppendCsv -Path $script:csv } | Should -Throw '*cannot determine the columns*'
    }

    It 'throws when a key column is not one of the columns' {
        { Export-AppendCsv -Path $script:csv -Column @('Id') -KeyColumn 'Missing' -Rows @([pscustomobject]@{ Id = '1' }) } |
            Should -Throw "*Key column 'Missing'*"
    }
}

Describe 'Get-CsvWatermark' {
    BeforeEach {
        $script:folder = New-TestFolder
        $script:csv = Join-Path $script:folder 'events.csv'
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'returns the latest timestamp in the file' {
        Export-AppendCsv -Path $script:csv -Rows @(
            [pscustomobject]@{ When = '2026-01-02T03:04:05Z'; Id = '1' }
            [pscustomobject]@{ When = '2026-03-04T05:06:07Z'; Id = '2' }
            [pscustomobject]@{ When = '2026-02-03T04:05:06Z'; Id = '3' }
        )

        $watermark = Get-CsvWatermark -Path $script:csv -Column 'When'
        $watermark | Should -BeOfType [datetime]
        $watermark.ToString('yyyy-MM-ddTHH:mm:ssZ') | Should -Be '2026-03-04T05:06:07Z'
    }

    It 'returns nothing for a file that does not exist' {
        Get-CsvWatermark -Path (Join-Path $script:folder 'missing.csv') -Column 'When' | Should -BeNullOrEmpty
    }

    It 'returns nothing for a header-only file' {
        Export-AppendCsv -Path $script:csv -Column @('When', 'Id')
        Get-CsvWatermark -Path $script:csv -Column 'When' | Should -BeNullOrEmpty
    }

    It 'returns nothing when the column is not in the file' {
        Export-AppendCsv -Path $script:csv -Rows @([pscustomobject]@{ When = '2026-01-01T00:00:00Z'; Id = '1' })
        Get-CsvWatermark -Path $script:csv -Column 'Missing' | Should -BeNullOrEmpty
    }

    It 'ignores values it cannot read as a timestamp' {
        Export-AppendCsv -Path $script:csv -Rows @(
            [pscustomobject]@{ When = 'not a date'; Id = '1' }
            [pscustomobject]@{ When = '2026-05-06T07:08:09Z'; Id = '2' }
            [pscustomobject]@{ When = ''; Id = '3' }
        )

        (Get-CsvWatermark -Path $script:csv -Column 'When').ToString('yyyy-MM-ddTHH:mm:ssZ') |
            Should -Be '2026-05-06T07:08:09Z'
    }
}

Describe 'Get-CsvLatestSnapshot' {
    BeforeEach {
        $script:folder = New-TestFolder
        $script:csv = Join-Path $script:folder 'guests.csv'
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'returns only the rows of the most recent RunDate' {
        Export-AppendCsv -Path $script:csv -Rows @(
            [pscustomobject]@{ RunDate = '2026-01-01'; Id = 'a' }
            [pscustomobject]@{ RunDate = '2026-02-01'; Id = 'b' }
            [pscustomobject]@{ RunDate = '2026-02-01'; Id = 'c' }
        )

        $rows = @(Get-CsvLatestSnapshot -Path $script:csv)
        $rows.Count | Should -Be 2
        $rows.Id | Should -Be @('b', 'c')
    }

    It 'returns nothing for a header-only file' {
        Export-AppendCsv -Path $script:csv -Column @('RunDate', 'Id')
        @(Get-CsvLatestSnapshot -Path $script:csv).Count | Should -Be 0
    }
}

Describe 'Write-CollectorLog' {
    BeforeEach { $script:folder = New-TestFolder }
    AfterEach { Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue }

    It 'appends to run.log in the output folder' {
        Write-CollectorLog -OutputPath $script:folder -Message 'first'
        Write-CollectorLog -OutputPath $script:folder -Message 'second' -Source 'guests'

        $lines = Get-Content -LiteralPath (Join-Path $script:folder 'run.log')
        $lines.Count | Should -Be 2
        $lines[0] | Should -Match '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z \[Info\] first$'
        $lines[1] | Should -Match '\[Info\] \[guests\] second$'
    }

    It 'creates the output folder when it is missing' {
        $nested = Join-Path $script:folder 'nested'
        Write-CollectorLog -OutputPath $nested -Message 'hello'
        Test-Path -LiteralPath (Join-Path $nested 'run.log') | Should -BeTrue
    }

    It 'records the level' {
        Write-CollectorLog -OutputPath $script:folder -Message 'careful' -Level Warning -WarningAction SilentlyContinue
        Get-Content -LiteralPath (Join-Path $script:folder 'run.log') | Should -Match '\[Warning\] careful'
    }
}

Describe 'Split-DateRange' {
    It 'covers the range in whole windows' {
        $start = [datetime]'2026-01-01T00:00:00Z'
        $windows = @(Split-DateRange -Start $start -End $start.AddDays(3) -WindowMinutes 1440)

        $windows.Count | Should -Be 3
        $windows[0].Start | Should -Be $start
        $windows[-1].End | Should -Be $start.AddDays(3)
    }

    It 'shortens the last window to land on the end' {
        $start = [datetime]'2026-01-01T00:00:00Z'
        $windows = @(Split-DateRange -Start $start -End $start.AddHours(30) -WindowMinutes 1440)

        $windows.Count | Should -Be 2
        $windows[-1].End | Should -Be $start.AddHours(30)
    }

    It 'returns nothing when the range is empty or inverted' {
        @(Split-DateRange -Start ([datetime]'2026-01-02') -End ([datetime]'2026-01-01')).Count | Should -Be 0
    }
}

Describe 'ConvertTo-CsvTimestamp' {
    It 'writes UTC in a sortable, round-trippable shape' {
        ConvertTo-CsvTimestamp ([datetime]::new(2026, 3, 4, 5, 6, 7, [System.DateTimeKind]::Utc)) |
            Should -Be '2026-03-04T05:06:07Z'
    }

    It 'returns an empty string for a missing value' {
        ConvertTo-CsvTimestamp $null | Should -Be ''
        ConvertTo-CsvTimestamp '' | Should -Be ''
    }

    It 'normalises a string timestamp' {
        ConvertTo-CsvTimestamp '2026-03-04T05:06:07+00:00' | Should -Be '2026-03-04T05:06:07Z'
    }
}

Describe 'Invoke-EntraUserCollector' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-MgGraph -ModuleName M365ReportLibrary -MockWith { }
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'writes users.csv with the library column order' {
        Mock Get-MgUser -ModuleName M365ReportLibrary -MockWith {
            [pscustomobject]@{
                Id                   = 'user-1'
                DisplayName          = 'Avery Abara'
                UserPrincipalName    = 'avery.abara@example.com'
                Mail                 = 'avery.abara@example.com'
                UserType             = 'Member'
                AccountEnabled       = $true
                CreatedDateTime      = [datetime]'2025-01-02T03:04:05Z'
                Department           = 'Engineering'
                JobTitle             = 'Engineering Specialist'
                City                 = 'Seattle'
                Country              = 'US'
                Manager              = [pscustomobject]@{
                    Id                   = 'user-0'
                    AdditionalProperties = @{ userPrincipalName = 'blair.bergstrom@example.com' }
                }
                AdditionalProperties = @{}
            }
        }

        Invoke-EntraUserCollector -OutputPath $script:folder -SkipConnect

        $csv = Join-Path $script:folder 'users.csv'
        (Get-CsvHeaderColumn -Path $csv) -join ',' | Should -Be ((Get-EntraUserCsvColumn) -join ',')

        $row = Import-Csv -LiteralPath $csv
        $row.Id | Should -Be 'user-1'
        $row.ManagerUserPrincipalName | Should -Be 'blair.bergstrom@example.com'
        $row.CreatedDateTime | Should -Be '2025-01-02T03:04:05Z'
        $row.RunDate | Should -Be ([datetime]::UtcNow.ToString('yyyy-MM-dd'))
    }

    It 'falls back to reading the manager per user when the expand is rejected' {
        Mock Get-MgUser -ModuleName M365ReportLibrary -MockWith { throw 'Invalid expand clause' } -ParameterFilter { $null -ne $ExpandProperty }
        Mock Get-MgUser -ModuleName M365ReportLibrary -MockWith {
            [pscustomobject]@{
                Id                   = 'user-1'; DisplayName = 'Avery Abara'
                UserPrincipalName    = 'avery.abara@example.com'; Mail = 'avery.abara@example.com'
                UserType             = 'Member'; AccountEnabled = $true
                CreatedDateTime      = [datetime]'2025-01-02T03:04:05Z'
                Department           = 'Engineering'; JobTitle = 'Specialist'; City = 'Seattle'; Country = 'US'
                AdditionalProperties = @{}
            }
        } -ParameterFilter { $null -eq $ExpandProperty }
        Mock Get-MgUserManager -ModuleName M365ReportLibrary -MockWith {
            [pscustomobject]@{ Id = 'user-0'; AdditionalProperties = @{ userPrincipalName = 'boss@example.com' } }
        }

        Invoke-EntraUserCollector -OutputPath $script:folder -SkipConnect -WarningAction SilentlyContinue

        (Import-Csv -LiteralPath (Join-Path $script:folder 'users.csv')).ManagerUserPrincipalName |
            Should -Be 'boss@example.com'
        Should -Invoke Get-MgUserManager -ModuleName M365ReportLibrary -Times 1 -Exactly
    }

    It 'writes the header only, and says why, when Graph refuses the request' {
        Mock Get-MgUser -ModuleName M365ReportLibrary -MockWith { throw 'Insufficient privileges to complete the operation.' }

        Invoke-EntraUserCollector -OutputPath $script:folder -SkipConnect -WarningAction SilentlyContinue

        $csv = Join-Path $script:folder 'users.csv'
        (Get-Content -LiteralPath $csv).Count | Should -Be 1
        (Get-CsvHeaderColumn -Path $csv) -join ',' | Should -Be ((Get-EntraUserCsvColumn) -join ',')
        Get-Content -LiteralPath (Join-Path $script:folder 'run.log') -Raw |
            Should -Match 'unavailable in this tenant or cloud'
    }

    It 'signs in unless it is told to reuse the session' {
        Mock Get-MgUser -ModuleName M365ReportLibrary -MockWith { }

        Invoke-EntraUserCollector -OutputPath $script:folder -Environment GCCHigh

        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter {
            $Environment -eq 'USGov'
        }
    }
}
