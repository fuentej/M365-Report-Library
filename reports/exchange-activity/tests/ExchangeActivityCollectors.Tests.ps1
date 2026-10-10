#Requires -Version 7.0

BeforeAll {
    $global:ExTest = @{}
    $script:Root = Join-Path $PSScriptRoot '../../..' | Resolve-Path | Select-Object -ExpandProperty Path
    $script:Collectors = Join-Path $script:Root 'reports/exchange-activity/collectors'
    $script:Samples = Join-Path $script:Root 'reports/exchange-activity/samples'

    . (Join-Path $script:Root 'shared/tests/TenantCmdletStubs.ps1')
    . (Join-Path $PSScriptRoot 'ExchangeStubs.ps1')
    Import-Module (Join-Path $script:Root 'shared/M365ReportLibrary.psm1') -Force
    . (Join-Path $script:Collectors 'ExchangeActivityHelpers.ps1')

    $script:Schema = Import-PowerShellDataFile -LiteralPath (Join-Path $script:Collectors 'ExchangeActivitySchema.psd1')

    function New-TestFolder {
        $path = Join-Path ([System.IO.Path]::GetTempPath()) ('exchange-' + [guid]::NewGuid().ToString('N'))
        New-Item -Path $path -ItemType Directory -Force | Out-Null
        return $path
    }

    function Get-HeaderText { param([string]$Path) return ((Get-CsvHeaderColumn -Path $Path) -join ',') }

    function Get-LogText { param([string]$Folder) return (Get-Content -LiteralPath (Join-Path $Folder 'run.log') -Raw) }

    function Invoke-CollectorScript {
        param([Parameter(Mandatory)][string]$Name, [hashtable]$Arguments = @{})
        & (Join-Path $script:Collectors $Name) @Arguments
    }

    # Header list of getMailboxUsageDetail, in the order the API page gives it.
    # https://learn.microsoft.com/graph/api/reportroot-getmailboxusagedetail
    $global:ExTest.UsageDetailHeader = 'Report Refresh Date,User Principal Name,Display Name,Is Deleted,Deleted Date,Created Date,Last Activity Date,Item Count,Storage Used (Byte),Issue Warning Quota (Byte),Prohibit Send Quota (Byte),Prohibit Send/Receive Quota (Byte),Deleted Item Count,Deleted Item Size (Byte),Deleted Item Quota (Byte),Has Archive,Report Period'

    function New-UsageDetailCsv {
        param([object[]]$Rows)
        $lines = @($global:ExTest.UsageDetailHeader)
        foreach ($r in $Rows) {
            $lines += ('2026-10-06,{0},{1},False,,2021-03-15,2026-10-05,100,{2},100,200,300,5,1024,2000,True,30' -f $r.Upn, $r.Name, $r.Used)
        }
        return ($lines -join "`n")
    }

    # Get-EXOMailbox: the property sets are at
    # https://learn.microsoft.com/powershell/exchange/cmdlet-property-sets. Minimum holds
    # ExternalDirectoryObjectId, PrimarySmtpAddress, RecipientType, RecipientTypeDetails and
    # UserPrincipalName; Quota holds the quota properties. The cmdlet page does not document
    # the output object, so the types follow the parameter descriptions.
    function New-MockMailbox {
        param([string]$Id = 'mbx-1', [string]$Upn = 'avery.abara@example.com', [string]$Type = 'UserMailbox')
        [pscustomobject]@{
            ExternalDirectoryObjectId = $Id
            UserPrincipalName         = $Upn
            PrimarySmtpAddress        = $Upn
            RecipientType             = 'UserMailbox'
            RecipientTypeDetails      = $Type
            IssueWarningQuota         = '45 GB (48,318,382,080 bytes)'
            ProhibitSendQuota         = '48 GB (51,539,607,552 bytes)'
            ProhibitSendReceiveQuota  = '50 GB (53,687,091,200 bytes)'
            RecoverableItemsQuota     = '30 GB (32,212,254,720 bytes)'
            ArchiveQuota              = '100 GB (107,374,182,400 bytes)'
            UseDatabaseQuotaDefaults  = $false
        }
    }

    # Get-MessageTraceV2: Received, SenderAddress, RecipientAddress, Status and MessageTraceId
    # are the names the cmdlet page uses for its parameters and for the Received Time and
    # Recipient address it says to carry into the next round. The page documents no output object.
    # https://learn.microsoft.com/powershell/module/exchangepowershell/get-messagetracev2
    function New-MockTrace {
        param([string]$Received, [string]$Recipient = 'blake.bishop@example.com', [string]$Sender = 'avery.abara@example.com', [string]$Id = ([guid]::NewGuid().ToString()))
        [pscustomobject]@{
            Received         = [datetime]::Parse($Received, [cultureinfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal)
            MessageTraceId   = $Id
            SenderAddress    = $Sender
            RecipientAddress = $Recipient
            Status           = 'Delivered'
        }
    }
}

AfterAll {
    Remove-Variable -Name ExTest -Scope Global -ErrorAction SilentlyContinue
    Remove-Module M365ReportLibrary -Force -ErrorAction SilentlyContinue
}

Describe 'Cloud availability' {
    BeforeEach {
        $script:Out = New-TestFolder
        Mock Connect-MgGraph -ModuleName M365ReportLibrary -MockWith { }
        Mock Connect-ExchangeOnline -ModuleName M365ReportLibrary -MockWith { }
        Mock Disconnect-ExchangeOnline { }
        Mock Get-MgReportMailboxUsageDetail { }
        Mock Get-MgReportMailboxUsageStorage { }
        Mock Get-MgReportEmailActivityUserDetail { }
        Mock Get-MgReportEmailAppUsageUserDetail { }
        Mock Get-EXOMailbox { New-MockMailbox }
        Mock Get-MessageTraceV2 { }
        Mock Get-MgAdminReportSetting { [pscustomobject]@{ DisplayConcealedNames = $false } }
    }
    AfterEach { Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue }

    It 'writes a header only for <Script> in GCC High and never calls the report' -ForEach @(
        @{ Script = 'Get-MailboxUsageDetail.ps1'; Csv = 'mailbox-usage-detail.csv'; Key = 'MailboxUsageDetail'; Cmdlet = 'Get-MgReportMailboxUsageDetail' }
        @{ Script = 'Get-MailboxUsageStorage.ps1'; Csv = 'mailbox-usage-storage.csv'; Key = 'MailboxUsageStorage'; Cmdlet = 'Get-MgReportMailboxUsageStorage' }
        @{ Script = 'Get-EmailActivityUserDetail.ps1'; Csv = 'email-activity-user-detail.csv'; Key = 'EmailActivityUserDetail'; Cmdlet = 'Get-MgReportEmailActivityUserDetail' }
        @{ Script = 'Get-EmailAppUsageUserDetail.ps1'; Csv = 'email-app-usage-user-detail.csv'; Key = 'EmailAppUsageUserDetail'; Cmdlet = 'Get-MgReportEmailAppUsageUserDetail' }
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
        @{ Script = 'Get-Mailboxes.ps1'; Cloud = 'GCC' }
        @{ Script = 'Get-Mailboxes.ps1'; Cloud = 'GCCHigh' }
        @{ Script = 'Get-MessageTrace.ps1'; Cloud = 'GCCHigh' }
    ) {
        Invoke-CollectorScript $Script @{ OutputPath = $script:Out; Environment = $Cloud; SkipConnect = $true }
        Get-LogText $script:Out | Should -Match "UNVERIFIED"
        Should -Invoke Get-EXOMailbox -Times 1 -Exactly
    }

    It 'attempts the report settings in GCC High and the Graph message trace in every cloud' {
        Invoke-CollectorScript 'Get-ReportSettings.ps1' @{ OutputPath = $script:Out; Environment = 'GCCHigh'; SkipConnect = $true }
        Should -Invoke Get-MgAdminReportSetting -Times 1 -Exactly
        Get-LogText $script:Out | Should -Match 'UNVERIFIED'
    }
}

Describe 'Connection endpoint per -Environment' {
    BeforeEach {
        $script:Out = New-TestFolder
        Mock Connect-MgGraph -ModuleName M365ReportLibrary -MockWith { }
        Mock Connect-ExchangeOnline -ModuleName M365ReportLibrary -MockWith { }
        Mock Disconnect-ExchangeOnline { }
        Mock Get-MgReportMailboxUsageDetail { Set-Content -LiteralPath $OutFile -Value (New-UsageDetailCsv @()) }
        Mock Get-EXOMailbox { New-MockMailbox }
        Mock Invoke-MgGraphRequest { @{ value = @() } }
    }
    AfterEach { Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue }

    It 'signs in to Graph with the <Graph> environment for <Cloud>' -ForEach @(
        @{ Cloud = 'Commercial'; Graph = 'Global' }
        @{ Cloud = 'GCC'; Graph = 'Global' }
    ) {
        Invoke-CollectorScript 'Get-MailboxUsageDetail.ps1' @{ OutputPath = $script:Out; Environment = $Cloud }
        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter { $Environment -eq $Graph }
    }

    It 'signs in to Graph with the USGov environment for GCC High' {
        Invoke-CollectorScript 'Get-GraphMessageTrace.ps1' @{ OutputPath = $script:Out; Environment = 'GCCHigh' }
        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter { $Environment -eq 'USGov' }
    }

    It 'signs in to Exchange Online with <Name> for <Cloud>' -ForEach @(
        @{ Cloud = 'Commercial'; Name = 'O365Default' }
        @{ Cloud = 'GCC'; Name = 'O365Default' }
        @{ Cloud = 'GCCHigh'; Name = 'O365USGovGCCHigh' }
    ) {
        Invoke-CollectorScript 'Get-Mailboxes.ps1' @{ OutputPath = $script:Out; Environment = $Cloud }
        Should -Invoke Connect-ExchangeOnline -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter { $ExchangeEnvironmentName -eq $Name }
        Should -Invoke Disconnect-ExchangeOnline -Times 1 -Exactly
    }

    It 'does not sign in when asked to reuse a session' {
        Invoke-CollectorScript 'Get-Mailboxes.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        Should -Invoke Connect-ExchangeOnline -ModuleName M365ReportLibrary -Times 0 -Exactly
        Should -Invoke Disconnect-ExchangeOnline -Times 0 -Exactly
    }
}

Describe 'Graph usage reports (sources 1 to 4)' {
    BeforeEach {
        $script:Out = New-TestFolder
        $script:RunDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')
        $global:ExTest.Files = [System.Collections.Generic.List[string]]::new()
    }
    AfterEach { Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue }

    It 'stamps mailbox usage with the run date and derives the quota status at or above each quota' {
        # Quotas in the mock file are 100, 200 and 300 bytes.
        Mock Get-MgReportMailboxUsageDetail {
            $global:ExTest.Files.Add($OutFile)
            Set-Content -LiteralPath $OutFile -Encoding utf8 -Value (New-UsageDetailCsv @(
                    @{ Upn = 'a@example.com'; Name = 'A'; Used = 99 }
                    @{ Upn = 'b@example.com'; Name = 'B'; Used = 100 }
                    @{ Upn = 'c@example.com'; Name = 'C'; Used = 200 }
                    @{ Upn = 'd@example.com'; Name = 'D'; Used = 300 }))
        }
        Invoke-CollectorScript 'Get-MailboxUsageDetail.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; Period = 'D30' }

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:Out 'mailbox-usage-detail.csv'))
        $rows.Count | Should -Be 4
        ($rows.RunDate | Sort-Object -Unique) | Should -Be $script:RunDate
        ($rows | Where-Object UserPrincipalName -EQ 'a@example.com').QuotaStatus | Should -Be 'Good'
        ($rows | Where-Object UserPrincipalName -EQ 'b@example.com').QuotaStatus | Should -Be 'Warning'
        ($rows | Where-Object UserPrincipalName -EQ 'c@example.com').QuotaStatus | Should -Be 'CantSend'
        ($rows | Where-Object UserPrincipalName -EQ 'd@example.com').QuotaStatus | Should -Be 'CantSendReceive'
        $rows[0].StorageUsedByte | Should -Be '99'
        $rows[0].HasArchive | Should -Be 'True'
        Should -Invoke Get-MgReportMailboxUsageDetail -Times 1 -Exactly -ParameterFilter { $Period -eq 'D30' }
        foreach ($file in $global:ExTest.Files) { Test-Path -LiteralPath $file | Should -BeFalse }
    }

    It 'leaves the two headers the example schema omits empty and the quota status empty' {
        $example = "Report Refresh Date,User Principal Name,Display Name,Is Deleted,Deleted Date,Created Date,Last Activity Date,Item Count,Storage Used (Byte),Issue Warning Quota (Byte),Prohibit Send Quota (Byte),Prohibit Send/Receive Quota (Byte),Deleted Item Count,Deleted Item Size (Byte),Report Period`n2026-10-06,a@example.com,A,False,,2021-03-15,2026-10-05,100,99,100,200,300,5,1024,30"
        Mock Get-MgReportMailboxUsageDetail { Set-Content -LiteralPath $OutFile -Value $example -Encoding utf8 }
        Invoke-CollectorScript 'Get-MailboxUsageDetail.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }

        $row = @(Import-Csv -LiteralPath (Join-Path $script:Out 'mailbox-usage-detail.csv'))[0]
        $row.DeletedItemQuotaByte | Should -Be ''
        $row.HasArchive | Should -Be ''
        $row.QuotaStatus | Should -Be 'Good'
    }

    It 'does not append a second copy of the same snapshot' {
        Mock Get-MgReportMailboxUsageDetail { Set-Content -LiteralPath $OutFile -Value (New-UsageDetailCsv @(@{ Upn = 'a@example.com'; Name = 'A'; Used = 1 })) -Encoding utf8 }
        Invoke-CollectorScript 'Get-MailboxUsageDetail.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        Invoke-CollectorScript 'Get-MailboxUsageDetail.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        @(Import-Csv -LiteralPath (Join-Path $script:Out 'mailbox-usage-detail.csv')).Count | Should -Be 1
    }

    It 'reads storage over time' {
        Mock Get-MgReportMailboxUsageStorage {
            Set-Content -LiteralPath $OutFile -Encoding utf8 -Value "Report Refresh Date,Storage Used (Byte),Report Date,Report Period`n2026-10-06,5000,2026-10-05,7`n2026-10-06,4900,2026-10-04,7"
        }
        Invoke-CollectorScript 'Get-MailboxUsageStorage.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; Period = 'D7' }

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:Out 'mailbox-usage-storage.csv'))
        $rows.Count | Should -Be 2
        $rows[0].StorageUsedByte | Should -Be '5000'
        $rows[0].ReportDate | Should -Be '2026-10-05'
        $rows[0].RunDate | Should -Be $script:RunDate
        Should -Invoke Get-MgReportMailboxUsageStorage -Times 1 -Exactly -ParameterFilter { $Period -eq 'D7' }
    }

    It 'requests 180 days of mailbox storage when no period is given' {
        # Source 2 is documented as period D180. D30 cannot backfill the earlier days.
        # https://learn.microsoft.com/graph/api/reportroot-getmailboxusagestorage
        Mock Get-MgReportMailboxUsageStorage {
            Set-Content -LiteralPath $OutFile -Encoding utf8 -Value "Report Refresh Date,Storage Used (Byte),Report Date,Report Period`n"
        }
        Invoke-CollectorScript 'Get-MailboxUsageStorage.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        Should -Invoke Get-MgReportMailboxUsageStorage -Times 1 -Exactly -ParameterFilter { $Period -eq 'D180' }
    }

    It 'describes the email app date window as 30 days' {
        # Source 4's date form reaches back 30 days, not the 28 days of source 3.
        # https://learn.microsoft.com/graph/api/reportroot-getemailappusageuserdetail
        $text = Get-Content -LiteralPath (Join-Path $script:Collectors 'Get-EmailAppUsageUserDetail.ps1') -Raw
        $text | Should -Match 'within the last 30 days'
        $text | Should -Not -Match 'within the last 28 days'
    }

    It 'sends the period for email activity, and the date alone when one is given' {
        Mock Get-MgReportEmailActivityUserDetail {
            Set-Content -LiteralPath $OutFile -Encoding utf8 -Value "Report Refresh Date,User Principal Name,Display Name,Is Deleted,Deleted Date,Last Activity Date,Send Count,Receive Count,Read Count,Meeting Created Count,Meeting Interacted Count,Assigned Products,Report Period`n2026-10-06,a@example.com,A,False,,2026-10-05,3,40,30,1,2,MICROSOFT 365 E3,30"
        }
        Invoke-CollectorScript 'Get-EmailActivityUserDetail.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; Period = 'D30' }
        Should -Invoke Get-MgReportEmailActivityUserDetail -Times 1 -Exactly -ParameterFilter { $Period -eq 'D30' -and -not $PSBoundParameters.ContainsKey('Date') }

        Invoke-CollectorScript 'Get-EmailActivityUserDetail.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; Date = [datetime]'2026-10-01' }
        Should -Invoke Get-MgReportEmailActivityUserDetail -Times 1 -Exactly -ParameterFilter { $Date -eq [datetime]'2026-10-01' -and -not $PSBoundParameters.ContainsKey('Period') }

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:Out 'email-activity-user-detail.csv'))
        $rows.Count | Should -Be 2
        ($rows | Where-Object QueryDate -EQ '').SendCount | Should -Be '3'
        ($rows | Where-Object QueryDate -EQ '2026-10-01').ReadCount | Should -Be '30'
    }

    It 'reads email app usage' {
        Mock Get-MgReportEmailAppUsageUserDetail {
            Set-Content -LiteralPath $OutFile -Encoding utf8 -Value "Report Refresh Date,User Principal Name,Display Name,Is Deleted,Deleted Date,Last Activity Date,Mail For Mac,Outlook For Mac,Outlook For Windows,Outlook For Mobile,Other For Mobile,Outlook For Web,POP3 App,IMAP4 App,SMTP App,Report Period`n2026-10-06,a@example.com,A,False,,2026-10-05,No,No,Yes,Yes,No,Yes,No,No,No,30"
        }
        Invoke-CollectorScript 'Get-EmailAppUsageUserDetail.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        $row = @(Import-Csv -LiteralPath (Join-Path $script:Out 'email-app-usage-user-detail.csv'))[0]
        $row.OutlookForWindows | Should -Be 'Yes'
        $row.SMTPApp | Should -Be 'No'
        $row.OutlookForWeb | Should -Be 'Yes'
    }

    It 'writes the header only and logs the refusal when the report is refused' {
        Mock Get-MgReportMailboxUsageDetail { throw 'Forbidden: Reports.Read.All is missing' }
        Invoke-CollectorScript 'Get-MailboxUsageDetail.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        Get-HeaderText (Join-Path $script:Out 'mailbox-usage-detail.csv') | Should -Be ($script:Schema.MailboxUsageDetail -join ',')
        @(Import-Csv -LiteralPath (Join-Path $script:Out 'mailbox-usage-detail.csv')).Count | Should -Be 0
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
        # The adminReportSettings object carries displayConcealedNames at its root.
        # https://learn.microsoft.com/graph/api/resources/adminreportsettings
        Mock Get-MgAdminReportSetting { [pscustomobject]@{ DisplayConcealedNames = $Value } }
        Invoke-CollectorScript 'Get-ReportSettings.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        $row = @(Import-Csv -LiteralPath (Join-Path $script:Out 'report-settings.csv'))[0]
        $row.DisplayConcealedNames | Should -Be $Expected
        if ($Value) { Get-LogText $script:Out | Should -Match 'conceal names' }
    }

    It 'writes the header only when the setting cannot be read' {
        Mock Get-MgAdminReportSetting { throw 'Forbidden' }
        Invoke-CollectorScript 'Get-ReportSettings.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        @(Import-Csv -LiteralPath (Join-Path $script:Out 'report-settings.csv')).Count | Should -Be 0
        Get-LogText $script:Out | Should -Match 'ReportSettings.Read.All'
    }
}

Describe 'Exchange mailbox sources (sources 6, 7 and 9)' {
    BeforeEach { $script:Out = New-TestFolder }
    AfterEach { Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue }

    It 'lists every mailbox with Unlimited and the Minimum and Quota property sets' {
        Mock Get-EXOMailbox { New-MockMailbox -Id 'm1'; New-MockMailbox -Id 'm2' -Upn 'finance.shared@example.com' -Type 'SharedMailbox' }
        Invoke-CollectorScript 'Get-Mailboxes.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }

        Should -Invoke Get-EXOMailbox -Times 1 -Exactly -ParameterFilter {
            $ResultSize -eq 'Unlimited' -and ($PropertySets -join ',') -eq 'Minimum,Quota'
        }
        $rows = @(Import-Csv -LiteralPath (Join-Path $script:Out 'mailboxes.csv'))
        $rows.Count | Should -Be 2
        ($rows | Where-Object ExternalDirectoryObjectId -EQ 'm2').RecipientTypeDetails | Should -Be 'SharedMailbox'
        $rows[0].ProhibitSendQuota | Should -Be '48 GB (51,539,607,552 bytes)'
        $rows[0].UseDatabaseQuotaDefaults | Should -Be 'False'
    }

    It 'honours -MailboxLimit' {
        Mock Get-EXOMailbox { 1..5 | ForEach-Object { New-MockMailbox -Id "m$_" -Upn "user$_@example.com" } }
        Invoke-CollectorScript 'Get-Mailboxes.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; MailboxLimit = 2 }
        @(Import-Csv -LiteralPath (Join-Path $script:Out 'mailboxes.csv')).Count | Should -Be 2
    }

    It 'writes the header only when mailboxes cannot be listed' {
        Mock Get-EXOMailbox { throw 'The role assigned to this user is not enough' }
        Invoke-CollectorScript 'Get-Mailboxes.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        @(Import-Csv -LiteralPath (Join-Path $script:Out 'mailboxes.csv')).Count | Should -Be 0
        Get-LogText $script:Out | Should -Match 'unavailable to this sign-in'
    }

    It 'reads statistics one mailbox per call by UPN, skips a mailbox with no UPN, and parses the byte count' {
        Mock Get-EXOMailbox { New-MockMailbox -Id 'm1'; New-MockMailbox -Id 'm2' -Upn $null; New-MockMailbox -Id 'm3' -Upn 'blake.bishop@example.com' }
        Mock Get-EXOMailboxStatistics {
            [pscustomobject]@{
                ItemCount = 1200; TotalItemSize = '1.5 GB (1,610,612,736 bytes)'; DeletedItemCount = 4; TotalDeletedItemSize = '2 MB (2,097,152 bytes)'
                StorageLimitStatus = 'BelowLimit'; LastLogonTime = [datetime]'2026-10-05T09:00:00Z'; LastLogoffTime = [datetime]'2026-10-05T17:00:00Z'
                LastLoggedOnUserAccount = 'EXAMPLE\user'; IsArchiveMailbox = $false
                DatabaseIssueWarningQuota = '99 GB (106,300,440,576 bytes)'; DatabaseProhibitSendQuota = '99 GB (106,300,440,576 bytes)'; DatabaseProhibitSendReceiveQuota = '100 GB (107,374,182,400 bytes)'
            }
        }
        Invoke-CollectorScript 'Get-MailboxStatistics.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }

        Should -Invoke Get-EXOMailboxStatistics -Times 2 -Exactly
        Should -Invoke Get-EXOMailboxStatistics -Times 1 -Exactly -ParameterFilter { $Identity -eq 'blake.bishop@example.com' -and ($PropertySets -join ',') -eq 'All' }
        $rows = @(Import-Csv -LiteralPath (Join-Path $script:Out 'mailbox-statistics.csv'))
        $rows.Count | Should -Be 2
        $rows[0].TotalItemSizeBytes | Should -Be '1610612736'
        $rows[0].TotalDeletedItemSizeBytes | Should -Be '2097152'
        $rows[0].LastLogonTime | Should -Be '2026-10-05T09:00:00Z'
        $rows[0].StorageLimitStatus | Should -Be 'BelowLimit'
        $rows[0].PSObject.Properties.Name | Should -Not -Contain 'LastUserActionTime'
    }

    It 'keeps going when one mailbox statistics call fails' {
        Mock Get-EXOMailbox { New-MockMailbox -Id 'm1' -Upn 'a@example.com'; New-MockMailbox -Id 'm2' -Upn 'b@example.com' }
        Mock Get-EXOMailboxStatistics { if ($Identity -eq 'a@example.com') { throw 'boom' } else { [pscustomobject]@{ ItemCount = 1; TotalItemSize = '1 KB (1,024 bytes)'; IsArchiveMailbox = $false } } }
        Invoke-CollectorScript 'Get-MailboxStatistics.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        @(Import-Csv -LiteralPath (Join-Path $script:Out 'mailbox-statistics.csv')).Count | Should -Be 1
        Get-LogText $script:Out | Should -Match 'a@example.com could not be read'
    }

    It 'stamps the run date on state snapshots' {
        Mock Get-EXOMailbox { New-MockMailbox }
        Invoke-CollectorScript 'Get-Mailboxes.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        @(Import-Csv -LiteralPath (Join-Path $script:Out 'mailboxes.csv'))[0].RunDate | Should -Be ([datetime]::UtcNow.ToString('yyyy-MM-dd'))
    }

    It 'asks for mobile devices by mailbox with no device-family switch' {
        Mock Get-EXOMailbox { New-MockMailbox -Upn 'avery.abara@example.com' }
        Mock Get-EXOMobileDeviceStatistics {
            # DeviceType, DeviceOS, DeviceAccessState and DeviceUserAgent are the properties Learn selects:
            # https://learn.microsoft.com/troubleshoot/exchange/administration/windows-mail-app-not-blocked
            [pscustomobject]@{ DeviceId = 'DEV1'; DeviceType = 'UniversalOutlook'; DeviceOS = 'WINDOWS'; DeviceAccessState = 'Unknown'; DeviceUserAgent = 'microsoft.windowscommunicationsapps' }
            [pscustomobject]@{ DeviceId = 'DEV2'; DeviceType = 'Outlook'; DeviceOS = 'iOS 17.6'; DeviceAccessState = 'Allowed'; DeviceUserAgent = 'Outlook-iOS/2.0' }
        }
        Invoke-CollectorScript 'Get-MobileDevices.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }

        Should -Invoke Get-EXOMobileDeviceStatistics -Times 1 -Exactly -ParameterFilter {
            $Mailbox -eq 'avery.abara@example.com' -and -not $ActiveSync -and -not $RestApi -and -not $OWAforDevices -and -not $UniversalOutlook
        }
        $rows = @(Import-Csv -LiteralPath (Join-Path $script:Out 'mobile-devices.csv'))
        $rows.Count | Should -Be 2
        $rows[0].MailboxUserPrincipalName | Should -Be 'avery.abara@example.com'
        ($rows | Where-Object DeviceId -EQ 'DEV2').DeviceOS | Should -Be 'iOS 17.6'
    }
}

Describe 'Exchange message trace (source 8)' {
    BeforeEach {
        $script:Out = New-TestFolder
        Mock Get-EXOMailbox { New-MockMailbox -Upn 'avery.abara@example.com' }
        Mock Start-Sleep { }
    }
    AfterEach { Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue }

    It 'continues a full round from the last row''s Received time and RecipientAddress' {
        $global:ExTest.Calls = [System.Collections.Generic.List[object]]::new()
        Mock Get-MessageTraceV2 {
            $global:ExTest.Calls.Add(@{ EndDate = $EndDate; Starting = $StartingRecipientAddress; Sender = $SenderAddress })
            if ($global:ExTest.Calls.Count -eq 1) {
                New-MockTrace -Received '2026-10-07T12:00:00Z' -Recipient 'a@example.com' -Id '00000000-0000-0000-0000-000000000001'
                New-MockTrace -Received '2026-10-07T10:00:00Z' -Recipient 'b@example.com' -Id '00000000-0000-0000-0000-000000000002'
            }
            else {
                New-MockTrace -Received '2026-10-07T09:00:00Z' -Recipient 'c@example.com' -Id '00000000-0000-0000-0000-000000000003'
            }
        }

        $rows = @(Invoke-MessageTraceV2Window -Role Sender -Address 'avery.abara@example.com' -Start ([datetime]'2026-10-06T00:00:00Z') -End ([datetime]'2026-10-08T00:00:00Z') -ResultSize 2)

        $rows.Count | Should -Be 3
        $global:ExTest.Calls.Count | Should -Be 2
        $global:ExTest.Calls[0].Starting | Should -BeNullOrEmpty
        $global:ExTest.Calls[1].Starting | Should -Be 'b@example.com'
        $global:ExTest.Calls[1].EndDate.ToUniversalTime() | Should -Be ([datetime]'2026-10-07T10:00:00Z').ToUniversalTime()
    }

    It 'continues from an Unspecified Received time without shifting it out of UTC' {
        # Output timestamps are UTC. Kind Unspecified is not the local zone.
        # https://learn.microsoft.com/powershell/module/exchangepowershell/get-messagetracev2
        $received = [datetime]::SpecifyKind([datetime]'2026-10-07T10:00:00', [DateTimeKind]::Unspecified)
        $global:ExTest.Calls = [System.Collections.Generic.List[object]]::new()
        Mock Get-MessageTraceV2 {
            $global:ExTest.Calls.Add(@{ EndDate = $EndDate })
            if ($global:ExTest.Calls.Count -eq 1) {
                $last = New-MockTrace -Received '2026-10-07T10:00:00Z' -Recipient 'b@example.com' -Id '00000000-0000-0000-0000-000000000002'
                $last.Received = $received
                New-MockTrace -Received '2026-10-07T12:00:00Z' -Recipient 'a@example.com' -Id '00000000-0000-0000-0000-000000000001'
                $last
            }
            else {
                New-MockTrace -Received '2026-10-07T09:00:00Z' -Recipient 'c@example.com' -Id '00000000-0000-0000-0000-000000000003'
            }
        }

        $rows = @(Invoke-MessageTraceV2Window -Role Sender -Address 'avery.abara@example.com' -Start ([datetime]'2026-10-06T00:00:00Z') -End ([datetime]'2026-10-08T00:00:00Z') -ResultSize 2)
        $rows.Count | Should -Be 3
        $global:ExTest.Calls[1].EndDate.ToString('yyyy-MM-ddTHH:mm:ss') | Should -Be '2026-10-07T10:00:00'
        $global:ExTest.Calls[1].EndDate.Kind | Should -Be ([DateTimeKind]::Utc)
    }

    It 're-queries a message that already has 1000 recipients by MessageTraceId' {
        # Over 1,000 recipients the query is incomplete unless MessageTraceId is set.
        # https://learn.microsoft.com/powershell/module/exchangepowershell/get-messagetracev2
        # https://learn.microsoft.com/exchange/monitoring/trace-an-email-message/message-trace-modern-eac
        $id = [guid]'00000000-0000-0000-0000-000000000010'
        $global:ExTest.Follow = 0
        Mock Get-MessageTraceV2 {
            if ($MessageTraceId -eq $id) {
                $global:ExTest.Follow++
                New-MockTrace -Received '2026-10-07T08:00:00Z' -Recipient 'extra@example.com' -Id $id
            }
            else {
                foreach ($n in 1..1000) {
                    New-MockTrace -Received '2026-10-07T12:00:00Z' -Recipient ('r{0:d4}@example.com' -f $n) -Id $id
                }
            }
        }

        $rows = @(Invoke-MessageTraceV2Window -Role Sender -Address 'avery.abara@example.com' -Start ([datetime]'2026-10-06T00:00:00Z') -End ([datetime]'2026-10-08T00:00:00Z') -ResultSize 5000)
        $rows.Count | Should -Be 1001
        $global:ExTest.Follow | Should -Be 1
        $rows.RecipientAddress | Should -Contain 'extra@example.com'
    }

    It 'stops when a full round adds nothing new instead of looping' {
        Mock Get-MessageTraceV2 {
            New-MockTrace -Received '2026-10-07T12:00:00Z' -Recipient 'a@example.com' -Id '00000000-0000-0000-0000-000000000001'
            New-MockTrace -Received '2026-10-07T10:00:00Z' -Recipient 'b@example.com' -Id '00000000-0000-0000-0000-000000000002'
        }
        $rows = @(Invoke-MessageTraceV2Window -Role Recipient -Address 'avery.abara@example.com' -Start ([datetime]'2026-10-06') -End ([datetime]'2026-10-08') -ResultSize 2)
        $rows.Count | Should -Be 2
        Should -Invoke Get-MessageTraceV2 -Times 2 -Exactly
    }

    It 'queries sent and received separately, each window at most 10 days, and no earlier than 90 days' {
        $global:ExTest.Spans = [System.Collections.Generic.List[double]]::new()
        $global:ExTest.Starts = [System.Collections.Generic.List[datetime]]::new()
        $global:ExTest.Roles = [System.Collections.Generic.List[string]]::new()
        Mock Get-MessageTraceV2 {
            $global:ExTest.Spans.Add(($EndDate - $StartDate).TotalDays)
            $global:ExTest.Starts.Add($StartDate)
            $global:ExTest.Roles.Add($(if ($SenderAddress) { 'Sender' } else { 'Recipient' }))
        }
        Invoke-CollectorScript 'Get-MessageTrace.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; LookbackDays = 90 }

        $global:ExTest.Spans.Count | Should -BeGreaterOrEqual 18
        ($global:ExTest.Spans | Measure-Object -Maximum).Maximum | Should -BeLessOrEqual 10
        ($global:ExTest.Starts | Measure-Object -Minimum).Minimum | Should -BeGreaterOrEqual ([datetime]::UtcNow.AddDays(-90).AddMinutes(-1))
        ($global:ExTest.Roles | Sort-Object -Unique) | Should -Be @('Recipient', 'Sender')
        Should -Invoke Get-MessageTraceV2 -ParameterFilter { $SenderAddress -eq 'avery.abara@example.com' -or $RecipientAddress -eq 'avery.abara@example.com' }
    }

    It 'resumes from the latest Received already in the file' {
        $path = Join-Path $script:Out 'message-trace.csv'
        Export-AppendCsv -Path $path -Column $script:Schema.MessageTrace -Rows @(
            [pscustomobject]@{ Received = ([datetime]::UtcNow.AddHours(-30)).ToString('yyyy-MM-ddTHH:mm:ssZ'); MessageTraceId = 'old-1'; SenderAddress = 'x@example.com'; RecipientAddress = 'y@example.com'; Status = 'Delivered' }
            [pscustomobject]@{ Received = '2026-01-01T00:00:00Z'; MessageTraceId = 'old-0'; SenderAddress = 'x@example.com'; RecipientAddress = 'y@example.com'; Status = 'Delivered' })
        $watermark = Get-CsvWatermark -Path $path -Column 'Received'

        $global:ExTest.FirstStart = $null
        Mock Get-MessageTraceV2 { if (-not $global:ExTest.FirstStart -or $StartDate -lt $global:ExTest.FirstStart) { $global:ExTest.FirstStart = $StartDate } }
        Invoke-CollectorScript 'Get-MessageTrace.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; LookbackDays = 90 }

        $global:ExTest.FirstStart.ToUniversalTime() | Should -Be $watermark.ToUniversalTime()
    }

    It 'writes one row per recipient, drops a repeat, and keeps no subject' {
        Mock Get-MessageTraceV2 {
            if ($SenderAddress) {
                New-MockTrace -Received ([datetime]::UtcNow.AddHours(-2).ToString('o')) -Recipient 'a@example.com' -Id '00000000-0000-0000-0000-000000000001'
                New-MockTrace -Received ([datetime]::UtcNow.AddHours(-2).ToString('o')) -Recipient 'b@example.com' -Id '00000000-0000-0000-0000-000000000001'
            }
        }
        Invoke-CollectorScript 'Get-MessageTrace.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        Invoke-CollectorScript 'Get-MessageTrace.ps1' @{ OutputPath = $script:Out; SkipConnect = $true }
        $rows = @(Import-Csv -LiteralPath (Join-Path $script:Out 'message-trace.csv'))
        $rows.Count | Should -Be 2
        $rows[0].PSObject.Properties.Name | Should -Not -Contain 'Subject'
        $rows[0].Received | Should -Match '^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\dZ$'
    }

    It 'waits before the 96th request in a 5-minute window' {
        Remove-Variable -Name TraceRequestTimes -Scope Script -ErrorAction SilentlyContinue
        1..95 | ForEach-Object { Wait-MessageTraceRateLimit }
        Should -Invoke Start-Sleep -Times 0 -Exactly
        Wait-MessageTraceRateLimit
        Should -Invoke Start-Sleep -Times 1 -Exactly
        Remove-Variable -Name TraceRequestTimes -Scope Script -ErrorAction SilentlyContinue
    }

    It 'counts requests that are still inside the window after a wait' {
        # Clearing the queue after the wait lets the next burst exceed 100 in 5 minutes.
        # https://learn.microsoft.com/powershell/module/exchangepowershell/get-messagetracev2
        Remove-Variable -Name TraceRequestTimes -Scope Script -ErrorAction SilentlyContinue
        1..95 | ForEach-Object { Wait-MessageTraceRateLimit }
        Wait-MessageTraceRateLimit
        Wait-MessageTraceRateLimit
        Should -Invoke Start-Sleep -Times 2 -Exactly
        Remove-Variable -Name TraceRequestTimes -Scope Script -ErrorAction SilentlyContinue
    }

    It 'resumes an unfinished mailbox trace from the original start, not the newest row' {
        # A newer row from a mailbox that finished must not hide an unfinished mailbox.
        $global:ExTest.Phase = 'fail'
        $global:ExTest.Starts = [System.Collections.Generic.List[datetime]]::new()
        Mock Get-EXOMailbox {
            New-MockMailbox -Id 'm1' -Upn 'first@example.com'
            New-MockMailbox -Id 'm2' -Upn 'second@example.com'
        }
        Mock Get-MessageTraceV2 {
            $global:ExTest.Starts.Add($StartDate)
            if ($global:ExTest.Phase -eq 'fail' -and $SenderAddress -eq 'second@example.com') { throw 'trace failed' }
            if ($SenderAddress -eq 'first@example.com') {
                New-MockTrace -Received ([datetime]::UtcNow.AddHours(-1).ToString('o')) -Sender 'first@example.com' -Recipient 'a@example.com' -Id '00000000-0000-0000-0000-00000000000a'
            }
        }

        Invoke-CollectorScript 'Get-MessageTrace.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; LookbackDays = 10 }
        Test-Path -LiteralPath (Join-Path $script:Out 'message-trace.pending') | Should -BeTrue

        $global:ExTest.Phase = 'ok'
        $global:ExTest.Starts.Clear()
        Invoke-CollectorScript 'Get-MessageTrace.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; LookbackDays = 10 }

        ($global:ExTest.Starts | Measure-Object -Minimum).Minimum | Should -BeLessOrEqual ([datetime]::UtcNow.AddDays(-9))
        Test-Path -LiteralPath (Join-Path $script:Out 'message-trace.pending') | Should -BeFalse
    }
}

Describe 'Graph message trace (source 11)' {
    BeforeEach {
        $script:Out = New-TestFolder
        Mock Connect-MgGraph -ModuleName M365ReportLibrary -MockWith { }
        Remove-Variable -Name TraceRequestTimes -Scope Script -ErrorAction SilentlyContinue
        # exchangeMessageTrace: id, senderAddress, recipientAddress, receivedDateTime, status, size.
        # https://learn.microsoft.com/exchange/monitoring/trace-an-email-message/graph-api-message-trace
        function global:New-GraphTrace {
            param([string]$Id, [string]$Received, [string]$Recipient = 'blake.bishop@example.com')
            @{ id = $Id; senderAddress = 'avery.abara@example.com'; recipientAddress = $Recipient; messageId = '<x@example.com>'; receivedDateTime = $Received; subject = 'must not be stored'; size = 45678; fromIP = '203.0.113.5'; toIP = ''; status = 'delivered' }
        }
    }
    AfterEach { Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue }

    It 'follows @odata.nextLink as returned until it is absent, with GET only' {
        $global:ExTest.Uris = [System.Collections.Generic.List[string]]::new()
        Mock Invoke-MgGraphRequest {
            $global:ExTest.Uris.Add($Uri)
            $Method | Should -Be 'GET'
            if ($Uri -like '/beta/admin/exchange/tracing/messageTraces*' -and $Uri -notlike '*skiptoken*') {
                @{ value = @(New-GraphTrace 'id-1' '2026-10-07T10:00:00Z'); '@odata.nextLink' = 'https://graph.microsoft.com/beta/admin/exchange/tracing/messageTraces?$skiptoken=abc' }
            }
            else {
                @{ value = @(New-GraphTrace 'id-2' '2026-10-07T11:00:00Z' 'c@example.com') }
            }
        }
        Invoke-CollectorScript 'Get-GraphMessageTrace.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; LookbackDays = 5 }

        $global:ExTest.Uris | Should -Contain 'https://graph.microsoft.com/beta/admin/exchange/tracing/messageTraces?$skiptoken=abc'
        @($global:ExTest.Uris | Where-Object { $_ -like '*skiptoken*' }).Count | Should -Be 1
        $rows = @(Import-Csv -LiteralPath (Join-Path $script:Out 'graph-message-trace.csv'))
        $rows.Count | Should -Be 2
        $rows[0].PSObject.Properties.Name | Should -Not -Contain 'Subject'
        ($rows | Where-Object Id -EQ 'id-2').RecipientAddress | Should -Be 'c@example.com'
        $rows[0].Size | Should -Be '45678'
    }

    It 'filters on receivedDateTime in ISO 8601 UTC, at most 10 days, never before 90 days ago, with $top=5000' {
        $global:ExTest.Uris = [System.Collections.Generic.List[string]]::new()
        Mock Invoke-MgGraphRequest { $global:ExTest.Uris.Add($Uri); @{ value = @() } }
        Invoke-CollectorScript 'Get-GraphMessageTrace.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; LookbackDays = 90 }

        $global:ExTest.Uris.Count | Should -BeGreaterOrEqual 9
        foreach ($uri in $global:ExTest.Uris) {
            $decoded = [uri]::UnescapeDataString($uri)
            $match = [regex]::Match($decoded, '^/beta/admin/exchange/tracing/messageTraces\?\$filter=receivedDateTime ge (\d{4}-\d\d-\d\dT\d\d:\d\d:\d\dZ) and receivedDateTime le (\d{4}-\d\d-\d\dT\d\d:\d\d:\d\dZ)&\$top=5000$')
            $match.Success | Should -BeTrue -Because $decoded
            $from = [datetime]::Parse($match.Groups[1].Value, [cultureinfo]::InvariantCulture, 'AssumeUniversal,AdjustToUniversal')
            $to = [datetime]::Parse($match.Groups[2].Value, [cultureinfo]::InvariantCulture, 'AssumeUniversal,AdjustToUniversal')
            ($to - $from).TotalDays | Should -BeLessOrEqual 10
            $from | Should -BeGreaterOrEqual ([datetime]::UtcNow.AddDays(-90).AddMinutes(-1))
        }
    }

    It 'resumes from the latest ReceivedDateTime already in the file' {
        $path = Join-Path $script:Out 'graph-message-trace.csv'
        $latest = [datetime]::UtcNow.AddHours(-20)
        Export-AppendCsv -Path $path -Column $script:Schema.GraphMessageTrace -Rows @(
            [pscustomobject]@{ ReceivedDateTime = $latest.ToString('yyyy-MM-ddTHH:mm:ssZ'); Id = 'old'; SenderAddress = 'a@example.com'; RecipientAddress = 'b@example.com'; Status = 'delivered'; Size = 1 })
        $global:ExTest.Uris = [System.Collections.Generic.List[string]]::new()
        Mock Invoke-MgGraphRequest { $global:ExTest.Uris.Add([uri]::UnescapeDataString($Uri)); @{ value = @() } }
        Invoke-CollectorScript 'Get-GraphMessageTrace.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; LookbackDays = 90 }

        $global:ExTest.Uris.Count | Should -Be 1
        $global:ExTest.Uris[0] | Should -Match ('receivedDateTime ge {0:yyyy-MM-ddTHH:mm:ss}Z' -f $latest)
    }

    It 'refuses a nextLink on a host that is not Microsoft Graph' {
        Mock Invoke-MgGraphRequest { @{ value = @(); '@odata.nextLink' = 'https://evil.example.com/steal?token=1' } }
        Invoke-CollectorScript 'Get-GraphMessageTrace.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; LookbackDays = 2 }
        Should -Invoke Invoke-MgGraphRequest -Times 1 -Exactly
        Get-LogText $script:Out | Should -Match 'outside Microsoft Graph'
    }

    It 'says a 401 is not an empty trace and leaves the file as it was' {
        Mock Invoke-MgGraphRequest { throw 'Service principal-less authentication failed: 401 Unauthorized' }
        Invoke-CollectorScript 'Get-GraphMessageTrace.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; LookbackDays = 2 }
        Get-LogText $script:Out | Should -Match 'a 401 is not an empty trace'
        Get-LogText $script:Out | Should -Match '8bd644d1-64a1-4d4b-ae52-2e0cbf64e373'
        @(Import-Csv -LiteralPath (Join-Path $script:Out 'graph-message-trace.csv')).Count | Should -Be 0
        Get-HeaderText (Join-Path $script:Out 'graph-message-trace.csv') | Should -Be ($script:Schema.GraphMessageTrace -join ',')
    }

    It 'does not skip an older window after a newer window was written and a later call failed' {
        # Newest-first writes a recent row, then a failed older window is past the watermark forever.
        # https://learn.microsoft.com/exchange/monitoring/trace-an-email-message/graph-api-message-trace
        Mock Invoke-MgGraphRequest {
            $decoded = [uri]::UnescapeDataString($Uri)
            $match = [regex]::Match($decoded, 'receivedDateTime ge (\S+) and receivedDateTime le (\S+)')
            $from = [datetime]::Parse($match.Groups[1].Value, [cultureinfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal)
            if ($from -lt [datetime]::UtcNow.AddDays(-9)) { throw 'older window failed' }
            @{ value = @(New-GraphTrace 'id-new' ([datetime]::UtcNow.AddHours(-1).ToString('yyyy-MM-ddTHH:mm:ssZ'))) }
        }
        Invoke-CollectorScript 'Get-GraphMessageTrace.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; LookbackDays = 15 }

        $global:ExTest.Uris = [System.Collections.Generic.List[string]]::new()
        Mock Invoke-MgGraphRequest {
            $global:ExTest.Uris.Add([uri]::UnescapeDataString($Uri))
            @{ value = @() }
        }
        Invoke-CollectorScript 'Get-GraphMessageTrace.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; LookbackDays = 15 }

        $starts = foreach ($uri in $global:ExTest.Uris) {
            $match = [regex]::Match($uri, 'receivedDateTime ge (\d{4}-\d\d-\d\dT\d\d:\d\d:\d\dZ)')
            [datetime]::Parse($match.Groups[1].Value, [cultureinfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal)
        }
        ($starts | Measure-Object -Minimum).Minimum | Should -BeLessOrEqual ([datetime]::UtcNow.AddDays(-14))
    }

    It 'retries a throttled Graph message trace instead of keeping an empty file' {
        # Retry after the window. The page does not name Retry-After. A 401 is a different error.
        # https://learn.microsoft.com/exchange/monitoring/trace-an-email-message/graph-api-message-trace
        Mock Start-Sleep { }
        $global:ExTest.Tries = 0
        Mock Invoke-MgGraphRequest {
            $global:ExTest.Tries++
            if ($global:ExTest.Tries -eq 1) { throw 'Your recent queries have surpassed the permitted limit, please try again later.' }
            @{ value = @(New-GraphTrace 'id-throttled' '2026-10-07T10:00:00Z') }
        }
        Invoke-CollectorScript 'Get-GraphMessageTrace.ps1' @{ OutputPath = $script:Out; SkipConnect = $true; LookbackDays = 2 }
        @(Import-Csv -LiteralPath (Join-Path $script:Out 'graph-message-trace.csv')).Count | Should -Be 1
        $global:ExTest.Tries | Should -BeGreaterOrEqual 2
    }
}

Describe 'Run-All' {
    BeforeEach {
        $script:Out = New-TestFolder
        Mock Connect-MgGraph -ModuleName M365ReportLibrary -MockWith { }
        Mock Connect-ExchangeOnline -ModuleName M365ReportLibrary -MockWith { }
        Mock Disconnect-ExchangeOnline { }
        Mock Get-MgAdminReportSetting { [pscustomobject]@{ DisplayConcealedNames = $false } }
        Mock Get-MgReportMailboxUsageDetail { throw 'should be skipped in GCC High' }
        Mock Get-MgReportMailboxUsageStorage { throw 'should be skipped in GCC High' }
        Mock Get-MgReportEmailActivityUserDetail { throw 'should be skipped in GCC High' }
        Mock Get-MgReportEmailAppUsageUserDetail { throw 'should be skipped in GCC High' }
        Mock Get-EXOMailbox { New-MockMailbox }
        Mock Get-EXOMailboxStatistics { [pscustomobject]@{ ItemCount = 1; TotalItemSize = '1 KB (1,024 bytes)'; IsArchiveMailbox = $false } }
        Mock Get-EXOMobileDeviceStatistics { }
        Mock Get-MessageTraceV2 { }
        Mock Invoke-MgGraphRequest { @{ value = @() } }
        Mock Start-Sleep { }
    }
    AfterEach { Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue }

    It 'leaves a header-only file for each usage report in GCC High and one file per other collector, after two sign-ins' {
        & (Join-Path $script:Collectors 'Run-All.ps1') -OutputPath $script:Out -Environment GCCHigh -MailboxLimit 1

        Should -Invoke Connect-ExchangeOnline -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter { $ExchangeEnvironmentName -eq 'O365USGovGCCHigh' }
        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter { $Environment -eq 'USGov' -and $Scopes -contains 'Reports.Read.All' -and $Scopes -contains 'ReportSettings.Read.All' }
        Should -Invoke Get-MgReportMailboxUsageDetail -Times 0 -Exactly
        foreach ($name in 'mailbox-usage-detail', 'mailbox-usage-storage', 'email-activity-user-detail', 'email-app-usage-user-detail') {
            @(Import-Csv -LiteralPath (Join-Path $script:Out "$name.csv")).Count | Should -Be 0
        }
        foreach ($name in 'report-settings', 'mailboxes', 'mailbox-statistics', 'message-trace', 'mobile-devices', 'graph-message-trace') {
            Test-Path -LiteralPath (Join-Path $script:Out "$name.csv") | Should -BeTrue -Because $name
        }
        Should -Invoke Disconnect-ExchangeOnline -Times 1 -Exactly
    }

    It 'asks for 180 days of mailbox storage unless -Period is passed, and 30 days of mailbox usage' {
        # https://learn.microsoft.com/graph/api/reportroot-getmailboxusagestorage
        Mock Get-MgReportMailboxUsageStorage { }
        Mock Get-MgReportMailboxUsageDetail { }
        & (Join-Path $script:Collectors 'Run-All.ps1') -OutputPath $script:Out -Environment Commercial -MailboxLimit 1
        Should -Invoke Get-MgReportMailboxUsageStorage -Times 1 -Exactly -ParameterFilter { $Period -eq 'D180' }
        Should -Invoke Get-MgReportMailboxUsageDetail -Times 1 -Exactly -ParameterFilter { $Period -eq 'D30' }

        & (Join-Path $script:Collectors 'Run-All.ps1') -OutputPath $script:Out -Environment Commercial -MailboxLimit 1 -Period D7
        Should -Invoke Get-MgReportMailboxUsageStorage -Times 1 -Exactly -ParameterFilter { $Period -eq 'D7' }
    }
}
