#Requires -Version 7.0

<#
    .SYNOPSIS
        Regenerates samples/: fake CSVs with exactly the columns the collectors write.

    .DESCRIPTION
        Every name is invented and every address is on example.com. Columns come from
        collectors/ExchangeActivitySchema.psd1, so a sample cannot drift from its collector.
        samples/gcchigh/ holds the header-only files a GCC High run leaves for the four Graph
        usage reports, which are not available in that cloud.

    .EXAMPLE
        ./New-SampleData.ps1
#>
[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path $PSScriptRoot 'samples')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot '../../shared/M365ReportLibrary.psm1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'collectors/ExchangeActivitySchema.psd1')

function Write-Sample {
    param([string]$Folder, [string]$Name, [string[]]$Column, [object[]]$Row)
    $path = Join-Path $Folder $Name
    if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
    Export-AppendCsv -Path $path -Rows $Row -Column $Column
}

if (-not (Test-Path -LiteralPath $OutputPath)) { New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null }
$gccHigh = Join-Path $OutputPath 'gcchigh'
if (-not (Test-Path -LiteralPath $gccHigh)) { New-Item -Path $gccHigh -ItemType Directory -Force | Out-Null }

$gb = 1GB
$people = @(
    @{ Id = '11111111-0000-0000-0000-000000000001'; Name = 'Avery Abara'; Used = 12 * $gb; Warn = 45 * $gb; Send = 48 * $gb; Both = 50 * $gb; Type = 'UserMailbox'; Activity = '2026-10-07' }
    @{ Id = '11111111-0000-0000-0000-000000000002'; Name = 'Blake Bishop'; Used = 46 * $gb; Warn = 45 * $gb; Send = 48 * $gb; Both = 50 * $gb; Type = 'UserMailbox'; Activity = '2026-10-06' }
    @{ Id = '11111111-0000-0000-0000-000000000003'; Name = 'Casey Cho'; Used = 49 * $gb; Warn = 45 * $gb; Send = 48 * $gb; Both = 50 * $gb; Type = 'UserMailbox'; Activity = '2026-09-02' }
    @{ Id = '11111111-0000-0000-0000-000000000004'; Name = 'Devon Diaz'; Used = 50 * $gb; Warn = 45 * $gb; Send = 48 * $gb; Both = 50 * $gb; Type = 'UserMailbox'; Activity = '' }
    @{ Id = '11111111-0000-0000-0000-000000000005'; Name = 'Finance Shared'; Used = 4 * $gb; Warn = 45 * $gb; Send = 48 * $gb; Both = 50 * $gb; Type = 'SharedMailbox'; Activity = '' }
    @{ Id = '11111111-0000-0000-0000-000000000006'; Name = 'Emery Ellis'; Used = 20 * $gb; Warn = 45 * $gb; Send = 48 * $gb; Both = 50 * $gb; Type = 'UserMailbox'; Activity = '2026-10-08' }
)
foreach ($p in $people) {
    $first, $last = $p.Name -split ' ', 2
    $p.Upn = if ($last) { ($first + '.' + $last).ToLowerInvariant() + '@example.com' } else { $first.ToLowerInvariant() + '@example.com' }
    $p.Upn = $p.Upn.Replace(' ', '.')
}
$people[4].Upn = 'finance.shared@example.com'

$runDates = '2026-10-01', '2026-10-08'

# Source 1
$rows = foreach ($run in $runDates) {
    foreach ($p in $people) {
        $growth = if ($run -eq '2026-10-08') { 1 * $gb } else { 0 }
        $used = [int64]($p.Used + $growth)
        [pscustomobject]@{
            RunDate = $run; ReportRefreshDate = ([datetime]$run).AddDays(-2).ToString('yyyy-MM-dd'); UserPrincipalName = $p.Upn; DisplayName = $p.Name
            IsDeleted = 'False'; DeletedDate = ''; CreatedDate = '2021-03-15'; LastActivityDate = $p.Activity; ItemCount = 40000 + $used / 1MB
            StorageUsedByte = $used; IssueWarningQuotaByte = [int64]$p.Warn; ProhibitSendQuotaByte = [int64]$p.Send; ProhibitSendReceiveQuotaByte = [int64]$p.Both
            DeletedItemCount = 120; DeletedItemSizeByte = 52428800; DeletedItemQuotaByte = 31457280000; HasArchive = if ($p.Type -eq 'UserMailbox') { 'True' } else { 'False' }
            ReportPeriod = 30
            QuotaStatus = ''
        }
    }
}
foreach ($r in $rows) {
    $status = if ([decimal]$r.StorageUsedByte -ge [decimal]$r.ProhibitSendReceiveQuotaByte) { 'CantSendReceive' }
    elseif ([decimal]$r.StorageUsedByte -ge [decimal]$r.ProhibitSendQuotaByte) { 'CantSend' }
    elseif ([decimal]$r.StorageUsedByte -ge [decimal]$r.IssueWarningQuotaByte) { 'Warning' }
    else { 'Good' }
    $r.QuotaStatus = $status
}
Write-Sample $OutputPath 'mailbox-usage-detail.csv' $schema.MailboxUsageDetail $rows
Write-Sample $gccHigh 'mailbox-usage-detail.csv' $schema.MailboxUsageDetail @()

# Source 2
$rows = foreach ($run in $runDates) {
    foreach ($offset in 0..3) {
        $day = ([datetime]$run).AddDays(-2 - $offset)
        [pscustomobject]@{ RunDate = $run; ReportRefreshDate = ([datetime]$run).AddDays(-2).ToString('yyyy-MM-dd'); StorageUsedByte = [int64](181 * $gb + $offset * 100MB * -1); ReportDate = $day.ToString('yyyy-MM-dd'); ReportPeriod = 7 }
    }
}
Write-Sample $OutputPath 'mailbox-usage-storage.csv' $schema.MailboxUsageStorage $rows
Write-Sample $gccHigh 'mailbox-usage-storage.csv' $schema.MailboxUsageStorage @()

# Source 3
$rows = foreach ($p in $people) {
    $active = [bool]$p.Activity
    [pscustomobject]@{
        RunDate = '2026-10-08'; QueryDate = ''; ReportRefreshDate = '2026-10-06'; UserPrincipalName = $p.Upn; DisplayName = $p.Name; IsDeleted = 'False'; DeletedDate = ''
        LastActivityDate = $p.Activity; SendCount = if ($active) { 35 } else { 0 }; ReceiveCount = if ($active) { 210 } else { 12 }; ReadCount = if ($active) { 180 } else { 0 }
        MeetingCreatedCount = if ($active) { 4 } else { 0 }; MeetingInteractedCount = if ($active) { 22 } else { 0 }; AssignedProducts = 'MICROSOFT 365 E3'; ReportPeriod = 30
    }
}
Write-Sample $OutputPath 'email-activity-user-detail.csv' $schema.EmailActivityUserDetail $rows
Write-Sample $gccHigh 'email-activity-user-detail.csv' $schema.EmailActivityUserDetail @()

# Source 4
$rows = foreach ($p in $people) {
    $active = [bool]$p.Activity
    [pscustomobject]@{
        RunDate = '2026-10-08'; QueryDate = ''; ReportRefreshDate = '2026-10-06'; UserPrincipalName = $p.Upn; DisplayName = $p.Name; IsDeleted = 'False'; DeletedDate = ''
        LastActivityDate = $p.Activity; MailForMac = 'No'; OutlookForMac = 'No'; OutlookForWindows = if ($active) { 'Yes' } else { 'No' }
        OutlookForMobile = if ($active) { 'Yes' } else { 'No' }; OtherForMobile = 'No'; OutlookForWeb = if ($active) { 'Yes' } else { 'No' }
        POP3App = 'No'; IMAP4App = 'No'; SMTPApp = 'No'; ReportPeriod = 30
    }
}
Write-Sample $OutputPath 'email-app-usage-user-detail.csv' $schema.EmailAppUsageUserDetail $rows
Write-Sample $gccHigh 'email-app-usage-user-detail.csv' $schema.EmailAppUsageUserDetail @()

# Source 5
Write-Sample $OutputPath 'report-settings.csv' $schema.ReportSettings @(
    [pscustomobject]@{ RunDate = '2026-10-01'; DisplayConcealedNames = 'False' }
    [pscustomobject]@{ RunDate = '2026-10-08'; DisplayConcealedNames = 'False' }
)

# Source 6
$rows = foreach ($p in $people) {
    [pscustomobject]@{
        RunDate = '2026-10-08'; ExternalDirectoryObjectId = $p.Id; UserPrincipalName = $p.Upn; PrimarySmtpAddress = $p.Upn; RecipientType = 'UserMailbox'; RecipientTypeDetails = $p.Type
        IssueWarningQuota = '45 GB (48,318,382,080 bytes)'; ProhibitSendQuota = '48 GB (51,539,607,552 bytes)'; ProhibitSendReceiveQuota = '50 GB (53,687,091,200 bytes)'
        RecoverableItemsQuota = '30 GB (32,212,254,720 bytes)'; ArchiveQuota = '100 GB (107,374,182,400 bytes)'; UseDatabaseQuotaDefaults = 'False'
    }
}
Write-Sample $OutputPath 'mailboxes.csv' $schema.Mailboxes $rows

# Source 7
$rows = foreach ($p in $people) {
    $size = [int64]($p.Used + 1 * $gb)
    [pscustomobject]@{
        RunDate = '2026-10-08'; ExternalDirectoryObjectId = $p.Id; UserPrincipalName = $p.Upn; ItemCount = 40000 + $size / 1MB
        TotalItemSize = ('{0} GB ({1:N0} bytes)' -f ($size / $gb), $size); TotalItemSizeBytes = $size; DeletedItemCount = 120
        TotalDeletedItemSize = '50 MB (52,428,800 bytes)'; TotalDeletedItemSizeBytes = 52428800
        StorageLimitStatus = 'BelowLimit'; LastLogonTime = if ($p.Activity) { "$($p.Activity)T09:12:44Z" } else { '' }; LastLogoffTime = if ($p.Activity) { "$($p.Activity)T17:40:02Z" } else { '' }
        LastLoggedOnUserAccount = if ($p.Activity) { 'EXAMPLE\' + ($p.Upn -split '@')[0] } else { '' }; IsArchiveMailbox = 'False'
        DatabaseIssueWarningQuota = '99 GB (106,300,440,576 bytes)'; DatabaseProhibitSendQuota = '99 GB (106,300,440,576 bytes)'; DatabaseProhibitSendReceiveQuota = '100 GB (107,374,182,400 bytes)'
    }
}
Write-Sample $OutputPath 'mailbox-statistics.csv' $schema.MailboxStatistics $rows

# Source 8
$rows = @(
    [pscustomobject]@{ Received = '2026-10-07T14:03:11Z'; MessageTraceId = 'aaaaaaaa-0000-0000-0000-000000000001'; SenderAddress = $people[0].Upn; RecipientAddress = $people[1].Upn; Status = 'Delivered' }
    [pscustomobject]@{ Received = '2026-10-07T14:03:11Z'; MessageTraceId = 'aaaaaaaa-0000-0000-0000-000000000001'; SenderAddress = $people[0].Upn; RecipientAddress = $people[5].Upn; Status = 'Delivered' }
    [pscustomobject]@{ Received = '2026-10-07T16:20:45Z'; MessageTraceId = 'aaaaaaaa-0000-0000-0000-000000000002'; SenderAddress = 'partner@fabrikam.example.com'; RecipientAddress = $people[0].Upn; Status = 'Delivered' }
    [pscustomobject]@{ Received = '2026-10-08T08:41:09Z'; MessageTraceId = 'aaaaaaaa-0000-0000-0000-000000000003'; SenderAddress = $people[5].Upn; RecipientAddress = 'partner@fabrikam.example.com'; Status = 'Delivered' }
)
Write-Sample $OutputPath 'message-trace.csv' $schema.MessageTrace $rows

# Source 9
$rows = @(
    [pscustomobject]@{ RunDate = '2026-10-08'; MailboxUserPrincipalName = $people[0].Upn; DeviceId = 'DEV0000000000000000000000001'; DeviceType = 'Outlook'; DeviceOS = 'iOS 17.6'; DeviceAccessState = 'Allowed'; DeviceUserAgent = 'Outlook-iOS/2.0' }
    [pscustomobject]@{ RunDate = '2026-10-08'; MailboxUserPrincipalName = $people[1].Upn; DeviceId = 'DEV0000000000000000000000002'; DeviceType = 'UniversalOutlook'; DeviceOS = 'WINDOWS'; DeviceAccessState = 'Unknown'; DeviceUserAgent = 'microsoft.windowscommunicationsapps' }
)
Write-Sample $OutputPath 'mobile-devices.csv' $schema.MobileDevices $rows

# Source 11
$rows = @(
    [pscustomobject]@{ ReceivedDateTime = '2026-10-07T14:03:11Z'; Id = 'bbbbbbbb-0000-0000-0000-000000000001'; SenderAddress = $people[0].Upn; RecipientAddress = $people[1].Upn; Status = 'delivered'; Size = 45678 }
    [pscustomobject]@{ ReceivedDateTime = '2026-10-07T16:20:45Z'; Id = 'bbbbbbbb-0000-0000-0000-000000000002'; SenderAddress = 'partner@fabrikam.example.com'; RecipientAddress = $people[0].Upn; Status = 'delivered'; Size = 12034 }
)
Write-Sample $OutputPath 'graph-message-trace.csv' $schema.GraphMessageTrace $rows

Remove-Module M365ReportLibrary -Force -ErrorAction SilentlyContinue
