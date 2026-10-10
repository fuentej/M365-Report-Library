#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes mailbox-usage-detail.csv: storage, item counts, quotas and last activity per mailbox, appended each run.

    .DESCRIPTION
        Source 1: Microsoft Graph getMailboxUsageDetail
        (https://learn.microsoft.com/graph/api/reportroot-getmailboxusagedetail), read with
        Get-MgReportMailboxUsageDetail -OutFile. Graph answers with a 302 to a preauthenticated
        CSV download valid for a few minutes; the cmdlet saves it to a temporary file, which
        this script reads and removes.

        A snapshot of a rolling period (D7, D30, D90 or D180) stamped with the run date.
        The header list includes Deleted Item Quota (Byte) and Has Archive, which the example
        schema on that page omits; a CSV without them leaves those columns empty. QuotaStatus is
        derived with the admin center's four categories (at or above a quota is the next one).
        The CSV has no recipient-type column, so join to mailboxes.csv to tell shared from user.

        Not available in GCC High (US Government L4 is marked unsupported on the API page).
        Needs Reports.Read.All; a delegated caller also needs a limited admin role such as
        Reports Reader. Global Reader and Usage Summary Reports Reader do not receive the
        detail rows. When the tenant conceals names (see report-settings.csv), the user columns
        hold concealed identifiers and cannot be joined to the Exchange mailbox files.

    .EXAMPLE
        ./Get-MailboxUsageDetail.ps1 -OutputPath ./out
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$OutputPath,

    [ValidateSet('Commercial', 'GCC', 'GCCHigh')]
    [string]$Environment = 'Commercial',

    [string]$AppId,
    [string]$CertificateThumbprint,
    [string]$TenantId,
    [string]$Organization,
    [ValidateSet('D7', 'D30', 'D90', 'D180')]
    [string]$Period = 'D30',

    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Imported by path without -Force: a forced re-import builds a fresh module session
# state, which would discard anything the caller has arranged around the module - and
# that is exactly how the tests put mocks in front of the tenant cmdlets.
Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

. (Join-Path $PSScriptRoot 'ExchangeActivityHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'ExchangeActivitySchema.psd1')
$columns = $schema.MailboxUsageDetail
$source = 'mailbox-usage-detail'
$csvPath = Join-Path $OutputPath 'mailbox-usage-detail.csv'

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

if (Test-ExchangeSourceSkipped -Source 'MailboxUsageDetail' -LogSource $source -CsvPath $csvPath -Column $columns `
        -OutputPath $OutputPath -Environment $Environment -Schema $schema) {
    return
}

if (-not $SkipConnect) {
    Connect-M365Service -Service Graph -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization
}

function Read-Column { param($Row, [string[]]$Header) Get-ReportColumnValue -Row $Row -Header $Header }

$map = {
    param($row, $runDate)
    $used = Read-Column $row 'Storage Used (Byte)'
    $warn = Read-Column $row 'Issue Warning Quota (Byte)'
    $send = Read-Column $row 'Prohibit Send Quota (Byte)'
    $both = Read-Column $row 'Prohibit Send/Receive Quota (Byte)'
    [pscustomobject]@{
        RunDate                      = $runDate
        ReportRefreshDate            = Read-Column $row 'Report Refresh Date'
        UserPrincipalName            = Read-Column $row 'User Principal Name'
        DisplayName                  = Read-Column $row 'Display Name'
        IsDeleted                    = Read-Column $row 'Is Deleted'
        DeletedDate                  = Read-Column $row 'Deleted Date'
        CreatedDate                  = Read-Column $row 'Created Date'
        LastActivityDate             = Read-Column $row 'Last Activity Date'
        ItemCount                    = Read-Column $row 'Item Count'
        StorageUsedByte              = $used
        IssueWarningQuotaByte        = $warn
        ProhibitSendQuotaByte        = $send
        ProhibitSendReceiveQuotaByte = $both
        DeletedItemCount             = Read-Column $row 'Deleted Item Count'
        DeletedItemSizeByte          = Read-Column $row 'Deleted Item Size (Byte)'
        DeletedItemQuotaByte         = Read-Column $row 'Deleted Item Quota (Byte)'
        HasArchive                   = Read-Column $row 'Has Archive'
        ReportPeriod                 = Read-Column $row 'Report Period'
        QuotaStatus                  = Get-MailboxQuotaStatus -StorageUsed $used -IssueWarning $warn -ProhibitSend $send -ProhibitSendReceive $both
    }
}

Invoke-UsageReportCollector -OutputPath $OutputPath -LogSource $source -CsvName 'mailbox-usage-detail.csv' -Column $columns `
    -KeyColumn @('RunDate', 'UserPrincipalName', 'ReportPeriod') -ReportName 'Exchange mailbox usage detail' -MapRow $map -Fetch {
        param($file)
        Get-MgReportMailboxUsageDetail -Period $Period -OutFile $file -ErrorAction Stop
    }

