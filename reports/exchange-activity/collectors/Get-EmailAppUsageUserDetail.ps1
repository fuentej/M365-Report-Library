#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes email-app-usage-user-detail.csv: the email apps each user connected with, appended each run.

    .DESCRIPTION
        Source 4: Microsoft Graph getEmailAppUsageUserDetail
        (https://learn.microsoft.com/graph/api/reportroot-getemailappusageuserdetail), read with
        Get-MgReportEmailAppUsageUserDetail -OutFile. Exactly one of -Period and -Date is sent; the
        date form reaches back 30 days. There is no Outlook version column; the admin center
        "Versions" chart is not in this API.

        Not available in GCC High (US Government L4 is marked unsupported on the API page).
        Needs Reports.Read.All; a delegated caller also needs a limited admin role such as
        Reports Reader. Global Reader and Usage Summary Reports Reader do not receive the
        detail rows. When the tenant conceals names (see report-settings.csv), the user columns
        hold concealed identifiers and cannot be joined to the Exchange mailbox files.

    .EXAMPLE
        ./Get-EmailAppUsageUserDetail.ps1 -OutputPath ./out
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

    # A single day, within the last 30 days. When set it is sent instead of -Period.
    [datetime]$Date,

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
$columns = $schema.EmailAppUsageUserDetail
$source = 'email-app-usage-user-detail'
$csvPath = Join-Path $OutputPath 'email-app-usage-user-detail.csv'

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

if (Test-ExchangeSourceSkipped -Source 'EmailAppUsageUserDetail' -LogSource $source -CsvPath $csvPath -Column $columns `
        -OutputPath $OutputPath -Environment $Environment -Schema $schema) {
    return
}

if (-not $SkipConnect) {
    Connect-M365Service -Service Graph -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization
}

# The day asked for, or empty when the period form is used. Exactly one of -Date and -Period is sent.
$queryDate = if ($PSBoundParameters.ContainsKey('Date')) { $Date.ToString('yyyy-MM-dd') } else { '' }

function Read-Column { param($Row, [string[]]$Header) Get-ReportColumnValue -Row $Row -Header $Header }

$map = {
    param($row, $runDate)
    [pscustomobject]@{
        RunDate            = $runDate
        QueryDate          = $queryDate
        ReportRefreshDate  = Read-Column $row 'Report Refresh Date'
        UserPrincipalName  = Read-Column $row 'User Principal Name'
        DisplayName        = Read-Column $row 'Display Name'
        IsDeleted          = Read-Column $row 'Is Deleted'
        DeletedDate        = Read-Column $row 'Deleted Date'
        LastActivityDate   = Read-Column $row 'Last Activity Date'
        MailForMac         = Read-Column $row 'Mail For Mac'
        OutlookForMac      = Read-Column $row 'Outlook For Mac'
        OutlookForWindows  = Read-Column $row 'Outlook For Windows'
        OutlookForMobile   = Read-Column $row 'Outlook For Mobile'
        OtherForMobile     = Read-Column $row 'Other For Mobile'
        OutlookForWeb      = Read-Column $row 'Outlook For Web'
        POP3App            = Read-Column $row 'POP3 App'
        IMAP4App           = Read-Column $row 'IMAP4 App'
        SMTPApp            = Read-Column $row 'SMTP App'
        ReportPeriod       = Read-Column $row 'Report Period'
    }
}

Invoke-UsageReportCollector -OutputPath $OutputPath -LogSource $source -CsvName 'email-app-usage-user-detail.csv' -Column $columns `
    -KeyColumn @('RunDate', 'QueryDate', 'UserPrincipalName', 'ReportPeriod') -ReportName 'Exchange email app usage' -MapRow $map -Fetch {
        param($file)
        if ($queryDate) {
            Get-MgReportEmailAppUsageUserDetail -Date $Date -OutFile $file -ErrorAction Stop
        }
        else {
            Get-MgReportEmailAppUsageUserDetail -Period $Period -OutFile $file -ErrorAction Stop
        }
    }

