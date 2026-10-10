#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes email-activity-user-detail.csv: send, receive and read counts per user, appended each run.

    .DESCRIPTION
        Source 3: Microsoft Graph getEmailActivityUserDetail
        (https://learn.microsoft.com/graph/api/reportroot-getemailactivityuserdetail), read with
        Get-MgReportEmailActivityUserDetail -OutFile. Exactly one of -Period and -Date is sent. The
        date form returns one day and reaches back 28 days only, so a daily run with -Date keeps
        daily rows contiguous only if it runs at least every 28 days (an inference, not a Learn
        statement). A per-day series is one -Date call per day.

        Not available in GCC High (US Government L4 is marked unsupported on the API page).
        Needs Reports.Read.All; a delegated caller also needs a limited admin role such as
        Reports Reader. Global Reader and Usage Summary Reports Reader do not receive the
        detail rows. When the tenant conceals names (see report-settings.csv), the user columns
        hold concealed identifiers and cannot be joined to the Exchange mailbox files.

    .EXAMPLE
        ./Get-EmailActivityUserDetail.ps1 -OutputPath ./out
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

    # A single day, within the last 28 days. When set it is sent instead of -Period.
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
$columns = $schema.EmailActivityUserDetail
$source = 'email-activity-user-detail'
$csvPath = Join-Path $OutputPath 'email-activity-user-detail.csv'

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

if (Test-ExchangeSourceSkipped -Source 'EmailActivityUserDetail' -LogSource $source -CsvPath $csvPath -Column $columns `
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
        RunDate                = $runDate
        QueryDate              = $queryDate
        ReportRefreshDate      = Read-Column $row 'Report Refresh Date'
        UserPrincipalName      = Read-Column $row 'User Principal Name'
        DisplayName            = Read-Column $row 'Display Name'
        IsDeleted              = Read-Column $row 'Is Deleted'
        DeletedDate            = Read-Column $row 'Deleted Date'
        LastActivityDate       = Read-Column $row 'Last Activity Date'
        SendCount              = Read-Column $row 'Send Count'
        ReceiveCount           = Read-Column $row 'Receive Count'
        ReadCount              = Read-Column $row 'Read Count'
        MeetingCreatedCount    = Read-Column $row 'Meeting Created Count'
        MeetingInteractedCount = Read-Column $row 'Meeting Interacted Count'
        AssignedProducts       = Read-Column $row 'Assigned Products'
        ReportPeriod           = Read-Column $row 'Report Period'
    }
}

Invoke-UsageReportCollector -OutputPath $OutputPath -LogSource $source -CsvName 'email-activity-user-detail.csv' -Column $columns `
    -KeyColumn @('RunDate', 'QueryDate', 'UserPrincipalName', 'ReportPeriod') -ReportName 'Exchange email activity' -MapRow $map -Fetch {
        param($file)
        if ($queryDate) {
            Get-MgReportEmailActivityUserDetail -Date $Date -OutFile $file -ErrorAction Stop
        }
        else {
            Get-MgReportEmailActivityUserDetail -Period $Period -OutFile $file -ErrorAction Stop
        }
    }

