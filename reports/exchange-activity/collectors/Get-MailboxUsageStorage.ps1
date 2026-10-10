#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes mailbox-usage-storage.csv: tenant-wide mailbox storage over time, appended each run.

    .DESCRIPTION
        Source 2: Microsoft Graph getMailboxUsageStorage
        (https://learn.microsoft.com/graph/api/reportroot-getmailboxusagestorage), read with
        Get-MgReportMailboxUsageStorage -OutFile. The default period is D180, the value in the
        contract, so one run holds 180 days of tenant storage. The admin center storage chart
        does not include archive mailboxes
        (https://learn.microsoft.com/microsoft-365/admin/activity-reports/mailbox-usage).

        Not available in GCC High (US Government L4 is marked unsupported on the API page).
        Needs Reports.Read.All; a delegated caller also needs a limited admin role such as
        Reports Reader. Global Reader and Usage Summary Reports Reader do not receive the
        detail rows.

    .EXAMPLE
        ./Get-MailboxUsageStorage.ps1 -OutputPath ./out
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
    # D180 is the contract's storage window. A shorter default cannot be backfilled later.
    [ValidateSet('D7', 'D30', 'D90', 'D180')]
    [string]$Period = 'D180',

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
$columns = $schema.MailboxUsageStorage
$source = 'mailbox-usage-storage'
$csvPath = Join-Path $OutputPath 'mailbox-usage-storage.csv'

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

if (Test-ExchangeSourceSkipped -Source 'MailboxUsageStorage' -LogSource $source -CsvPath $csvPath -Column $columns `
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
    [pscustomobject]@{
        RunDate           = $runDate
        ReportRefreshDate = Read-Column $row 'Report Refresh Date'
        StorageUsedByte   = Read-Column $row 'Storage Used (Byte)'
        ReportDate        = Read-Column $row 'Report Date'
        ReportPeriod      = Read-Column $row 'Report Period'
    }
}

Invoke-UsageReportCollector -OutputPath $OutputPath -LogSource $source -CsvName 'mailbox-usage-storage.csv' -Column $columns `
    -KeyColumn @('RunDate', 'ReportDate', 'ReportPeriod') -ReportName 'Exchange mailbox usage storage' -MapRow $map -Fetch {
        param($file)
        Get-MgReportMailboxUsageStorage -Period $Period -OutFile $file -ErrorAction Stop
    }

