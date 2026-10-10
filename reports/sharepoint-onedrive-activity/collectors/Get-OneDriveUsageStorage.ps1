#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes onedrive-usage-storage.csv: OneDrive storage, one row per day, appended each run.

    .DESCRIPTION
        Source 4: Microsoft Graph getOneDriveUsageStorage
        (https://learn.microsoft.com/graph/api/reportroot-getonedriveusagestorage), read with
        Get-MgReportOneDriveUsageStorage -OutFile.
        One row per report date per run and period. The default period is D30; the API accepts D7, D30,
        D90 and D180 and has no date form. The page says storage allocated and consumed but lists no
        allocated column, so none is written.

        Not available in GCC High (US Government L4 is marked unsupported on the API page). A GCC High
        run writes the header only and logs why. Needs Reports.Read.All; a delegated caller also needs
        a limited admin role such as Reports Reader. Global Reader and Usage Summary Reports Reader do
        not receive the detail rows. When the tenant conceals names (see report-settings.csv), the
        identifier columns hold concealed values and cannot be joined to drive-quota.csv or users.csv.

    .EXAMPLE
        ./Get-OneDriveUsageStorage.ps1 -OutputPath ./out
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

. (Join-Path $PSScriptRoot 'SharePointOneDriveHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'SharePointOneDriveSchema.psd1')
$columns = $schema.OneDriveUsageStorage
$source = 'onedrive-usage-storage'
$csvPath = Join-Path $OutputPath 'onedrive-usage-storage.csv'

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

if (Test-SharePointSourceSkipped -Source 'OneDriveUsageStorage' -LogSource $source -CsvPath $csvPath -Column $columns `
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
        RunDate                    = $runDate
        ReportRefreshDate          = Read-Column $row 'Report Refresh Date'
        SiteType                   = Read-Column $row 'Site Type'
        StorageUsedByte            = Read-Column $row 'Storage Used (Byte)'
        ReportDate                 = Read-Column $row 'Report Date'
        ReportPeriod               = Read-Column $row 'Report Period'
    }
}

Invoke-UsageReportCollector -OutputPath $OutputPath -LogSource $source -CsvName 'onedrive-usage-storage.csv' -Column $columns `
    -KeyColumn @('RunDate', 'SiteType', 'ReportDate', 'ReportPeriod') -ReportName 'OneDrive usage storage' -MapRow $map -Fetch {
        param($file)
        Get-MgReportOneDriveUsageStorage -Period $Period -OutFile $file -ErrorAction Stop
    }
