#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes onedrive-usage-account-detail.csv: storage, file counts and last activity per OneDrive account, appended each run.

    .DESCRIPTION
        Source 2: Microsoft Graph getOneDriveUsageAccountDetail
        (https://learn.microsoft.com/graph/api/reportroot-getonedriveusageaccountdetail), read with
        Get-MgReportOneDriveUsageAccountDetail -OutFile.
        One row per account per run and period. The default period is D30. The Graph CSV has no Site Id
        column. Last Activity Date is blank when the OneDrive has no file activity, and a deleted
        account takes at least seven days to be marked deleted
        (https://learn.microsoft.com/microsoft-365/admin/activity-reports/onedrive-for-business-usage).

        Not available in GCC High (US Government L4 is marked unsupported on the API page). A GCC High
        run writes the header only and logs why. Needs Reports.Read.All; a delegated caller also needs
        a limited admin role such as Reports Reader. Global Reader and Usage Summary Reports Reader do
        not receive the detail rows. When the tenant conceals names (see report-settings.csv), the
        identifier columns hold concealed values and cannot be joined to drive-quota.csv or users.csv.

    .EXAMPLE
        ./Get-OneDriveUsageAccountDetail.ps1 -OutputPath ./out
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
$columns = $schema.OneDriveUsageAccountDetail
$source = 'onedrive-usage-account-detail'
$csvPath = Join-Path $OutputPath 'onedrive-usage-account-detail.csv'

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

if (Test-SharePointSourceSkipped -Source 'OneDriveUsageAccountDetail' -LogSource $source -CsvPath $csvPath -Column $columns `
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
        SiteUrl                    = Read-Column $row 'Site URL'
        OwnerDisplayName           = Read-Column $row 'Owner Display Name'
        IsDeleted                  = Read-Column $row 'Is Deleted'
        LastActivityDate           = Read-Column $row 'Last Activity Date'
        FileCount                  = Read-Column $row 'File Count'
        ActiveFileCount            = Read-Column $row 'Active File Count'
        StorageUsedByte            = Read-Column $row 'Storage Used (Byte)'
        StorageAllocatedByte       = Read-Column $row 'Storage Allocated (Byte)'
        OwnerPrincipalName         = Read-Column $row 'Owner Principal Name'
        ReportPeriod               = Read-Column $row 'Report Period'
    }
}

Invoke-UsageReportCollector -OutputPath $OutputPath -LogSource $source -CsvName 'onedrive-usage-account-detail.csv' -Column $columns `
    -KeyColumn @('RunDate', 'OwnerPrincipalName', 'SiteUrl', 'ReportPeriod') -ReportName 'OneDrive usage account' -MapRow $map -Fetch {
        param($file)
        Get-MgReportOneDriveUsageAccountDetail -Period $Period -OutFile $file -ErrorAction Stop
    }
