#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes report-settings.csv: whether usage reports show names or concealed identifiers, appended each run.

    .DESCRIPTION
        Source 7: Microsoft Graph GET /admin/reportSettings
        (https://learn.microsoft.com/graph/api/adminreportsettings-get), read with
        Get-MgAdminReportSetting. The result is one adminReportSettings object; displayConcealedNames
        is a property of that object
        (https://learn.microsoft.com/graph/api/resources/adminreportsettings). True means names are
        concealed: site ids, site URLs and user principal names in sharepoint-site-usage-detail.csv,
        onedrive-usage-account-detail.csv, sharepoint-activity-user-detail.csv and
        onedrive-activity-user-detail.csv then hold concealed values and cannot be joined to
        drive-quota.csv, spo-sites.csv or users.csv. The setting is Settings, Org settings, Services,
        Reports, "Conceal user, group, and site names in all reports" in the Microsoft 365 admin
        center. This script only reads it.

        Needs ReportSettings.Read.All. Availability in GCC High is UNVERIFIED: the Graph page marks
        US Government L4 unsupported while the usage reports overview says the setting has an API in
        all environments. Run this collector first (Run-All.ps1 does).

    .EXAMPLE
        ./Get-ReportSettings.ps1 -OutputPath ./out
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
$columns = $schema.ReportSettings
$source = 'report-settings'
$csvPath = Join-Path $OutputPath 'report-settings.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

if (Test-SharePointSourceSkipped -Source 'ReportSettings' -LogSource $source -CsvPath $csvPath -Column $columns `
        -OutputPath $OutputPath -Environment $Environment -Schema $schema) {
    return
}

if (-not $SkipConnect) {
    Connect-M365Service -Service Graph -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization
}

try {
    $settings = Invoke-ReadWithThrottleRetry -OutputPath $OutputPath -Source $source -Action {
        Get-MgAdminReportSetting -ErrorAction Stop
    }
}
catch {
    if (Test-GraphThrottleStatus -ErrorRecord $_) {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            'The report settings are still throttled after retries ({0}). A 429 or 503 is not an empty report. Writing the header only.' -f $_.Exception.Message)
    }
    else {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            'The report settings are unavailable to this sign-in or cloud ({0}). They need ReportSettings.Read.All. Writing the header only.' -f $_.Exception.Message)
    }
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}

$concealed = ''
if ($null -ne $settings) {
    $concealed = Get-CsvBoolean (Get-GraphAdditionalProperty -Object $settings -Name 'displayConcealedNames')
}
if ($concealed -eq 'True') {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        'Usage reports conceal names in this tenant. The usage report rows cannot be joined to the site list.')
}

$row = [pscustomobject]@{ RunDate = $runDate; DisplayConcealedNames = $concealed }
$result = Export-AppendCsv -Path $csvPath -Rows @($row) -Column $columns -KeyColumn @('RunDate') -PassThru
Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'report-settings.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
