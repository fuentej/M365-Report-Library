#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes report-settings.csv: whether usage reports show names or concealed identifiers, appended each run.

    .DESCRIPTION
        Source 5: Microsoft Graph GET /admin/reportSettings
        (https://learn.microsoft.com/graph/api/adminreportsettings-get), read with
        Get-MgAdminReportSetting. displayConcealedNames is a property of the one adminReportSettings
        object (https://learn.microsoft.com/graph/api/resources/adminreportsettings). True means
        names are concealed: the user columns in teams-user-activity-user-detail.csv and
        teams-device-usage-user-detail.csv, and the team name in the team-activity.csv written by the
        Teams and Groups lifecycle report, then hold concealed identifiers and cannot be joined to the
        users or groups files. The setting is Settings, Org settings, Services, Reports, "Conceal user,
        group, and site names in all reports" in the Microsoft 365 admin center
        (https://learn.microsoft.com/microsoft-365/admin/activity-reports/activity-reports#show-user-group-or-site-details-in-usage-reports).
        This script only reads it.

        Needs ReportSettings.Read.All. Availability in GCC and GCC High is UNVERIFIED: the Graph page marks
        US Government L4 unsupported while the usage reports overview says the setting has an API in
        all environments. Run this collector first.

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

. (Join-Path $PSScriptRoot 'TeamsActivityHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'TeamsActivitySchema.psd1')
$columns = $schema.ReportSettings
$source = 'report-settings'
$csvPath = Join-Path $OutputPath 'report-settings.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

if (Test-TeamsSourceSkipped -Source 'ReportSettings' -LogSource $source -CsvPath $csvPath -Column $columns `
        -OutputPath $OutputPath -Environment $Environment -Schema $schema) {
    return
}

if (-not $SkipConnect) {
    Connect-M365Service -Service Graph -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization
}

try {
    $settings = Get-MgAdminReportSetting -ErrorAction Stop
}
catch {
    Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
        'The report settings are unavailable to this sign-in or cloud ({0}). They need ReportSettings.Read.All. Writing the header only.' -f $_.Exception.Message)
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}

$concealed = ''
if ($null -ne $settings) {
    $concealed = Get-CsvBoolean (Get-GraphAdditionalProperty -Object $settings -Name 'displayConcealedNames')
}
if ($concealed -eq 'True') {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        'Usage reports conceal names in this tenant. The user columns cannot be joined to the users file.')
}

$row = [pscustomobject]@{ RunDate = $runDate; DisplayConcealedNames = $concealed }
$result = Export-AppendCsv -Path $csvPath -Rows @($row) -Column $columns -KeyColumn @('RunDate') -PassThru
Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'report-settings.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
