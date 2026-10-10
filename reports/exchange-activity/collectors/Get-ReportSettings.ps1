#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes report-settings.csv: whether usage reports show names or concealed identifiers, appended each run.

    .DESCRIPTION
        Source 5: Microsoft Graph GET /admin/reportSettings
        (https://learn.microsoft.com/graph/api/adminreportsettings-get), read with
        Get-MgAdminReportSetting. The result is one adminReportSettings object; displayConcealedNames
        is a property of that object
        (https://learn.microsoft.com/graph/api/resources/adminreportsettings). True means names are
        concealed: the user columns in mailbox-usage-detail.csv, email-activity-user-detail.csv and
        email-app-usage-user-detail.csv then hold concealed identifiers and cannot be joined to
        mailboxes.csv. The setting is Settings, Org settings, Services, Reports, "Conceal user, group,
        and site names in all reports" in the Microsoft 365 admin center. This script only reads it.

        Needs ReportSettings.Read.All. Availability in GCC High is UNVERIFIED: the Graph page marks
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

. (Join-Path $PSScriptRoot 'ExchangeActivityHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'ExchangeActivitySchema.psd1')
$columns = $schema.ReportSettings
$source = 'report-settings'
$csvPath = Join-Path $OutputPath 'report-settings.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

if (Test-ExchangeSourceSkipped -Source 'ReportSettings' -LogSource $source -CsvPath $csvPath -Column $columns `
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
        'Usage reports conceal names in this tenant. The user columns cannot be joined to the Exchange mailbox files.')
}

$row = [pscustomobject]@{ RunDate = $runDate; DisplayConcealedNames = $concealed }
$result = Export-AppendCsv -Path $csvPath -Rows @($row) -Column $columns -KeyColumn @('RunDate') -PassThru
Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'report-settings.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
