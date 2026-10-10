#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes report-settings.csv: whether the Microsoft 365 usage reports show user
        names or conceal them. A snapshot, appended each run.

    .DESCRIPTION
        Source: Microsoft Graph get adminReportSettings
        (https://learn.microsoft.com/graph/api/adminreportsettings-get, resource:
        https://learn.microsoft.com/graph/api/resources/adminreportsettings), read with
        Get-MgAdminReportSetting. DisplayConcealedNames true means the usage reports hide
        usernames, groups and sites, and by default they do. The same switch is the
        Microsoft 365 admin center checkbox Settings, Org settings, Services, Reports,
        "Conceal user, group, and site names in all reports"
        (https://learn.microsoft.com/microsoft-365/admin/activity-reports/activity-reports).

        This collector only reads the setting. Changing it needs
        ReportSettings.ReadWrite.All and is out of scope: the library is read-only.

        Run this collector before the usage collectors. When it says names are
        concealed, the usage collectors log that their rows cannot be joined to users.

        Needs ReportSettings.Read.All; signed in, an Entra limited admin role
        (https://learn.microsoft.com/graph/reportroot-authorization). In GCC High the
        Graph page marks the API unavailable while the activity-reports page says it
        works in all environments, so it is UNVERIFIED there: attempted, with a warning.

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

    # An alternate schema file. The tests use it to exercise the NotAvailable path.
    [string]$SchemaPath = (Join-Path $PSScriptRoot 'LicenseUtilizationSchema.psd1'),

    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Imported by path without -Force: a forced re-import builds a fresh module session
# state, which would discard anything the caller has arranged around the module - and
# that is exactly how the tests put mocks in front of the tenant cmdlets.
Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

. (Join-Path $PSScriptRoot 'LicenseUtilizationHelpers.ps1')


Invoke-LicenseUtilizationSnapshot -Source 'ReportSettings' -CsvName 'report-settings.csv' `
    -Description 'Usage report settings' `
    -License 'ReportSettings.Read.All (an Entra limited admin role when signed in)' `
    -KeyColumn @('RunDate') -OutputPath $OutputPath -Environment $Environment -AppId $AppId `
    -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization `
    -SchemaPath $SchemaPath -SkipConnect:$SkipConnect `
    -Fetch { Get-MgAdminReportSetting -ErrorAction Stop } `
    -Map {
        param($settings, $runDate)

        [pscustomobject]@{
            RunDate               = $runDate
            DisplayConcealedNames = [string](Get-GraphAdditionalProperty -Object $settings -Name 'displayConcealedNames')
        }
    }
