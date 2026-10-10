#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes retention-reference.csv: how far back the sign-in and audit logs reach at each
        licence level. A snapshot stamped with the run date. Makes no tenant call.

    .DESCRIPTION
        Source 5 of docs/candidates/entra-activity.md: the public Microsoft Entra data
        retention page
        (https://learn.microsoft.com/entra/identity/monitoring-health/reference-reports-data-retention).
        Sign-in and audit logs are 7 days on Free and 30 days on P1 and P2; risky sign-ins
        (not collected by this report) are 7, 30 and 90. The page names no cloud, so
        RetentionStatus is UNVERIFIED for Commercial, GCC and GCC High alike, and the
        collector logs a warning saying so. A run should not promise more history than the
        window it can read.

        No sign-in happens and no cmdlet from a tenant module is called. -Environment only
        labels the rows.

    .EXAMPLE
        ./Get-RetentionReference.ps1 -OutputPath ./out -Environment GCCHigh
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$OutputPath,

    [ValidateSet('Commercial', 'GCC', 'GCCHigh')]
    [string]$Environment = 'Commercial',

    # Accepted so Run-All.ps1 can pass the same sign-in arguments to every collector;
    # this collector does not sign in.
    [string]$AppId,
    [string]$CertificateThumbprint,
    [string]$TenantId,
    [string]$Organization,

    # An alternate schema file. The tests use it to exercise the NotAvailable path.
    [string]$SchemaPath = (Join-Path $PSScriptRoot 'EntraActivitySchema.psd1'),

    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Imported by path without -Force: see Get-InteractiveSignIns.ps1.
Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

. (Join-Path $PSScriptRoot 'EntraActivityHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath $SchemaPath
$columns = [string[]]$schema.RetentionReference
$source = 'retention-reference'
$csvPath = Join-Path $OutputPath 'retention-reference.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-EntraActivitySourceAvailability -Source 'RetentionReference' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping retention-reference.csv. $($availability.Reason) $($availability.Reference)")
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}
if ($availability.Status -eq 'Unverified') {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Log retention in $Environment is UNVERIFIED: the Microsoft retention page names no cloud. The 7 and 30 day windows are written as documented and flagged. $($availability.Reference)")
}

$status = if ($availability.Status -eq 'Unverified') { 'UNVERIFIED' } else { $availability.Status }

$rows = foreach ($level in $schema.RetentionLevels) {
    [pscustomobject]@{
        RunDate                  = $runDate
        Environment              = $Environment
        LicenseLevel             = $level.LicenseLevel
        SignInRetentionDays      = $level.SignInRetentionDays
        AuditRetentionDays       = $level.AuditRetentionDays
        RiskySignInRetentionDays = $level.RiskySignInRetentionDays
        RetentionStatus          = $status
        Reference                = $availability.Reference
    }
}

$result = Export-AppendCsv -Path $csvPath -Rows @($rows) -Column $columns -KeyColumn 'RunDate', 'Environment', 'LicenseLevel' -PassThru
Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'retention-reference.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
