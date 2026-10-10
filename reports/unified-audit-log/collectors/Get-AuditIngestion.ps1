#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes audit-ingestion.csv: whether unified audit log ingestion is on, one row per run.

    .DESCRIPTION
        Source 4: Get-AdminAuditLogConfig | UnifiedAuditLogIngestionEnabled, run in Exchange Online
        PowerShell. In Security & Compliance PowerShell the property is always False even when search is
        on, so this collector connects to Exchange Online only
        (https://learn.microsoft.com/purview/audit-log-enable-disable).

        State source: a snapshot stamped with the run date, appended each run. After auditing is turned
        on it can take up to 60 minutes to take effect and several hours before events are searchable,
        so a search that returns nothing in that window is not proof of no activity.

        Role: View-Only Audit Logs or Audit Logs
        (https://learn.microsoft.com/purview/audit-log-search-script). Available in Commercial.
        UNVERIFIED in GCC and GCC High: attempted, and a refusal is logged.

    .EXAMPLE
        ./Get-AuditIngestion.ps1 -OutputPath ./out
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

. (Join-Path $PSScriptRoot 'UnifiedAuditLogHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'UnifiedAuditLogSchema.psd1')
$columns = $schema.AuditIngestion
$source = 'audit-ingestion'
$csvPath = Join-Path $OutputPath 'audit-ingestion.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

if (Test-AuditSourceSkipped -Source 'AuditIngestion' -LogSource $source -CsvPath $csvPath -Column $columns `
        -OutputPath $OutputPath -Environment $Environment -Schema $schema) {
    return
}

$connectedHere = -not $SkipConnect
if ($connectedHere) {
    Connect-M365Service -Service ExchangeOnline -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization
}

try {
    try {
        $config = Get-AdminAuditLogConfig -ErrorAction Stop
    }
    catch {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            'Get-AdminAuditLogConfig is unavailable to this sign-in or cloud ({0}). It needs View-Only Audit Logs or Audit Logs in Exchange Online. Writing the header only.' -f $_.Exception.Message)
        Export-AppendCsv -Path $csvPath -Column $columns
        return
    }

    $row = [pscustomobject]@{
        RunDate                         = $runDate
        UnifiedAuditLogIngestionEnabled = Get-ObjectText -Object $config -Name 'UnifiedAuditLogIngestionEnabled'
    }
    $result = Export-AppendCsv -Path $csvPath -Rows @($row) -Column $columns -KeyColumn @('RunDate') -PassThru
    Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
        'audit-ingestion.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
}
finally {
    if ($connectedHere) {
        Disconnect-ExchangeOnline -Confirm:$false
    }
}
