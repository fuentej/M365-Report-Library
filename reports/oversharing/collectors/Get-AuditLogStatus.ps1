#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes audit-log-status.csv: whether the unified audit log is on, appended each run.

    .DESCRIPTION
        Source: Get-AdminAuditLogConfig UnifiedAuditLogIngestionEnabled in Exchange Online
        PowerShell
        (https://learn.microsoft.com/purview/audit-log-enable-disable). The same property is
        always False in Security & Compliance PowerShell even when auditing is on, so a false
        from that module is not "auditing is off"; this collector reads it through Exchange
        Online only. A search returns nothing when ingestion is off, so an empty
        anonymous-link-events.csv or sharing-events.csv is read against this file. Microsoft
        365 Business Basic, Business Standard and Business Premium, and unmanaged trial
        tenants, do not have auditing on by default (same page).

        Role: the audit roles as in Get-AnonymousLinkEvents.ps1; the role for
        Get-AdminAuditLogConfig itself is UNVERIFIED. Per-cloud availability of the cmdlet is
        UNVERIFIED in GCC and GCC High, so those clouds are attempted with a warning.

    .EXAMPLE
        ./Get-AuditLogStatus.ps1 -OutputPath ./out
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

. (Join-Path $PSScriptRoot 'OversharingHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'OversharingSchema.psd1')
$columns = $schema.AuditLogStatus
$source = 'audit-log-status'
$csvPath = Join-Path $OutputPath 'audit-log-status.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-OversharingSourceAvailability -Source 'AuditLogStatus' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping audit-log-status.csv. $($availability.Reason) $($availability.Reference)")
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}
if ($availability.Status -eq 'Unverified') {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message $availability.Reason
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
            'Get-AdminAuditLogConfig is unavailable to this sign-in ({0}). It needs an audit role; the exact role is UNVERIFIED. Writing the header only.' -f $_.Exception.Message)
        Export-AppendCsv -Path $csvPath -Column $columns
        return
    }

    $row = [pscustomobject]@{
        RunDate                         = $runDate
        UnifiedAuditLogIngestionEnabled = Get-CsvBoolean $config.UnifiedAuditLogIngestionEnabled
    }

    $result = Export-AppendCsv -Path $csvPath -Rows @($row) -Column $columns -KeyColumn 'RunDate' -PassThru
    Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
        'audit-log-status.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
}
finally {
    if ($connectedHere) {
        Disconnect-ExchangeOnline -Confirm:$false
    }
}
