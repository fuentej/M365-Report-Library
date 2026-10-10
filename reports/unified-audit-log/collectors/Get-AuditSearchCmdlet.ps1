#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes audit-search-cmdlet.csv: unified audit log records from Search-UnifiedAuditLog,
        appended from where the last run stopped.

    .DESCRIPTION
        Source 1: Search-UnifiedAuditLog in Exchange Online PowerShell
        (https://learn.microsoft.com/powershell/module/exchangepowershell/search-unifiedauditlog).
        Event source. A run starts at the latest CreationTime already in the file, or -LookbackDays
        back on the first run, and reads the range in -SliceMinutes slices, oldest first. Each slice
        is one ReturnLargeSet session (-ResultSize 5000, one SessionId, the same command for every
        call) read until a call returns nothing. A slice that reaches the 50,000-record cap is not
        the full slice, so it is read again as two halves. A slice is written only when it has been
        read completely, so the file's newest CreationTime never runs ahead of a slice that failed.

        Start and end are UTC dates with a time. Records can take hours to become searchable
        (https://learn.microsoft.com/purview/audit-log-enable-disable); a record that appears after
        a later one was exported is not picked up by a run that has moved past it. Pass a larger
        -LookbackDays on a first run, or delete the file to read again.

        Role: Exchange Online role View-Only Audit Logs or Audit Logs
        (https://learn.microsoft.com/purview/audit-log-search-script). Audit (Standard) covers the
        cmdlet. Availability in GCC and GCC High is UNVERIFIED: no page read names the cmdlet there,
        so it is attempted and a refusal is logged.

    .EXAMPLE
        ./Get-AuditSearchCmdlet.ps1 -OutputPath ./out
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

    # How far back the first run looks, in days. Later runs resume from the file.
    [ValidateRange(1, 180)]
    [int]$LookbackDays = 7,

    # The slice each session reads. The audit log search script article uses 60 minutes.
    [ValidateRange(1, 1440)]
    [int]$SliceMinutes = 60,

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

$columns = $schema.AuditSearchCmdlet
$source = 'audit-search-cmdlet'
$csvPath = Join-Path $OutputPath 'audit-search-cmdlet.csv'

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

if (Test-AuditSourceSkipped -Source 'AuditSearchCmdlet' -LogSource $source -CsvPath $csvPath -Column $columns `
        -OutputPath $OutputPath -Environment $Environment -Schema $schema) {
    return
}

$connectedHere = -not $SkipConnect
if ($connectedHere) {
    Connect-M365Service -Service ExchangeOnline -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization
}

try {
    $end = [datetime]::UtcNow
    $start = Get-AuditRunStart -CsvPath $csvPath -WatermarkColumn 'CreationTime' -End $end -LookbackDays $LookbackDays
    Export-AppendCsv -Path $csvPath -Column $columns
    $total = [pscustomobject]@{ Written = 0; Skipped = 0 }

    try {
        foreach ($slice in Split-DateRange -Start $start -End $end -WindowMinutes $SliceMinutes) {
            $records = @(Get-AuditSearchSlice -Start $slice.Start -End $slice.End -OutputPath $OutputPath -LogSource $source)
            $rows = @($records | ForEach-Object { ConvertTo-AuditSearchRow -Record $_ })
            $result = Export-AppendCsv -Path $csvPath -Rows $rows -Column $columns -KeyColumn @('RecordId') -PassThru
            $total.Written += $result.Written
            $total.Skipped += $result.Skipped
        }
        Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
            'audit-search-cmdlet.csv: {0} rows written, {1} skipped. Window {2:yyyy-MM-ddTHH:mm:ssZ} to {3:yyyy-MM-ddTHH:mm:ssZ}.' -f $total.Written, $total.Skipped, $start, $end)
    }
    catch {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            'Search-UnifiedAuditLog failed ({0}). It needs the View-Only Audit Logs or Audit Logs role, and auditing must be on. Slices already written are kept; the next run resumes from the latest CreationTime.' -f $_.Exception.Message)
    }
}
finally {
    if ($connectedHere) {
        Disconnect-ExchangeOnline -Confirm:$false
    }
}
