#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes audit-retention-policies.csv: the audit log retention policies in force, one snapshot
        per run.

    .DESCRIPTION
        Source 5: Get-UnifiedAuditLogRetentionPolicy in Security & Compliance PowerShell
        (https://learn.microsoft.com/powershell/module/exchangepowershell/get-unifiedauditlogretentionpolicy).
        Properties: Priority, Name, RecordTypes, Operations, UserIds, RetentionDuration. List values
        are joined with a semicolon. RetentionDuration is written as returned: the portal offers
        durations (7 Days, 30 Days, 3, 5 and 7 Years) that are not among the five names in the cmdlet's
        accepted list, and none is dropped.

        State source: a snapshot stamped with the run date. The cmdlet does not return the default
        audit log retention policy, so no rows is not "no one-year retention"
        (https://learn.microsoft.com/purview/audit-log-retention-policies). Retention policies are an
        Audit (Premium) capability. The cmdlet page does not name the read role. GCC and GCC High are
        UNVERIFIED for the cmdlet: attempted, and a refusal is logged.

    .EXAMPLE
        ./Get-AuditRetentionPolicies.ps1 -OutputPath ./out
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
$columns = $schema.AuditRetentionPolicies
$source = 'audit-retention-policies'
$csvPath = Join-Path $OutputPath 'audit-retention-policies.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

if (Test-AuditSourceSkipped -Source 'AuditRetentionPolicies' -LogSource $source -CsvPath $csvPath -Column $columns `
        -OutputPath $OutputPath -Environment $Environment -Schema $schema) {
    return
}

$connectedHere = -not $SkipConnect
if ($connectedHere) {
    Connect-M365Service -Service SecurityCompliance -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization
}

try {
    try {
        $policies = @(Get-UnifiedAuditLogRetentionPolicy -ErrorAction Stop | Where-Object { $null -ne $_ })
    }
    catch {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            'Get-UnifiedAuditLogRetentionPolicy is unavailable to this sign-in or cloud ({0}). Writing the header only.' -f $_.Exception.Message)
        Export-AppendCsv -Path $csvPath -Column $columns
        return
    }

    $rows = foreach ($policy in $policies) {
        [pscustomobject]@{
            RunDate           = $runDate
            Priority          = Get-ObjectText -Object $policy -Name 'Priority'
            Name              = Get-ObjectText -Object $policy -Name 'Name'
            RecordTypes       = Get-ObjectText -Object $policy -Name 'RecordTypes'
            Operations        = Get-ObjectText -Object $policy -Name 'Operations'
            UserIds           = Get-ObjectText -Object $policy -Name 'UserIds'
            RetentionDuration = Get-ObjectText -Object $policy -Name 'RetentionDuration'
        }
    }
    $result = Export-AppendCsv -Path $csvPath -Rows @($rows) -Column $columns -KeyColumn @('RunDate', 'Priority', 'Name') -PassThru
    if ($policies.Count -eq 0) {
        Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
            'No custom retention policy was returned. The cmdlet does not return the default policy, so this is not proof that records are not kept for a year.')
    }
    Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
        'audit-retention-policies.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
}
finally {
    if ($connectedHere) {
        Disconnect-ExchangeOnline -Confirm:$false
    }
}
