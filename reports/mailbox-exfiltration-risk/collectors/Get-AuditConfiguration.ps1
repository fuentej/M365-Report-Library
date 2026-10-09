#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes audit-configuration.csv: whether auditing is on for the organization,
        whether unified audit log ingestion is on, and which actions each mailbox logs,
        appended each run.

    .DESCRIPTION
        Sources, in Exchange Online PowerShell:
        * Get-OrganizationConfig AuditDisabled. False means mailbox auditing on by default
          is on, and that setting overrides a mailbox-level off.
        * Get-AdminAuditLogConfig UnifiedAuditLogIngestionEnabled. The same property is
          always False in Security & Compliance PowerShell even when ingestion is on, so
          this is read through Exchange Online.
        * Per mailbox, DefaultAuditSet, AuditAdmin, AuditDelegate and AuditOwner, from
          Get-EXOMailbox -ResultSize Unlimited -PropertySets Minimum, Audit (the Audit
          property set of the cmdlet Learn recommends over Get-Mailbox,
          https://learn.microsoft.com/powershell/exchange/cmdlet-property-sets).
          DefaultAuditSet of 'Admin, Delegate, Owner' means the default actions.
        Get-Mailbox always shows AuditEnabled True when on-by-default is on, so it is not a
        per-mailbox on/off switch and is not collected
        (https://learn.microsoft.com/purview/audit-mailboxes#verify-mailbox-auditing-on-by-default-is-turned-on).

        Availability of Get-OrganizationConfig and Get-AdminAuditLogConfig in GCC and
        GCC High is UNVERIFIED, so those clouds are attempted with a warning. Each call that
        is refused is logged and leaves its columns empty; the roles are the audit roles
        (Audit Reader) for the second and the organization-configuration viewing roles
        (UNVERIFIED) for the first.

    .EXAMPLE
        ./Get-AuditConfiguration.ps1 -OutputPath ./out
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

. (Join-Path $PSScriptRoot 'MailboxExfiltrationHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'MailboxExfiltrationSchema.psd1')
$columns = $schema.AuditConfiguration
$source = 'audit-configuration'
$csvPath = Join-Path $OutputPath 'audit-configuration.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-MailboxSourceAvailability -Source 'AuditConfiguration' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping audit-configuration.csv. $($availability.Reason) $($availability.Reference)")
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
    $rows = [System.Collections.Generic.List[object]]::new()
    $problems = 0

    $auditDisabled = ''
    try {
        $orgConfig = Get-OrganizationConfig -ErrorAction Stop
        $auditDisabled = Get-CsvBoolean $orgConfig.AuditDisabled
    }
    catch {
        $problems++
        Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
            'Get-OrganizationConfig is unavailable to this sign-in ({0}); AuditDisabled stays empty.' -f $_.Exception.Message)
    }

    $ingestion = ''
    try {
        $config = Get-AdminAuditLogConfig -ErrorAction Stop
        $ingestion = Get-CsvBoolean $config.UnifiedAuditLogIngestionEnabled
    }
    catch {
        $problems++
        Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
            'Get-AdminAuditLogConfig is unavailable to this sign-in ({0}); UnifiedAuditLogIngestionEnabled stays empty.' -f $_.Exception.Message)
    }

    $rows.Add([pscustomobject]@{
            RunDate                         = $runDate
            Scope                           = 'Organization'
            Identity                        = 'Organization'
            AuditDisabled                   = $auditDisabled
            UnifiedAuditLogIngestionEnabled = $ingestion
            DefaultAuditSet                 = ''
            AuditAdmin                      = ''
            AuditDelegate                   = ''
            AuditOwner                      = ''
        })

    try {
        foreach ($mailbox in @(Get-EXOMailbox -ResultSize Unlimited -PropertySets Minimum, Audit -ErrorAction Stop)) {
            $rows.Add([pscustomobject]@{
                    RunDate                         = $runDate
                    Scope                           = 'Mailbox'
                    Identity                        = [string]$mailbox.UserPrincipalName
                    AuditDisabled                   = ''
                    UnifiedAuditLogIngestionEnabled = ''
                    DefaultAuditSet                 = Join-ListValue $mailbox.DefaultAuditSet
                    AuditAdmin                      = Join-ListValue $mailbox.AuditAdmin
                    AuditDelegate                   = Join-ListValue $mailbox.AuditDelegate
                    AuditOwner                      = Join-ListValue $mailbox.AuditOwner
                })
        }
    }
    catch {
        $problems++
        Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
            'Get-EXOMailbox is unavailable to this sign-in ({0}); no per-mailbox audit actions were collected.' -f $_.Exception.Message)
    }

    if ($problems -ge 3) {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            'Every audit configuration call was refused. Writing the header only.')
        Export-AppendCsv -Path $csvPath -Column $columns
        return
    }

    $result = Export-AppendCsv -Path $csvPath -Rows $rows.ToArray() -Column $columns -KeyColumn @('RunDate', 'Scope', 'Identity') -PassThru
    Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
        'audit-configuration.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
}
finally {
    if ($connectedHere) {
        Disconnect-ExchangeOnline -Confirm:$false
    }
}
