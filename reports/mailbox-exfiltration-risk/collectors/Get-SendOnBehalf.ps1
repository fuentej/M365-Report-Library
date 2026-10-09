#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes send-on-behalf.csv: one row per Send on Behalf grant, appended each run.

    .DESCRIPTION
        Source: GrantSendOnBehalfTo, returned by Get-EXOMailbox -ResultSize Unlimited
        -PropertySets Minimum, Delivery
        (https://learn.microsoft.com/exchange/recipients-in-exchange-online/manage-permissions-for-recipients).
        Full Access is in Get-MailboxFullAccess.ps1 and Send As in
        Get-SendAsPermissions.ps1. A user with both Send As and Send on Behalf always uses
        Send As.

        The matching role, View-Only Recipients, is an inference and is UNVERIFIED.

    .EXAMPLE
        ./Get-SendOnBehalf.ps1 -OutputPath ./out
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
$columns = $schema.SendOnBehalf
$source = 'send-on-behalf'
$csvPath = Join-Path $OutputPath 'send-on-behalf.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-MailboxSourceAvailability -Source 'SendOnBehalf' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping send-on-behalf.csv. $($availability.Reason) $($availability.Reference)")
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
        $mailboxes = @(Get-EXOMailbox -ResultSize Unlimited -PropertySets Minimum, Delivery -ErrorAction Stop)
    
        $rows = foreach ($mailbox in $mailboxes) {
            foreach ($delegate in @($mailbox.GrantSendOnBehalfTo | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })) {
                [pscustomobject]@{
                    RunDate                   = $runDate
                    ExternalDirectoryObjectId = [string]$mailbox.ExternalDirectoryObjectId
                    UserPrincipalName         = [string]$mailbox.UserPrincipalName
                    PrimarySmtpAddress        = [string]$mailbox.PrimarySmtpAddress
                    Delegate                  = [string]$delegate
                }
            }
        }
    
        $result = Export-AppendCsv -Path $csvPath -Rows @($rows) -Column $columns -KeyColumn @('RunDate', 'ExternalDirectoryObjectId', 'Delegate') -PassThru
        Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
            'send-on-behalf.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
    }
    catch {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            'Get-EXOMailbox is unavailable to this sign-in ({0}). It needs a role that can view recipients (View-Only Recipients matches; the least privileged role is UNVERIFIED). Writing the header only.' -f $_.Exception.Message)
        Export-AppendCsv -Path $csvPath -Column $columns
        return
    }
}
finally {
    if ($connectedHere) {
        Disconnect-ExchangeOnline -Confirm:$false
    }
}
