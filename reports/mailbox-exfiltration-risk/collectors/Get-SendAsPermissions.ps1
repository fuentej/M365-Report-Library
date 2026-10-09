#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes send-as-permissions.csv: who holds Send As on which mailbox or group,
        appended each run.

    .DESCRIPTION
        Source: Get-RecipientPermission -ResultSize Unlimited
        (https://learn.microsoft.com/powershell/module/exchangepowershell/get-recipientpermission).
        A cloud-only cmdlet; Learn recommends Get-EXORecipientPermission. ResultSize
        defaults to 1000, so a call without Unlimited drops grants past that cap. Rows
        are written as returned, with AccessControlType and IsInherited so a report can
        filter the default entries.

        The matching role, View-Only Recipients, is an inference and is UNVERIFIED.

    .EXAMPLE
        ./Get-SendAsPermissions.ps1 -OutputPath ./out
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
$columns = $schema.SendAsPermissions
$source = 'send-as-permissions'
$csvPath = Join-Path $OutputPath 'send-as-permissions.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-MailboxSourceAvailability -Source 'SendAsPermissions' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping send-as-permissions.csv. $($availability.Reason) $($availability.Reference)")
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
        $permissions = @(Get-RecipientPermission -ResultSize Unlimited -ErrorAction Stop)
    
        $rows = foreach ($permission in $permissions) {
            [pscustomobject]@{
                RunDate           = $runDate
                Identity          = [string]$permission.Identity
                Trustee           = [string]$permission.Trustee
                AccessRights      = Join-ListValue $permission.AccessRights
                AccessControlType = [string]$permission.AccessControlType
                IsInherited       = Get-CsvBoolean $permission.IsInherited
            }
        }
    
        $result = Export-AppendCsv -Path $csvPath -Rows @($rows) -Column $columns -KeyColumn @('RunDate', 'Identity', 'Trustee', 'AccessControlType') -PassThru
        Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
            'send-as-permissions.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
    }
    catch {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            'Get-RecipientPermission is unavailable to this sign-in ({0}). It needs a role that can view recipients (View-Only Recipients matches; the least privileged role is UNVERIFIED). Writing the header only.' -f $_.Exception.Message)
        Export-AppendCsv -Path $csvPath -Column $columns
        return
    }
}
finally {
    if ($connectedHere) {
        Disconnect-ExchangeOnline -Confirm:$false
    }
}
