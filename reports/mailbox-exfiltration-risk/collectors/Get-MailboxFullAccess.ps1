#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes mailbox-full-access.csv: who holds Full Access on which mailbox, appended
        each run.

    .DESCRIPTION
        Source: Get-MailboxPermission -Identity <mailbox> -ResultSize Unlimited, one call
        per mailbox
        (https://learn.microsoft.com/powershell/module/exchangepowershell/get-mailboxpermission).
        -Identity is mandatory. A row is kept only when AccessRights is like 'Full*',
        Deny is false and IsInherited is false, and NT AUTHORITY\SELF is dropped: the cmdlet
        page assigns FullAccess to SELF by default, inherited entries for Administrator
        and Organization Management appear to allow FullAccess, and a Deny entry removes
        it (https://learn.microsoft.com/exchange/recipients-in-exchange-online/manage-permissions-for-recipients).
        The mailbox list comes from Get-EXOMailbox -ResultSize Unlimited.

        The matching role, View-Only Recipients, is an inference and is UNVERIFIED. A
        mailbox the sign-in cannot read is logged and skipped.

    .PARAMETER MailboxLimit
        Read at most this many mailboxes. For a trial run; 0 means all.

    .EXAMPLE
        ./Get-MailboxFullAccess.ps1 -OutputPath ./out
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

    [ValidateRange(0, 1000000)]
    [int]$MailboxLimit = 0,

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
$columns = $schema.MailboxFullAccess
$source = 'mailbox-full-access'
$csvPath = Join-Path $OutputPath 'mailbox-full-access.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-MailboxSourceAvailability -Source 'MailboxFullAccess' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping mailbox-full-access.csv. $($availability.Reason) $($availability.Reference)")
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
        $mailboxes = @(Get-EXOMailbox -ResultSize Unlimited -ErrorAction Stop)
        if ($MailboxLimit -gt 0) { $mailboxes = @($mailboxes | Select-Object -First $MailboxLimit) }
    
        $rows = [System.Collections.Generic.List[object]]::new()
        $failed = 0
        foreach ($mailbox in $mailboxes) {
            $permissions = $null
            try {
                $permissions = @(Get-MailboxPermission -Identity ([string]$mailbox.UserPrincipalName) -ResultSize Unlimited -ErrorAction Stop)
            }
            catch {
                $failed++
                Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
                    'Get-MailboxPermission failed for {0} ({1}); skipping the mailbox.' -f $mailbox.UserPrincipalName, $_.Exception.Message)
                continue
            }
    
            foreach ($permission in $permissions) {
                if (-not (Test-FullAccessRow -Row $permission)) { continue }
                $rows.Add([pscustomobject]@{
                        RunDate                          = $runDate
                        MailboxExternalDirectoryObjectId = [string]$mailbox.ExternalDirectoryObjectId
                        MailboxUserPrincipalName         = [string]$mailbox.UserPrincipalName
                        User                             = [string]$permission.User
                        AccessRights                     = Join-ListValue $permission.AccessRights
                    })
            }
        }
    
        if ($failed -gt 0 -and $failed -eq $mailboxes.Count) {
            Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
                'Get-MailboxPermission failed for every mailbox. It needs a role that can view recipients (the least privileged role is UNVERIFIED). Writing the header only.')
            Export-AppendCsv -Path $csvPath -Column $columns
            return
        }
    
        $result = Export-AppendCsv -Path $csvPath -Rows $rows.ToArray() -Column $columns `
            -KeyColumn @('RunDate', 'MailboxExternalDirectoryObjectId', 'User') -PassThru
        Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
            'mailbox-full-access.csv: {0} rows written, {1} skipped; {2} mailbox(es) could not be read.' -f $result.Written, $result.Skipped, $failed)
    }
    catch {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            'Get-EXOMailbox is unavailable to this sign-in ({0}). It needs a role that can view recipients (the least privileged role is UNVERIFIED). Writing the header only.' -f $_.Exception.Message)
        Export-AppendCsv -Path $csvPath -Column $columns
        return
    }
}
finally {
    if ($connectedHere) {
        Disconnect-ExchangeOnline -Confirm:$false
    }
}
