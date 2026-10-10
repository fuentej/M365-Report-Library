#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes mailbox-statistics.csv: size, item count, storage limit status and last logon per mailbox, from Exchange, appended each run.

    .DESCRIPTION
        Source 7: Get-EXOMailboxStatistics -Identity <upn> -PropertySets All
        (https://learn.microsoft.com/powershell/module/exchangepowershell/get-exomailboxstatistics).
        One call per mailbox: the cmdlet takes one -Identity, and a $null or non-existent -Identity
        returns every object, so a mailbox with no UPN is skipped. The mailbox list comes from
        Get-EXOMailbox -ResultSize Unlimited -PropertySets Minimum. The All set adds
        StorageLimitStatus, LastLogonTime, LastLogoffTime, LastLoggedOnUserAccount, IsArchiveMailbox
        and the database quotas
        (https://learn.microsoft.com/powershell/exchange/cmdlet-property-sets).

        LastUserActionTime is not read: Learn says it is being deprecated and is not the last
        active time (https://learn.microsoft.com/powershell/module/exchangepowershell/get-mailboxstatistics).
        LastLogonTime is kept in its own column and is not assumed equal to the Graph Last Activity Date.
        Size values keep the cmdlet's text and the byte count read from it. A run is one call per
        mailbox, so size it for the largest tenant; -MailboxLimit makes a trial run.

        The cmdlet pages defer to the permissions page for the role. App-only sign-in needs
        Exchange.ManageAsApp plus an Exchange role. Organization Management and Recipient
        Management are listed for managing recipients
        (https://learn.microsoft.com/exchange/permissions-exo/feature-permissions); the least
        privileged role that can run this cmdlet is UNVERIFIED until Get-ManagementRole is run
        in the tenant. Availability in GCC and GCC High is UNVERIFIED at cmdlet level: the
        service description says only that remote PowerShell is available
        (https://learn.microsoft.com/office365/servicedescriptions/office-365-platform-service-description/office-365-us-government/exchange-online-for-us-government-environments).

    .EXAMPLE
        ./Get-MailboxStatistics.ps1 -OutputPath ./out
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

    # Read at most this many mailboxes. For a trial run; 0 reads them all.
    [int]$MailboxLimit = 0,

    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Imported by path without -Force: a forced re-import builds a fresh module session
# state, which would discard anything the caller has arranged around the module - and
# that is exactly how the tests put mocks in front of the tenant cmdlets.
Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

. (Join-Path $PSScriptRoot 'ExchangeActivityHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'ExchangeActivitySchema.psd1')
$columns = $schema.MailboxStatistics
$source = 'mailbox-statistics'
$csvPath = Join-Path $OutputPath 'mailbox-statistics.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

if (Test-ExchangeSourceSkipped -Source 'MailboxStatistics' -LogSource $source -CsvPath $csvPath -Column $columns `
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
        $mailboxes = @(Get-EXOMailbox -ResultSize Unlimited -PropertySets Minimum -ErrorAction Stop)
    }
    catch {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            'Get-EXOMailbox is unavailable to this sign-in ({0}). Writing the header only.' -f $_.Exception.Message)
        Export-AppendCsv -Path $csvPath -Column $columns
        return
    }
    if ($MailboxLimit -gt 0) { $mailboxes = @($mailboxes | Select-Object -First $MailboxLimit) }

    $rows = [System.Collections.Generic.List[object]]::new()
    $failed = 0
    foreach ($mailbox in $mailboxes) {
        $upn = [string]$mailbox.UserPrincipalName
        if ([string]::IsNullOrWhiteSpace($upn)) {
            Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
                'Skipping a mailbox with no UserPrincipalName: a $null -Identity would return every mailbox.')
            continue
        }
        try {
            $statistics = @(Get-EXOMailboxStatistics -Identity $upn -PropertySets All -ErrorAction Stop)
        }
        catch {
            $failed++
            Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
                'Statistics for {0} could not be read ({1}).' -f $upn, $_.Exception.Message)
            continue
        }
        foreach ($item in $statistics) {
            $rows.Add([pscustomobject]@{
                    RunDate                          = $runDate
                    ExternalDirectoryObjectId        = [string]$mailbox.ExternalDirectoryObjectId
                    UserPrincipalName                = $upn
                    ItemCount                        = Get-ObjectText $item 'ItemCount'
                    TotalItemSize                    = Get-ObjectText $item 'TotalItemSize'
                    TotalItemSizeBytes               = ConvertTo-ByteCount (Get-ObjectValue $item 'TotalItemSize')
                    DeletedItemCount                 = Get-ObjectText $item 'DeletedItemCount'
                    TotalDeletedItemSize             = Get-ObjectText $item 'TotalDeletedItemSize'
                    TotalDeletedItemSizeBytes        = ConvertTo-ByteCount (Get-ObjectValue $item 'TotalDeletedItemSize')
                    StorageLimitStatus               = Get-ObjectText $item 'StorageLimitStatus'
                    LastLogonTime                    = ConvertTo-CsvTimestamp (Get-ObjectValue $item 'LastLogonTime')
                    LastLogoffTime                   = ConvertTo-CsvTimestamp (Get-ObjectValue $item 'LastLogoffTime')
                    LastLoggedOnUserAccount          = Get-ObjectText $item 'LastLoggedOnUserAccount'
                    IsArchiveMailbox                 = Get-CsvBoolean (Get-ObjectValue $item 'IsArchiveMailbox')
                    DatabaseIssueWarningQuota        = Get-ObjectText $item 'DatabaseIssueWarningQuota'
                    DatabaseProhibitSendQuota        = Get-ObjectText $item 'DatabaseProhibitSendQuota'
                    DatabaseProhibitSendReceiveQuota = Get-ObjectText $item 'DatabaseProhibitSendReceiveQuota'
                })
        }
    }

    $result = Export-AppendCsv -Path $csvPath -Rows $rows.ToArray() -Column $columns -KeyColumn @('RunDate', 'ExternalDirectoryObjectId', 'IsArchiveMailbox') -PassThru
    Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
        'mailbox-statistics.csv: {0} rows written, {1} skipped, {2} mailboxes failed.' -f $result.Written, $result.Skipped, $failed)
}
finally {
    if ($connectedHere) {
        Disconnect-ExchangeOnline -Confirm:$false
    }
}
