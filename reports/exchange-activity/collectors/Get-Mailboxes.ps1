#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes mailboxes.csv: every mailbox with its recipient type and the quotas as configured, appended each run.

    .DESCRIPTION
        Source 6: Get-EXOMailbox -ResultSize Unlimited -PropertySets Minimum, Quota
        (https://learn.microsoft.com/powershell/module/exchangepowershell/get-exomailbox).
        ResultSize defaults to 1000, so Unlimited is required or every mailbox after the first
        1000 is dropped. Minimum carries UserPrincipalName, RecipientType, RecipientTypeDetails and
        ExternalDirectoryObjectId; Quota carries IssueWarningQuota, ProhibitSendQuota,
        ProhibitSendReceiveQuota, RecoverableItemsQuota, ArchiveQuota and UseDatabaseQuotaDefaults
        (https://learn.microsoft.com/powershell/exchange/cmdlet-property-sets). Inactive and
        soft-deleted mailboxes are omitted, as the cmdlet omits them without
        -IncludeInactiveMailbox and -SoftDeletedMailbox.

        The Graph usage report has no recipient-type column; join this file on UserPrincipalName
        to tell shared mailboxes from user mailboxes.

        The cmdlet pages defer to the permissions page for the role. App-only sign-in needs
        Exchange.ManageAsApp plus an Exchange role. Organization Management and Recipient
        Management are listed for managing recipients
        (https://learn.microsoft.com/exchange/permissions-exo/feature-permissions); the least
        privileged role that can run this cmdlet is UNVERIFIED until Get-ManagementRole is run
        in the tenant. Availability in GCC and GCC High is UNVERIFIED at cmdlet level: the
        service description says only that remote PowerShell is available
        (https://learn.microsoft.com/office365/servicedescriptions/office-365-platform-service-description/office-365-us-government/exchange-online-for-us-government-environments).

    .EXAMPLE
        ./Get-Mailboxes.ps1 -OutputPath ./out
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
$columns = $schema.Mailboxes
$source = 'mailboxes'
$csvPath = Join-Path $OutputPath 'mailboxes.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

if (Test-ExchangeSourceSkipped -Source 'Mailboxes' -LogSource $source -CsvPath $csvPath -Column $columns `
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
        $mailboxes = @(Get-EXOMailbox -ResultSize Unlimited -PropertySets Minimum, Quota -ErrorAction Stop)
    }
    catch {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            'Get-EXOMailbox is unavailable to this sign-in ({0}). It needs a role that can view recipients (UNVERIFIED which one is least). Writing the header only.' -f $_.Exception.Message)
        Export-AppendCsv -Path $csvPath -Column $columns
        return
    }
    if ($MailboxLimit -gt 0) { $mailboxes = @($mailboxes | Select-Object -First $MailboxLimit) }

    $rows = foreach ($mailbox in $mailboxes) {
        [pscustomobject]@{
            RunDate                   = $runDate
            ExternalDirectoryObjectId = Get-ObjectText $mailbox 'ExternalDirectoryObjectId'
            UserPrincipalName         = Get-ObjectText $mailbox 'UserPrincipalName'
            PrimarySmtpAddress        = Get-ObjectText $mailbox 'PrimarySmtpAddress'
            RecipientType             = Get-ObjectText $mailbox 'RecipientType'
            RecipientTypeDetails      = Get-ObjectText $mailbox 'RecipientTypeDetails'
            IssueWarningQuota         = Get-ObjectText $mailbox 'IssueWarningQuota'
            ProhibitSendQuota         = Get-ObjectText $mailbox 'ProhibitSendQuota'
            ProhibitSendReceiveQuota  = Get-ObjectText $mailbox 'ProhibitSendReceiveQuota'
            RecoverableItemsQuota     = Get-ObjectText $mailbox 'RecoverableItemsQuota'
            ArchiveQuota              = Get-ObjectText $mailbox 'ArchiveQuota'
            UseDatabaseQuotaDefaults  = Get-CsvBoolean (Get-ObjectValue $mailbox 'UseDatabaseQuotaDefaults')
        }
    }

    $result = Export-AppendCsv -Path $csvPath -Rows @($rows) -Column $columns -KeyColumn @('RunDate', 'ExternalDirectoryObjectId') -PassThru
    Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
        'mailboxes.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
}
finally {
    if ($connectedHere) {
        Disconnect-ExchangeOnline -Confirm:$false
    }
}
