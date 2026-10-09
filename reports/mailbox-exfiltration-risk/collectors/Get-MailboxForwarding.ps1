#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes mailbox-forwarding.csv: every mailbox that forwards mail, with the
        destination and whether it is external, appended each run.

    .DESCRIPTION
        Source: Get-EXOMailbox -ResultSize Unlimited -PropertySets Minimum, Delivery
        (https://learn.microsoft.com/powershell/module/exchangepowershell/get-exomailbox).
        The default property set (Minimum) does not include ForwardingAddress,
        ForwardingSmtpAddress or DeliverToMailboxAndForward; they are in the Delivery set
        (https://learn.microsoft.com/powershell/exchange/cmdlet-property-sets). Minimum is
        requested too, for the identity columns. ResultSize defaults to 1000, so
        Unlimited is required.

        ForwardingAddress is an internal recipient, not an SMTP forward. External means
        the domain of ForwardingSmtpAddress is not an accepted domain
        (Get-AcceptedDomain). If the accepted domains cannot be read, IsExternal is
        empty rather than guessed.

        The cmdlet page defers to the permissions page for the role; the matching role,
        View-Only Recipients, is an inference and is UNVERIFIED.

    .EXAMPLE
        ./Get-MailboxForwarding.ps1 -OutputPath ./out
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
$columns = $schema.MailboxForwarding
$source = 'mailbox-forwarding'
$csvPath = Join-Path $OutputPath 'mailbox-forwarding.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-MailboxSourceAvailability -Source 'MailboxForwarding' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping mailbox-forwarding.csv. $($availability.Reason) $($availability.Reference)")
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
        $accepted = Get-AcceptedDomainName -OutputPath $OutputPath -Source $source
    
        $rows = foreach ($mailbox in $mailboxes) {
            $forwardingAddress = [string]$mailbox.ForwardingAddress
            $forwardingSmtp = [string]$mailbox.ForwardingSmtpAddress
            if ([string]::IsNullOrWhiteSpace($forwardingAddress) -and [string]::IsNullOrWhiteSpace($forwardingSmtp)) { continue }
    
            $smtpDomain = ''
            $isExternal = ''
            $address = @(Get-SmtpAddress $forwardingSmtp) | Select-Object -First 1
            if ($address) {
                $smtpDomain = Get-AddressDomain $address
                $external = Test-ExternalDomain -Domain $smtpDomain -AcceptedDomain $accepted
                if ($null -ne $external) { $isExternal = $external.ToString() }
            }
            elseif (-not [string]::IsNullOrWhiteSpace($forwardingAddress)) {
                $isExternal = 'False'
            }
    
            [pscustomobject]@{
                RunDate                    = $runDate
                ExternalDirectoryObjectId  = [string]$mailbox.ExternalDirectoryObjectId
                UserPrincipalName          = [string]$mailbox.UserPrincipalName
                PrimarySmtpAddress         = [string]$mailbox.PrimarySmtpAddress
                ForwardingAddress          = $forwardingAddress
                ForwardingSmtpAddress      = $forwardingSmtp
                ForwardingSmtpDomain       = $smtpDomain
                DeliverToMailboxAndForward = Get-CsvBoolean $mailbox.DeliverToMailboxAndForward
                IsExternal                 = $isExternal
            }
        }
    
        $result = Export-AppendCsv -Path $csvPath -Rows @($rows) -Column $columns -KeyColumn @('RunDate', 'ExternalDirectoryObjectId') -PassThru
        Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
            'mailbox-forwarding.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
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
