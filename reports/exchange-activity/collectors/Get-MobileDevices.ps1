#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes mobile-devices.csv: the mobile devices syncing to each mailbox, appended each run.

    .DESCRIPTION
        Source 9: Get-EXOMobileDeviceStatistics -Mailbox <upn>
        (https://learn.microsoft.com/powershell/module/exchangepowershell/get-exomobiledevicestatistics).
        No device-family switch is passed: -ActiveSync, -RestApi, -OWAforDevices and
        -UniversalOutlook are filters, and passing only -ActiveSync would drop the other families.
        This is mobile sync only, not Outlook for Windows, Mac or web.

        The cmdlet page lists no output properties. DeviceType, DeviceOS, DeviceAccessState and
        DeviceUserAgent are the properties Learn's troubleshooting page selects from the same
        statistics
        (https://learn.microsoft.com/troubleshoot/exchange/administration/windows-mail-app-not-blocked);
        DeviceId is the identity the Identity parameter describes. A property the cmdlet does not
        return is left empty.

        The cmdlet pages defer to the permissions page for the role. App-only sign-in needs
        Exchange.ManageAsApp plus an Exchange role. Organization Management and Recipient
        Management are listed for managing recipients
        (https://learn.microsoft.com/exchange/permissions-exo/feature-permissions); the least
        privileged role that can run this cmdlet is UNVERIFIED until Get-ManagementRole is run
        in the tenant. Availability in GCC and GCC High is UNVERIFIED at cmdlet level: the
        service description says only that remote PowerShell is available
        (https://learn.microsoft.com/office365/servicedescriptions/office-365-platform-service-description/office-365-us-government/exchange-online-for-us-government-environments).

    .EXAMPLE
        ./Get-MobileDevices.ps1 -OutputPath ./out
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
$columns = $schema.MobileDevices
$source = 'mobile-devices'
$csvPath = Join-Path $OutputPath 'mobile-devices.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

if (Test-ExchangeSourceSkipped -Source 'MobileDevices' -LogSource $source -CsvPath $csvPath -Column $columns `
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
    foreach ($mailbox in $mailboxes) {
        $upn = [string]$mailbox.UserPrincipalName
        if ([string]::IsNullOrWhiteSpace($upn)) { continue }
        try {
            $devices = @(Get-EXOMobileDeviceStatistics -Mailbox $upn -ErrorAction Stop)
        }
        catch {
            Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
                'Mobile devices for {0} could not be read ({1}).' -f $upn, $_.Exception.Message)
            continue
        }
        foreach ($device in $devices) {
            $rows.Add([pscustomobject]@{
                    RunDate                  = $runDate
                    MailboxUserPrincipalName = $upn
                    DeviceId                 = Get-ObjectText $device 'DeviceId'
                    DeviceType               = Get-ObjectText $device 'DeviceType'
                    DeviceOS                 = Get-ObjectText $device 'DeviceOS'
                    DeviceAccessState        = Get-ObjectText $device 'DeviceAccessState'
                    DeviceUserAgent          = Get-ObjectText $device 'DeviceUserAgent'
                })
        }
    }

    $result = Export-AppendCsv -Path $csvPath -Rows $rows.ToArray() -Column $columns -KeyColumn @('RunDate', 'MailboxUserPrincipalName', 'DeviceId', 'DeviceType', 'DeviceUserAgent') -PassThru
    Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
        'mobile-devices.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
}
finally {
    if ($connectedHere) {
        Disconnect-ExchangeOnline -Confirm:$false
    }
}
