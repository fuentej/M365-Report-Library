#Requires -Version 7.0

<#
    .SYNOPSIS
        Runs the ten Exchange activity collectors into one output folder.

    .DESCRIPTION
        Signs in to Exchange Online once and to Microsoft Graph once, then runs every
        collector against those sessions (-SkipConnect), so an interactive run prompts twice,
        not ten times. report-settings.csv runs first: it says whether the usage reports
        conceal names.

        A collector whose source is refused, or documented as unavailable in this cloud,
        leaves a header-only CSV and a line in run.log without stopping the others. In GCC High
        the four Graph usage reports are skipped that way.

        Reports.Read.All and ReportSettings.Read.All are requested for an interactive sign-in. The
        Graph message trace needs the application permission ExchangeMessageTrace.Read.All, so it
        only returns data on an app-only sign-in (-AppId and -CertificateThumbprint).

    .PARAMETER Organization
        The tenant's *.onmicrosoft.com domain. Required for app-only sign-in to Exchange Online.

    .PARAMETER MailboxLimit
        Read at most this many mailboxes in the per-mailbox collectors. For a trial run.

    .EXAMPLE
        ./Run-All.ps1 -OutputPath ./out

    .EXAMPLE
        ./Run-All.ps1 -OutputPath ./out -Environment GCCHigh -AppId $appId `
            -CertificateThumbprint $thumbprint -TenantId $tenantId -Organization contoso.onmicrosoft.com
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

    # Sources 1, 3 and 4. Source 2 stays on D180 unless this is passed explicitly.
    [ValidateSet('D7', 'D30', 'D90', 'D180')]
    [string]$Period = 'D30',

    [ValidateRange(1, 90)]
    [int]$LookbackDays = 10,

    [int]$MailboxLimit = 0
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Imported by path without -Force: a forced re-import builds a fresh module session
# state, which would discard anything the caller has arranged around the module - and
# that is exactly how the tests put mocks in front of the tenant cmdlets.
Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}
$OutputPath = (Resolve-Path -LiteralPath $OutputPath).Path

$auth = @{
    Environment           = $Environment
    AppId                 = $AppId
    CertificateThumbprint = $CertificateThumbprint
    TenantId              = $TenantId
    Organization          = $Organization
}

Write-CollectorLog -OutputPath $OutputPath -Source 'run-all' -Message (
    'Starting the Exchange activity collectors against the {0} cloud.' -f $Environment)

$exchangeConnected = $false

try {
    try {
        Connect-M365Service -Service ExchangeOnline @auth
        $exchangeConnected = $true
    }
    catch {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source 'run-all' -Message (
            'Exchange Online sign-in failed ({0}). The Exchange-based collectors will write headers only.' -f $_.Exception.Message)
    }
    try {
        Connect-M365Service -Service Graph @auth -Scopes (@(Get-DefaultGraphScope) + 'Reports.Read.All', 'ReportSettings.Read.All')
    }
    catch {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source 'run-all' -Message (
            'Microsoft Graph sign-in failed ({0}). The Graph-based collectors will write headers only.' -f $_.Exception.Message)
    }

    $common = @{ OutputPath = $OutputPath; Environment = $Environment; SkipConnect = $true }
    $usage = $common + @{ Period = $Period }
    # getMailboxUsageStorage is documented as period D180. A D30 snapshot cannot backfill the
    # earlier days. An explicit -Period applies to storage as well.
    $storagePeriod = if ($PSBoundParameters.ContainsKey('Period')) { $Period } else { 'D180' }
    $mailboxScoped = @{ MailboxLimit = $MailboxLimit }
    $trace = @{ LookbackDays = $LookbackDays }

    $steps = @(
        @{ Name = 'report-settings'; Script = 'Get-ReportSettings.ps1'; Arguments = $common }
        @{ Name = 'mailbox-usage-detail'; Script = 'Get-MailboxUsageDetail.ps1'; Arguments = $usage }
        @{ Name = 'mailbox-usage-storage'; Script = 'Get-MailboxUsageStorage.ps1'; Arguments = ($common + @{ Period = $storagePeriod }) }
        @{ Name = 'email-activity-user-detail'; Script = 'Get-EmailActivityUserDetail.ps1'; Arguments = $usage }
        @{ Name = 'email-app-usage-user-detail'; Script = 'Get-EmailAppUsageUserDetail.ps1'; Arguments = $usage }
        @{ Name = 'mailboxes'; Script = 'Get-Mailboxes.ps1'; Arguments = $common + $mailboxScoped }
        @{ Name = 'mailbox-statistics'; Script = 'Get-MailboxStatistics.ps1'; Arguments = $common + $mailboxScoped }
        @{ Name = 'message-trace'; Script = 'Get-MessageTrace.ps1'; Arguments = $common + $mailboxScoped + $trace }
        @{ Name = 'mobile-devices'; Script = 'Get-MobileDevices.ps1'; Arguments = $common + $mailboxScoped }
        @{ Name = 'graph-message-trace'; Script = 'Get-GraphMessageTrace.ps1'; Arguments = $common + $trace }
    )

    foreach ($step in $steps) {
        try {
            $arguments = $step.Arguments
            & (Join-Path $PSScriptRoot $step.Script) @arguments
        }
        catch {
            Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $step.Name -Message (
                'The {0} collector stopped ({1}). The remaining collectors still run.' -f $step.Name, $_.Exception.Message)
        }
    }
}
finally {
    if ($exchangeConnected) {
        Disconnect-ExchangeOnline -Confirm:$false
    }
}

Write-CollectorLog -OutputPath $OutputPath -Source 'run-all' -Message 'Finished.'
