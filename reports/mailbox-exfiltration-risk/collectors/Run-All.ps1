#Requires -Version 7.0

<#
    .SYNOPSIS
        Runs the Entra user collector and all twelve mailbox exfiltration risk collectors
        into one output folder.

    .DESCRIPTION
        Signs in to Exchange Online once and to Microsoft Graph once, then runs every
        collector against those sessions (-SkipConnect), so an interactive run prompts
        twice, not thirteen times. users.csv comes from the shared Entra users collector.

        A collector whose source is refused or unavailable in this cloud leaves a
        header-only CSV and a line in run.log without stopping the others.

    .PARAMETER Organization
        The tenant's *.onmicrosoft.com domain. Required for app-only sign-in to Exchange
        Online.

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

    [datetime]$StartDate,
    [datetime]$EndDate,

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

$range = @{}
if ($PSBoundParameters.ContainsKey('StartDate')) { $range['StartDate'] = $StartDate }
if ($PSBoundParameters.ContainsKey('EndDate')) { $range['EndDate'] = $EndDate }

Write-CollectorLog -OutputPath $OutputPath -Source 'run-all' -Message (
    'Starting the mailbox exfiltration risk collectors against the {0} cloud.' -f $Environment)

$exchangeConnected = $false
$graphConnected = $false

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
        Connect-M365Service -Service Graph @auth -Scopes 'User.Read.All', 'Directory.Read.All', 'Application.Read.All'
        $graphConnected = $true
    }
    catch {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source 'run-all' -Message (
            'Microsoft Graph sign-in failed ({0}). The Graph-based collectors will write headers only.' -f $_.Exception.Message)
    }

    $common = @{ OutputPath = $OutputPath; Environment = $Environment; SkipConnect = $true }
    $mailboxScoped = @{ MailboxLimit = $MailboxLimit }

    $steps = @(
        @{ Name = 'users'; Script = $null; Arguments = $null }
        @{ Name = 'accepted-domains'; Script = 'Get-AcceptedDomains.ps1'; Arguments = $common }
        @{ Name = 'mailbox-forwarding'; Script = 'Get-MailboxForwarding.ps1'; Arguments = $common }
        @{ Name = 'send-on-behalf'; Script = 'Get-SendOnBehalf.ps1'; Arguments = $common }
        @{ Name = 'inbox-rules'; Script = 'Get-InboxRules.ps1'; Arguments = $common + $mailboxScoped }
        @{ Name = 'transport-rules'; Script = 'Get-TransportRules.ps1'; Arguments = $common }
        @{ Name = 'mailbox-full-access'; Script = 'Get-MailboxFullAccess.ps1'; Arguments = $common + $mailboxScoped }
        @{ Name = 'send-as-permissions'; Script = 'Get-SendAsPermissions.ps1'; Arguments = $common }
        @{ Name = 'delegated-consents'; Script = 'Get-DelegatedConsents.ps1'; Arguments = $common }
        @{ Name = 'app-role-assignments'; Script = 'Get-AppRoleAssignments.ps1'; Arguments = $common }
        @{ Name = 'audit-configuration'; Script = 'Get-AuditConfiguration.ps1'; Arguments = $common }
        @{ Name = 'mailbox-change-events'; Script = 'Get-MailboxChangeEvents.ps1'; Arguments = $common + $range }
        @{ Name = 'mail-access-events'; Script = 'Get-MailAccessEvents.ps1'; Arguments = $common + $range }
    )

    $failed = 0
    foreach ($step in $steps) {
        try {
            if ($null -eq $step.Script) {
                Invoke-EntraUserCollector -OutputPath $OutputPath -Environment $Environment -SkipConnect:$graphConnected `
                    -AppId $AppId -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization
            }
            else {
                $arguments = $step.Arguments
                & (Join-Path $PSScriptRoot $step.Script) @arguments
            }
        }
        catch {
            # One collector failing outright must not stop the rest of the run.
            $failed++
            Write-CollectorLog -OutputPath $OutputPath -Level Error -Source 'run-all' -Message (
                'The {0} collector stopped with an error: {1}' -f $step.Name, $_.Exception.Message)
        }
    }

    Write-CollectorLog -OutputPath $OutputPath -Source 'run-all' -Message 'Finished.'

    if ($failed -gt 0) {
        throw ('{0} collector(s) stopped with an error. See run.log.' -f $failed)
    }
}
finally {
    if ($exchangeConnected) { Disconnect-ExchangeOnline -Confirm:$false }
}
