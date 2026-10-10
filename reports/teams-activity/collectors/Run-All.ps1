#Requires -Version 7.0

<#
    .SYNOPSIS
        Runs the Teams activity collectors into one output folder.

    .DESCRIPTION
        Signs in to Microsoft Graph once and to Exchange Online once, then runs every collector against
        those sessions (-SkipConnect), so an interactive run prompts twice, not once per collector.
        report-settings.csv runs first: it says whether the usage reports conceal names.

        The team activity report (source 4 of the contract) is not rebuilt here. It is the same API
        the Teams and Groups lifecycle report already collects, so this script runs that report's
        Get-TeamActivity.ps1 into the same folder and its team-activity.csv is the source 4 file.

        A collector whose source is refused, or documented as unavailable in this cloud, leaves a
        header-only CSV and a line in run.log without stopping the others. In GCC High the three usage
        report CSVs and team-activity.csv are skipped that way.

        Reports.Read.All and ReportSettings.Read.All are requested for an interactive sign-in.
        The call records need the application permission CallRecords.Read.All, so they only return
        data on an app-only sign-in (-AppId and -CertificateThumbprint).

    .PARAMETER Organization
        The tenant's *.onmicrosoft.com domain. Required for app-only sign-in to Exchange Online.

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

    [ValidateSet('D7', 'D30', 'D90', 'D180')]
    [string]$Period = 'D30',

    [ValidateRange(1, 180)]
    [int]$LookbackDays = 30
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
    'Starting the Teams activity collectors against the {0} cloud.' -f $Environment)

$exchangeConnected = $false

try {
    try {
        Connect-M365Service -Service Graph @auth -Scopes (@(Get-DefaultGraphScope) + 'Reports.Read.All', 'ReportSettings.Read.All', 'CallRecords.Read.All')
    }
    catch {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source 'run-all' -Message (
            'Microsoft Graph sign-in failed ({0}). The Graph-based collectors will write headers only.' -f $_.Exception.Message)
    }
    try {
        Connect-M365Service -Service ExchangeOnline @auth
        $exchangeConnected = $true
    }
    catch {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source 'run-all' -Message (
            'Exchange Online sign-in failed ({0}). The audit collector will write a header only.' -f $_.Exception.Message)
    }

    $common = @{ OutputPath = $OutputPath; Environment = $Environment; SkipConnect = $true }
    $usage = $common + @{ Period = $Period }
    $lifecycle = Join-Path $PSScriptRoot '../../teams-groups-lifecycle/collectors/Get-TeamActivity.ps1'

    $steps = @(
        @{ Name = 'report-settings'; Script = Join-Path $PSScriptRoot 'Get-ReportSettings.ps1'; Arguments = $common }
        @{ Name = 'teams-user-activity-user-detail'; Script = Join-Path $PSScriptRoot 'Get-TeamsUserActivityUserDetail.ps1'; Arguments = $usage }
        @{ Name = 'teams-user-activity-counts'; Script = Join-Path $PSScriptRoot 'Get-TeamsUserActivityCounts.ps1'; Arguments = $usage }
        @{ Name = 'teams-device-usage-user-detail'; Script = Join-Path $PSScriptRoot 'Get-TeamsDeviceUsageUserDetail.ps1'; Arguments = $usage }
        @{ Name = 'team-activity'; Script = $lifecycle; Arguments = $usage }
        @{ Name = 'call-records'; Script = Join-Path $PSScriptRoot 'Get-CallRecords.ps1'; Arguments = $common + @{ LookbackDays = [math]::Min($LookbackDays, 30) } }
        @{ Name = 'teams-audit-events'; Script = Join-Path $PSScriptRoot 'Get-TeamsAuditEvents.ps1'; Arguments = $common + @{ LookbackDays = $LookbackDays } }
    )

    foreach ($step in $steps) {
        try {
            $arguments = $step.Arguments
            & $step.Script @arguments
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
