#Requires -Version 7.0

<#
    .SYNOPSIS
        Runs the twelve SharePoint and OneDrive activity collectors into one output folder.

    .DESCRIPTION
        Signs in to Microsoft Graph once, to Exchange Online once and, when -AdminUrl is given, to
        SharePoint Online once, then runs every collector against those sessions (-SkipConnect), so
        an interactive run prompts three times, not twelve. report-settings.csv runs first: it says
        whether the usage reports conceal names.

        A collector whose source is refused, or documented as unavailable in this cloud, leaves a
        header-only CSV and a line in run.log without stopping the others. In GCC High the six Graph
        usage reports are skipped that way; drive-quota.csv, file-events.csv and (UNVERIFIED)
        site-activity.csv are what that cloud gets by script.

        Reports.Read.All, ReportSettings.Read.All, Sites.Read.All and Files.Read.All are requested for
        an interactive Graph sign-in. getAllSites supports application permissions only, so
        drive-quota.csv and site-activity.csv return data only on an app-only sign-in (-AppId and
        -CertificateThumbprint).

    .PARAMETER AdminUrl
        The SharePoint admin center URL, for tenant-storage.csv and spo-sites.csv. Without it those
        two write headers only.

    .PARAMETER Organization
        The tenant's *.onmicrosoft.com domain. Required for app-only sign-in to Exchange Online.

    .PARAMETER SiteLimit
        Read at most this many sites in the per-site collectors. For a trial run.

    .EXAMPLE
        ./Run-All.ps1 -OutputPath ./out -AdminUrl https://contoso-admin.sharepoint.com

    .EXAMPLE
        ./Run-All.ps1 -OutputPath ./out -Environment GCCHigh -AdminUrl https://contoso-admin.sharepoint.us `
            -AppId $appId -CertificateThumbprint $thumbprint -TenantId $tenantId -Organization contoso.onmicrosoft.com
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
    [string]$AdminUrl,

    # Sources 1, 2, 5 and 6. Sources 3 and 4 stay on D30 too: the trend is one row per day in the period.
    [ValidateSet('D7', 'D30', 'D90', 'D180')]
    [string]$Period = 'D30',

    [ValidateRange(1, 180)]
    [int]$LookbackDays = 30,

    [int]$SiteLimit = 0
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Imported by path without -Force: a forced re-import builds a fresh module session
# state, which would discard anything the caller has arranged around the module - and
# that is exactly how the tests put mocks in front of the tenant cmdlets.
Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')
. (Join-Path $PSScriptRoot 'SharePointOneDriveHelpers.ps1')

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
    'Starting the SharePoint and OneDrive activity collectors against the {0} cloud.' -f $Environment)

$exchangeConnected = $false
$sharePointConnected = $false

try {
    try {
        Connect-M365Service -Service ExchangeOnline @auth
        $exchangeConnected = $true
    }
    catch {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source 'run-all' -Message (
            'Exchange Online sign-in failed ({0}). file-events.csv will be a header only.' -f $_.Exception.Message)
    }
    try {
        Connect-M365Service -Service Graph @auth -Scopes (@(Get-DefaultGraphScope) + 'Reports.Read.All', 'ReportSettings.Read.All', 'Sites.Read.All', 'Files.Read.All')
    }
    catch {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source 'run-all' -Message (
            'Microsoft Graph sign-in failed ({0}). The Graph-based collectors will write headers only.' -f $_.Exception.Message)
    }
    if ([string]::IsNullOrWhiteSpace($AdminUrl)) {
        Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source 'run-all' -Message (
            'No -AdminUrl was given. tenant-storage.csv and spo-sites.csv will write headers only.')
    }
    else {
        try {
            Connect-SharePointAdmin -AdminUrl $AdminUrl -Environment $Environment -AppId $AppId `
                -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -OutputPath $OutputPath -Source 'run-all'
            $sharePointConnected = $true
        }
        catch {
            Write-CollectorLog -OutputPath $OutputPath -Level Error -Source 'run-all' -Message (
                'SharePoint Online sign-in failed ({0}). tenant-storage.csv and spo-sites.csv will write headers only.' -f $_.Exception.Message)
        }
    }

    $common = @{ OutputPath = $OutputPath; Environment = $Environment; SkipConnect = $true }
    $usage = $common + @{ Period = $Period }
    $perSite = $common + @{ SiteLimit = $SiteLimit }

    $steps = @(
        @{ Name = 'report-settings'; Script = 'Get-ReportSettings.ps1'; Arguments = $common }
        @{ Name = 'sharepoint-site-usage-detail'; Script = 'Get-SharePointSiteUsageDetail.ps1'; Arguments = $usage }
        @{ Name = 'onedrive-usage-account-detail'; Script = 'Get-OneDriveUsageAccountDetail.ps1'; Arguments = $usage }
        @{ Name = 'sharepoint-site-usage-storage'; Script = 'Get-SharePointSiteUsageStorage.ps1'; Arguments = $usage }
        @{ Name = 'onedrive-usage-storage'; Script = 'Get-OneDriveUsageStorage.ps1'; Arguments = $usage }
        @{ Name = 'sharepoint-activity-user-detail'; Script = 'Get-SharePointActivityUserDetail.ps1'; Arguments = $usage }
        @{ Name = 'onedrive-activity-user-detail'; Script = 'Get-OneDriveActivityUserDetail.ps1'; Arguments = $usage }
        @{ Name = 'tenant-storage'; Script = 'Get-TenantStorage.ps1'; Arguments = $common; NeedsSharePoint = $true }
        @{ Name = 'spo-sites'; Script = 'Get-SpoSites.ps1'; Arguments = $perSite; NeedsSharePoint = $true }
        @{ Name = 'drive-quota'; Script = 'Get-DriveQuota.ps1'; Arguments = $perSite }
        @{ Name = 'file-events'; Script = 'Get-FileEvents.ps1'; Arguments = ($common + @{ LookbackDays = $LookbackDays }) }
        @{ Name = 'site-activity'; Script = 'Get-SiteActivity.ps1'; Arguments = ($perSite + @{ LookbackDays = [math]::Min($LookbackDays, 89) }) }
    )

    foreach ($step in $steps) {
        try {
            if ($step.ContainsKey('NeedsSharePoint') -and -not $sharePointConnected) {
                # Without a session the SharePoint Online cmdlets cannot run. Leave the header
                # and say why, as a collector that is refused does.
                $schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'SharePointOneDriveSchema.psd1')
                $key = if ($step.Name -eq 'tenant-storage') { 'TenantStorage' } else { 'SpoSites' }
                Export-AppendCsv -Path (Join-Path $OutputPath ($step.Name + '.csv')) -Column $schema[$key]
                Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $step.Name -Message (
                    'No SharePoint Online session. Wrote the header only.')
                continue
            }
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
    if ($sharePointConnected) { Disconnect-SPOService }
    if ($exchangeConnected) { Disconnect-ExchangeOnline -Confirm:$false }
}

Write-CollectorLog -OutputPath $OutputPath -Source 'run-all' -Message 'Finished.'
