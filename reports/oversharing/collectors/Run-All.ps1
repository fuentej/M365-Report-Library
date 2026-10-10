#Requires -Version 7.0

<#
    .SYNOPSIS
        Runs all eleven oversharing collectors into one output folder.

    .DESCRIPTION
        Signs in to Microsoft Graph, Exchange Online and SharePoint Online once each, and to
        Security & Compliance PowerShell once when no -LabelGuid is given (to list the
        sensitivity labels), then runs every collector against those sessions
        (-SkipConnect), so an interactive run prompts three or four times, not eleven.

        The contract doc joins to no user table, so this report writes no users.csv. If a
        page needs one, run Invoke-EntraUserCollector from the shared module into the same
        folder; the Guest and external access report already covers guests.

        A collector whose source is refused or unavailable in this cloud leaves a header-only
        CSV and a line in run.log without stopping the others. The Data access governance
        collectors are asynchronous: a report still running when -WaitMinutes ends is left
        running, and the next run exports it.

    .PARAMETER AdminUrl
        The SharePoint admin center URL, such as https://contoso-admin.sharepoint.com. Without
        it the SharePoint Online collectors write headers only.

    .PARAMETER LabelGuid
        Sensitivity label GUIDs for labeled-file-sites.csv. Without them the labels are listed
        with Get-Label.

    .PARAMETER SkipItemPermissions
        Leave out item-permissions.csv, which makes one Graph call per file and folder.

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

    [guid[]]$LabelGuid,

    [datetime]$StartDate,
    [datetime]$EndDate,

    [int]$MaxSites = 0,

    [int]$MaxItemsPerDrive = 0,

    [switch]$LinksOnly,

    [switch]$SkipItemPermissions,

    [ValidateRange(0, [int]::MaxValue)]
    [int]$WaitMinutes = 30
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Imported by path without -Force: a forced re-import builds a fresh module session
# state, which would discard anything the caller has arranged around the module - and
# that is exactly how the tests put mocks in front of the tenant cmdlets.
Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

. (Join-Path $PSScriptRoot 'OversharingHelpers.ps1')

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
    'Starting the oversharing collectors against the {0} cloud.' -f $Environment)

$graphConnected = $false
$exchangeConnected = $false
$sharePointConnected = $false
$complianceConnected = $false

try {
    try {
        Connect-M365Service -Service Graph @auth -Scopes 'Sites.Read.All', 'Files.Read.All'
        $graphConnected = $true
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
            'Exchange Online sign-in failed ({0}). The audit collectors will write headers only.' -f $_.Exception.Message)
    }
    if ([string]::IsNullOrWhiteSpace($AdminUrl)) {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source 'run-all' -Message (
            'No -AdminUrl was given. The SharePoint Online collectors will write headers only.')
    }
    else {
        try {
            Connect-SharePointAdmin -AdminUrl $AdminUrl -Environment $Environment -AppId $AppId `
                -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -OutputPath $OutputPath -Source 'run-all'
            $sharePointConnected = $true
        }
        catch {
            Write-CollectorLog -OutputPath $OutputPath -Level Error -Source 'run-all' -Message (
                'SharePoint Online sign-in failed ({0}). The SharePoint Online collectors will write headers only.' -f $_.Exception.Message)
        }
    }

    $common = @{ OutputPath = $OutputPath; Environment = $Environment; SkipConnect = $true }
    $sharePoint = $common + @{ WaitMinutes = $WaitMinutes }
    $labels = @{}
    if ($LabelGuid) { $labels['LabelGuid'] = $LabelGuid }

    $steps = @(
        @{ Name = 'sites'; Script = 'Get-Sites.ps1'; Arguments = $common }
    )
    if (-not $SkipItemPermissions) {
        $steps += @{ Name = 'item-permissions'; Script = 'Get-ItemSharingPermissions.ps1'
            Arguments = $common + @{ MaxSites = $MaxSites; MaxItemsPerDrive = $MaxItemsPerDrive; LinksOnly = $LinksOnly } }
    }
    $steps += @(
        @{ Name = 'site-permission-breadth'; Script = 'Get-SitePermissionBreadth.ps1'; Arguments = $sharePoint }
        @{ Name = 'everyone-item-exposure'; Script = 'Get-EveryoneItemExposure.ps1'; Arguments = $sharePoint }
        @{ Name = 'sharing-link-activity'; Script = 'Get-SharingLinkActivity.ps1'; Arguments = $sharePoint }
        @{ Name = 'eeeu-activity'; Script = 'Get-EeeuActivity.ps1'; Arguments = $sharePoint }
        @{ Name = 'labeled-file-sites'; Script = 'Get-LabeledFileSites.ps1'; Arguments = $sharePoint + $labels }
        @{ Name = 'site-sharing-settings'; Script = 'Get-SiteSharingSettings.ps1'; Arguments = $common }
        @{ Name = 'anonymous-link-events'; Script = 'Get-AnonymousLinkEvents.ps1'; Arguments = $common + $range }
        @{ Name = 'sharing-events'; Script = 'Get-SharingEvents.ps1'; Arguments = $common + $range }
        @{ Name = 'audit-log-status'; Script = 'Get-AuditLogStatus.ps1'; Arguments = $common }
    )

    # Listing the labels needs Security & Compliance PowerShell, which is signed in to once
    # for the whole run rather than once inside the collector.
    if (-not $LabelGuid) {
        try {
            Connect-M365Service -Service SecurityCompliance @auth
            $complianceConnected = $true
        }
        catch {
            Write-CollectorLog -OutputPath $OutputPath -Level Error -Source 'run-all' -Message (
                'Security & Compliance sign-in failed ({0}). labeled-file-sites.csv will be a header only unless -LabelGuid is given.' -f $_.Exception.Message)
        }
    }

    $failed = 0
    foreach ($step in $steps) {
        try {
            $arguments = $step.Arguments
            & (Join-Path $PSScriptRoot $step.Script) @arguments
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
    if ($sharePointConnected) { Disconnect-SPOService }
    if ($exchangeConnected -or $complianceConnected) { Disconnect-ExchangeOnline -Confirm:$false }
}
