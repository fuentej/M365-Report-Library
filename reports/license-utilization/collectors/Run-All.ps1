#Requires -Version 7.0

<#
    .SYNOPSIS
        Runs the Entra user collector and the license utilization collectors into one
        output folder.

    .DESCRIPTION
        Each collector signs in for itself, so a source that is unavailable in this cloud
        or unlicensed in this tenant leaves a header-only CSV and a line in run.log
        without stopping the others. In GCC High the seven usage reports are not
        available through Graph and are written as header-only files; see the README.

        report-settings.csv runs before the usage reports, so they can log that names are
        concealed. The per-user licence detail collector (source 3) makes one call per
        user and is optional: pass -IncludeLicenseDetails, and sign in as a person, since
        the API does not support application permissions. Source 7 (the product-name
        reference) and source 8 (the admin center fallback) are not collectors.

    .EXAMPLE
        ./Run-All.ps1 -OutputPath ./out

    .EXAMPLE
        ./Run-All.ps1 -OutputPath ./out -Environment GCCHigh -AppId $appId `
            -CertificateThumbprint $thumbprint -TenantId $tenantId
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

    [switch]$IncludeLicenseDetails
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
    'Starting the license utilization collectors against the {0} cloud.' -f $Environment)

$steps = @(
    @{ Name = 'users'; Script = $null }
    @{ Name = 'report-settings'; Script = 'Get-ReportSettings.ps1' }
    @{ Name = 'subscribed-skus'; Script = 'Get-SubscribedSkus.ps1' }
    @{ Name = 'user-licenses'; Script = 'Get-UserLicenses.ps1' }
)
if ($IncludeLicenseDetails) {
    $steps += @{ Name = 'license-details'; Script = 'Get-LicenseDetails.ps1' }
}
$steps += @(
    @{ Name = 'user-signin-activity'; Script = 'Get-UserSignInActivity.ps1' }
    @{ Name = 'usage-active-users'; Script = 'Get-ActiveUserUsage.ps1' }
    @{ Name = 'usage-email-activity'; Script = 'Get-EmailActivityUsage.ps1' }
    @{ Name = 'usage-teams-activity'; Script = 'Get-TeamsActivityUsage.ps1' }
    @{ Name = 'usage-sharepoint-activity'; Script = 'Get-SharePointActivityUsage.ps1' }
    @{ Name = 'usage-onedrive-activity'; Script = 'Get-OneDriveActivityUsage.ps1' }
    @{ Name = 'usage-m365-apps'; Script = 'Get-M365AppUsage.ps1' }
    @{ Name = 'usage-copilot'; Script = 'Get-CopilotUsage.ps1' }
)

$failed = 0
foreach ($step in $steps) {
    try {
        if ($null -eq $step.Script) {
            Invoke-EntraUserCollector -OutputPath $OutputPath @auth
        }
        else {
            $arguments = @{ OutputPath = $OutputPath } + $auth
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
