#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes sites.csv: every SharePoint and OneDrive site in the tenant, appended each run.

    .DESCRIPTION
        Source: Microsoft Graph GET /sites/getAllSites
        (https://learn.microsoft.com/graph/api/site-getallsites), read through
        Invoke-MgGraphRequest and paged by following @odata.nextLink until it is absent
        (https://learn.microsoft.com/graph/paging). The returned URL is requested as it is:
        the sample nextLink changes the path to oneDrive.getAllSites, so rebuilding
        /sites/getAllSites from it would ask for the wrong page. For a large tenant, read
        https://learn.microsoft.com/onedrive/developer/rest-api/concepts/scan-guidance first.

        Least privileged permission: Application Sites.Read.All. Delegated is not supported,
        so an interactive sign-in cannot read this source and the collector warns when it is
        run that way; sign in app-only with -AppId, -CertificateThumbprint and -TenantId.
        Available in GCC and GCC High on the endpoint the cloud's Graph sign-in uses.

        This is the list of sites the Item sharing permissions collector walks.

    .EXAMPLE
        ./Get-Sites.ps1 -OutputPath ./out -AppId $appId -CertificateThumbprint $thumbprint -TenantId $tenantId
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

. (Join-Path $PSScriptRoot 'OversharingHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'OversharingSchema.psd1')
$columns = $schema.Sites
$source = 'sites'
$csvPath = Join-Path $OutputPath 'sites.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-OversharingSourceAvailability -Source 'Sites' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping sites.csv. $($availability.Reason) $($availability.Reference)")
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}
if ($availability.Status -eq 'Unverified') {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message $availability.Reason
}

if (-not $SkipConnect) {
    if ([string]::IsNullOrWhiteSpace($AppId)) {
        Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
            'GET /sites/getAllSites supports application permissions only (Sites.Read.All); an interactive sign-in is refused. Sign in app-only with -AppId, -CertificateThumbprint and -TenantId.')
    }
    Connect-M365Service -Service Graph -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization `
        -Scopes 'Sites.Read.All'
}

try {
    $rows = @(Get-AllSiteRow -RunDate $runDate)

    $result = Export-AppendCsv -Path $csvPath -Rows $rows -Column $columns -KeyColumn @('RunDate', 'SiteId') -PassThru
    Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
        'sites.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
}
catch {
    Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
        'Microsoft Graph list sites (getAllSites) is unavailable to this sign-in ({0}). It needs the Sites.Read.All application permission. Writing the header only.' -f $_.Exception.Message)
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}
