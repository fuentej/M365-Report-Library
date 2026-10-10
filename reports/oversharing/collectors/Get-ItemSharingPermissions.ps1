#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes item-permissions.csv: the sharing permissions on the files and folders of every
        site's drives, appended each run.

    .DESCRIPTION
        Source: Microsoft Graph GET /drives/{drive-id}/items/{item-id}/permissions
        (https://learn.microsoft.com/graph/api/driveitem-list-permissions). It is one call
        per item, so a tenant scan is a walk of every drive: for each site, its drives
        (GET /sites/{site-id}/drives, https://learn.microsoft.com/graph/api/drive-list), then
        every folder's children (https://learn.microsoft.com/graph/api/driveitem-list-children),
        every collection paged by following @odata.nextLink. Learn's scan guidance is the page
        to read before sizing it:
        https://learn.microsoft.com/onedrive/developer/rest-api/concepts/scan-guidance

        link.scope is anonymous (anyone), organization (everyone signed in to the tenant),
        users (specific people) or existingAccess
        (https://learn.microsoft.com/graph/api/resources/sharinglink#scope-options).
        inheritedFrom separates an inherited permission from a direct one. Learn does not list
        the creator on a permission, so who created a link comes from the audit events.
        link.webUrl and shareId are secrets and are not written.

        What an app-only caller sees is UNVERIFIED: the page says it returns all sharing
        permissions to the item's owner and only those that apply to the caller otherwise,
        and says nothing about an app-only caller. The collector logs a warning on every run,
        so a scan is not read as proof that it saw every link.

        Least privileged permission: Application Files.Read.All; delegated Files.Read.

        The sites come from the latest snapshot in sites.csv in the output folder, or from
        GET /sites/getAllSites when there is none. -SiteId reads only those sites,
        -MaxSites the first N, -MaxItemsPerDrive stops each drive's walk after N items.

    .PARAMETER LinksOnly
        Write only the permissions that are sharing links.

    .EXAMPLE
        ./Get-ItemSharingPermissions.ps1 -OutputPath ./out -MaxSites 5 -MaxItemsPerDrive 500
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

    [string[]]$SiteId,

    [int]$MaxSites = 0,

    [int]$MaxItemsPerDrive = 0,

    [switch]$LinksOnly,

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
$columns = $schema.ItemPermissions
$source = 'item-permissions'
$csvPath = Join-Path $OutputPath 'item-permissions.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-OversharingSourceAvailability -Source 'ItemPermissions' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping item-permissions.csv. $($availability.Reason) $($availability.Reference)")
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}
if ($availability.Status -eq 'Unverified') {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message $availability.Reason
}

if (-not $SkipConnect) {
    Connect-M365Service -Service Graph -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization `
        -Scopes 'Files.Read.All', 'Sites.Read.All'
}

Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
    'What an app-only caller sees from GET .../permissions is UNVERIFIED: Learn says a non-owner caller gets only the permissions that apply to it. Do not read this file as proof that every sharing link was seen.')

try {
    $siteIds = @()
    if ($SiteId) {
        $siteIds = @($SiteId)
    }
    else {
        $latest = @(Get-CsvLatestSnapshot -Path (Join-Path $OutputPath 'sites.csv'))
        if ($latest.Count -gt 0) {
            $siteIds = @($latest | ForEach-Object { [string]$_.SiteId })
        }
        else {
            $siteIds = @(Get-AllSiteRow -RunDate $runDate | ForEach-Object { $_.SiteId })
        }
    }
    if ($MaxSites -gt 0) { $siteIds = @($siteIds | Select-Object -First $MaxSites) }
}
catch {
    Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
        'The list of sites to scan is unavailable to this sign-in ({0}). It needs the Sites.Read.All application permission. Writing the header only.' -f $_.Exception.Message)
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}

$scanned = 0
$failed = 0
$totalWritten = 0

foreach ($site in $siteIds) {
    try {
        $drives = @(Get-GraphPagedValue -Uri ('/v1.0/sites/{0}/drives?$select=id,name,driveType' -f $site))
    }
    catch {
        $failed++
        Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
            'The drives of site {0} could not be read ({1}). Skipping the site.' -f $site, $_.Exception.Message)
        continue
    }

    foreach ($drive in $drives) {
        $driveId = [string](Get-JsonProperty -Object $drive -Name 'id')
        $rows = [System.Collections.Generic.List[object]]::new()
        $itemFailures = 0

        try {
            $items = @(Get-DriveItemWalk -DriveId $driveId -MaxItems $MaxItemsPerDrive)
        }
        catch {
            $failed++
            Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
                'The items of drive {0} could not be listed ({1}). Skipping the drive.' -f $driveId, $_.Exception.Message)
            continue
        }

        if ($MaxItemsPerDrive -gt 0 -and $items.Count -ge $MaxItemsPerDrive) {
            Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
                'Drive {0} was cut off after {1} items (-MaxItemsPerDrive); the rest were not scanned.' -f $driveId, $MaxItemsPerDrive)
        }

        foreach ($item in $items) {
            try {
                $permissions = @(Get-GraphPagedValue -Uri ('/v1.0/drives/{0}/items/{1}/permissions' -f $driveId, $item.Id))
            }
            catch {
                $itemFailures++
                Write-Verbose ('Permissions of {0}/{1} not read: {2}' -f $driveId, $item.Id, $_.Exception.Message)
                continue
            }

            foreach ($permission in $permissions) {
                $row = ConvertTo-ItemPermissionRow -RunDate $runDate -SiteId $site -DriveId $driveId -Item $item -Permission $permission
                if ($LinksOnly -and -not $row.LinkScope) { continue }
                $rows.Add($row)
            }
        }

        if ($itemFailures -gt 0) {
            Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
                '{0} item(s) in drive {1} refused the permissions read and were skipped.' -f $itemFailures, $driveId)
        }

        $result = Export-AppendCsv -Path $csvPath -Rows $rows.ToArray() -Column $columns -KeyColumn @('RunDate', 'DriveId', 'ItemId', 'PermissionId') -PassThru
        $totalWritten += $result.Written
        $scanned++
    }
}

if ($scanned -eq 0) {
    Export-AppendCsv -Path $csvPath -Column $columns
    if ($failed -gt 0) {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            'No drive could be read. It needs the Files.Read.All application permission. Writing the header only.')
    }
}

Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'item-permissions.csv: {0} rows written from {1} drive(s) in {2} site(s).' -f $totalWritten, $scanned, @($siteIds).Count)
