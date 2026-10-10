#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes drive-quota.csv: used and total bytes per drive for every site, appended each run.

    .DESCRIPTION
        Source 11: Microsoft Graph GET /sites/getAllSites
        (https://learn.microsoft.com/graph/api/site-getallsites) then GET /sites/{id}/drives
        (https://learn.microsoft.com/graph/api/drive-list). Both pages are followed by the
        @odata.nextLink Graph returns, as returned: the getAllSites sample nextLink changes the
        path to oneDrive.getAllSites, so the URL is never rebuilt. Sites can have more than one
        drive. The quota object (https://learn.microsoft.com/graph/api/resources/quota) holds used,
        total, remaining, deleted and state in bytes and has no file count. The drive's
        lastModifiedDateTime is when the drive was modified, not the usage report's Last Activity
        Date. This is the only storage source for GCC High, where the Graph usage reports are
        unavailable.

        Personal (OneDrive) sites that getAllSites returns are read through GET /sites/{id}/drives
        like any other site. GET /users/{id}/drives is not called.

        Needs the application permissions Sites.Read.All (getAllSites does not support delegated)
        and Files.Read.All (or Sites.Read.All) for the drive list. Available in Commercial, GCC and
        GCC High.

    .PARAMETER SiteLimit
        Read the drives of only the first N sites. For a trial run.

    .EXAMPLE
        ./Get-DriveQuota.ps1 -OutputPath ./out
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

    [int]$SiteLimit = 0,

    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Imported by path without -Force: a forced re-import builds a fresh module session
# state, which would discard anything the caller has arranged around the module - and
# that is exactly how the tests put mocks in front of the tenant cmdlets.
Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

. (Join-Path $PSScriptRoot 'SharePointOneDriveHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'SharePointOneDriveSchema.psd1')
$columns = $schema.DriveQuota
$source = 'drive-quota'
$csvPath = Join-Path $OutputPath 'drive-quota.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

if (Test-SharePointSourceSkipped -Source 'DriveQuota' -LogSource $source -CsvPath $csvPath -Column $columns `
        -OutputPath $OutputPath -Environment $Environment -Schema $schema) {
    return
}

if (-not $SkipConnect) {
    Connect-M365Service -Service Graph -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization `
        -Scopes (@(Get-DefaultGraphScope) + 'Sites.Read.All', 'Files.Read.All')
}

$rows = [System.Collections.Generic.List[object]]::new()
$skipped = 0
try {
    $count = 0
    foreach ($site in Get-SiteRow) {
        if ($SiteLimit -gt 0 -and $count -ge $SiteLimit) { break }
        $count++
        try {
            foreach ($row in Get-SiteDriveRow -Site $site -RunDate $runDate) { $rows.Add($row) }
        }
        catch {
            $skipped++
            Write-Verbose ('Drives of site {0} not read: {1}' -f (Get-GraphJsonValue -Object $site -Name 'id'), $_.Exception.Message)
        }
    }
}
catch {
    Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
        'The site list is unavailable to this sign-in ({0}). getAllSites needs the application permission Sites.Read.All. Writing the header only.' -f $_.Exception.Message)
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}

if ($skipped -gt 0) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        '{0} site(s) could not have their drives read and were skipped. Drive list needs Files.Read.All or Sites.Read.All.' -f $skipped)
}

$result = Export-AppendCsv -Path $csvPath -Rows $rows.ToArray() -Column $columns -KeyColumn @('RunDate', 'SiteId', 'DriveId') -PassThru
Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'drive-quota.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
