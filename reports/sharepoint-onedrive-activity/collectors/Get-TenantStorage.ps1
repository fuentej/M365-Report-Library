#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes tenant-storage.csv: the tenant's SharePoint storage quota and resource quota, appended each run.

    .DESCRIPTION
        Source 9, route (b): SharePoint Online PowerShell Get-SPOTenant
        (https://learn.microsoft.com/powershell/module/microsoft.online.sharepoint.powershell/get-spotenant),
        which the cmdlet page says "returns organization-level site collection properties such as
        StorageQuota, StorageQuotaAllocated, ResourceQuota, ResourceQuotaAllocated and
        SiteCreationMode" and, in its example, OneDriveStorageQuota. A property the output lacks is
        left empty. Route (a), the SharePoint Storage report in the Microsoft 365 admin center, is a
        manual export and is not a collector. No Graph endpoint for the tenant quota was found.

        Whether the cmdlet works in any cloud is UNVERIFIED, so every cloud is attempted with a
        warning. Role: SharePoint Online administrator. Needs Connect-SPOService.

    .PARAMETER AdminUrl
        The SharePoint admin center URL. Required unless -SkipConnect is used.

    .EXAMPLE
        ./Get-TenantStorage.ps1 -OutputPath ./out -AdminUrl https://contoso-admin.sharepoint.com
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
$columns = $schema.TenantStorage
$source = 'tenant-storage'
$csvPath = Join-Path $OutputPath 'tenant-storage.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

if (Test-SharePointSourceSkipped -Source 'TenantStorage' -LogSource $source -CsvPath $csvPath -Column $columns `
        -OutputPath $OutputPath -Environment $Environment -Schema $schema) {
    return
}

$connectedHere = -not $SkipConnect
if ($connectedHere) {
    if ([string]::IsNullOrWhiteSpace($AdminUrl)) {
        throw '-AdminUrl is required to sign in to SharePoint Online. Pass the SharePoint admin center URL, or -SkipConnect to use an existing session.'
    }
    Connect-SharePointAdmin -AdminUrl $AdminUrl -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -OutputPath $OutputPath -Source $source
}

try {
    try {
        $tenant = Get-SPOTenant -ErrorAction Stop
    }
    catch {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            'Get-SPOTenant is unavailable to this sign-in or cloud ({0}). It needs the SharePoint Online administrator role. Writing the header only.' -f $_.Exception.Message)
        Export-AppendCsv -Path $csvPath -Column $columns
        return
    }

    $row = [ordered]@{ RunDate = $runDate }
    foreach ($name in $columns | Where-Object { $_ -ne 'RunDate' }) {
        $row[$name] = Get-ObjectText -Object $tenant -Name $name
    }
    $missing = @($columns | Where-Object { $_ -ne 'RunDate' -and [string]::IsNullOrEmpty($row[$_]) })
    if ($missing.Count -gt 0) {
        Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
            'Get-SPOTenant did not return: {0}. Those columns are empty.' -f ($missing -join ', '))
    }

    $result = Export-AppendCsv -Path $csvPath -Rows @([pscustomobject]$row) -Column $columns -KeyColumn @('RunDate') -PassThru
    Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
        'tenant-storage.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
}
finally {
    if ($connectedHere) {
        Disconnect-SPOService
    }
}
