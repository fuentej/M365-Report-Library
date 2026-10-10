#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes spo-sites.csv: every site collection, including OneDrive sites, with the storage the cmdlet returns, appended each run.

    .DESCRIPTION
        Source 10: SharePoint Online PowerShell Get-SPOSite -Limit ALL -IncludePersonalSite $true
        (https://learn.microsoft.com/powershell/module/microsoft.online.sharepoint.powershell/get-sposite).
        -Limit defaults to 200, so a call that omits ALL drops the sites past that, and
        -IncludePersonalSite defaults to $false, so OneDrive sites are left out unless it is $true.
        -Detailed is deprecated from module version 5361 and is not used; it would add
        StorageUsageCurrent, ResourceUsageCurrent and WebsCount. This collector reads those three
        when the cmdlet returns them without -Detailed and leaves them empty when it does not (the
        contract asks for that to be confirmed). The cmdlet lists no file count, active file count
        or last-activity date. Sites in the recycle bin are not returned.

        Whether the cmdlet works in any cloud is UNVERIFIED, so every cloud is attempted with a
        warning. Role: SharePoint Online administrator and site collection administrator. Needs
        Connect-SPOService (-Region ITAR in GCC High).

    .PARAMETER AdminUrl
        The SharePoint admin center URL. Required unless -SkipConnect is used.

    .PARAMETER SiteLimit
        Write only the first N sites. For a trial run.

    .EXAMPLE
        ./Get-SpoSites.ps1 -OutputPath ./out -AdminUrl https://contoso-admin.sharepoint.com
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
$columns = $schema.SpoSites
$source = 'spo-sites'
$csvPath = Join-Path $OutputPath 'spo-sites.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

if (Test-SharePointSourceSkipped -Source 'SpoSites' -LogSource $source -CsvPath $csvPath -Column $columns `
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
        $sites = @(Get-SPOSite -Limit ALL -IncludePersonalSite $true -ErrorAction Stop)
    }
    catch {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            'Get-SPOSite is unavailable to this sign-in or cloud ({0}). It needs the SharePoint Online administrator role and site collection administrator access. Writing the header only.' -f $_.Exception.Message)
        Export-AppendCsv -Path $csvPath -Column $columns
        return
    }

    if ($SiteLimit -gt 0) { $sites = @($sites | Select-Object -First $SiteLimit) }

    $rows = foreach ($site in $sites) {
        [pscustomobject]@{
            RunDate              = $runDate
            Url                  = Get-ObjectText -Object $site -Name 'Url'
            Title                = Get-ObjectText -Object $site -Name 'Title'
            Template             = Get-ObjectText -Object $site -Name 'Template'
            StorageUsageCurrent  = Get-ObjectText -Object $site -Name 'StorageUsageCurrent'
            ResourceUsageCurrent = Get-ObjectText -Object $site -Name 'ResourceUsageCurrent'
            WebsCount            = Get-ObjectText -Object $site -Name 'WebsCount'
        }
    }
    $rows = @($rows)

    if ($rows.Count -gt 0 -and @($rows | Where-Object { $_.StorageUsageCurrent -ne '' }).Count -eq 0) {
        Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
            'Get-SPOSite returned no StorageUsageCurrent without -Detailed. The storage columns are empty; use drive-quota.csv for storage.')
    }

    $result = Export-AppendCsv -Path $csvPath -Rows $rows -Column $columns -KeyColumn @('RunDate', 'Url') -PassThru
    Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
        'spo-sites.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
}
finally {
    if ($connectedHere) {
        Disconnect-SPOService
    }
}
