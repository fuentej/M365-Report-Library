#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes site-sharing-settings.csv: the sharing level, default link type and whether
        company-wide links are disabled, for every site and for the tenant, appended each run.

    .DESCRIPTION
        Source: SharePoint Online PowerShell Get-SPOSite -Limit All -IncludePersonalSite $true
        (https://learn.microsoft.com/powershell/module/microsoft.online.sharepoint.powershell/get-sposite).
        -Limit defaults to 200, so a call that omits All drops the sites past that, and
        -IncludePersonalSite defaults to $false, so OneDrive sites are left out unless it is
        $true. Learn says SharingCapability, DefaultSharingLinkType,
        DisableCompanyWideSharingLinks, SensitivityLabel and others are not populated when
        -Limit or -Filter is used, so each site is then read again with -Identity (the page's
        example 2). -Detailed is deprecated and is not used.

        SharingCapability is ExternalUserAndGuestSharing (the default), Disabled,
        ExternalUserSharingOnly or ExistingExternalUserSharingOnly
        (https://learn.microsoft.com/powershell/module/microsoft.online.sharepoint.powershell/set-spotenant).
        DefaultSharingLinkType is None (the widest scope the other settings allow, not "no
        link"), Direct (Specific people), Internal (organization) or AnonymousAccess (Anyone).
        Set-SPOTenant is never called; this only reads.

        The tenant row comes from Get-SPOTenant. That it returns SharingCapability is not
        stated on the Get-SPOTenant page, so it is UNVERIFIED: the call is attempted, and a
        property the output lacks is left empty.

        Role: SharePoint Online administrator and site collection administrator. No licence
        is named. Whether the cmdlet works in GCC and GCC High is UNVERIFIED, so every cloud is
        attempted with a warning. Needs Connect-SPOService.

    .PARAMETER AdminUrl
        The SharePoint admin center URL. Required unless -SkipConnect is used.

    .PARAMETER SiteLimit
        Read only the first N sites in the per-site loop. For a trial run.

    .EXAMPLE
        ./Get-SiteSharingSettings.ps1 -OutputPath ./out -AdminUrl https://contoso-admin.sharepoint.com
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

. (Join-Path $PSScriptRoot 'OversharingHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'OversharingSchema.psd1')
$columns = $schema.SiteSharingSettings
$source = 'site-sharing-settings'
$csvPath = Join-Path $OutputPath 'site-sharing-settings.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-OversharingSourceAvailability -Source 'SiteSharingSettings' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping site-sharing-settings.csv. $($availability.Reason) $($availability.Reference)")
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}
if ($availability.Status -eq 'Unverified') {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message $availability.Reason
}

$connectedHere = -not $SkipConnect
if ($connectedHere) {
    if ([string]::IsNullOrWhiteSpace($AdminUrl)) {
        throw '-AdminUrl is required to sign in to SharePoint Online. Pass the SharePoint admin center URL, or -SkipConnect to use an existing session.'
    }
    Connect-SharePointAdmin -AdminUrl $AdminUrl -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -OutputPath $OutputPath -Source $source
}

function Get-OptionalValue {
    param($Object, [string]$Name)
    $property = $Object.PSObject.Properties[$Name]
    if ($property -and $null -ne $property.Value) { return [string]$property.Value }
    return ''
}

try {
    $rows = [System.Collections.Generic.List[object]]::new()

    try {
        $tenant = Get-SPOTenant -ErrorAction Stop
        $rows.Add([pscustomobject]@{
                RunDate                        = $runDate
                Scope                          = 'Tenant'
                Url                            = ''
                Title                          = ''
                Template                       = ''
                SharingCapability              = Get-OptionalValue $tenant 'SharingCapability'
                DefaultSharingLinkType         = Get-OptionalValue $tenant 'DefaultSharingLinkType'
                DisableCompanyWideSharingLinks = Get-OptionalValue $tenant 'DisableCompanyWideSharingLinks'
                SensitivityLabel               = ''
            })
    }
    catch {
        Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
            'Get-SPOTenant is unavailable to this sign-in ({0}); no tenant row was written. That it returns SharingCapability is UNVERIFIED.' -f $_.Exception.Message)
    }

    try {
        $sites = @(Get-SPOSite -Limit All -IncludePersonalSite $true -ErrorAction Stop)
    }
    catch {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            'Get-SPOSite is unavailable to this sign-in ({0}). It needs the SharePoint Online administrator role and site collection administrator access. Writing the header only.' -f $_.Exception.Message)
        Export-AppendCsv -Path $csvPath -Column $columns
        return
    }

    if ($SiteLimit -gt 0) { $sites = @($sites | Select-Object -First $SiteLimit) }

    $unread = 0
    foreach ($listed in $sites) {
        # -Limit leaves SharingCapability and the other sharing properties unpopulated, so
        # the site is read again by its URL.
        try {
            $site = Get-SPOSite -Identity ([string]$listed.Url) -ErrorAction Stop
        }
        catch {
            $unread++
            Write-Verbose ('Site {0} not read: {1}' -f $listed.Url, $_.Exception.Message)
            continue
        }

        $rows.Add([pscustomobject]@{
                RunDate                        = $runDate
                Scope                          = 'Site'
                Url                            = [string]$site.Url
                Title                          = Get-OptionalValue $site 'Title'
                Template                       = Get-OptionalValue $site 'Template'
                SharingCapability              = Get-OptionalValue $site 'SharingCapability'
                DefaultSharingLinkType         = Get-OptionalValue $site 'DefaultSharingLinkType'
                DisableCompanyWideSharingLinks = Get-CsvBoolean (Get-OptionalValue $site 'DisableCompanyWideSharingLinks')
                SensitivityLabel               = Get-OptionalValue $site 'SensitivityLabel'
            })
    }

    if ($unread -gt 0) {
        Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
            '{0} site(s) could not be read by URL and were skipped.' -f $unread)
    }

    $result = Export-AppendCsv -Path $csvPath -Rows $rows.ToArray() -Column $columns -KeyColumn @('RunDate', 'Scope', 'Url') -PassThru
    Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
        'site-sharing-settings.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
}
finally {
    if ($connectedHere) {
        Disconnect-SPOService
    }
}
