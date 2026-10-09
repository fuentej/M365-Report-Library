#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes delegated-consents.csv: every delegated permission grant, appended each run.

    .DESCRIPTION
        Source: Microsoft Graph list oauth2PermissionGrants
        (https://learn.microsoft.com/graph/api/oauth2permissiongrant-list), read with
        Get-MgOauth2PermissionGrant -All, which follows @odata.nextLink until it is absent
        (https://learn.microsoft.com/graph/paging). A recently created grant can be missing
        until replication catches up.

        HasMailScope is True when a scope starts with Mail. or MailboxSettings.; that
        prefix rule is this collector's, not Microsoft's. ConsentType is AllPrincipals for
        tenant-wide admin consent, in which case PrincipalId is empty and the grant does
        not name who approved it
        (https://learn.microsoft.com/graph/api/resources/oauth2permissiongrant). Who
        consented to a tenant-wide grant is therefore not collected.

        Least privileged permission: Directory.Read.All (application and delegated).
        Signed in, Global Reader and Directory Readers are among the least privileged roles.

    .EXAMPLE
        ./Get-DelegatedConsents.ps1 -OutputPath ./out
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

. (Join-Path $PSScriptRoot 'MailboxExfiltrationHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'MailboxExfiltrationSchema.psd1')
$columns = $schema.DelegatedConsents
$source = 'delegated-consents'
$csvPath = Join-Path $OutputPath 'delegated-consents.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-MailboxSourceAvailability -Source 'DelegatedConsents' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping delegated-consents.csv. $($availability.Reason) $($availability.Reference)")
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}
if ($availability.Status -eq 'Unverified') {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message $availability.Reason
}

if (-not $SkipConnect) {
    Connect-M365Service -Service Graph -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization `
        -Scopes 'Directory.Read.All'
}

try {
    $grants = @(Get-MgOauth2PermissionGrant -All -ErrorAction Stop)
    
    $rows = foreach ($grant in $grants) {
        $scope = [string]$grant.Scope
        $tokens = @($scope -split '\s+' | Where-Object { $_ })
        [pscustomobject]@{
            RunDate      = $runDate
            Id           = [string]$grant.Id
            ClientId     = [string]$grant.ClientId
            ConsentType  = [string]$grant.ConsentType
            PrincipalId  = [string]$grant.PrincipalId
            ResourceId   = [string]$grant.ResourceId
            Scope        = $scope.Trim()
            HasMailScope = [bool](@($tokens | Where-Object { $_ -like 'Mail.*' -or $_ -like 'MailboxSettings.*' }).Count -gt 0) -as [string]
        }
    }
    
    $result = Export-AppendCsv -Path $csvPath -Rows @($rows) -Column $columns -KeyColumn @('RunDate', 'Id') -PassThru
    Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
        'delegated-consents.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
}
catch {
    Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
        'Microsoft Graph list oauth2PermissionGrants is unavailable to this sign-in ({0}). It needs the Directory.Read.All permission (or the Global Reader or Directory Readers role). Writing the header only.' -f $_.Exception.Message)
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}
