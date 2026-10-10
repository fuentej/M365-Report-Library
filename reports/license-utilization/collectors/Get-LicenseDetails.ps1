#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes license-details.csv: the service plans in each licence a user holds. A
        snapshot, appended each run. Optional.

    .DESCRIPTION
        Source: Microsoft Graph list licenseDetails
        (https://learn.microsoft.com/graph/api/user-list-licensedetails), read with
        Get-MgUserLicenseDetail, one call per user. The users are those in the latest
        snapshot of user-licenses.csv in the output folder, so run Get-UserLicenses.ps1
        first. Needed only when the join of subscribed-skus and user-licenses is not
        enough; the overlap derivation reads sku-service-plans.csv first.

        The page lists Application permissions as not supported, so this collector needs
        a delegated sign-in with LicenseAssignment.Read.All and one of the roles Guest
        Inviter, Directory Readers, Directory Writers, License Administrator or User
        Administrator. Under -AppId and -CertificateThumbprint every call is refused;
        that is logged and the file holds its header only.

        A user whose call fails is logged and skipped. If every call fails the collector
        fails.

    .EXAMPLE
        ./Get-LicenseDetails.ps1 -OutputPath ./out
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

    # An alternate schema file. The tests use it to exercise the NotAvailable path.
    [string]$SchemaPath = (Join-Path $PSScriptRoot 'LicenseUtilizationSchema.psd1'),

    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Imported by path without -Force: a forced re-import builds a fresh module session
# state, which would discard anything the caller has arranged around the module - and
# that is exactly how the tests put mocks in front of the tenant cmdlets.
Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

. (Join-Path $PSScriptRoot 'LicenseUtilizationHelpers.ps1')


Invoke-LicenseUtilizationSnapshot -Source 'LicenseDetails' -CsvName 'license-details.csv' `
    -Description 'Per-user licence details' `
    -License 'a delegated sign-in with LicenseAssignment.Read.All (application permissions are not supported)' `
    -KeyColumn @('RunDate', 'UserId', 'SkuId', 'ServicePlanId') -OutputPath $OutputPath -Environment $Environment -AppId $AppId `
    -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization `
    -SchemaPath $SchemaPath -SkipConnect:$SkipConnect `
    -Fetch {
        $latest = @(Get-CsvLatestSnapshot -Path (Join-Path $OutputPath 'user-licenses.csv'))
        $userIds = @($latest | Where-Object { $_ } | ForEach-Object { $_.UserId } | Sort-Object -Unique)
        if ($userIds.Count -eq 0) {
            Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source 'license-details' -Message (
                'user-licenses.csv holds no snapshot to read users from. Run Get-UserLicenses.ps1 first. Writing the header only.')
            return
        }

        $failures = 0
        $lastError = ''
        foreach ($userId in $userIds) {
            try {
                foreach ($detail in @(Get-MgUserLicenseDetail -UserId $userId -All -ErrorAction Stop)) {
                    [pscustomobject]@{ UserId = $userId; Detail = $detail }
                }
            }
            catch {
                $failures++
                $lastError = $_.Exception.Message
                Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source 'license-details' -Message (
                    "Licence details for user $userId were refused ($lastError).")
            }
        }
        if ($failures -eq $userIds.Count) { throw $lastError }
    } `
    -Map {
        param($item, $runDate)

        $detail = $item.Detail
        foreach ($plan in @(Get-GraphAdditionalProperty -Object $detail -Name 'servicePlans' | Where-Object { $null -ne $_ })) {
            [pscustomobject]@{
                RunDate            = $runDate
                UserId             = $item.UserId
                SkuId              = [string](Get-GraphAdditionalProperty -Object $detail -Name 'skuId')
                SkuPartNumber      = Get-GraphAdditionalProperty -Object $detail -Name 'skuPartNumber'
                ServicePlanId      = [string](Get-GraphAdditionalProperty -Object $plan -Name 'servicePlanId')
                ServicePlanName    = Get-GraphAdditionalProperty -Object $plan -Name 'servicePlanName'
                ProvisioningStatus = [string](Get-GraphAdditionalProperty -Object $plan -Name 'provisioningStatus')
            }
        }
    }
