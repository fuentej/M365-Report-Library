#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes subscribed-skus.csv and sku-service-plans.csv: the licences the tenant owns
        and the service plans each one holds. A snapshot, appended each run.

    .DESCRIPTION
        Source: Microsoft Graph list subscribedSkus
        (https://learn.microsoft.com/graph/api/subscribedsku-list, resource:
        https://learn.microsoft.com/graph/api/resources/subscribedsku), read with
        Get-MgSubscribedSku. The call returns every SKU and supports only $select.

        ConsumedUnits is the number of licences assigned. FreeUnits is derived:
        prepaidUnits.enabled minus consumedUnits. Suspended, warning and locked-out
        units are separate columns and are not folded into it. Only SKUs whose
        AppliesTo is User can be assigned to a user.

        Needs LicenseAssignment.Read.All (Directory.Read.All and Organization.Read.All
        are higher privileged alternatives). Signed in, Global Reader or Directory
        Readers is enough.

    .EXAMPLE
        ./Get-SubscribedSkus.ps1 -OutputPath ./out
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


Invoke-LicenseUtilizationSnapshot -Source 'SubscribedSkus' -CsvName 'subscribed-skus.csv' `
    -Description 'Subscribed SKUs' `
    -License 'LicenseAssignment.Read.All (Global Reader or Directory Readers when signed in)' `
    -KeyColumn @('RunDate', 'SkuId') -OutputPath $OutputPath -Environment $Environment -AppId $AppId `
    -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization `
    -SchemaPath $SchemaPath -SkipConnect:$SkipConnect `
    -Fetch { Get-MgSubscribedSku -All -ErrorAction Stop } `
    -Map {
        param($sku, $runDate)

        $prepaid = Get-GraphAdditionalProperty -Object $sku -Name 'prepaidUnits'
        $enabled = Get-GraphAdditionalProperty -Object $prepaid -Name 'enabled'
        $consumed = Get-GraphAdditionalProperty -Object $sku -Name 'consumedUnits'

        [pscustomobject]@{
            RunDate          = $runDate
            SkuId            = [string]$sku.SkuId
            SkuPartNumber    = $sku.SkuPartNumber
            AppliesTo        = [string](Get-GraphAdditionalProperty -Object $sku -Name 'appliesTo')
            CapabilityStatus = [string](Get-GraphAdditionalProperty -Object $sku -Name 'capabilityStatus')
            ConsumedUnits    = $consumed
            PrepaidEnabled   = $enabled
            PrepaidSuspended = Get-GraphAdditionalProperty -Object $prepaid -Name 'suspended'
            PrepaidWarning   = Get-GraphAdditionalProperty -Object $prepaid -Name 'warning'
            PrepaidLockedOut = Get-GraphAdditionalProperty -Object $prepaid -Name 'lockedOut'
            FreeUnits        = if ($null -ne $enabled -and $null -ne $consumed) { [int]$enabled - [int]$consumed } else { '' }
        }
    } `
    -Secondary @{
        Source    = 'SkuServicePlans'
        CsvName   = 'sku-service-plans.csv'
        KeyColumn = @('RunDate', 'SkuId', 'ServicePlanId')
        Map       = {
            param($sku, $runDate)

            foreach ($plan in @(Get-GraphAdditionalProperty -Object $sku -Name 'servicePlans' | Where-Object { $null -ne $_ })) {
                [pscustomobject]@{
                    RunDate            = $runDate
                    SkuId              = [string]$sku.SkuId
                    ServicePlanId      = [string](Get-GraphAdditionalProperty -Object $plan -Name 'servicePlanId')
                    ServicePlanName    = Get-GraphAdditionalProperty -Object $plan -Name 'servicePlanName'
                    ProvisioningStatus = [string](Get-GraphAdditionalProperty -Object $plan -Name 'provisioningStatus')
                    AppliesTo          = [string](Get-GraphAdditionalProperty -Object $plan -Name 'appliesTo')
                }
            }
        }
    }
