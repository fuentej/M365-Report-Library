#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes user-licenses.csv: the licences each user holds, one row per assignment,
        with how it was assigned and whether it is working. A snapshot, appended each run.

    .DESCRIPTION
        Source: Microsoft Graph list users
        (https://learn.microsoft.com/graph/api/user-list), read with Get-MgUser -All and
        an explicit property list, because assignedLicenses and licenseAssignmentStates
        return only when selected. -All follows @odata.nextLink until it is absent
        (https://learn.microsoft.com/graph/paging).

        assignedLicenses includes group-assigned licences and does not say which.
        licenseAssignmentStates does: AssignedByGroup is empty for a direct assignment
        and the group's id otherwise, State is Active, ActiveWithError, Disabled or
        Error, and Error holds the failure
        (https://learn.microsoft.com/graph/api/resources/licenseassignmentstate). A SKU
        found only in assignedLicenses is still written, with State and AssignmentType
        left empty. A user with no licence has no row.

        Department, job title, city and country come from users.csv, which the shared
        Entra users collector (Invoke-EntraUserCollector) writes; they are not read again
        here.

        Needs User.Read.All, which the shared sign-in already requests. The list page
        does not say whether User.ReadBasic.All returns assignedLicenses.

    .EXAMPLE
        ./Get-UserLicenses.ps1 -OutputPath ./out
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


Invoke-LicenseUtilizationSnapshot -Source 'UserLicenses' -CsvName 'user-licenses.csv' `
    -Description 'User licence assignments' `
    -License 'User.Read.All' `
    -KeyColumn @('RunDate', 'UserId', 'SkuId', 'AssignedByGroup') -OutputPath $OutputPath -Environment $Environment -AppId $AppId `
    -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization `
    -SchemaPath $SchemaPath -SkipConnect:$SkipConnect `
    -Fetch {
        Get-MgUser -All -ErrorAction Stop -Property @(
            'id', 'userPrincipalName', 'assignedLicenses', 'licenseAssignmentStates'
            'usageLocation', 'officeLocation', 'accountEnabled')
    } `
    -Map {
        param($user, $runDate)

        $states = @(Get-GraphAdditionalProperty -Object $user -Name 'licenseAssignmentStates' | Where-Object { $null -ne $_ })
        $assigned = @(Get-GraphAdditionalProperty -Object $user -Name 'assignedLicenses' | Where-Object { $null -ne $_ })

        $common = [ordered]@{
            RunDate           = $runDate
            UserId            = $user.Id
            UserPrincipalName = $user.UserPrincipalName
            AccountEnabled    = Get-GraphAdditionalProperty -Object $user -Name 'accountEnabled'
            UsageLocation     = Get-GraphAdditionalProperty -Object $user -Name 'usageLocation'
            OfficeLocation    = Get-GraphAdditionalProperty -Object $user -Name 'officeLocation'
        }

        foreach ($state in $states) {
            $group = [string](Get-GraphAdditionalProperty -Object $state -Name 'assignedByGroup')
            [pscustomobject]($common + [ordered]@{
                SkuId           = [string](Get-GraphAdditionalProperty -Object $state -Name 'skuId')
                AssignedByGroup = $group
                AssignmentType  = if ($group) { 'Group' } else { 'Direct' }
                State           = [string](Get-GraphAdditionalProperty -Object $state -Name 'state')
                Error           = [string](Get-GraphAdditionalProperty -Object $state -Name 'error')
                DisabledPlans   = Join-ListValue (Get-GraphAdditionalProperty -Object $state -Name 'disabledPlans')
            })
        }

        $covered = @($states | ForEach-Object { [string](Get-GraphAdditionalProperty -Object $_ -Name 'skuId') })
        foreach ($license in $assigned) {
            $skuId = [string](Get-GraphAdditionalProperty -Object $license -Name 'skuId')
            if ($covered -contains $skuId) { continue }
            [pscustomobject]($common + [ordered]@{
                SkuId           = $skuId
                AssignedByGroup = ''
                AssignmentType  = ''
                State           = ''
                Error           = ''
                DisabledPlans   = Join-ListValue (Get-GraphAdditionalProperty -Object $license -Name 'disabledPlans')
            })
        }
    }
