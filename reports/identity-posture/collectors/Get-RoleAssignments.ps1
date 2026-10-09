#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes role-assignments.csv: active directory role assignments without PIM.
        A snapshot, appended each run.

    .DESCRIPTION
        Source: Microsoft Graph list roleAssignments
        (https://learn.microsoft.com/graph/api/rbacapplication-list-roleassignments).
        Use this for active holders in a tenant that cannot call the PIM schedule
        instances. It does not return eligible assignments.

        Needs the RoleManagement.Read.Directory permission. Signed in: Directory Readers,
        Global Reader or Privileged Role Administrator. The API page names no licence.

    .EXAMPLE
        ./Get-RoleAssignments.ps1 -OutputPath ./out
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
    [string]$SchemaPath = (Join-Path $PSScriptRoot 'IdentityPostureSchema.psd1'),

    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Imported by path without -Force: a forced re-import builds a fresh module session
# state, which would discard anything the caller has arranged around the module - and
# that is exactly how the tests put mocks in front of the tenant cmdlets.
Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

. (Join-Path $PSScriptRoot 'IdentityPostureHelpers.ps1')


Invoke-IdentityPostureSnapshot -Source 'RoleAssignments' -CsvName 'role-assignments.csv' `
    -Description 'Directory role assignments' `
    -License 'the RoleManagement.Read.Directory permission (no Microsoft Entra licence is named on the API page)' `
    -KeyColumn @('RunDate', 'Id') -OutputPath $OutputPath -Environment $Environment -AppId $AppId `
    -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization `
    -SchemaPath $SchemaPath -SkipConnect:$SkipConnect `
    -Fetch { Get-MgRoleManagementDirectoryRoleAssignment -All -ErrorAction Stop } `
    -Map {
        param($assignment, $runDate)

        [pscustomobject]@{
            RunDate          = $runDate
            Id               = $assignment.Id
            PrincipalId      = $assignment.PrincipalId
            RoleDefinitionId = $assignment.RoleDefinitionId
            DirectoryScopeId = $assignment.DirectoryScopeId
            AppScopeId       = $assignment.AppScopeId
        }
    }
