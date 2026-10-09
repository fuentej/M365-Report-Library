#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes role-assignments-eligible.csv: eligible (not yet activated) directory role
        assignments. A snapshot, appended each run.

    .DESCRIPTION
        Source: Microsoft Graph list roleEligibilityScheduleInstances
        (https://learn.microsoft.com/graph/api/rbacapplication-list-roleeligibilityscheduleinstances).
        An empty EndDateTime is a permanent eligibility.

        Needs Microsoft Entra ID P2 or Microsoft Entra ID Governance and the
        RoleEligibilitySchedule.Read.Directory permission. When that licence expires
        eligible assignments are removed
        (https://learn.microsoft.com/entra/id-governance/licensing-fundamentals), so an
        unlicensed tenant gets a header-only file and a warning, not a zero.

    .EXAMPLE
        ./Get-EligibleRoleAssignments.ps1 -OutputPath ./out
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


Invoke-IdentityPostureSnapshot -Source 'EligibleRoleAssignments' -CsvName 'role-assignments-eligible.csv' `
    -Description 'Eligible role assignments (PIM)' `
    -License 'Microsoft Entra ID P2 or Microsoft Entra ID Governance and the RoleEligibilitySchedule.Read.Directory permission' `
    -KeyColumn @('RunDate', 'Id') -OutputPath $OutputPath -Environment $Environment -AppId $AppId `
    -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization `
    -SchemaPath $SchemaPath -SkipConnect:$SkipConnect `
    -Fetch { Get-MgRoleManagementDirectoryRoleEligibilityScheduleInstance -All -ErrorAction Stop } `
    -Map {
        param($assignment, $runDate)

        [pscustomobject]@{
            RunDate          = $runDate
            Id               = $assignment.Id
            PrincipalId      = $assignment.PrincipalId
            RoleDefinitionId = $assignment.RoleDefinitionId
            DirectoryScopeId = $assignment.DirectoryScopeId
            AppScopeId       = $assignment.AppScopeId
            MemberType                = [string]$assignment.MemberType
            StartDateTime             = ConvertTo-CsvTimestamp $assignment.StartDateTime
            EndDateTime               = ConvertTo-CsvTimestamp $assignment.EndDateTime
            RoleEligibilityScheduleId = $assignment.RoleEligibilityScheduleId
        }
    }
