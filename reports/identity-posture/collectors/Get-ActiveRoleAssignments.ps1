#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes role-assignments-active.csv: active directory role assignments, including
        PIM activations. A snapshot, appended each run.

    .DESCRIPTION
        Source: Microsoft Graph list roleAssignmentScheduleInstances
        (https://learn.microsoft.com/graph/api/rbacapplication-list-roleassignmentscheduleinstances).
        It covers assignments made through PIM and directly through the role assignments
        API. AssignmentType is Assigned or Activated; MemberType is Inherited, Direct or
        Group. An empty EndDateTime is a permanent assignment.

        Join RoleDefinitionId to the role's name and its isPrivileged flag. That flag is
        only on the beta endpoint, which this library does not call; see the README.

        Needs Microsoft Entra ID P2 or Microsoft Entra ID Governance and the
        RoleAssignmentSchedule.Read.Directory permission. A tenant without that licence
        gets a header-only file and a warning; role-assignments.csv
        (Get-RoleAssignments.ps1) still lists the active holders.

    .EXAMPLE
        ./Get-ActiveRoleAssignments.ps1 -OutputPath ./out
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


Invoke-IdentityPostureSnapshot -Source 'ActiveRoleAssignments' -CsvName 'role-assignments-active.csv' `
    -Description 'Active role assignments (PIM)' `
    -License 'Microsoft Entra ID P2 or Microsoft Entra ID Governance and the RoleAssignmentSchedule.Read.Directory permission' `
    -KeyColumn @('RunDate', 'Id') -OutputPath $OutputPath -Environment $Environment -AppId $AppId `
    -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization `
    -SchemaPath $SchemaPath -SkipConnect:$SkipConnect `
    -Fetch { Get-MgRoleManagementDirectoryRoleAssignmentScheduleInstance -All -ErrorAction Stop } `
    -Map {
        param($assignment, $runDate)

        [pscustomobject]@{
            RunDate          = $runDate
            Id               = $assignment.Id
            PrincipalId      = $assignment.PrincipalId
            RoleDefinitionId = $assignment.RoleDefinitionId
            DirectoryScopeId = $assignment.DirectoryScopeId
            AppScopeId       = $assignment.AppScopeId
            AssignmentType           = [string]$assignment.AssignmentType
            MemberType               = [string]$assignment.MemberType
            StartDateTime            = ConvertTo-CsvTimestamp $assignment.StartDateTime
            EndDateTime              = ConvertTo-CsvTimestamp $assignment.EndDateTime
            RoleAssignmentOriginId   = $assignment.RoleAssignmentOriginId
            RoleAssignmentScheduleId = $assignment.RoleAssignmentScheduleId
        }
    }
