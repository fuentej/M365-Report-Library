#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes user-signin-activity.csv: the last sign-in times Entra ID holds for each
        user. A snapshot, appended each run.

    .DESCRIPTION
        Source: Microsoft Graph list users with $select=signInActivity
        (https://learn.microsoft.com/graph/api/user-list#example-11-get-users-including-their-last-sign-in-time,
        resource: https://learn.microsoft.com/graph/api/resources/signinactivity).
        Selecting signInActivity caps a page at 500 users; -All follows @odata.nextLink
        until it is absent.

        An empty cell means Graph returned no value. Graph returns 0001-01-01T00:00:00Z
        for lastNonInteractiveSignInDateTime and "" for lastSuccessfulSignInDateTime when
        there is none; both are written as an empty cell, never as a date. An empty
        LastSuccessfulSignInDateTime alone is not proof the account never signed in (it
        is not backfilled), and a blank LastSignInDateTime means the account never signed
        in or its last attempt was before April 2020
        (https://learn.microsoft.com/entra/identity/monitoring-health/howto-manage-inactive-user-accounts).

        Needs Microsoft Entra ID P1 or P2, User.Read.All and AuditLog.Read.All. In GCC
        High the property is UNVERIFIED: it is requested anyway, with a warning, and a
        refusal is recorded in run.log with a header-only file.

    .EXAMPLE
        ./Get-UserSignInActivity.ps1 -OutputPath ./out
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


Invoke-IdentityPostureSnapshot -Source 'UserSignInActivity' -CsvName 'user-signin-activity.csv' `
    -Description 'User sign-in activity' `
    -License 'Microsoft Entra ID P1 or P2, User.Read.All and AuditLog.Read.All (Reports Reader when signed in)' `
    -KeyColumn @('RunDate', 'UserId') -OutputPath $OutputPath -Environment $Environment -AppId $AppId `
    -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization `
    -SchemaPath $SchemaPath -SkipConnect:$SkipConnect `
    -Fetch { Get-MgUser -All -Property @('id', 'userPrincipalName', 'signInActivity') -ErrorAction Stop } `
    -Map {
        param($user, $runDate)

        $activity = Get-GraphAdditionalProperty -Object $user -Name 'signInActivity'

        [pscustomobject]@{
            RunDate                          = $runDate
            UserId                           = $user.Id
            UserPrincipalName                = $user.UserPrincipalName
            LastSignInDateTime               = ConvertTo-SignInTimestamp (Get-GraphAdditionalProperty -Object $activity -Name 'lastSignInDateTime')
            LastNonInteractiveSignInDateTime = ConvertTo-SignInTimestamp (Get-GraphAdditionalProperty -Object $activity -Name 'lastNonInteractiveSignInDateTime')
            LastSuccessfulSignInDateTime     = ConvertTo-SignInTimestamp (Get-GraphAdditionalProperty -Object $activity -Name 'lastSuccessfulSignInDateTime')
        }
    }
