#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes risky-users.csv: the users Entra ID Protection flags as risky. A snapshot,
        appended each run.

    .DESCRIPTION
        Source: Microsoft Graph list riskyUsers
        (https://learn.microsoft.com/graph/api/riskyuser-list). RiskLevel is low, medium,
        high, hidden, none or unknownFutureValue; RiskState is none, confirmedSafe,
        remediated, dismissed, atRisk, confirmedCompromised or unknownFutureValue
        (https://learn.microsoft.com/graph/api/resources/riskyuser). Values are written
        as Graph returns them. A page holds at most 500 users; -All follows
        @odata.nextLink.

        Needs Microsoft Entra ID P2 (P1 is not enough) and the IdentityRiskyUser.Read.All
        permission. An unlicensed tenant gets a header-only file and a warning, not zero
        risky users.

    .EXAMPLE
        ./Get-RiskyUsers.ps1 -OutputPath ./out
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


Invoke-IdentityPostureSnapshot -Source 'RiskyUsers' -CsvName 'risky-users.csv' `
    -Description 'Risky users' `
    -License 'Microsoft Entra ID P2 and the IdentityRiskyUser.Read.All permission' `
    -KeyColumn @('RunDate', 'Id') -OutputPath $OutputPath -Environment $Environment -AppId $AppId `
    -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization `
    -SchemaPath $SchemaPath -SkipConnect:$SkipConnect `
    -Fetch { Get-MgRiskyUser -All -ErrorAction Stop } `
    -Map {
        param($riskyUser, $runDate)

        [pscustomobject]@{
            RunDate                 = $runDate
            Id                      = $riskyUser.Id
            UserPrincipalName       = $riskyUser.UserPrincipalName
            UserDisplayName         = $riskyUser.UserDisplayName
            RiskLevel               = [string]$riskyUser.RiskLevel
            RiskState               = [string]$riskyUser.RiskState
            RiskDetail              = [string]$riskyUser.RiskDetail
            RiskLastUpdatedDateTime = ConvertTo-CsvTimestamp $riskyUser.RiskLastUpdatedDateTime
            IsDeleted               = $riskyUser.IsDeleted
            IsProcessing            = $riskyUser.IsProcessing
        }
    }
