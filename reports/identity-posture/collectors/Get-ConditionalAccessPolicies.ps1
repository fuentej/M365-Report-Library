#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes conditional-access-policies.csv: every Conditional Access policy with its
        state, who and what it targets, and its grant controls. A snapshot, appended each run.

    .DESCRIPTION
        Source: Microsoft Graph list policies
        (https://learn.microsoft.com/graph/api/conditionalaccessroot-list-policies).
        State is enabled, disabled or enabledForReportingButNotEnforced (report-only in the
        portal) and is written exactly as Graph returns it
        (https://learn.microsoft.com/graph/api/resources/conditionalaccesspolicy).

        A change shows up as a difference between two RunDate snapshots; the API records
        no one who made it. List cells (users, groups, roles, apps, controls) hold their
        values separated by semicolons.

        Needs Microsoft Entra ID P1 and the Policy.Read.All permission. Risk-based
        policies (sign-in risk, user risk) need P2 to create but are read the same way.

    .EXAMPLE
        ./Get-ConditionalAccessPolicies.ps1 -OutputPath ./out
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


Invoke-IdentityPostureSnapshot -Source 'ConditionalAccessPolicies' -CsvName 'conditional-access-policies.csv' `
    -Description 'Conditional Access policies' `
    -License 'Microsoft Entra ID P1 (or Microsoft 365 Business Premium) and the Policy.Read.All permission' `
    -KeyColumn @('RunDate', 'Id') -OutputPath $OutputPath -Environment $Environment -AppId $AppId `
    -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization `
    -SchemaPath $SchemaPath -SkipConnect:$SkipConnect `
    -Fetch { Get-MgIdentityConditionalAccessPolicy -All -ErrorAction Stop } `
    -Map {
        param($policy, $runDate)

        $conditions = Get-GraphAdditionalProperty -Object $policy -Name 'conditions'
        $users = Get-GraphAdditionalProperty -Object $conditions -Name 'users'
        $apps = Get-GraphAdditionalProperty -Object $conditions -Name 'applications'
        $grant = Get-GraphAdditionalProperty -Object $policy -Name 'grantControls'

        [pscustomobject]@{
            RunDate             = $runDate
            Id                  = $policy.Id
            DisplayName         = $policy.DisplayName
            State               = [string]$policy.State
            CreatedDateTime     = ConvertTo-CsvTimestamp $policy.CreatedDateTime
            ModifiedDateTime    = ConvertTo-CsvTimestamp $policy.ModifiedDateTime
            IncludeUsers        = Join-ListValue (Get-GraphAdditionalProperty -Object $users -Name 'includeUsers')
            ExcludeUsers        = Join-ListValue (Get-GraphAdditionalProperty -Object $users -Name 'excludeUsers')
            IncludeGroups       = Join-ListValue (Get-GraphAdditionalProperty -Object $users -Name 'includeGroups')
            ExcludeGroups       = Join-ListValue (Get-GraphAdditionalProperty -Object $users -Name 'excludeGroups')
            IncludeRoles        = Join-ListValue (Get-GraphAdditionalProperty -Object $users -Name 'includeRoles')
            ExcludeRoles        = Join-ListValue (Get-GraphAdditionalProperty -Object $users -Name 'excludeRoles')
            IncludeApplications = Join-ListValue (Get-GraphAdditionalProperty -Object $apps -Name 'includeApplications')
            ExcludeApplications = Join-ListValue (Get-GraphAdditionalProperty -Object $apps -Name 'excludeApplications')
            ClientAppTypes      = Join-ListValue (Get-GraphAdditionalProperty -Object $conditions -Name 'clientAppTypes')
            SignInRiskLevels    = Join-ListValue (Get-GraphAdditionalProperty -Object $conditions -Name 'signInRiskLevels')
            UserRiskLevels      = Join-ListValue (Get-GraphAdditionalProperty -Object $conditions -Name 'userRiskLevels')
            GrantOperator               = [string](Get-GraphAdditionalProperty -Object $grant -Name 'operator')
            BuiltInControls             = Join-ListValue (Get-GraphAdditionalProperty -Object $grant -Name 'builtInControls')
            CustomAuthenticationFactors = Join-ListValue (Get-GraphAdditionalProperty -Object $grant -Name 'customAuthenticationFactors')
            TermsOfUse                  = Join-ListValue (Get-GraphAdditionalProperty -Object $grant -Name 'termsOfUse')
            AuthenticationStrength      = [string](Get-GraphAdditionalProperty -Object (Get-GraphAdditionalProperty -Object $grant -Name 'authenticationStrength') -Name 'displayName')
        }
    }
