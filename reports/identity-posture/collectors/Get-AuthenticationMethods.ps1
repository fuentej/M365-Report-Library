#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes authentication-methods.csv: per user, which authentication methods are
        registered and whether the user is MFA registered and capable, passwordless
        capable and registered for self-service password reset. A snapshot, appended each run.

    .DESCRIPTION
        Source: Microsoft Graph list userRegistrationDetails
        (https://learn.microsoft.com/graph/api/authenticationmethodsroot-list-userregistrationdetails).
        Disabled users are not returned, so a "no MFA" count covers enabled users only;
        join users.csv to list disabled users separately.

        Needs Microsoft Entra ID P1 or P2. Without it the API refuses the call and the
        collector writes the header only and says so in run.log. Needs AuditLog.Read.All
        and, to read the tenant licence reliably, Directory.Read.All. Signed in, the
        Reports Reader role is enough.

    .EXAMPLE
        ./Get-AuthenticationMethods.ps1 -OutputPath ./out
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


Invoke-IdentityPostureSnapshot -Source 'AuthenticationMethods' -CsvName 'authentication-methods.csv' `
    -Description 'Authentication method registration' `
    -License 'Microsoft Entra ID P1 or P2, AuditLog.Read.All and the Reports Reader role (Directory.Read.All as well, to read the tenant licence)' `
    -KeyColumn @('RunDate', 'UserId') -OutputPath $OutputPath -Environment $Environment -AppId $AppId `
    -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization `
    -SchemaPath $SchemaPath -SkipConnect:$SkipConnect `
    -Fetch { Get-MgReportAuthenticationMethodUserRegistrationDetail -All -ErrorAction Stop } `
    -Map {
        param($detail, $runDate)

        [pscustomobject]@{
            RunDate                                       = $runDate
            UserId                                        = $detail.Id
            UserPrincipalName                             = $detail.UserPrincipalName
            UserDisplayName                               = $detail.UserDisplayName
            UserType                                      = [string]$detail.UserType
            IsAdmin                                       = $detail.IsAdmin
            IsMfaRegistered                               = $detail.IsMfaRegistered
            IsMfaCapable                                  = $detail.IsMfaCapable
            IsPasswordlessCapable                         = $detail.IsPasswordlessCapable
            IsSsprEnabled                                 = $detail.IsSsprEnabled
            IsSsprRegistered                              = $detail.IsSsprRegistered
            IsSsprCapable                                 = $detail.IsSsprCapable
            IsSystemPreferredAuthenticationMethodEnabled  = $detail.IsSystemPreferredAuthenticationMethodEnabled
            UserPreferredMethodForSecondaryAuthentication = [string]$detail.UserPreferredMethodForSecondaryAuthentication
            MethodsRegistered                             = Join-ListValue $detail.MethodsRegistered
            SystemPreferredAuthenticationMethods          = Join-ListValue $detail.SystemPreferredAuthenticationMethods
            LastUpdatedDateTime                           = ConvertTo-CsvTimestamp $detail.LastUpdatedDateTime
        }
    }
