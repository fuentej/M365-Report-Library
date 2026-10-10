#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes signin-conditional-access.csv: the Conditional Access result recorded on
        each sign-in, one row per applied policy. Appended from the last exported timestamp.

    .DESCRIPTION
        Source 3 of docs/candidates/entra-activity.md. There is no separate Graph call:
        conditionalAccessStatus and appliedConditionalAccessPolicies are properties of the
        sign-in (https://learn.microsoft.com/graph/api/signin-list). This collector reads
        the v1.0 sign-in log, the same endpoint as Get-InteractiveSignIns.ps1, and keeps
        only the Conditional Access part. -IncludeNonInteractive adds the beta
        non-interactive stream (source 2).

        appliedConditionalAccessPolicies comes back only when the caller holds
        Policy.Read.All, Policy.Read.ConditionalAccess or Policy.ReadWrite.ConditionalAccess
        (https://learn.microsoft.com/graph/api/signin-list#permissions). With
        AuditLog.Read.All alone the status is still returned and the policy list is
        dropped WITHOUT an error, so an empty list is not proof that no policy applied.
        This collector asks for Policy.Read.All when it signs in interactively, and each
        row records PolicyDetailReadable from the permissions the session holds.

        The report-only results (reportOnlySuccess and three more) are returned only with
        the header "Prefer: include-unknown-enum-members"
        (https://learn.microsoft.com/graph/api/resources/appliedconditionalaccesspolicy);
        the SDK cmdlet call here does not send it, so a report-only result can read as
        unknownFutureValue. Conditional Access itself needs Entra ID P1.

    .EXAMPLE
        ./Get-SignInConditionalAccess.ps1 -OutputPath ./out -LookbackDays 7
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
    [string]$SchemaPath = (Join-Path $PSScriptRoot 'EntraActivitySchema.psd1'),

    [datetime]$StartDate,
    [datetime]$EndDate,

    # The sign-in log holds 7 days (Free) or 30 days (P1, P2) at most.
    [ValidateRange(1, 30)]
    [int]$LookbackDays = 30,

    [ValidateRange(1, 24)]
    [int]$WindowHours = 24,

    # Also read the beta non-interactive stream.
    [switch]$IncludeNonInteractive,

    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Imported by path without -Force: see Get-InteractiveSignIns.ps1.
Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

. (Join-Path $PSScriptRoot 'EntraActivityHelpers.ps1')

$range = @{}
if ($PSBoundParameters.ContainsKey('StartDate')) { $range['StartDate'] = $StartDate }
if ($PSBoundParameters.ContainsKey('EndDate')) { $range['EndDate'] = $EndDate }

$eventType = (Import-PowerShellDataFile -LiteralPath $SchemaPath).NonInteractiveEventType
$includeBeta = [bool]$IncludeNonInteractive

Invoke-EntraActivityEventCollector -Source SignInConditionalAccess -CsvName 'signin-conditional-access.csv' `
    -OutputPath $OutputPath -Environment $Environment -AppId $AppId `
    -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization `
    -SchemaPath $SchemaPath -LookbackDays $LookbackDays -WindowHours $WindowHours -SkipConnect:$SkipConnect `
    -KeyColumn 'SignInId', 'PolicyId', 'PolicyDisplayName' -WatermarkColumn 'CreatedDateTime' `
    -Description 'the Conditional Access result on each sign-in' `
    -License 'Microsoft Entra ID P1 or P2, the AuditLog.Read.All permission, a Conditional Access read permission (Policy.Read.All) for the per-policy detail and the Reports Reader role' `
    -BeforeFetch { param($schema) Test-ConditionalAccessReadable -Schema $schema } `
    -Fetch {
        param($from, $to)
        Get-MgAuditLogSignIn -All -Filter "createdDateTime ge $from and createdDateTime lt $to" -ErrorAction Stop
        if ($includeBeta) {
            $filter = "(createdDateTime ge $from and createdDateTime lt $to) and signInEventTypes/any(t: t eq '$eventType')"
            Get-MgBetaAuditLogSignIn -All -Filter $filter -ErrorAction Stop
        }
    } `
    -Map {
        param($signIn, $readable)

        $base = [ordered]@{
            CreatedDateTime         = ConvertTo-CsvTimestamp $signIn.CreatedDateTime
            SignInId                = $signIn.Id
            UserId                  = $signIn.UserId
            IsInteractive           = $signIn.IsInteractive
            ConditionalAccessStatus = [string]$signIn.ConditionalAccessStatus
            PolicyDetailReadable    = [bool]$readable
        }

        $policies = @(Get-GraphAdditionalProperty -Object $signIn -Name 'appliedConditionalAccessPolicies' | Where-Object { $null -ne $_ })
        if ($policies.Count -eq 0) {
            # Keep the status. Whether "no policies" means none applied depends on
            # PolicyDetailReadable.
            [pscustomobject]($base + [ordered]@{
                    PolicyId = ''; PolicyDisplayName = ''; PolicyResult = ''
                    EnforcedGrantControls = ''; EnforcedSessionControls = ''
                })
            return
        }

        foreach ($policy in $policies) {
            [pscustomobject]($base + [ordered]@{
                    PolicyId                = Get-PropertyValue $policy 'id'
                    PolicyDisplayName       = Get-PropertyValue $policy 'displayName'
                    PolicyResult            = Get-PropertyValue $policy 'result'
                    EnforcedGrantControls   = Join-ListValue (Get-GraphAdditionalProperty -Object $policy -Name 'enforcedGrantControls')
                    EnforcedSessionControls = Join-ListValue (Get-GraphAdditionalProperty -Object $policy -Name 'enforcedSessionControls')
                })
        }
    } `
    @range
