#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes attack-simulations.csv: the attack simulation campaigns the tenant has run
        or scheduled, so the License value page can say whether the E5 step-up feature
        "Defender for Office P2" (Attack simulation training) is in use. A snapshot,
        appended each run.

    .DESCRIPTION
        Source 10: Microsoft Graph list simulations
        (https://learn.microsoft.com/graph/api/attacksimulationroot-list-simulations), a
        GET to /v1.0/security/attackSimulation/simulations. @odata.nextLink is followed
        until it is absent, and a next link on another host is refused.

        Needs AttackSimulation.Read.All. Signed in, Security Reader and Security Operator
        can read it, and Security Administrator and Attack Simulation Administrator can
        too (https://learn.microsoft.com/defender-office-365/attack-simulation-training-get-started).
        The feature needs Microsoft 365 E5 or Microsoft Defender for Office 365 Plan 2.

        Not available in GCC High (US Government L4): the header is written and the
        reason logged. The page does not mention GCC, so there the source is UNVERIFIED:
        it is attempted and a warning is logged.

    .EXAMPLE
        ./Get-AttackSimulations.ps1 -OutputPath ./out
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

$base = (Get-M365ServiceEndpoint -Service Graph -Environment $Environment).ResourceEndpoint
$uri = '{0}/v1.0/security/attackSimulation/simulations' -f $base

Invoke-LicenseUtilizationSnapshot -Source 'AttackSimulations' -CsvName 'attack-simulations.csv' `
    -Description 'Attack simulation campaigns' `
    -License 'AttackSimulation.Read.All and, when signed in, Security Reader; Microsoft 365 E5 or Defender for Office 365 Plan 2' `
    -KeyColumn @('RunDate', 'SimulationId') -OutputPath $OutputPath -Environment $Environment -AppId $AppId `
    -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization `
    -SchemaPath $SchemaPath -SkipConnect:$SkipConnect `
    -Fetch { Get-GraphReportJson -Uri $uri } `
    -Map {
        param($simulation, $runDate)

        $id = [string](Get-GraphAdditionalProperty -Object $simulation -Name 'id')
        if ([string]::IsNullOrWhiteSpace($id)) { return }

        [pscustomobject]@{
            RunDate                 = $runDate
            SimulationId            = $id
            DisplayName             = [string](Get-GraphAdditionalProperty -Object $simulation -Name 'displayName')
            AttackType              = [string](Get-GraphAdditionalProperty -Object $simulation -Name 'attackType')
            AttackTechnique         = [string](Get-GraphAdditionalProperty -Object $simulation -Name 'attackTechnique')
            Status                  = [string](Get-GraphAdditionalProperty -Object $simulation -Name 'status')
            IsAutomated             = [string](Get-GraphAdditionalProperty -Object $simulation -Name 'isAutomated')
            PayloadDeliveryPlatform = [string](Get-GraphAdditionalProperty -Object $simulation -Name 'payloadDeliveryPlatform')
            LaunchDateTime          = ConvertTo-CsvTimestamp (Get-GraphAdditionalProperty -Object $simulation -Name 'launchDateTime')
            CompletionDateTime      = ConvertTo-CsvTimestamp (Get-GraphAdditionalProperty -Object $simulation -Name 'completionDateTime')
        }
    }
