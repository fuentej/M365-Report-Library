#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes signins-interactive.csv: interactive sign-ins from the Entra ID sign-in log.
        Appended from the last exported timestamp.

    .DESCRIPTION
        Source 1 of docs/candidates/entra-activity.md: Microsoft Graph list signIns, v1.0
        (https://learn.microsoft.com/graph/api/signin-list), available in all three clouds.
        v1.0 returns sign-ins that are interactive in nature plus successful federated
        sign-ins; non-interactive history is Get-NonInteractiveSignIns.ps1. Each page is
        read with Invoke-MgGraphRequest and the full @odata.nextLink is followed. A page
        holds at most 1,000 sign-ins, so stopping after the first page would drop the rest.
        The request sends Prefer: include-unknown-enum-members on every page.

        Each window is filtered on createdDateTime (UTC), as the list page advises, so no
        request asks for an unbounded span. Retention is 7 days (Free) or 30 days (P1, P2)
        and is UNVERIFIED in GCC and GCC High; a daily run builds a longer history than
        the source keeps. Rows are keyed by sign-in Id.

        Needs Entra ID P1 or P2 (the Learn pages disagree about Free, so a refusal is
        logged as "not licensed"), AuditLog.Read.All and the Reports Reader role.
        riskDetail and the risk levels read "hidden" without P2.

    .EXAMPLE
        ./Get-InteractiveSignIns.ps1 -OutputPath ./out -LookbackDays 7
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

    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Imported by path without -Force: a forced re-import builds a fresh module session
# state, which would discard anything the caller has arranged around the module - and
# that is exactly how the tests put mocks in front of the tenant cmdlets.
Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

. (Join-Path $PSScriptRoot 'EntraActivityHelpers.ps1')

$range = @{}
if ($PSBoundParameters.ContainsKey('StartDate')) { $range['StartDate'] = $StartDate }
if ($PSBoundParameters.ContainsKey('EndDate')) { $range['EndDate'] = $EndDate }

Invoke-EntraActivityEventCollector -Source InteractiveSignIns -CsvName 'signins-interactive.csv' `
    -OutputPath $OutputPath -Environment $Environment -AppId $AppId `
    -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization `
    -SchemaPath $SchemaPath -LookbackDays $LookbackDays -WindowHours $WindowHours -SkipConnect:$SkipConnect `
    -ColumnKey 'SignIns' -KeyColumn 'Id' -WatermarkColumn 'CreatedDateTime' -Description 'the interactive sign-in log' `
    -License 'Microsoft Entra ID P1 or P2, the AuditLog.Read.All permission and the Reports Reader role' `
    -Fetch {
        param($from, $to)
        Get-EntraActivityPagedValues -Version 'v1.0' -RelativePath 'auditLogs/signIns' `
            -Filter "createdDateTime ge $from and createdDateTime lt $to" `
            -OutputPath $OutputPath -LogSource 'signins-interactive'
    } `
    -Map { param($signIn) ConvertTo-SignInRow -SignIn $signIn } `
    @range
