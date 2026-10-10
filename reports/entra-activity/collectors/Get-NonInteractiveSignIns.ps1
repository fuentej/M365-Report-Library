#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes signins-noninteractive.csv: non-interactive user sign-ins from the Entra ID
        sign-in log. Appended from the last exported timestamp.

    .DESCRIPTION
        Source 2 of docs/candidates/entra-activity.md: the BETA list signIns API
        (https://learn.microsoft.com/graph/api/signin-list?view=graph-rest-beta),
        Get-MgBetaAuditLogSignIn, available in all three clouds. The beta page says the
        API returns only interactive sign-ins unless signInEventTypes is filtered, so each
        window is filtered to 'nonInteractiveUser' (spelled as the list examples spell it).
        The filter is not `ne 'interactiveUser'`, which would also return service principal
        and managed identity sign-ins. Pages are followed on @odata.nextLink, and every
        page sends Prefer: include-unknown-enum-members.

        CAVEAT: Microsoft does not support beta APIs in production applications. A report
        that cannot accept that ships signins-interactive.csv alone and states that it
        covers interactive sign-ins only.

        Paging, windows, retention, licence and permissions are as for
        Get-InteractiveSignIns.ps1. The beta path is /beta/auditLogs/signIns on the
        signed-in Graph host; it does not need the Microsoft.Graph.Beta.Reports module.

    .EXAMPLE
        ./Get-NonInteractiveSignIns.ps1 -OutputPath ./out -LookbackDays 7
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

# Imported by path without -Force: see Get-InteractiveSignIns.ps1.
Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

. (Join-Path $PSScriptRoot 'EntraActivityHelpers.ps1')

$range = @{}
if ($PSBoundParameters.ContainsKey('StartDate')) { $range['StartDate'] = $StartDate }
if ($PSBoundParameters.ContainsKey('EndDate')) { $range['EndDate'] = $EndDate }

$eventType = (Import-PowerShellDataFile -LiteralPath $SchemaPath).NonInteractiveEventType

Invoke-EntraActivityEventCollector -Source NonInteractiveSignIns -CsvName 'signins-noninteractive.csv' `
    -OutputPath $OutputPath -Environment $Environment -AppId $AppId `
    -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization `
    -SchemaPath $SchemaPath -LookbackDays $LookbackDays -WindowHours $WindowHours -SkipConnect:$SkipConnect `
    -ColumnKey 'SignIns' -KeyColumn 'Id' -WatermarkColumn 'CreatedDateTime' -Description 'the non-interactive sign-in log (beta)' `
    -License 'Microsoft Entra ID P1 or P2, the AuditLog.Read.All permission and the Reports Reader role' `
    -Fetch {
        param($from, $to)
        $filter = "(createdDateTime ge $from and createdDateTime lt $to) and signInEventTypes/any(t: t eq '$eventType')"
        Get-EntraActivityPagedValues -Version 'beta' -RelativePath 'auditLogs/signIns' -Filter $filter `
            -OutputPath $OutputPath -LogSource 'signins-noninteractive'
    } `
    -Map { param($signIn) ConvertTo-SignInRow -SignIn $signIn } `
    @range
