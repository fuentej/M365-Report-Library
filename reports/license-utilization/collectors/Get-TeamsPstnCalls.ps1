#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes teams-pstn-calls.csv: the PSTN calls Teams Phone logged in the last -Days
        days, so the License value page can say whether the E5 step-up feature "Teams
        Phone" is in use. A snapshot, appended each run.

    .DESCRIPTION
        Source 11: Microsoft Graph callRecord getPstnCalls
        (https://learn.microsoft.com/graph/api/callrecords-callrecord-getpstncalls), a GET
        to /v1.0/communications/callRecords/getPstnCalls(fromDateTime=...,toDateTime=...).
        The window cannot exceed 90 days, so -Days is limited to 1 through 90. A page holds
        up to 1,000 rows; @odata.nextLink is followed until it is absent, and a next link
        on another host is refused. The page says the API returns nothing for Telstra
        calling plans. Caller and callee numbers are not written to the CSV.

        Application permission only: CallRecords.Read.All. The page lists delegated access
        as not supported, so run this collector with -AppId and -CertificateThumbprint. A
        signed-in user's token is refused and logged. Available in US Government L4 and L5.

        Each run appends the calls in its window under that run's RunDate, so a call
        collected on two runs appears twice, once per RunDate.

    .PARAMETER Days
        How many days back from now the window starts. Defaults to 30; at most 90.

    .EXAMPLE
        ./Get-TeamsPstnCalls.ps1 -OutputPath ./out -AppId $appId -CertificateThumbprint $thumbprint -TenantId $tenantId
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

    [ValidateRange(1, 90)]
    [int]$Days = 30,

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
$to = [datetime]::UtcNow
$from = $to.AddDays(-$Days)
$uri = '{0}/v1.0/communications/callRecords/getPstnCalls(fromDateTime={1},toDateTime={2})' -f $base,
    $from.ToString('yyyy-MM-ddTHH:mm:ssZ'), $to.ToString('yyyy-MM-ddTHH:mm:ssZ')

Invoke-LicenseUtilizationSnapshot -Source 'TeamsPstnCalls' -CsvName 'teams-pstn-calls.csv' `
    -Description 'Teams PSTN call log' `
    -License 'CallRecords.Read.All as an application permission (delegated is not supported); a Teams Phone licence is the population' `
    -KeyColumn @('RunDate', 'CallRowId') -OutputPath $OutputPath -Environment $Environment -AppId $AppId `
    -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization `
    -SchemaPath $SchemaPath -SkipConnect:$SkipConnect `
    -Fetch { Get-GraphReportJson -Uri $uri } `
    -Map {
        param($call, $runDate)

        $id = [string](Get-GraphAdditionalProperty -Object $call -Name 'id')
        if ([string]::IsNullOrWhiteSpace($id)) { return }

        [pscustomobject]@{
            RunDate           = $runDate
            CallRowId         = $id
            UserId            = [string](Get-GraphAdditionalProperty -Object $call -Name 'userId')
            UserPrincipalName = [string](Get-GraphAdditionalProperty -Object $call -Name 'userPrincipalName')
            StartDateTime     = ConvertTo-CsvTimestamp (Get-GraphAdditionalProperty -Object $call -Name 'startDateTime')
            DurationSeconds   = [string](Get-GraphAdditionalProperty -Object $call -Name 'duration')
            CallType          = [string](Get-GraphAdditionalProperty -Object $call -Name 'callType')
            UsageCountryCode  = [string](Get-GraphAdditionalProperty -Object $call -Name 'usageCountryCode')
            LicenseCapability = [string](Get-GraphAdditionalProperty -Object $call -Name 'licenseCapability')
            InventoryType     = [string](Get-GraphAdditionalProperty -Object $call -Name 'inventoryType')
        }
    }
