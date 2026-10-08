#Requires -Version 7.0

<#
    .SYNOPSIS
        Runs the Entra user collector and all nine Teams and Groups lifecycle collectors
        into one output folder.

    .DESCRIPTION
        Order matters: groups.csv is written first, because Get-GroupOwners.ps1,
        Get-GroupLifecycleCoverage.ps1 and Get-ArchivedTeams.ps1 read the groups it lists,
        and Get-GroupLifecycleCoverage.ps1 also reads group-lifecycle-policies.csv.

        Each collector signs in for itself, so a collector whose service is unavailable
        in this cloud or unlicensed in this tenant leaves a header-only CSV and a line in
        run.log without stopping the others.

    .PARAMETER Organization
        The tenant's *.onmicrosoft.com domain. Required for app-only sign-in to Exchange
        Online, which Get-GroupCreationEvents.ps1 uses.

    .EXAMPLE
        ./Run-All.ps1 -OutputPath ./out

    .EXAMPLE
        ./Run-All.ps1 -OutputPath ./out -Environment GCCHigh -AppId $appId `
            -CertificateThumbprint $thumbprint -TenantId $tenantId -Organization contoso.onmicrosoft.com
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

    [datetime]$StartDate,
    [datetime]$EndDate
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Imported by path without -Force: a forced re-import builds a fresh module session
# state, which would discard anything the caller has arranged around the module - and
# that is exactly how the tests put mocks in front of the tenant cmdlets.
Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}
$OutputPath = (Resolve-Path -LiteralPath $OutputPath).Path

$auth = @{
    Environment           = $Environment
    AppId                 = $AppId
    CertificateThumbprint = $CertificateThumbprint
    TenantId              = $TenantId
    Organization          = $Organization
}

$range = @{}
if ($PSBoundParameters.ContainsKey('StartDate')) { $range['StartDate'] = $StartDate }
if ($PSBoundParameters.ContainsKey('EndDate')) { $range['EndDate'] = $EndDate }

Write-CollectorLog -OutputPath $OutputPath -Source 'run-all' -Message (
    'Starting the Teams and Groups lifecycle collectors against the {0} cloud.' -f $Environment)

$steps = @(
    @{ Name = 'users'; Script = $null; Range = $false }
    @{ Name = 'groups'; Script = 'Get-Groups.ps1'; Range = $false }
    @{ Name = 'group-owners'; Script = 'Get-GroupOwners.ps1'; Range = $false }
    @{ Name = 'deleted-groups'; Script = 'Get-DeletedGroups.ps1'; Range = $false }
    @{ Name = 'group-lifecycle-policies'; Script = 'Get-GroupLifecyclePolicies.ps1'; Range = $false }
    @{ Name = 'group-lifecycle-coverage'; Script = 'Get-GroupLifecycleCoverage.ps1'; Range = $false }
    @{ Name = 'team-archive-status'; Script = 'Get-ArchivedTeams.ps1'; Range = $false }
    @{ Name = 'team-activity'; Script = 'Get-TeamActivity.ps1'; Range = $false }
    @{ Name = 'group-activity'; Script = 'Get-GroupActivity.ps1'; Range = $false }
    @{ Name = 'group-creation-events'; Script = 'Get-GroupCreationEvents.ps1'; Range = $true }
)

$failed = 0
foreach ($step in $steps) {
    try {
        if ($null -eq $step.Script) {
            Invoke-EntraUserCollector -OutputPath $OutputPath @auth
        }
        else {
            $arguments = @{ OutputPath = $OutputPath } + $auth
            if ($step.Range) { $arguments += $range }
            & (Join-Path $PSScriptRoot $step.Script) @arguments
        }
    }
    catch {
        # One collector failing outright must not stop the rest of the run.
        $failed++
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source 'run-all' -Message (
            'The {0} collector stopped with an error: {1}' -f $step.Name, $_.Exception.Message)
    }
}

Write-CollectorLog -OutputPath $OutputPath -Source 'run-all' -Message 'Finished.'

if ($failed -gt 0) {
    throw ('{0} collector(s) stopped with an error. See run.log.' -f $failed)
}
