#Requires -Version 7.0

<#
    .SYNOPSIS
        Runs the Entra user collector and all five Entra activity collectors into one
        output folder.

    .DESCRIPTION
        users.csv comes from Invoke-EntraUserCollector in the shared module (the report
        joins to user attributes through it; this report does not rebuild the identity
        posture collectors). Each collector signs in for itself, so a collector whose
        source is refused leaves a header-only CSV and a line in run.log without stopping
        the others.

        The sign-in log is read three times (interactive, non-interactive, and the
        Conditional Access detail) because each source is its own CSV. Graph throttles the
        identity and access reports to five requests per 10 seconds per app per tenant
        (https://learn.microsoft.com/graph/throttling-limits#identity-and-access-reports-service-limits);
        the SDK waits the Retry-After seconds on a 429. If throttling persists, shorten
        -StartDate/-EndDate.

    .EXAMPLE
        ./Run-All.ps1 -OutputPath ./out

    .EXAMPLE
        ./Run-All.ps1 -OutputPath ./out -Environment GCCHigh -AppId $appId `
            -CertificateThumbprint $thumbprint -TenantId $tenantId
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
    'Starting the Entra activity collectors against the {0} cloud.' -f $Environment)

$steps = @(
    @{ Name = 'users'; Script = $null; Range = $false }
    @{ Name = 'signins-interactive'; Script = 'Get-InteractiveSignIns.ps1'; Range = $true }
    @{ Name = 'signins-noninteractive'; Script = 'Get-NonInteractiveSignIns.ps1'; Range = $true }
    @{ Name = 'signin-conditional-access'; Script = 'Get-SignInConditionalAccess.ps1'; Range = $true }
    @{ Name = 'directory-audits'; Script = 'Get-DirectoryAudits.ps1'; Range = $true }
    @{ Name = 'retention-reference'; Script = 'Get-RetentionReference.ps1'; Range = $false }
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
