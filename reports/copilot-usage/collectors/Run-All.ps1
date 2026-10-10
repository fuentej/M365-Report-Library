#Requires -Version 7.0

<#
    .SYNOPSIS
        Runs the Entra user collector and all six Copilot usage collectors into one output
        folder.

    .DESCRIPTION
        users.csv comes from Invoke-EntraUserCollector in the shared module and is written
        first, because Get-CopilotInteractions.ps1 reads the users from it. Licence
        inventory (who holds a Copilot licence) and the report-settings check for concealed
        names belong to the license utilization report and are not run here.

        Each collector signs in for itself, so a collector whose source is unavailable in this
        cloud or refused leaves a header-only CSV and a line in run.log without stopping the
        others. In GCC High the three usage report collectors skip: Microsoft marks them
        unavailable there.

        Get-CopilotInteractions.ps1 reads prompt and response METADATA for every user and
        needs the application permission AiEnterpriseInteraction.Read.All, which lets the
        caller read every prompt in the tenant. It is run last and is easy to leave out; see
        -SkipInteractions.

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
    [datetime]$EndDate,

    # Leave out the interaction export, the most sensitive source in the report.
    [switch]$SkipInteractions
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
    'Starting the Copilot usage collectors against the {0} cloud.' -f $Environment)

$steps = @(
    @{ Name = 'users'; Script = $null; Range = $false }
    @{ Name = 'copilot-usage-user-detail'; Script = 'Get-CopilotUsageUserDetail.ps1'; Range = $false }
    @{ Name = 'copilot-user-count-summary'; Script = 'Get-CopilotUserCountSummary.ps1'; Range = $false }
    @{ Name = 'copilot-user-count-trend'; Script = 'Get-CopilotUserCountTrend.ps1'; Range = $false }
    @{ Name = 'copilot-audit-events'; Script = 'Get-CopilotAuditEvents.ps1'; Range = $true }
    @{ Name = 'copilot-feature-availability'; Script = 'Get-CopilotFeatureAvailability.ps1'; Range = $false }
)
if (-not $SkipInteractions) {
    $steps += @{ Name = 'copilot-interactions'; Script = 'Get-CopilotInteractions.ps1'; Range = $true }
}

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
