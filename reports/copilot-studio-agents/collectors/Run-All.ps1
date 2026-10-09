#Requires -Version 7.0

<#
    .SYNOPSIS
        Runs the six Copilot Studio agents collectors into one output folder.

    .DESCRIPTION
        Four of the six sign in as a person, interactively: the three that read the
        Power Platform inventory (the API refuses service principals) and, through Az,
        the two that read Dataverse. Run this at a keyboard. -AppId and
        -CertificateThumbprint apply only to the audit collector, which follows the
        library's normal connection pattern.

        Each collector signs in for itself, so a source that is unavailable in this cloud
        or unlicensed in this tenant leaves a header-only CSV and a line in run.log
        without stopping the others.

    .PARAMETER DataverseUrl
        The Dataverse URL of each environment to read for agent-components.csv and
        agent-modifications.csv. Without it those two write a header only.

    .PARAMETER ApiHost
        The Power Platform API base URL. Required for GCC and GCC High (UNVERIFIED).

    .EXAMPLE
        ./Run-All.ps1 -OutputPath ./out -DataverseUrl https://org12345.crm.dynamics.com
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$OutputPath,

    [ValidateSet('Commercial', 'GCC', 'GCCHigh')]
    [string]$Environment = 'Commercial',

    [string]$TenantId,
    [string]$ApiHost,
    [string[]]$DataverseUrl,

    [string]$AppId,
    [string]$CertificateThumbprint,
    [string]$Organization,

    [datetime]$StartDate,
    [datetime]$EndDate
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}
$OutputPath = (Resolve-Path -LiteralPath $OutputPath).Path

Write-CollectorLog -OutputPath $OutputPath -Source 'run-all' -Message (
    'Starting the Copilot Studio agents collectors against the {0} cloud.' -f $Environment)

$inventory = @{ OutputPath = $OutputPath; Environment = $Environment; TenantId = $TenantId; ApiHost = $ApiHost }
$dataverse = @{ OutputPath = $OutputPath; Environment = $Environment; TenantId = $TenantId; DataverseUrl = $DataverseUrl }
$audit = @{
    OutputPath = $OutputPath; Environment = $Environment; AppId = $AppId
    CertificateThumbprint = $CertificateThumbprint; TenantId = $TenantId; Organization = $Organization
}
if ($PSBoundParameters.ContainsKey('StartDate')) { $audit['StartDate'] = $StartDate }
if ($PSBoundParameters.ContainsKey('EndDate')) { $audit['EndDate'] = $EndDate }

$steps = @(
    @{ Name = 'environments'; Script = 'Get-PowerPlatformEnvironments.ps1'; Arguments = $inventory }
    @{ Name = 'agents'; Script = 'Get-CopilotStudioAgents.ps1'; Arguments = $inventory }
    @{ Name = 'agent-connectors'; Script = 'Get-AgentConnectors.ps1'; Arguments = $inventory }
    @{ Name = 'agent-components'; Script = 'Get-AgentComponents.ps1'; Arguments = $dataverse }
    @{ Name = 'agent-modifications'; Script = 'Get-AgentModifications.ps1'; Arguments = $dataverse }
    @{ Name = 'agent-audit-events'; Script = 'Get-AgentAuditEvents.ps1'; Arguments = $audit }
)

$failed = 0
foreach ($step in $steps) {
    try {
        $arguments = $step.Arguments
        & (Join-Path $PSScriptRoot $step.Script) @arguments
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
