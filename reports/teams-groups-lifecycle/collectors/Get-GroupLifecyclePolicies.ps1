#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes group-lifecycle-policies.csv: the group expiration policy, appended each run.

    .DESCRIPTION
        Source: Microsoft Graph list groupLifecyclePolicies
        (https://learn.microsoft.com/graph/api/grouplifecyclepolicy-list). A tenant has
        at most one policy. A tenant with no policy writes the header only and says so
        in run.log.

        Expiration needs Microsoft Entra ID P1 or P2 for the members of every group the
        policy applies to (https://learn.microsoft.com/entra/identity/users/groups-lifecycle).
        A group is deleted one day after its expirationDateTime, and the 30-day restore
        window starts then.

        Needs Directory.Read.All.

    .EXAMPLE
        ./Get-GroupLifecyclePolicies.ps1 -OutputPath ./out
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

    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Imported by path without -Force: a forced re-import builds a fresh module session
# state, which would discard anything the caller has arranged around the module - and
# that is exactly how the tests put mocks in front of the tenant cmdlets.
Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

. (Join-Path $PSScriptRoot 'TeamsGroupsHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'TeamsGroupsSchema.psd1')
$columns = $schema.GroupLifecyclePolicies
$source = 'group-lifecycle-policies'
$csvPath = Join-Path $OutputPath 'group-lifecycle-policies.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-TeamsGroupsSourceAvailability -Source 'GroupLifecyclePolicies' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping group-lifecycle-policies.csv. $($availability.Reason) $($availability.Reference)")
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}
if ($availability.Status -eq 'Unverified') {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message $availability.Reason
}

if (-not $SkipConnect) {
    Connect-M365Service -Service Graph -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization
}

$policies = $null
try {
    $policies = @(Get-MgGroupLifecyclePolicy -All -ErrorAction Stop)
}
catch {
    Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
        'Group lifecycle policies are unavailable to this sign-in ({0}). They need the Directory.Read.All permission, and expiration needs Entra ID P1 or P2. Writing the header only.' -f $_.Exception.Message)
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}

if ($policies.Count -eq 0) {
    Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
        'The tenant has no group expiration policy. Writing the header only.')
}

$rows = foreach ($policy in $policies) {
    [pscustomobject]@{
        RunDate                     = $runDate
        Id                          = $policy.Id
        GroupLifetimeInDays         = $policy.GroupLifetimeInDays
        ManagedGroupTypes           = $policy.ManagedGroupTypes
        AlternateNotificationEmails = $policy.AlternateNotificationEmails
    }
}

$result = Export-AppendCsv -Path $csvPath -Rows @($rows) -Column $columns -KeyColumn @('RunDate', 'Id') -PassThru
Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'group-lifecycle-policies.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
