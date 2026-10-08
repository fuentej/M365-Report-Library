#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes group-lifecycle-coverage.csv: whether the expiration policy applies to each
        Microsoft 365 group, appended each run.

    .DESCRIPTION
        Source: the policy's managedGroupTypes from group-lifecycle-policies.csv, plus,
        for a policy scoped to Selected groups, Microsoft Graph list groupLifecyclePolicies
        for a group (https://learn.microsoft.com/graph/api/group-list-grouplifecyclepolicies).
        All covers every group and None covers none, so only Selected costs one call per
        group. Selected scope holds at most 500 groups.

        Run Get-Groups.ps1 and Get-GroupLifecyclePolicies.ps1 first. Near-expiry is
        ExpirationDateTime in groups.csv.

        Needs Directory.Read.All.

    .EXAMPLE
        ./Get-GroupLifecycleCoverage.ps1 -OutputPath ./out
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
$columns = $schema.GroupLifecycleCoverage
$source = 'group-lifecycle-coverage'
$csvPath = Join-Path $OutputPath 'group-lifecycle-coverage.csv'
$groupsPath = Join-Path $OutputPath 'groups.csv'
$policiesPath = Join-Path $OutputPath 'group-lifecycle-policies.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-TeamsGroupsSourceAvailability -Source 'GroupLifecyclePolicies' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping group-lifecycle-coverage.csv. $($availability.Reason) $($availability.Reference)")
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}
if ($availability.Status -eq 'Unverified') {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message $availability.Reason
}

$groups = @(Get-CsvLatestSnapshot -Path $groupsPath | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Id) })
if ($groups.Count -eq 0) {
    Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
        'groups.csv holds no groups, so there is no coverage to work out. Run Get-Groups.ps1 first. Writing the header only.')
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}

$policy = @(Get-CsvLatestSnapshot -Path $policiesPath | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Id) }) | Select-Object -First 1

$rows = [System.Collections.Generic.List[object]]::new()

function New-CoverageRow {
    param([string]$GroupId, [string]$PolicyId, [string]$Status)
    [pscustomobject]@{ RunDate = $runDate; GroupId = $GroupId; PolicyId = $PolicyId; CoverageStatus = $Status }
}

if ($null -eq $policy) {
    Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
        'group-lifecycle-policies.csv holds no policy, so no group is covered.')
    foreach ($group in $groups) { $rows.Add((New-CoverageRow -GroupId $group.Id -PolicyId '' -Status 'NotCovered')) }
}
elseif ($policy.ManagedGroupTypes -eq 'All') {
    foreach ($group in $groups) { $rows.Add((New-CoverageRow -GroupId $group.Id -PolicyId $policy.Id -Status 'Covered')) }
}
elseif ($policy.ManagedGroupTypes -eq 'None') {
    foreach ($group in $groups) { $rows.Add((New-CoverageRow -GroupId $group.Id -PolicyId $policy.Id -Status 'NotCovered')) }
}
else {
    if (-not $SkipConnect) {
        Connect-M365Service -Service Graph -Environment $Environment -AppId $AppId `
            -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization
    }

    $failed = 0
    foreach ($group in $groups) {
        try {
            $applied = @(Get-MgGroupLifecyclePolicyByGroup -GroupId $group.Id -All -ErrorAction Stop)
        }
        catch {
            $failed++
            Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
                'Reading the lifecycle policy for group {0} failed: {1}' -f $group.Id, $_.Exception.Message)
            $rows.Add((New-CoverageRow -GroupId $group.Id -PolicyId $policy.Id -Status 'Unknown'))
            continue
        }
        $status = if ($applied.Count -gt 0) { 'Covered' } else { 'NotCovered' }
        $rows.Add((New-CoverageRow -GroupId $group.Id -PolicyId $policy.Id -Status $status))
    }

    if ($failed -eq $groups.Count) {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            'The lifecycle policy could not be read for any group. It needs the Directory.Read.All permission. Writing the header only.')
        Export-AppendCsv -Path $csvPath -Column $columns
        return
    }
}

$result = Export-AppendCsv -Path $csvPath -Rows $rows.ToArray() -Column $columns -KeyColumn @('RunDate', 'GroupId') -PassThru
Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'group-lifecycle-coverage.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
