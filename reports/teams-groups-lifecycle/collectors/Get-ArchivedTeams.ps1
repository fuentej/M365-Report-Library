#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes team-archive-status.csv: whether each team is archived, appended each run.

    .DESCRIPTION
        Source: Microsoft Graph get team (https://learn.microsoft.com/graph/api/team-get),
        which returns isArchived. GET /groups does not return isArchived, even with
        $select, so every Microsoft 365 group in the latest groups.csv snapshot is asked.
        Asking every group, not only those flagged Team, is deliberate: certain unused old
        teams do not carry resourceProvisioningOptions 'Team', and a filter on it would
        drop them. A group that is not a team answers 404 and is skipped.

        Run Get-Groups.ps1 first.

        Needs Team.ReadBasic.All (delegated). If a token with only that permission omits
        isArchived, use TeamSettings.Read.All. Application: TeamSettings.Read.Group
        (resource-specific consent) or Team.ReadBasic.All for an organization-wide read.

    .EXAMPLE
        ./Get-ArchivedTeams.ps1 -OutputPath ./out
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
$columns = $schema.TeamArchiveStatus
$source = 'team-archive-status'
$csvPath = Join-Path $OutputPath 'team-archive-status.csv'
$groupsPath = Join-Path $OutputPath 'groups.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-TeamsGroupsSourceAvailability -Source 'TeamArchiveStatus' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping team-archive-status.csv. $($availability.Reason) $($availability.Reference)")
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

$groups = @(Get-CsvLatestSnapshot -Path $groupsPath | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Id) })

if ($groups.Count -eq 0) {
    Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
        'groups.csv holds no groups, so there are no teams to check. Run Get-Groups.ps1 first. Writing the header only.')
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}

$rows = [System.Collections.Generic.List[object]]::new()
$failed = 0

foreach ($group in $groups) {
    try {
        $team = Get-MgTeam -TeamId $group.Id -ErrorAction Stop
    }
    catch {
        if (Get-GraphErrorIsNotFound -Message $_.Exception.Message) {
            # Not every Microsoft 365 group is a team.
            continue
        }
        $failed++
        Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
            'Reading team {0} failed: {1}' -f $group.Id, $_.Exception.Message)
        continue
    }

    if ($null -eq $team) { continue }

    $rows.Add([pscustomobject]@{
            RunDate     = $runDate
            TeamId      = $team.Id
            DisplayName = $team.DisplayName
            IsArchived  = $team.IsArchived
        })
}

if ($failed -gt 0 -and $rows.Count -eq 0) {
    Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
        'Team archive status could not be read for any team. It needs Team.ReadBasic.All (or TeamSettings.Read.All if isArchived is missing). Writing the header only.')
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}

$result = Export-AppendCsv -Path $csvPath -Rows $rows.ToArray() -Column $columns -KeyColumn @('RunDate', 'TeamId') -PassThru
Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'team-archive-status.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
