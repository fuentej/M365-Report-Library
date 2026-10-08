#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes group-owners.csv: the owners of each Microsoft 365 group, appended each run.

    .DESCRIPTION
        Source: Microsoft Graph list group owners
        (https://learn.microsoft.com/graph/api/group-list-owners). Owners are paged, so
        -All follows @odata.nextLink.

        The groups come from the latest snapshot in groups.csv, so run Get-Groups.ps1 first.

        Owners are not available for groups created in Exchange, distribution groups, or
        groups synchronized from on-premises. A group whose owners cannot be read is
        written once with OwnerListStatus Unknown - "owner unknown" - and never as None,
        which means the call succeeded and returned nobody.

        Needs GroupMember.Read.All. A delegated caller also needs one of the roles on the
        list-owners page (group owners, Member users, Guest users (limited), or
        Directory Readers).

    .EXAMPLE
        ./Get-GroupOwners.ps1 -OutputPath ./out
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
$columns = $schema.GroupOwners
$source = 'group-owners'
$csvPath = Join-Path $OutputPath 'group-owners.csv'
$groupsPath = Join-Path $OutputPath 'groups.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-TeamsGroupsSourceAvailability -Source 'GroupOwners' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping group-owners.csv. $($availability.Reason) $($availability.Reference)")
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
        'groups.csv holds no groups, so there are no owners to read. Run Get-Groups.ps1 first. Writing the header only.')
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}

function New-OwnerRow {
    param([string]$GroupId, [string]$Status, $Owner)

    [pscustomobject]@{
        RunDate                = $runDate
        GroupId                = $GroupId
        OwnerListStatus        = $Status
        OwnerId                = if ($Owner) { [string]$Owner.Id } else { '' }
        OwnerType              = if ($Owner) { ([string](Get-GraphAdditionalProperty -Object $Owner -Name '@odata.type')) -replace '^#microsoft\.graph\.', '' } else { '' }
        OwnerDisplayName       = if ($Owner) { [string](Get-GraphAdditionalProperty -Object $Owner -Name 'displayName') } else { '' }
        OwnerUserPrincipalName = if ($Owner) { [string](Get-GraphAdditionalProperty -Object $Owner -Name 'userPrincipalName') } else { '' }
    }
}

$rows = [System.Collections.Generic.List[object]]::new()
$failed = 0

foreach ($group in $groups) {
    if ($group.OnPremisesSyncEnabled -eq 'True') {
        # Graph documents that owners are not available for groups synchronized from
        # on-premises: the answer is "unknown", not "none".
        $rows.Add((New-OwnerRow -GroupId $group.Id -Status 'Unknown' -Owner $null))
        continue
    }

    try {
        $owners = @(Get-MgGroupOwner -GroupId $group.Id -All -ErrorAction Stop)
    }
    catch {
        $failed++
        Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
            'Reading owners for group {0} failed: {1}' -f $group.Id, $_.Exception.Message)
        $rows.Add((New-OwnerRow -GroupId $group.Id -Status 'Unknown' -Owner $null))
        continue
    }

    if ($owners.Count -eq 0) {
        $rows.Add((New-OwnerRow -GroupId $group.Id -Status 'None' -Owner $null))
        continue
    }

    foreach ($owner in $owners) {
        $rows.Add((New-OwnerRow -GroupId $group.Id -Status 'Listed' -Owner $owner))
    }
}

if ($failed -eq $groups.Count) {
    Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
        'Owners could not be read for any group. Delegated sign-in needs GroupMember.Read.All plus a role listed on the list-owners page; app-only needs GroupMember.Read.All. Writing the header only.')
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}

$result = Export-AppendCsv -Path $csvPath -Rows $rows.ToArray() -Column $columns -KeyColumn @('RunDate', 'GroupId', 'OwnerId') -PassThru
Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'group-owners.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
