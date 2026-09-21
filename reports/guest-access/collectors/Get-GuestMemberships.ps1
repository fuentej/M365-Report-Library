#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes guest-memberships.csv: the groups each guest belongs to, appended each run.

    .DESCRIPTION
        Source: Microsoft Graph list a user's memberOf
        (https://learn.microsoft.com/graph/api/user-list-memberof), available in all clouds.

        IsTeam is true when the group's resourceProvisioningOptions contains 'Team'
        (https://learn.microsoft.com/graph/api/resources/group). Those properties are
        requested with $select. App-only sign-in needs the Directory.Read.All application
        permission for another user's memberOf; GroupMember.Read.All is not enough, and
        Graph then returns id-only objects instead of failing
        (https://learn.microsoft.com/graph/api/user-list-memberof).

        The guests come from the latest snapshot in guests.csv, so run Get-Guests.ps1 first.

    .EXAMPLE
        ./Get-GuestMemberships.ps1 -OutputPath ./out
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

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'GuestAccessSchema.psd1')
$columns = $schema.GuestMemberships
$source = 'guest-memberships'
$csvPath = Join-Path $OutputPath 'guest-memberships.csv'
$guestsPath = Join-Path $OutputPath 'guests.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

if (-not $SkipConnect) {
    Connect-M365Service -Service Graph -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization
}

$guests = @(Get-CsvLatestSnapshot -Path $guestsPath | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Id) })

if ($guests.Count -eq 0) {
    Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
        'guests.csv holds no guests, so there are no memberships to read. Run Get-Guests.ps1 first. Writing the header only.')
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}

$rows = [System.Collections.Generic.List[object]]::new()
$failed = 0
$limitedInfo = 0

foreach ($guest in $guests) {
    try {
        $memberships = @(Get-MgUserMemberOfAsGroup -UserId $guest.Id -All -Property @(
                'id', 'displayName', 'visibility', 'resourceProvisioningOptions'
            ) -ErrorAction Stop)
    }
    catch {
        $failed++
        Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
            'Reading group memberships for {0} failed: {1}' -f $guest.Id, $_.Exception.Message)
        continue
    }

    foreach ($membership in $memberships) {
        $odataType = [string](Get-GraphAdditionalProperty -Object $membership -Name '@odata.type')
        if ($odataType -and $odataType -ne '#microsoft.graph.group') { continue }

        $groupId = [string](Get-GraphAdditionalProperty -Object $membership -Name 'id')
        if ([string]::IsNullOrWhiteSpace($groupId)) { continue }

        $displayName = [string](Get-GraphAdditionalProperty -Object $membership -Name 'displayName')
        if ([string]::IsNullOrWhiteSpace($displayName)) {
            # Limited-information payload: Graph returns id and @odata.type when the
            # sign-in cannot read the group, and does not throw.
            $limitedInfo++
            Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
                'Group {0} for guest {1} came back with no display name. App-only memberOf of another user needs the Directory.Read.All application permission.' -f $groupId, $guest.Id)
            continue
        }

        $provisioning = @(Get-GraphAdditionalProperty -Object $membership -Name 'resourceProvisioningOptions')

        $rows.Add([pscustomobject]@{
                RunDate          = $runDate
                GuestId          = $guest.Id
                GroupId          = $groupId
                GroupDisplayName = $displayName
                IsTeam           = $provisioning -contains 'Team'
                Visibility       = [string](Get-GraphAdditionalProperty -Object $membership -Name 'visibility')
            })
    }
}

if ($failed -eq $guests.Count) {
    Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
        'Group memberships could not be read for any guest. Delegated sign-in needs GroupMember.Read.All or Directory.Read.All; app-only needs Directory.Read.All. Writing the header only.')
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}

if ($limitedInfo -gt 0 -and $rows.Count -eq 0) {
    Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
        'Group memberships came back with ids only. App-only sign-in needs the Directory.Read.All application permission for another user''s memberOf; Graph then returns limited objects instead of failing. Writing the header only.')
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}

$result = Export-AppendCsv -Path $csvPath -Rows $rows.ToArray() -Column $columns -KeyColumn @('RunDate', 'GuestId', 'GroupId') -PassThru
Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'guest-memberships.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
