#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes guest-memberships.csv: the groups each guest belongs to, appended each run.

    .DESCRIPTION
        Source: Microsoft Graph list a user's memberOf
        (https://learn.microsoft.com/graph/api/user-list-memberof), available in all clouds.

        IsTeam is true when the group's resourceProvisioningOptions contains 'Team'
        (https://learn.microsoft.com/graph/api/resources/group).

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

foreach ($guest in $guests) {
    try {
        $memberships = @(Get-MgUserMemberOf -UserId $guest.Id -All -ErrorAction Stop)
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

        $provisioning = @(Get-GraphAdditionalProperty -Object $membership -Name 'resourceProvisioningOptions')

        $rows.Add([pscustomobject]@{
                RunDate          = $runDate
                GuestId          = $guest.Id
                GroupId          = [string](Get-GraphAdditionalProperty -Object $membership -Name 'id')
                GroupDisplayName = [string](Get-GraphAdditionalProperty -Object $membership -Name 'displayName')
                IsTeam           = $provisioning -contains 'Team'
                Visibility       = [string](Get-GraphAdditionalProperty -Object $membership -Name 'visibility')
            })
    }
}

if ($failed -eq $guests.Count) {
    Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
        'Group memberships could not be read for any guest. The sign-in needs the GroupMember.Read.All or Directory.Read.All permission. Writing the header only.')
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}

$result = Export-AppendCsv -Path $csvPath -Rows $rows.ToArray() -Column $columns -KeyColumn @('RunDate', 'GuestId', 'GroupId') -PassThru
Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'guest-memberships.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
