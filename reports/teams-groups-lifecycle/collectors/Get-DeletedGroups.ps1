#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes deleted-groups.csv: groups in the soft-delete container, appended each run.

    .DESCRIPTION
        Source: Microsoft Graph list deleted items, groups
        (https://learn.microsoft.com/graph/api/directory-deleteditems-list).

        Microsoft 365 and security groups can be restored for 30 days after deletion, so
        PurgeDateTime is deletedDateTime plus 30 days
        (https://learn.microsoft.com/graph/api/group-delete). Distribution groups are
        permanently deleted immediately and never appear here. A soft-deleted security
        group returns securityEnabled false; GroupTypes (Unified or empty) tells
        Microsoft 365 groups from security groups.

        Needs Group.Read.All.

    .EXAMPLE
        ./Get-DeletedGroups.ps1 -OutputPath ./out
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
$columns = $schema.DeletedGroups
$source = 'deleted-groups'
$csvPath = Join-Path $OutputPath 'deleted-groups.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')
$restoreWindowDays = 30

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-TeamsGroupsSourceAvailability -Source 'DeletedGroups' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping deleted-groups.csv. $($availability.Reason) $($availability.Reference)")
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

$select = @(
    'id', 'displayName', 'groupTypes', 'securityEnabled', 'mailEnabled'
    'createdDateTime', 'deletedDateTime', 'resourceProvisioningOptions'
)

$deleted = $null
try {
    $deleted = @(Get-MgDirectoryDeletedItemAsGroup -All -Property $select -ErrorAction Stop)
}
catch {
    Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
        'The deleted groups container is unavailable to this sign-in ({0}). It needs the Group.Read.All permission. Writing the header only.' -f $_.Exception.Message)
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}

$rows = foreach ($group in $deleted) {
    $provisioning = Get-GraphAdditionalProperty -Object $group -Name 'resourceProvisioningOptions'
    $deletedAt = ConvertTo-CsvTimestamp $group.DeletedDateTime
    $purgeAt = ''
    if ($deletedAt) {
        $purgeAt = ConvertTo-CsvTimestamp ([datetime]::Parse($deletedAt, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AdjustToUniversal).AddDays($restoreWindowDays))
    }

    [pscustomobject]@{
        RunDate                     = $runDate
        Id                          = $group.Id
        DisplayName                 = $group.DisplayName
        GroupTypes                  = Join-ListValue (Get-GraphAdditionalProperty -Object $group -Name 'groupTypes')
        SecurityEnabled             = $group.SecurityEnabled
        MailEnabled                 = $group.MailEnabled
        CreatedDateTime             = ConvertTo-CsvTimestamp $group.CreatedDateTime
        DeletedDateTime             = $deletedAt
        PurgeDateTime               = $purgeAt
        ResourceProvisioningOptions = Join-ListValue $provisioning
        IsTeam                      = Test-ListContains -Value $provisioning -Item 'Team'
    }
}

$result = Export-AppendCsv -Path $csvPath -Rows @($rows) -Column $columns -KeyColumn @('RunDate', 'Id') -PassThru
Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'deleted-groups.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
