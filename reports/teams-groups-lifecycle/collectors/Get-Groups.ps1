#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes groups.csv: a snapshot of every Microsoft 365 group (a Team is one of
        them), appended each run.

    .DESCRIPTION
        Source: Microsoft Graph list groups (https://learn.microsoft.com/graph/api/group-list),
        filtered to groupTypes/any(c:c eq 'Unified'). Graph returns 100 groups a page by
        default and 999 at most; -All follows @odata.nextLink until it is absent
        (https://learn.microsoft.com/graph/paging).

        A team is usually a group whose resourceProvisioningOptions contains 'Team'
        (https://learn.microsoft.com/graph/teams-list-all-teams), but certain unused old
        teams do not carry that value. IsTeam here is only that hint; Get-ArchivedTeams.ps1
        asks Graph about every Microsoft 365 group, so it does not depend on it.

        expirationDateTime is always UTC and null for security groups
        (https://learn.microsoft.com/graph/api/resources/group).

        Needs Group.Read.All.

    .EXAMPLE
        ./Get-Groups.ps1 -OutputPath ./out
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
$columns = $schema.Groups
$source = 'groups'
$csvPath = Join-Path $OutputPath 'groups.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-TeamsGroupsSourceAvailability -Source 'Groups' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping groups.csv. $($availability.Reason) $($availability.Reference)")
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
    'id', 'displayName', 'mail', 'groupTypes', 'securityEnabled', 'mailEnabled', 'visibility'
    'createdDateTime', 'renewedDateTime', 'expirationDateTime', 'deletedDateTime'
    'onPremisesSyncEnabled', 'resourceProvisioningOptions'
)

$groups = $null
try {
    $groups = @(Get-MgGroup -All -Filter "groupTypes/any(c:c eq 'Unified')" -Property $select -ErrorAction Stop)
}
catch {
    Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
        'Microsoft Graph list groups is unavailable to this sign-in ({0}). It needs the Group.Read.All permission. Writing the header only.' -f $_.Exception.Message)
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}

$rows = foreach ($group in $groups) {
    $provisioning = Get-GraphAdditionalProperty -Object $group -Name 'resourceProvisioningOptions'
    $groupTypes = Get-GraphAdditionalProperty -Object $group -Name 'groupTypes'

    [pscustomobject]@{
        RunDate                     = $runDate
        Id                          = $group.Id
        DisplayName                 = $group.DisplayName
        Mail                        = $group.Mail
        GroupTypes                  = Join-ListValue $groupTypes
        SecurityEnabled             = $group.SecurityEnabled
        MailEnabled                 = $group.MailEnabled
        Visibility                  = $group.Visibility
        CreatedDateTime             = ConvertTo-CsvTimestamp $group.CreatedDateTime
        RenewedDateTime             = ConvertTo-CsvTimestamp $group.RenewedDateTime
        ExpirationDateTime          = ConvertTo-CsvTimestamp $group.ExpirationDateTime
        DeletedDateTime             = ConvertTo-CsvTimestamp $group.DeletedDateTime
        OnPremisesSyncEnabled       = $group.OnPremisesSyncEnabled
        ResourceProvisioningOptions = Join-ListValue $provisioning
        IsTeam                      = Test-ListContains -Value $provisioning -Item 'Team'
    }
}

$result = Export-AppendCsv -Path $csvPath -Rows @($rows) -Column $columns -KeyColumn @('RunDate', 'Id') -PassThru
Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'groups.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
