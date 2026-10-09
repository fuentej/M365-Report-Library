#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes app-role-assignments.csv: which applications hold app roles on the Microsoft
        Graph service principal, appended each run.

    .DESCRIPTION
        Source: Microsoft Graph list appRoleAssignedTo on the Microsoft Graph service
        principal (appId 00000003-0000-0000-c000-000000000000)
        (https://learn.microsoft.com/graph/api/serviceprincipal-list-approleassignedto),
        read with Get-MgServicePrincipalAppRoleAssignedTo -All, which follows
        @odata.nextLink. A recently granted assignment can be missing until replication
        catches up. The role name comes from the service principal's AppRoles, matched by
        AppRoleId.

        IsMailRole is True when the role value starts with Mail. or MailboxSettings.; that
        prefix rule is this collector's, not Microsoft's.

        Least privileged permission: Application.Read.All. Signed in, the least privileged
        role that can read application assignments is Directory Readers; Global Reader is
        not in that page's role list.

    .EXAMPLE
        ./Get-AppRoleAssignments.ps1 -OutputPath ./out
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

. (Join-Path $PSScriptRoot 'MailboxExfiltrationHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'MailboxExfiltrationSchema.psd1')
$columns = $schema.AppRoleAssignments
$source = 'app-role-assignments'
$csvPath = Join-Path $OutputPath 'app-role-assignments.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-MailboxSourceAvailability -Source 'AppRoleAssignments' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping app-role-assignments.csv. $($availability.Reason) $($availability.Reference)")
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}
if ($availability.Status -eq 'Unverified') {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message $availability.Reason
}

if (-not $SkipConnect) {
    Connect-M365Service -Service Graph -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization `
        -Scopes 'Application.Read.All'
}

try {
    $graphAppId = '00000003-0000-0000-c000-000000000000'
    $servicePrincipal = @(Get-MgServicePrincipal -Filter "appId eq '$graphAppId'" -ErrorAction Stop) | Select-Object -First 1
    if ($null -eq $servicePrincipal) { throw 'The Microsoft Graph service principal was not found in this tenant.' }
    
    $roleValues = @{}
    foreach ($role in @($servicePrincipal.AppRoles)) {
        if ($null -ne $role) { $roleValues[[string]$role.Id] = [string]$role.Value }
    }
    
    $assignments = @(Get-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $servicePrincipal.Id -All -ErrorAction Stop)
    
    $rows = foreach ($assignment in $assignments) {
        $value = if ($roleValues.ContainsKey([string]$assignment.AppRoleId)) { $roleValues[[string]$assignment.AppRoleId] } else { '' }
        [pscustomobject]@{
            RunDate              = $runDate
            AssignmentId         = [string]$assignment.Id
            AppRoleId            = [string]$assignment.AppRoleId
            AppRoleValue         = $value
            PrincipalId          = [string]$assignment.PrincipalId
            PrincipalDisplayName = [string]$assignment.PrincipalDisplayName
            PrincipalType        = [string]$assignment.PrincipalType
            ResourceId           = [string]$assignment.ResourceId
            ResourceDisplayName  = [string]$assignment.ResourceDisplayName
            CreatedDateTime      = ConvertTo-CsvTimestamp $assignment.CreatedDateTime
            IsMailRole           = [bool]($value -like 'Mail.*' -or $value -like 'MailboxSettings.*') -as [string]
        }
    }
    
    $result = Export-AppendCsv -Path $csvPath -Rows @($rows) -Column $columns -KeyColumn @('RunDate', 'AssignmentId') -PassThru
    Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
        'app-role-assignments.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
}
catch {
    Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
        'Microsoft Graph list appRoleAssignedTo is unavailable to this sign-in ({0}). It needs the Application.Read.All permission (or the Directory Readers role). Writing the header only.' -f $_.Exception.Message)
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}
