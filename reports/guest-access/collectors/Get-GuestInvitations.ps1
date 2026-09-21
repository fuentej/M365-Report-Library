#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes guest-invitations.csv: guest invitation and redemption events from the
        Entra ID directory audit log.

    .DESCRIPTION
        Source: Microsoft Graph list directoryAudits
        (https://learn.microsoft.com/graph/api/directoryaudit-list), available in all clouds.

        The activityDisplayName values collected are listed in GuestAccessSchema.psd1 and
        in the report README, verified against
        https://learn.microsoft.com/entra/identity/monitoring-health/reference-audit-activities.

        Directory audit retention is 7 days on Entra ID Free and 30 days on P1/P2
        (https://learn.microsoft.com/entra/identity/monitoring-health/reference-reports-data-retention),
        so a run started after the retention window has passed cannot recover the gap.

    .PARAMETER StartDate
        Overrides the start of the query window. By default the collector resumes from
        the latest ActivityDateTime already in guest-invitations.csv, or -LookbackDays
        ago when the file is new.

    .EXAMPLE
        ./Get-GuestInvitations.ps1 -OutputPath ./out
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

    [datetime]$StartDate,
    [datetime]$EndDate,

    [ValidateRange(1, 30)]
    [int]$LookbackDays = 30,

    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Imported by path without -Force: a forced re-import builds a fresh module session
# state, which would discard anything the caller has arranged around the module - and
# that is exactly how the tests put mocks in front of the tenant cmdlets.
Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'GuestAccessSchema.psd1')
$columns = $schema.GuestInvitations
$activities = $schema.InvitationActivities
$source = 'guest-invitations'
$csvPath = Join-Path $OutputPath 'guest-invitations.csv'

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

if (-not $SkipConnect) {
    Connect-M365Service -Service Graph -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization
}

$watermark = Get-CsvWatermark -Path $csvPath -Column 'ActivityDateTime'

$start = if ($PSBoundParameters.ContainsKey('StartDate')) { $StartDate.ToUniversalTime() }
elseif ($null -ne $watermark) { $watermark }
else { [datetime]::UtcNow.AddDays(-$LookbackDays) }

$end = if ($PSBoundParameters.ContainsKey('EndDate')) { $EndDate.ToUniversalTime() } else { [datetime]::UtcNow }

Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'Querying directory audits from {0} to {1} (watermark: {2}).' -f
    (ConvertTo-CsvTimestamp $start), (ConvertTo-CsvTimestamp $end),
    $(if ($null -eq $watermark) { 'none' } else { ConvertTo-CsvTimestamp $watermark }))

$dateFilter = "activityDateTime ge {0} and activityDateTime lt {1}" -f (ConvertTo-CsvTimestamp $start), (ConvertTo-CsvTimestamp $end)
$activityFilter = '(' + (($activities | ForEach-Object { "activityDisplayName eq '{0}'" -f ($_ -replace "'", "''") }) -join ' or ') + ')'

$audits = $null
try {
    $audits = @(Get-MgAuditLogDirectoryAudit -All -Filter "$dateFilter and $activityFilter" -ErrorAction Stop)
}
catch {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        'Filtering on activityDisplayName failed ({0}). Retrying with the date filter only and matching the activity names in the script.' -f $_.Exception.Message)
    try {
        $audits = @(Get-MgAuditLogDirectoryAudit -All -Filter $dateFilter -ErrorAction Stop |
                Where-Object { $activities -contains $_.ActivityDisplayName })
    }
    catch {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            'The Entra ID directory audit log is unavailable to this sign-in ({0}). It needs the AuditLog.Read.All permission and the Reports Reader role. Writing the header only.' -f $_.Exception.Message)
        Export-AppendCsv -Path $csvPath -Column $columns
        return
    }
}

$rows = foreach ($audit in $audits) {
    $initiatedByUpn = ''
    $initiatedByApp = ''
    $initiatedBy = $audit.PSObject.Properties['InitiatedBy']
    if ($initiatedBy -and $null -ne $initiatedBy.Value) {
        $user = Get-GraphAdditionalProperty -Object $initiatedBy.Value -Name 'user'
        if ($null -ne $user) {
            $initiatedByUpn = [string](Get-GraphAdditionalProperty -Object $user -Name 'userPrincipalName')
        }
        $app = Get-GraphAdditionalProperty -Object $initiatedBy.Value -Name 'app'
        if ($null -ne $app) {
            $initiatedByApp = [string](Get-GraphAdditionalProperty -Object $app -Name 'displayName')
        }
    }

    # The invited or redeeming account is the first target resource of type User.
    $targetId = ''
    $targetUpn = ''
    $targets = $audit.PSObject.Properties['TargetResources']
    if ($targets -and $null -ne $targets.Value) {
        foreach ($target in @($targets.Value)) {
            $type = [string](Get-GraphAdditionalProperty -Object $target -Name 'type')
            if ($type -and $type -ne 'User') { continue }
            $targetId = [string](Get-GraphAdditionalProperty -Object $target -Name 'id')
            $targetUpn = [string](Get-GraphAdditionalProperty -Object $target -Name 'userPrincipalName')
            break
        }
    }

    [pscustomobject]@{
        ActivityDateTime             = ConvertTo-CsvTimestamp $audit.ActivityDateTime
        Id                           = $audit.Id
        ActivityDisplayName          = $audit.ActivityDisplayName
        Result                       = [string]$audit.Result
        InitiatedByUserPrincipalName = $initiatedByUpn
        InitiatedByAppDisplayName    = $initiatedByApp
        TargetUserId                 = $targetId
        TargetUserPrincipalName      = $targetUpn
    }
}

$result = Export-AppendCsv -Path $csvPath -Rows @($rows) -Column $columns -KeyColumn 'Id' -PassThru
Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'guest-invitations.csv: {0} rows written, {1} skipped as already collected.' -f $result.Written, $result.Skipped)
