#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes directory-audits.csv: Entra ID directory audit events - who changed users,
        groups, roles, applications and policies, when, and the result. Appended from the
        last exported timestamp.

    .DESCRIPTION
        Source 4 of docs/candidates/entra-activity.md: Microsoft Graph list directoryAudits
        (https://learn.microsoft.com/graph/api/directoryaudit-list), available in all three
        clouds. The page does not state a page size, so -All follows @odata.nextLink until
        it is absent. Each window is filtered on activityDateTime (UTC).

        Every category (Conditional Access policy changes are category Policy),
        activityDisplayName, operationType and targetResources.type is kept as returned;
        the documented lists are examples and the reference may lag the service. Ids that
        are not GUIDs are kept.

        Retention is 7 days (Free) or 30 days (P1, P2) and is UNVERIFIED in GCC and GCC
        High. Needs AuditLog.Read.All and the Reports Reader, Security Administrator or
        Security Reader role. The licensing page lists audit logs as available on Free.

    .EXAMPLE
        ./Get-DirectoryAudits.ps1 -OutputPath ./out -LookbackDays 7
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

    # An alternate schema file. The tests use it to exercise the NotAvailable path.
    [string]$SchemaPath = (Join-Path $PSScriptRoot 'EntraActivitySchema.psd1'),

    [datetime]$StartDate,
    [datetime]$EndDate,

    # The audit log holds 7 days (Free) or 30 days (P1, P2) at most.
    [ValidateRange(1, 30)]
    [int]$LookbackDays = 30,

    [ValidateRange(1, 24)]
    [int]$WindowHours = 24,

    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Imported by path without -Force: see Get-InteractiveSignIns.ps1.
Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

. (Join-Path $PSScriptRoot 'EntraActivityHelpers.ps1')

$range = @{}
if ($PSBoundParameters.ContainsKey('StartDate')) { $range['StartDate'] = $StartDate }
if ($PSBoundParameters.ContainsKey('EndDate')) { $range['EndDate'] = $EndDate }

Invoke-EntraActivityEventCollector -Source DirectoryAudits -CsvName 'directory-audits.csv' `
    -OutputPath $OutputPath -Environment $Environment -AppId $AppId `
    -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization `
    -SchemaPath $SchemaPath -LookbackDays $LookbackDays -WindowHours $WindowHours -SkipConnect:$SkipConnect `
    -KeyColumn 'Id' -WatermarkColumn 'ActivityDateTime' -Description 'the directory audit log' `
    -License 'the AuditLog.Read.All permission and the Reports Reader, Security Administrator or Security Reader role' `
    -Fetch {
        param($from, $to)
        Get-MgAuditLogDirectoryAudit -All -Filter "activityDateTime ge $from and activityDateTime lt $to" -ErrorAction Stop
    } `
    -Map {
        param($audit)

        $initiatedBy = $audit.PSObject.Properties['InitiatedBy']
        $initiator = if ($initiatedBy) { $initiatedBy.Value } else { $null }
        $user = if ($initiator) { Get-GraphAdditionalProperty -Object $initiator -Name 'user' } else { $null }
        $app = if ($initiator) { Get-GraphAdditionalProperty -Object $initiator -Name 'app' } else { $null }

        $targetsProperty = $audit.PSObject.Properties['TargetResources']
        $targets = if ($targetsProperty -and $null -ne $targetsProperty.Value) { @($targetsProperty.Value) } else { @() }
        $modified = @($targets | ForEach-Object {
                $props = Get-GraphAdditionalProperty -Object $_ -Name 'modifiedProperties'
                foreach ($p in @($props)) {
                    if ($null -eq $p) { continue }
                    [ordered]@{
                        displayName = Get-PropertyValue $p 'displayName'
                        oldValue    = Get-PropertyValue $p 'oldValue'
                        newValue    = Get-PropertyValue $p 'newValue'
                    }
                }
            })

        [pscustomobject]@{
            ActivityDateTime                 = ConvertTo-CsvTimestamp $audit.ActivityDateTime
            Id                               = $audit.Id
            ActivityDisplayName              = $audit.ActivityDisplayName
            Category                         = $audit.Category
            OperationType                    = $audit.OperationType
            Result                           = [string]$audit.Result
            ResultReason                     = $audit.ResultReason
            LoggedByService                  = $audit.LoggedByService
            CorrelationId                    = $audit.CorrelationId
            InitiatedByUserId                = Get-PropertyValue $user 'id'
            InitiatedByUserPrincipalName     = Get-PropertyValue $user 'userPrincipalName'
            InitiatedByAppId                 = Get-PropertyValue $app 'appId'
            InitiatedByAppDisplayName        = Get-PropertyValue $app 'displayName'
            TargetResourceTypes              = Join-ListValue ($targets | ForEach-Object { Get-PropertyValue $_ 'type' })
            TargetResourceIds                = Join-ListValue ($targets | ForEach-Object { Get-PropertyValue $_ 'id' })
            TargetResourceDisplayNames       = Join-ListValue ($targets | ForEach-Object { Get-PropertyValue $_ 'displayName' })
            TargetResourceUserPrincipalNames = Join-ListValue ($targets | ForEach-Object { Get-PropertyValue $_ 'userPrincipalName' })
            ModifiedProperties               = if ($modified.Count -gt 0) { ConvertTo-Json -InputObject $modified -Compress -Depth 4 } else { '' }
        }
    } `
    @range
