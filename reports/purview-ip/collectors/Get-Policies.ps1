#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes policies.csv: a snapshot of the information protection policy
        configuration, appended each run.

    .DESCRIPTION
        Source: Security & Compliance PowerShell, one table with an ObjectType
        discriminator, collecting:

            SensitivityLabel    Get-Label
            LabelPolicy         Get-LabelPolicy
            AutoLabelingPolicy  Get-AutoSensitivityLabelPolicy
            DlpPolicy           Get-DlpCompliancePolicy
            DlpRule             Get-DlpComplianceRule
            RetentionPolicy     Get-RetentionCompliancePolicy
            RetentionLabel      Get-ComplianceTag

        One table keeps the report's policy dimension to a single file; filter
        on ObjectType to get back any one of the seven collections. Columns
        that do not apply to a row are empty.

        AppliesToCopilot marks a DLP policy scoped to Microsoft 365 Copilot or
        Copilot Chat, detected from the CopilotExperiences enforcement plane or
        the Copilot location GUID in the policy's Locations JSON.

        Configuration is a snapshot of current state, keyed on RunDate plus
        ObjectType and ObjectId, so running twice in one day does not
        duplicate a snapshot.

    .PARAMETER SkipConnect
        Use an existing Security & Compliance PowerShell session instead of
        signing in.

    .EXAMPLE
        ./Get-Policies.ps1 -OutputPath ./out -Environment GCCHigh
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

. (Join-Path $PSScriptRoot 'PurviewIpHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'PurviewIpSchema.psd1')
$columns = $schema.Policies
$source = 'policies'
$csvPath = Join-Path $OutputPath 'policies.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-PurviewSourceAvailability -Source 'Policies' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping policies.csv. $($availability.Reason) $($availability.Reference)")
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}
if ($availability.Status -eq 'Unverified') {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message $availability.Reason
}

$connectedHere = -not $SkipConnect
if ($connectedHere) {
    Connect-M365Service -Service SecurityCompliance -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization
}

function New-PolicyRow {
    param([Parameter(Mandatory)][string]$ObjectType, [Parameter(Mandatory)][hashtable]$Values)

    $row = @{
        RunDate    = $runDate
        ObjectType = $ObjectType
    }
    foreach ($entry in $Values.GetEnumerator()) { $row[$entry.Key] = $entry.Value }
    [pscustomobject]$row
}

$rows = [System.Collections.Generic.List[object]]::new()

try {
    # --- Sensitivity labels -------------------------------------------------
    $labels = @(Get-Label)
    $labelNameById = @{}
    foreach ($label in $labels) {
        $guid = [string](Get-PurviewProperty $label 'Guid' '')
        if ($guid) { $labelNameById[$guid] = [string](Get-PurviewProperty $label 'Name' '') }
    }

    foreach ($label in $labels) {
        $parentId = [string](Get-PurviewProperty $label 'ParentId' '')
        $encryption = Get-PurviewProperty $label 'EncryptionEnabled'
        if ($null -eq $encryption) {
            # Older tenants expose encryption through the Settings collection only.
            $settings = ConvertTo-PurviewListValue (Get-PurviewProperty $label 'Settings')
            if ($settings) { $encryption = [bool]($settings -match 'encrypt') }
        }

        $rows.Add((New-PolicyRow -ObjectType 'SensitivityLabel' -Values @{
                    ObjectId          = Get-PurviewProperty $label 'Guid'
                    Name              = Get-PurviewProperty $label 'Name'
                    DisplayName       = Get-PurviewProperty $label 'DisplayName'
                    ParentId          = $parentId
                    ParentName        = if ($parentId -and $labelNameById.ContainsKey($parentId)) { $labelNameById[$parentId] } else { '' }
                    Priority          = Get-PurviewProperty $label 'Priority'
                    Enabled           = -not [bool](Get-PurviewProperty $label 'Disabled' $false)
                    Workload          = Get-PurviewProperty $label 'Workload'
                    IsDefaultLabel    = Get-PurviewProperty $label 'IsDefault'
                    EncryptionEnabled = $encryption
                    Comment           = Get-PurviewProperty $label 'Comment'
                    WhenCreatedUtc    = ConvertTo-CsvTimestamp (Get-PurviewProperty $label 'WhenCreatedUTC' (Get-PurviewProperty $label 'WhenCreated'))
                    WhenChangedUtc    = ConvertTo-CsvTimestamp (Get-PurviewProperty $label 'WhenChangedUTC' (Get-PurviewProperty $label 'WhenChanged'))
                }))
    }

    # --- Label policies -----------------------------------------------------
    foreach ($policy in @(Get-LabelPolicy)) {
        $labelIds = @(Get-PurviewProperty $policy 'Labels' @())
        $rows.Add((New-PolicyRow -ObjectType 'LabelPolicy' -Values @{
                    ObjectId       = Get-PurviewProperty $policy 'Guid'
                    Name           = Get-PurviewProperty $policy 'Name'
                    Mode           = Get-PurviewProperty $policy 'Mode'
                    Enabled        = Get-PurviewProperty $policy 'Enabled'
                    Workload       = Get-PurviewProperty $policy 'Workload'
                    # A label policy can publish to SharePoint, OneDrive, public
                    # folders and Skype as well; leaving those out made a
                    # SharePoint-only policy look unscoped.
                    Locations      = ConvertTo-PurviewListValue @(
                        Get-PurviewProperty $policy 'ExchangeLocation' @()
                        Get-PurviewProperty $policy 'ModernGroupLocation' @()
                        Get-PurviewProperty $policy 'SharePointLocation' @()
                        Get-PurviewProperty $policy 'OneDriveLocation' @()
                        Get-PurviewProperty $policy 'PublicFolderLocation' @()
                        Get-PurviewProperty $policy 'SkypeLocation' @()
                    )
                    LabelIds       = ConvertTo-PurviewListValue $labelIds
                    LabelNames     = ConvertTo-PurviewListValue @($labelIds | ForEach-Object {
                            $id = [string]$_
                            if ($labelNameById.ContainsKey($id)) { $labelNameById[$id] } else { $id }
                        })
                    Comment        = Get-PurviewProperty $policy 'Comment'
                    WhenCreatedUtc = ConvertTo-CsvTimestamp (Get-PurviewProperty $policy 'WhenCreatedUTC' (Get-PurviewProperty $policy 'WhenCreated'))
                    WhenChangedUtc = ConvertTo-CsvTimestamp (Get-PurviewProperty $policy 'WhenChangedUTC' (Get-PurviewProperty $policy 'WhenChanged'))
                }))
    }

    # --- Auto-labeling policies ---------------------------------------------
    foreach ($policy in @(Get-AutoSensitivityLabelPolicy)) {
        $labelId = [string](Get-PurviewProperty $policy 'ApplySensitivityLabel' '')
        $rows.Add((New-PolicyRow -ObjectType 'AutoLabelingPolicy' -Values @{
                    ObjectId       = Get-PurviewProperty $policy 'Guid'
                    Name           = Get-PurviewProperty $policy 'Name'
                    Priority       = Get-PurviewProperty $policy 'Priority'
                    Mode           = Get-PurviewProperty $policy 'Mode'
                    Enabled        = Get-PurviewProperty $policy 'Enabled'
                    Locations      = ConvertTo-PurviewListValue @(
                        Get-PurviewProperty $policy 'ExchangeLocation' @()
                        Get-PurviewProperty $policy 'SharePointLocation' @()
                        Get-PurviewProperty $policy 'OneDriveLocation' @()
                        Get-PurviewProperty $policy 'ModernGroupLocation' @()
                        Get-PurviewProperty $policy 'TeamsLocation' @()
                    )
                    LabelIds       = $labelId
                    LabelNames     = if ($labelId -and $labelNameById.ContainsKey($labelId)) { $labelNameById[$labelId] } else { $labelId }
                    Comment        = Get-PurviewProperty $policy 'Comment'
                    WhenCreatedUtc = ConvertTo-CsvTimestamp (Get-PurviewProperty $policy 'WhenCreatedUTC' (Get-PurviewProperty $policy 'WhenCreated'))
                    WhenChangedUtc = ConvertTo-CsvTimestamp (Get-PurviewProperty $policy 'WhenChangedUTC' (Get-PurviewProperty $policy 'WhenChanged'))
                }))
    }

    # --- DLP policies -------------------------------------------------------
    $dlpPolicies = @(Get-DlpCompliancePolicy)
    $dlpPolicyNameById = @{}

    foreach ($policy in $dlpPolicies) {
        $guid = [string](Get-PurviewProperty $policy 'Guid' '')
        $name = [string](Get-PurviewProperty $policy 'Name' '')
        if ($guid) { $dlpPolicyNameById[$guid] = $name }

        $enforcementPlanes = Get-PurviewProperty $policy 'EnforcementPlanes'
        $locations = Get-PurviewProperty $policy 'Locations'
        $flattenedLocations = @(
            Get-PurviewProperty $policy 'ExchangeLocation' @()
            Get-PurviewProperty $policy 'SharePointLocation' @()
            Get-PurviewProperty $policy 'OneDriveLocation' @()
            Get-PurviewProperty $policy 'TeamsLocation' @()
            Get-PurviewProperty $policy 'EndpointDlpLocation' @()
            Get-PurviewProperty $policy 'PowerBIDlpLocation' @()
            Get-PurviewProperty $policy 'ThirdPartyAppDlpLocation' @()
            $locations
        )

        $rows.Add((New-PolicyRow -ObjectType 'DlpPolicy' -Values @{
                    ObjectId         = $guid
                    Name             = $name
                    Priority         = Get-PurviewProperty $policy 'Priority'
                    Mode             = Get-PurviewProperty $policy 'Mode'
                    Enabled          = Get-PurviewProperty $policy 'Enabled'
                    Locations        = ConvertTo-PurviewListValue $flattenedLocations
                    AppliesToCopilot = Get-PurviewCopilotScope -EnforcementPlanes $enforcementPlanes -Locations $flattenedLocations -Schema $schema
                    Comment          = Get-PurviewProperty $policy 'Comment'
                    WhenCreatedUtc   = ConvertTo-CsvTimestamp (Get-PurviewProperty $policy 'WhenCreatedUTC' (Get-PurviewProperty $policy 'WhenCreated'))
                    WhenChangedUtc   = ConvertTo-CsvTimestamp (Get-PurviewProperty $policy 'WhenChangedUTC' (Get-PurviewProperty $policy 'WhenChanged'))
                }))
    }

    # --- DLP rules ----------------------------------------------------------
    foreach ($rule in @(Get-DlpComplianceRule)) {
        $parentId = [string](Get-PurviewProperty $rule 'ParentPolicyId' '')
        $parentName = [string](Get-PurviewProperty $rule 'Policy' '')
        if (-not $parentName -and $parentId -and $dlpPolicyNameById.ContainsKey($parentId)) {
            $parentName = $dlpPolicyNameById[$parentId]
        }

        # Handles both the flat and the grouped condition shapes, and picks up a
        # sensitivity-label condition (a Copilot rule often has no SIT at all).
        $sitNames = @(Get-PurviewSensitiveInformationTypeName (Get-PurviewProperty $rule 'ContentContainsSensitiveInformation'))
        if ($sitNames.Count -eq 0) {
            $sitNames = @(Get-PurviewSensitiveInformationTypeName (Get-PurviewProperty $rule 'AdvancedRule'))
        }

        $rows.Add((New-PolicyRow -ObjectType 'DlpRule' -Values @{
                    ObjectId                  = Get-PurviewProperty $rule 'Guid'
                    Name                      = Get-PurviewProperty $rule 'Name'
                    ParentId                  = $parentId
                    ParentName                = $parentName
                    Priority                  = Get-PurviewProperty $rule 'Priority'
                    Enabled                   = -not [bool](Get-PurviewProperty $rule 'Disabled' $false)
                    SensitiveInformationTypes = ConvertTo-PurviewListValue $sitNames
                    BlockAccess               = Get-PurviewProperty $rule 'BlockAccess'
                    BlockAccessScope          = Get-PurviewProperty $rule 'BlockAccessScope'
                    NotifyUser                = Get-PurviewProperty $rule 'NotifyUser'
                    AllowOverride             = Get-PurviewProperty $rule 'NotifyAllowOverride'
                    RequireJustification      = ((ConvertTo-PurviewListValue (Get-PurviewProperty $rule 'NotifyAllowOverride')) -match 'WithJustification')
                    GenerateIncidentReport    = Get-PurviewProperty $rule 'GenerateIncidentReport'
                    Comment                   = Get-PurviewProperty $rule 'Comment'
                    WhenCreatedUtc            = ConvertTo-CsvTimestamp (Get-PurviewProperty $rule 'WhenCreatedUTC' (Get-PurviewProperty $rule 'WhenCreated'))
                    WhenChangedUtc            = ConvertTo-CsvTimestamp (Get-PurviewProperty $rule 'WhenChangedUTC' (Get-PurviewProperty $rule 'WhenChanged'))
                }))
    }

    # --- Retention policies -------------------------------------------------
    foreach ($policy in @(Get-RetentionCompliancePolicy)) {
        $rows.Add((New-PolicyRow -ObjectType 'RetentionPolicy' -Values @{
                    ObjectId       = Get-PurviewProperty $policy 'Guid'
                    Name           = Get-PurviewProperty $policy 'Name'
                    Mode           = Get-PurviewProperty $policy 'Mode'
                    Enabled        = Get-PurviewProperty $policy 'Enabled'
                    Workload       = Get-PurviewProperty $policy 'Workload'
                    Locations      = ConvertTo-PurviewListValue @(
                        Get-PurviewProperty $policy 'ExchangeLocation' @()
                        Get-PurviewProperty $policy 'SharePointLocation' @()
                        Get-PurviewProperty $policy 'OneDriveLocation' @()
                        Get-PurviewProperty $policy 'TeamsChatLocation' @()
                        Get-PurviewProperty $policy 'TeamsChannelLocation' @()
                        Get-PurviewProperty $policy 'ModernGroupLocation' @()
                    )
                    Comment        = Get-PurviewProperty $policy 'Comment'
                    WhenCreatedUtc = ConvertTo-CsvTimestamp (Get-PurviewProperty $policy 'WhenCreatedUTC' (Get-PurviewProperty $policy 'WhenCreated'))
                    WhenChangedUtc = ConvertTo-CsvTimestamp (Get-PurviewProperty $policy 'WhenChangedUTC' (Get-PurviewProperty $policy 'WhenChanged'))
                }))
    }

    # --- Retention labels ---------------------------------------------------
    foreach ($tag in @(Get-ComplianceTag)) {
        $rows.Add((New-PolicyRow -ObjectType 'RetentionLabel' -Values @{
                    ObjectId          = Get-PurviewProperty $tag 'Guid'
                    Name              = Get-PurviewProperty $tag 'Name'
                    DisplayName       = Get-PurviewProperty $tag 'DisplayName'
                    Priority          = Get-PurviewProperty $tag 'Priority'
                    RetentionAction   = Get-PurviewProperty $tag 'RetentionAction'
                    RetentionDuration = Get-PurviewProperty $tag 'RetentionDuration'
                    RetentionType     = Get-PurviewProperty $tag 'RetentionType'
                    IsRecordLabel     = Get-PurviewProperty $tag 'IsRecordLabel'
                    Comment           = Get-PurviewProperty $tag 'Comment'
                    WhenCreatedUtc    = ConvertTo-CsvTimestamp (Get-PurviewProperty $tag 'WhenCreatedUTC' (Get-PurviewProperty $tag 'WhenCreated'))
                    WhenChangedUtc    = ConvertTo-CsvTimestamp (Get-PurviewProperty $tag 'WhenChangedUTC' (Get-PurviewProperty $tag 'WhenChanged'))
                }))
    }
}
catch {
    Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
        "Security & Compliance PowerShell policy configuration is unavailable to this sign-in ({0}). Writing the header only." -f $_.Exception.Message)
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}
finally {
    if ($connectedHere) {
        Disconnect-ExchangeOnline -Confirm:$false
    }
}

$result = Export-AppendCsv -Path $csvPath -Rows $rows.ToArray() -Column $columns -KeyColumn @('RunDate', 'ObjectType', 'ObjectId') -PassThru
Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'policies.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
