#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes copilot-feature-availability.csv: which Copilot features each cloud offers, as a
        dated copy of a public Microsoft page. A snapshot stamped with the run date. Makes no
        tenant call.

    .DESCRIPTION
        Source 10 of docs/candidates/copilot-usage.md: the feature availability table of the
        Microsoft 365 Copilot service description
        (https://learn.microsoft.com/office365/servicedescriptions/office-365-platform-service-description/microsoft-365-copilot#feature-availability).
        The page changes, so the copy is dated: PageReadDate is the day the contract read it,
        and the rows in CopilotUsageSchema.psd1 are the ones the contract records, not the
        whole table. NotStated means the contract does not say, not that a feature is absent.
        Refreshing the copy means re-reading the page and editing FeatureRows.

        It decides which app columns can ever be non-empty in a cloud: a GCC High tenant with no
        Teams or SharePoint Copilot activity is expected, not a collection gap.

        No sign-in happens. -Environment only decides whether the run is recorded as skipped.

    .EXAMPLE
        ./Get-CopilotFeatureAvailability.ps1 -OutputPath ./out
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$OutputPath,

    [ValidateSet('Commercial', 'GCC', 'GCCHigh')]
    [string]$Environment = 'Commercial',

    # Accepted so Run-All.ps1 can pass the same sign-in arguments to every collector;
    # this collector does not sign in.
    [string]$AppId,
    [string]$CertificateThumbprint,
    [string]$TenantId,
    [string]$Organization,

    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Imported by path without -Force: see Get-CopilotUsageUserDetail.ps1.
Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

. (Join-Path $PSScriptRoot 'CopilotUsageHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'CopilotUsageSchema.psd1')
$columns = [string[]]$schema.CopilotFeatureAvailability
$source = 'copilot-feature-availability'
$csvPath = Join-Path $OutputPath 'copilot-feature-availability.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

if (Test-CopilotSourceSkipped -Source 'CopilotFeatureAvailability' -LogSource $source -CsvPath $csvPath -Column $columns `
        -OutputPath $OutputPath -Environment $Environment -Schema $schema) {
    return
}

$rows = foreach ($feature in $schema.FeatureRows) {
    [pscustomobject]@{
        RunDate      = $runDate
        Feature      = $feature.Feature
        Commercial   = $feature.Commercial
        GCC          = $feature.GCC
        GCCHigh      = $feature.GCCHigh
        Note         = $feature.Note
        PageReadDate = $schema.FeaturePageReadDate
        Reference    = $schema.FeaturePage
    }
}

$result = Export-AppendCsv -Path $csvPath -Rows @($rows) -Column $columns -KeyColumn 'RunDate', 'Feature' -PassThru
Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'copilot-feature-availability.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
