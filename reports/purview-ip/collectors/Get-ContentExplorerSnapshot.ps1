#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes content-explorer-snapshot.csv: item counts per tag per workload,
        appended each run.

    .DESCRIPTION
        Asks Export-ContentExplorerData for the aggregate item count of every
        sensitivity label, retention label and sensitive information type, in
        every workload, and writes one row per (tag, workload) stamped with the
        run date - so repeated runs give the report a trend rather than a
        single point.

        Export-ContentExplorerData returns an array whose first element carries
        TotalCount for the query, so the count is read from there rather than
        by paging through the item records.

        The tag names come from the tenant's own label and SIT definitions, so
        this collector reads them with Get-Label, Get-ComplianceTag and
        Get-DlpSensitiveInformationType first. Every cmdlet it calls is
        read-only.

    .PARAMETER Workload
        Workloads to snapshot. Defaults to all four Export-ContentExplorerData
        accepts. Teams is dropped automatically where it is documented as
        unavailable (see PurviewIpSchema.psd1's SourceAvailability.ContentExplorerTeams).

    .PARAMETER SkipConnect
        Use an existing Security & Compliance PowerShell session instead of
        signing in.

    .EXAMPLE
        ./Get-ContentExplorerSnapshot.ps1 -OutputPath ./out -Environment GCC

    .LINK
        https://learn.microsoft.com/powershell/module/exchangepowershell/export-contentexplorerdata
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

    [ValidateSet('EXO', 'ODB', 'SPO', 'Teams')]
    [string[]]$Workload = @('EXO', 'ODB', 'SPO', 'Teams'),

    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

. (Join-Path $PSScriptRoot 'PurviewIpHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'PurviewIpSchema.psd1')
$columns = $schema.ContentExplorerSnapshot
$source = 'content-explorer-snapshot'
$csvPath = Join-Path $OutputPath 'content-explorer-snapshot.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-PurviewSourceAvailability -Source 'ContentExplorer' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping content-explorer-snapshot.csv. $($availability.Reason) $($availability.Reference)")
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

try {
    $tags = [System.Collections.Generic.List[object]]::new()

    foreach ($label in @(Get-Label)) {
        $name = Get-PurviewProperty $label 'Name'
        if ($name) { $tags.Add([pscustomobject]@{ TagType = 'Sensitivity'; TagName = [string]$name }) }
    }
    foreach ($tag in @(Get-ComplianceTag)) {
        $name = Get-PurviewProperty $tag 'Name'
        if ($name) { $tags.Add([pscustomobject]@{ TagType = 'Retention'; TagName = [string]$name }) }
    }
    foreach ($sit in @(Get-DlpSensitiveInformationType)) {
        $name = Get-PurviewProperty $sit 'Name'
        if ($name) { $tags.Add([pscustomobject]@{ TagType = 'SensitiveInformationType'; TagName = [string]$name }) }
    }

    # Teams is not covered by Content Explorer in every cloud. Drop it where it
    # is documented as unavailable rather than recording a run of zero counts
    # that reads like the tenant has no labelled Teams content.
    $workloads = @($Workload)
    if ('Teams' -in $workloads) {
        $teams = Get-PurviewSourceAvailability -Source 'ContentExplorerTeams' -Environment $Environment -Schema $schema
        if ($teams.ShouldSkip) {
            Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
                "Leaving the Teams workload out. $($teams.Reason) $($teams.Reference)")
            $workloads = @($workloads | Where-Object { $_ -ne 'Teams' })
        }
    }

    Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
        "Snapshotting $($tags.Count) tag(s) across $($workloads.Count) workload(s).")

    $rows = [System.Collections.Generic.List[object]]::new()

    foreach ($tag in $tags) {
        foreach ($workloadName in $workloads) {
            $total = $null
            try {
                # Microsoft recommends one query per tag per workload rather
                # than one query covering many tags.
                $response = @(Export-ContentExplorerData -TagType $tag.TagType -TagName $tag.TagName -Workload $workloadName -PageSize 1)
                if ($response.Count -gt 0) {
                    $total = Get-PurviewProperty $response[0] 'TotalCount'
                }
            }
            catch {
                Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
                    "No count for $($tag.TagType)/$($tag.TagName) in $workloadName : $($_.Exception.Message)")
                continue
            }

            if ($null -eq $total) { continue }

            $rows.Add([pscustomobject]@{
                    RunDate    = $runDate
                    TagType    = $tag.TagType
                    TagName    = $tag.TagName
                    Workload   = $workloadName
                    # [long], not [int]: a tenant-wide count for a common
                    # sensitive information type can exceed 2^31, and an
                    # overflow here would abort the whole snapshot.
                    TotalCount = [long]$total
                })
        }
    }
}
catch {
    Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
        "Content Explorer is unavailable to this sign-in ({0}). Writing the header only." -f $_.Exception.Message)
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}
finally {
    if ($connectedHere) {
        Disconnect-ExchangeOnline -Confirm:$false
    }
}

$result = Export-AppendCsv -Path $csvPath -Rows $rows.ToArray() -Column $columns -KeyColumn @('RunDate', 'TagType', 'TagName', 'Workload') -PassThru
Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'content-explorer-snapshot.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
