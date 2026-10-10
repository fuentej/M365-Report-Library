#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes eeeu-activity.csv: the sites and items newly shared with "Everyone except
        external users" (EEEU) in the last 28 days, appended each run.

    .DESCRIPTION
        Source: SharePoint Online PowerShell
        Start-SPODataAccessGovernanceInsight -ReportEntity EveryoneExceptExternalUsersAtSite
        and EveryoneExceptExternalUsersForItems with -Workload SharePoint or
        OneDriveForBusiness, -ReportType RecentActivity and -Name <name>
        (https://learn.microsoft.com/sharepoint/powershell-for-data-access-governance#identify-content-shared-with-everyone-except-external-users-in-last-28-days).
        The EEEU parameter set marks -Name and -Workload mandatory
        (https://learn.microsoft.com/powershell/module/microsoft.online.sharepoint.powershell/start-spodataaccessgovernanceinsight).
        OneDrive supports the item level only, so the site-level report is asked for
        SharePoint alone.

        Windowed, licensed and rolled up as in Get-SharingLinkActivity.ps1: a rolling 28-day
        window appended with the run date, data collection that must be InProgress without a
        SharePoint Advanced Management license (read and logged, never started), and ReportRow
        holding the whole exported row because Learn does not list this CSV's columns.
        Which -ReportEntity string Get-SPOAuditDataCollectionStatusForActivityInsights accepts
        is UNVERIFIED.

    .PARAMETER AdminUrl
        The SharePoint admin center URL. Required unless -SkipConnect is used.

    .EXAMPLE
        ./Get-EeeuActivity.ps1 -OutputPath ./out -AdminUrl https://contoso-admin.sharepoint.com
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

    [string]$AdminUrl,

    [ValidateRange(0, [int]::MaxValue)]
    [int]$MaxReportAgeHours = 24,

    [ValidateRange(0, [int]::MaxValue)]
    [int]$WaitMinutes = 30,

    [ValidateRange(0, [int]::MaxValue)]
    [int]$PollSeconds = 60,

    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Imported by path without -Force: a forced re-import builds a fresh module session
# state, which would discard anything the caller has arranged around the module - and
# that is exactly how the tests put mocks in front of the tenant cmdlets.
Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

. (Join-Path $PSScriptRoot 'OversharingHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'OversharingSchema.psd1')
$columns = $schema.EeeuActivity
$source = 'eeeu-activity'
$csvPath = Join-Path $OutputPath 'eeeu-activity.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-OversharingSourceAvailability -Source 'EeeuActivity' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping eeeu-activity.csv. $($availability.Reason) $($availability.Reference)")
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}
if ($availability.Status -eq 'Unverified') {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message $availability.Reason
}

$connectedHere = -not $SkipConnect
if ($connectedHere) {
    if ([string]::IsNullOrWhiteSpace($AdminUrl)) {
        throw '-AdminUrl is required to sign in to SharePoint Online. Pass the SharePoint admin center URL, or -SkipConnect to use an existing session.'
    }
    Connect-SharePointAdmin -AdminUrl $AdminUrl -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -OutputPath $OutputPath -Source $source
}

try {
    $exported = 0
    $pending = 0

    $reports = @(
        @{ Entity = 'EveryoneExceptExternalUsersAtSite'; Workload = 'SharePoint' }
        @{ Entity = 'EveryoneExceptExternalUsersForItems'; Workload = 'SharePoint' }
        @{ Entity = 'EveryoneExceptExternalUsersForItems'; Workload = 'OneDriveForBusiness' }
    )

    foreach ($entity in @($reports | ForEach-Object { $_.Entity } | Select-Object -Unique)) {
        Test-ActivityDataCollection -ReportEntity $entity -OutputPath $OutputPath -Source $source
    }

    foreach ($wanted in $reports) {
        $entity = $wanted.Entity
        $workload = $wanted.Workload

        try {
            $report = Invoke-DagReport -ReportEntity $entity -ReportType 'RecentActivity' -Workload $workload `
                -StartArgument @{ Name = ('Oversharing-{0}-{1}-{2}' -f $entity, $workload, $runDate) } `
                -MaxReportAgeHours $MaxReportAgeHours -WaitMinutes $WaitMinutes -PollSeconds $PollSeconds `
                -OutputPath $OutputPath -Source $source
        }
        catch {
            Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
                'The {0} ({1}) report is unavailable to this sign-in ({2}). It needs the SharePoint Administrator role and either a SharePoint Advanced Management license or data collection turned on.' -f $entity, $workload, $_.Exception.Message)
            continue
        }

        if ($report.State -eq 'Pending') { $pending++ }
        if ($report.State -ne 'Exported') { continue }
        $exported++

        $start = ConvertTo-CsvTimestamp (Get-JsonProperty -Object $report.Report -Name 'ReportStartTime')
        $end = ConvertTo-CsvTimestamp (Get-JsonProperty -Object $report.Report -Name 'ReportEndTime')

        $rows = @(foreach ($row in $report.Rows) {
            [pscustomobject]@{
                RunDate         = $runDate
                ReportEntity    = $entity
                Workload        = $workload
                ReportId        = $report.ReportId
                ReportStartTime = $start
                ReportEndTime   = $end
                SiteId          = Get-CsvField -Row $row -Name 'Site ID', 'SiteId'
                SiteUrl         = Get-CsvField -Row $row -Name 'Site URL', 'SiteUrl'
                ReportRow       = ConvertTo-ReportRowJson -Row $row
            }
        })

        Write-DagExportCap -Count $rows.Count -Cap 10000 -OutputPath $OutputPath -Source $source `
            -ReportName "$entity ($workload)" `
            -Reference 'https://learn.microsoft.com/sharepoint/data-access-governance-reports'

        $result = Export-AppendCsv -Path $csvPath -Rows $rows -Column $columns -KeyColumn @('ReportId', 'ReportEntity', 'Workload', 'ReportRow') -PassThru
        Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
            'eeeu-activity.csv ({0}, {1}): {2} rows written, {3} skipped.' -f $entity, $workload, $result.Written, $result.Skipped)
    }

    Export-AppendCsv -Path $csvPath -Column $columns

    if ($exported -eq 0 -and $pending -eq 0) {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            'No EEEU activity report could be read. Writing the header only.')
    }
}
finally {
    if ($connectedHere) {
        Disconnect-SPOService
    }
}
