#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes sharing-link-activity.csv: the sites that created Anyone, People in your
        organization and Specific people sharing links in the last 28 days, appended each run.

    .DESCRIPTION
        Source: SharePoint Online PowerShell
        Start-SPODataAccessGovernanceInsight -ReportEntity SharingLinks_Anyone,
        SharingLinks_PeopleInYourOrg and SharingLinks_Guests, each with -Workload SharePoint
        and OneDriveForBusiness and -ReportType RecentActivity
        (https://learn.microsoft.com/sharepoint/powershell-for-data-access-governance#generate-sharing-link-activity-reports).
        This parameter set takes no -Name.

        The report is over a rolling 28-day window, not a stream, so each export is appended
        with the run date as a state source is. A gap longer than 28 days between runs loses
        events; the audit log (anonymous-link-events.csv, sharing-events.csv) is the longer
        record. A report can be run again every 24 hours, so -MaxReportAgeHours defaults to 24.
        ReportStartTime and ReportEndTime are the period the report covers
        (Get-SPODataAccessGovernanceInsight).

        Without a SharePoint Advanced Management license, data collection must be on before a
        report can be generated: Get-SPOAuditDataCollectionStatusForActivityInsights must say
        InProgress, and the report holds only data from when collection began. This collector
        reads that status and logs it, and does not call
        Start-SPOAuditDataCollectionForActivityInsights, which changes the tenant. Which
        -ReportEntity string that cmdlet accepts (SharingLinksAnyone or SharingLinks_Anyone)
        is UNVERIFIED, so both are tried when reading the status. Activity reports for
        Microsoft 365 E5 without SharePoint Advanced Management return at most 10,000 sites
        (https://learn.microsoft.com/sharepoint/data-access-governance-reports).

        Learn does not list the columns of this CSV, so each exported row is kept whole as
        JSON in ReportRow, and SiteId and SiteUrl are copied out when the export has columns
        of that name. Role and licence as in Get-SitePermissionBreadth.ps1. The module needs
        Connect-SPOService without -Credential.

    .PARAMETER AdminUrl
        The SharePoint admin center URL. Required unless -SkipConnect is used.

    .EXAMPLE
        ./Get-SharingLinkActivity.ps1 -OutputPath ./out -AdminUrl https://contoso-admin.sharepoint.com
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
$columns = $schema.SharingLinkActivity
$source = 'sharing-link-activity'
$csvPath = Join-Path $OutputPath 'sharing-link-activity.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-OversharingSourceAvailability -Source 'SharingLinkActivity' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping sharing-link-activity.csv. $($availability.Reason) $($availability.Reference)")
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

    foreach ($entity in 'SharingLinks_Anyone', 'SharingLinks_PeopleInYourOrg', 'SharingLinks_Guests') {
        Test-ActivityDataCollection -ReportEntity $entity -OutputPath $OutputPath -Source $source

        foreach ($workload in 'SharePoint', 'OneDriveForBusiness') {
            try {
                $report = Invoke-DagReport -ReportEntity $entity -ReportType 'RecentActivity' -Workload $workload `
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

            $rows = foreach ($row in $report.Rows) {
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
            }

            $result = Export-AppendCsv -Path $csvPath -Rows @($rows) -Column $columns -KeyColumn @('ReportId', 'ReportEntity', 'Workload', 'ReportRow') -PassThru
            Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
                'sharing-link-activity.csv ({0}, {1}): {2} rows written, {3} skipped.' -f $entity, $workload, $result.Written, $result.Skipped)
        }
    }

    Export-AppendCsv -Path $csvPath -Column $columns

    if ($exported -eq 0 -and $pending -eq 0) {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            'No sharing link activity report could be read. Writing the header only.')
    }
}
finally {
    if ($connectedHere) {
        Disconnect-SPOService
    }
}
