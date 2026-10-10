#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes labeled-file-sites.csv: the sites that hold files with a given sensitivity
        label, one report per label, appended each run.

    .DESCRIPTION
        Source: SharePoint Online PowerShell
        Start-SPODataAccessGovernanceInsight -ReportEntity SensitivityLabelForFiles
        -Workload SharePoint -ReportType Snapshot -FileSensitivityLabelGUID <guid>, with
        -FileSensitivityLabelName <name> when the name is known
        (https://learn.microsoft.com/sharepoint/powershell-for-data-access-governance#generate-a-sensitivity-label-in-files-report-with-powershell).
        One GUID per run. OneDrive accounts with labeled files are not supported, so the
        report is asked for SharePoint alone.

        -LabelGuid names the labels. Without it the collector lists them with Get-Label in
        Security & Compliance PowerShell, keeping the labels whose ContentType includes File,
        because a report can only be made for a label with file scope
        (https://learn.microsoft.com/powershell/module/exchangepowershell/get-label). Get-Label
        says permissions are required and names no role, so the role that runs it is UNVERIFIED.

        A snapshot report. Joined to site-permission-breadth.csv on site it answers which sites
        hold a chosen label and are also open to many users; sensitivity label coverage
        itself is the Purview information protection report's job. Role and licence as in
        Get-SitePermissionBreadth.ps1, and the feature table adds "Requires E5 or G5" for the
        sensitivity labels row. A report can be run again every 24 hours, so
        -MaxReportAgeHours defaults to 24. A listed report is reused only when its listing
        carries the label GUID; Learn's example output does not, so a new report is normally
        started per label per run.

        Learn does not list the columns of this CSV, so ReportRow holds the whole exported row
        as JSON, and SiteId and SiteUrl are copied out when the export has columns of that
        name.

    .PARAMETER AdminUrl
        The SharePoint admin center URL. Required unless -SkipConnect is used.

    .PARAMETER LabelGuid
        The sensitivity label GUIDs to report on.

    .EXAMPLE
        ./Get-LabeledFileSites.ps1 -OutputPath ./out -AdminUrl https://contoso-admin.sharepoint.com `
            -LabelGuid 11111111-2222-3333-4444-555555555555
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

    [guid[]]$LabelGuid,

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
$columns = $schema.LabeledFileSites
$source = 'labeled-file-sites'
$csvPath = Join-Path $OutputPath 'labeled-file-sites.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-OversharingSourceAvailability -Source 'LabeledFileSites' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping labeled-file-sites.csv. $($availability.Reason) $($availability.Reference)")
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}
if ($availability.Status -eq 'Unverified') {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message $availability.Reason
}

if (-not $SkipConnect -and [string]::IsNullOrWhiteSpace($AdminUrl)) {
    throw '-AdminUrl is required to sign in to SharePoint Online. Pass the SharePoint admin center URL, or -SkipConnect to use an existing session.'
}

$labels = @()
if ($LabelGuid) {
    $labels = @($LabelGuid | ForEach-Object { [pscustomobject]@{ Guid = $_.ToString(); Name = '' } })
}
else {
    $complianceHere = -not $SkipConnect
    try {
        if ($complianceHere) {
            Connect-M365Service -Service SecurityCompliance -Environment $Environment -AppId $AppId `
                -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization
        }
        $labels = @(Get-Label -ErrorAction Stop | Where-Object { (Join-ListValue $_.ContentType) -match 'File' } | ForEach-Object {
                [pscustomobject]@{ Guid = ([string]$_.Guid).ToLowerInvariant(); Name = [string]$_.DisplayName }
            })
    }
    catch {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            'Get-Label is unavailable to this sign-in ({0}). The role that runs it is UNVERIFIED; pass -LabelGuid to name the labels instead. Writing the header only.' -f $_.Exception.Message)
        Export-AppendCsv -Path $csvPath -Column $columns
        return
    }
    finally {
        if ($complianceHere) { Disconnect-ExchangeOnline -Confirm:$false }
    }
}

if ($labels.Count -eq 0) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        'No sensitivity label with file scope was found, so no report was started.')
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}

$connectedHere = -not $SkipConnect
if ($connectedHere) {
    Connect-SharePointAdmin -AdminUrl $AdminUrl -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -OutputPath $OutputPath -Source $source
}

try {
    $exported = 0
    $pending = 0

    foreach ($label in $labels) {
        $start = @{ FileSensitivityLabelGUID = [guid]$label.Guid }
        if ($label.Name) { $start['FileSensitivityLabelName'] = $label.Name }

        try {
            $report = Invoke-DagReport -ReportEntity 'SensitivityLabelForFiles' -ReportType 'Snapshot' -Workload 'SharePoint' `
                -StartArgument $start -MatchProperty @{ FileSensitivityLabelGUID = $label.Guid } `
                -MaxReportAgeHours $MaxReportAgeHours -WaitMinutes $WaitMinutes -PollSeconds $PollSeconds `
                -OutputPath $OutputPath -Source $source
        }
        catch {
            Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
                'The sensitivity label report for {0} is unavailable to this sign-in ({1}). It needs the SharePoint Administrator role and a SharePoint Advanced Management license.' -f $label.Guid, $_.Exception.Message)
            continue
        }

        if ($report.State -eq 'Pending') { $pending++ }
        if ($report.State -ne 'Exported') { continue }
        $exported++

        $created = ConvertTo-CsvTimestamp (Get-JsonProperty -Object $report.Report -Name 'CreatedDateTime')

        $rows = foreach ($row in $report.Rows) {
            [pscustomobject]@{
                RunDate               = $runDate
                LabelGuid             = $label.Guid
                LabelName             = $label.Name
                Workload              = 'SharePoint'
                ReportId              = $report.ReportId
                ReportCreatedDateTime = $created
                SiteId                = Get-CsvField -Row $row -Name 'Site ID', 'SiteId'
                SiteUrl               = Get-CsvField -Row $row -Name 'Site URL', 'SiteUrl'
                ReportRow             = ConvertTo-ReportRowJson -Row $row
            }
        }

        $result = Export-AppendCsv -Path $csvPath -Rows @($rows) -Column $columns -KeyColumn @('ReportId', 'LabelGuid', 'ReportRow') -PassThru
        Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
            'labeled-file-sites.csv ({0}): {1} rows written, {2} skipped.' -f $label.Guid, $result.Written, $result.Skipped)
    }

    Export-AppendCsv -Path $csvPath -Column $columns

    if ($exported -eq 0 -and $pending -eq 0) {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            'No sensitivity label report could be read. Writing the header only.')
    }
}
finally {
    if ($connectedHere) {
        Disconnect-SPOService
    }
}
