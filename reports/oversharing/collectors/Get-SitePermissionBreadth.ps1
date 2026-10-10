#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes site-permission-breadth.csv: every site ranked by how many users can reach it,
        from the Data access governance site permissions report, appended each run.

    .DESCRIPTION
        Source: SharePoint Online PowerShell
        Start-SPODataAccessGovernanceInsight -ReportEntity PermissionedUsers -ReportType Snapshot
        -Workload SharePoint -Name <name> -CountOfUsersMoreThan 0, repeated with
        -Workload OneDriveForBusiness
        (https://learn.microsoft.com/sharepoint/powershell-for-data-access-governance#generate-a-site-permission-state-report).
        On this parameter set -Name and -CountOfUsersMoreThan are both mandatory
        (https://learn.microsoft.com/powershell/module/microsoft.online.sharepoint.powershell/start-spodataaccessgovernanceinsight).
        The article uses 0 for sites at least one user can reach; the cmdlet example says
        "more than 1000 users" next to a 0, so the threshold is a parameter here
        (-CountOfUsersMoreThan, minimum 0). Then Get-SPODataAccessGovernanceInsight for the
        status and Export-SPODataAccessGovernanceInsight to download the CSV.

        The report is asynchronous. The first one takes up to five days, later ones up to 24
        hours, the data is up to 48 hours old, and a report can be run again once every 30
        days. A completed report no older than -MaxReportAgeHours (default 720, 30 days) is
        reused rather than started again; a report still running after -WaitMinutes is left
        running and the next run exports it. Each report is appended once, tagged with the
        run date it was collected.

        Needs module Microsoft.Online.SharePoint.PowerShell 16.0.25409 or later, and
        Connect-SPOService without -Credential. Role: SharePoint Administrator, or SharePoint
        Advanced Management Administrator. Licence: SharePoint Advanced Management (at least
        one Microsoft Copilot license, or the Plan 1 add-on); Microsoft 365 E5 alone gets no
        snapshot reports
        (https://learn.microsoft.com/sharepoint/data-access-governance-reports).

        The CSV columns are the ones the site permissions report page lists
        (https://learn.microsoft.com/sharepoint/data-access-governance-site-permissions-report#download-the-site-permissions-for-your-organization-reports).
        Reports might not work when "Display concealed user, group, and site names in all
        reports" is cleared in the Microsoft 365 admin center, and are unavailable for
        Microsoft 365 operated by 21Vianet.

    .PARAMETER AdminUrl
        The SharePoint admin center URL, such as https://contoso-admin.sharepoint.com.
        Required unless -SkipConnect is used.

    .EXAMPLE
        ./Get-SitePermissionBreadth.ps1 -OutputPath ./out -AdminUrl https://contoso-admin.sharepoint.com
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
    [int]$CountOfUsersMoreThan = 0,

    [ValidateRange(0, [int]::MaxValue)]
    [int]$MaxReportAgeHours = 720,

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
$columns = $schema.SitePermissionBreadth
$source = 'site-permission-breadth'
$csvPath = Join-Path $OutputPath 'site-permission-breadth.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-OversharingSourceAvailability -Source 'SitePermissionBreadth' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping site-permission-breadth.csv. $($availability.Reason) $($availability.Reference)")
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

    foreach ($workload in 'SharePoint', 'OneDriveForBusiness') {
        try {
            $report = Invoke-DagReport -ReportEntity 'PermissionedUsers' -ReportType 'Snapshot' -Workload $workload `
                -StartArgument @{ Name = ('OversharingSitePermissions-{0}-{1}' -f $workload, $runDate); CountOfUsersMoreThan = $CountOfUsersMoreThan } `
                -MatchProperty @{ CountOfUsersMoreThan = $CountOfUsersMoreThan } `
                -MaxReportAgeHours $MaxReportAgeHours -WaitMinutes $WaitMinutes -PollSeconds $PollSeconds `
                -OutputPath $OutputPath -Source $source
        }
        catch {
            Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
                'The {0} site permissions report is unavailable to this sign-in ({1}). It needs the SharePoint Administrator role and a SharePoint Advanced Management license.' -f $workload, $_.Exception.Message)
            continue
        }

        if ($report.State -eq 'Pending') { $pending++ }
        if ($report.State -ne 'Exported') { continue }
        $exported++

        $rows = foreach ($row in $report.Rows) {
            [pscustomobject]@{
                RunDate                        = $runDate
                Workload                       = $workload
                ReportId                       = $report.ReportId
                ReportDate                     = ConvertTo-CsvTimestamp (Get-CsvField -Row $row -Name 'Report Date')
                SiteId                         = Get-CsvField -Row $row -Name 'Site ID', 'SiteId'
                SiteName                       = Get-CsvField -Row $row -Name 'Site Name'
                SiteUrl                        = Get-CsvField -Row $row -Name 'Site URL'
                SiteTemplate                   = Get-CsvField -Row $row -Name 'Site Template'
                PrimaryAdmin                   = Get-CsvField -Row $row -Name 'Primary admin'
                PrimaryAdminEmail              = Get-CsvField -Row $row -Name 'Primary admin email'
                ExternalSharing                = Get-CsvField -Row $row -Name 'ExternalSharing'
                SitePrivacy                    = Get-CsvField -Row $row -Name 'Site Privacy'
                SiteSensitivity                = Get-CsvField -Row $row -Name 'Site Sensitivity'
                UsersWithAccess                = Get-CsvField -Row $row -Name 'Number of users having access'
                GuestUserPermissions           = Get-CsvField -Row $row -Name 'Guest user permissions'
                ExternalParticipantPermissions = Get-CsvField -Row $row -Name 'External participant permissions'
                EntraGroupPermissions          = Get-CsvField -Row $row -Name 'Microsoft Entra group permission count'
                FileCount                      = Get-CsvField -Row $row -Name 'File count'
                ItemsWithUniquePermissions     = Get-CsvField -Row $row -Name 'Items with unique permissions count'
                PeopleInYourOrgLinks           = Get-CsvField -Row $row -Name 'People In Your Org link count'
                AnyoneLinks                    = Get-CsvField -Row $row -Name 'Anyone link count'
                EeeuPermissions                = Get-CsvField -Row $row -Name 'EEEU permission count'
                EveryonePermissions            = Get-CsvField -Row $row -Name 'Everyone permission count'
            }
        }

        $result = Export-AppendCsv -Path $csvPath -Rows @($rows) -Column $columns -KeyColumn @('ReportId', 'SiteId') -PassThru
        Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
            'site-permission-breadth.csv ({0}): {1} rows written, {2} skipped.' -f $workload, $result.Written, $result.Skipped)
    }

    Export-AppendCsv -Path $csvPath -Column $columns

    if ($exported -eq 0 -and $pending -eq 0) {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            'No site permissions report could be read. Writing the header only.')
    }
}
finally {
    if ($connectedHere) {
        Disconnect-SPOService
    }
}
