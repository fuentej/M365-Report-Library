#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes everyone-item-exposure.csv: the sites, folders and files that "Everyone except
        external users" (EEEU) and "Everyone" can reach, and how access was granted, appended
        each run.

    .DESCRIPTION
        Source: SharePoint Online PowerShell
        Start-SPODataAccessGovernanceInsight -ReportEntity EveryoneExceptExternalUsers
        -ReportType Snapshot, and -ReportEntity Everyone
        (https://learn.microsoft.com/sharepoint/powershell-for-data-access-governance#generate-a-report-on-sites-and-files-shared-via-special-sharepoint-groups).
        This parameter set takes no -Workload and no -Name: the output covers SharePoint and
        OneDrive, is limited to 1 million rows, and leaves out permissions granted to these
        groups on system files and system groups. The report page recommends running the site
        permissions report at least once first.

        Needs module Microsoft.Online.SharePoint.PowerShell 16.0.27215.12000 or later, and the
        SharePoint Advanced Management Administrator role, which a Global Administrator
        assigns (https://learn.microsoft.com/sharepoint/data-access-governance-detailed-eeeu-everyone-permissions-report).
        Licence as in Get-SitePermissionBreadth.ps1; a snapshot report, so Microsoft 365 E5
        without SharePoint Advanced Management does not get it.

        Whether this item-level report is available in GCC and GCC High is UNVERIFIED: the
        EEEU insights feature row says Yes for both, but no page found says this snapshot
        report is available there and the report is newer than the feature table. Those
        clouds are attempted with a warning.

        Asynchronous, and reused or left running as in Get-SitePermissionBreadth.ps1; a
        report can be run again once every 30 days, so -MaxReportAgeHours defaults to 720.
        The CSV columns are the ones the report page lists
        (https://learn.microsoft.com/sharepoint/data-access-governance-detailed-eeeu-everyone-permissions-report#download-the-report).

    .PARAMETER AdminUrl
        The SharePoint admin center URL. Required unless -SkipConnect is used.

    .EXAMPLE
        ./Get-EveryoneItemExposure.ps1 -OutputPath ./out -AdminUrl https://contoso-admin.sharepoint.com
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
$columns = $schema.EveryoneItemExposure
$source = 'everyone-item-exposure'
$csvPath = Join-Path $OutputPath 'everyone-item-exposure.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-OversharingSourceAvailability -Source 'EveryoneItemExposure' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping everyone-item-exposure.csv. $($availability.Reason) $($availability.Reference)")
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

    foreach ($entity in 'EveryoneExceptExternalUsers', 'Everyone') {
        try {
            $report = Invoke-DagReport -ReportEntity $entity -ReportType 'Snapshot' `
                -MaxReportAgeHours $MaxReportAgeHours -WaitMinutes $WaitMinutes -PollSeconds $PollSeconds `
                -OutputPath $OutputPath -Source $source
        }
        catch {
            Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
                'The {0} item report is unavailable to this sign-in ({1}). It needs the SharePoint Advanced Management Administrator role, assigned by a Global Administrator, and module version 16.0.27215.12000 or later.' -f $entity, $_.Exception.Message)
            continue
        }

        if ($report.State -eq 'Pending') { $pending++ }
        if ($report.State -ne 'Exported') { continue }
        $exported++

        $rows = @(foreach ($row in $report.Rows) {
            [pscustomobject]@{
                RunDate         = $runDate
                ReportEntity    = $entity
                ReportId        = $report.ReportId
                ReportDate      = ConvertTo-CsvTimestamp (Get-CsvField -Row $row -Name 'ReportDate')
                SiteId          = Get-CsvField -Row $row -Name 'SiteId'
                WebId           = Get-CsvField -Row $row -Name 'WebId'
                ListId          = Get-CsvField -Row $row -Name 'ListId'
                ScopeId         = Get-CsvField -Row $row -Name 'ScopeId'
                UniqueId        = Get-CsvField -Row $row -Name 'UniqueId'
                ListItemId      = Get-CsvField -Row $row -Name 'ListItemId'
                ItemType        = Get-CsvField -Row $row -Name 'ItemType'
                ItemUrl         = Get-CsvField -Row $row -Name 'Item Url'
                RoleDefinition  = Get-CsvField -Row $row -Name 'Role definition'
                LinkId          = Get-CsvField -Row $row -Name 'LinkId'
                LinkScope       = Get-CsvField -Row $row -Name 'LinkScope'
                Recipient       = Get-CsvField -Row $row -Name 'Recipient'
                ParentObjectId  = Get-CsvField -Row $row -Name 'ParentObjectID'
                ParentGroupName = Get-CsvField -Row $row -Name 'ParentGroupName'
                ParentGroupEmail = Get-CsvField -Row $row -Name 'ParentGroupEmail'
                ParentGroupType = Get-CsvField -Row $row -Name 'ParentGroupType'
                TotalUserCount  = Get-CsvField -Row $row -Name 'TotalUserCount'
            }
        })

        Write-DagExportCap -Count $rows.Count -Cap 1000000 -OutputPath $OutputPath -Source $source `
            -ReportName "Everyone and EEEU items ($entity)" `
            -Reference 'https://learn.microsoft.com/sharepoint/data-access-governance-detailed-eeeu-everyone-permissions-report'

        $result = Export-AppendCsv -Path $csvPath -Rows $rows -Column $columns `
            -KeyColumn @('ReportId', 'ReportEntity', 'SiteId', 'WebId', 'ListId', 'ScopeId', 'UniqueId', 'LinkId', 'ParentObjectId', 'RoleDefinition', 'Recipient') -PassThru
        Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
            'everyone-item-exposure.csv ({0}): {1} rows written, {2} skipped.' -f $entity, $result.Written, $result.Skipped)
    }

    Export-AppendCsv -Path $csvPath -Column $columns

    if ($exported -eq 0 -and $pending -eq 0) {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            'No Everyone or EEEU item report could be read. Writing the header only.')
    }
}
finally {
    if ($connectedHere) {
        Disconnect-SPOService
    }
}
