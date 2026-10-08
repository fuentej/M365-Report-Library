#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes group-activity.csv: Microsoft 365 group activity by group, appended each run.

    .DESCRIPTION
        Source: Microsoft Graph getOffice365GroupsActivityDetail
        (https://learn.microsoft.com/graph/api/reportroot-getoffice365groupsactivitydetail),
        read with Get-MgReportOffice365GroupActivityDetail -OutFile. Graph answers with a
        302 to a preauthenticated CSV download that is valid for a few minutes.

        It is a snapshot of a rolling period (D7, D30, D90 or D180). Last Activity Date is
        not capped at the period.

        The external-member column is "External Member Count" in the header list and
        "Guest Count" in the schema example; whichever is present is read into
        ExternalMemberCount. Group names are blank when the organization setting that
        conceals user, group and site names is on.

        Not available in GCC High (US Government L4). GCC is UNVERIFIED and is attempted
        with a warning. Needs Reports.Read.All; a delegated caller also needs a limited
        admin role such as Reports Reader. Global Reader and Usage Summary Reports Reader
        do not receive the detail rows.

    .PARAMETER Period
        The report period. Defaults to D180, the longest.

    .EXAMPLE
        ./Get-GroupActivity.ps1 -OutputPath ./out
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
    [ValidateSet('D7', 'D30', 'D90', 'D180')]
    [string]$Period = 'D180',

    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Imported by path without -Force: a forced re-import builds a fresh module session
# state, which would discard anything the caller has arranged around the module - and
# that is exactly how the tests put mocks in front of the tenant cmdlets.
Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

. (Join-Path $PSScriptRoot 'TeamsGroupsHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'TeamsGroupsSchema.psd1')
$columns = $schema.GroupActivity
$aliases = $schema.ReportHeaderAliases
$source = 'group-activity'
$csvPath = Join-Path $OutputPath 'group-activity.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-TeamsGroupsSourceAvailability -Source 'GroupActivity' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping group-activity.csv. $($availability.Reason) $($availability.Reference)")
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}
if ($availability.Status -eq 'Unverified') {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message $availability.Reason
}

if (-not $SkipConnect) {
    Connect-M365Service -Service Graph -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization
}

$download = Join-Path ([System.IO.Path]::GetTempPath()) ('group-activity-' + [guid]::NewGuid().ToString('N') + '.csv')
$report = $null
try {
    try {
        Get-MgReportOffice365GroupActivityDetail -Period $Period -OutFile $download -ErrorAction Stop
        $report = if (Test-Path -LiteralPath $download) { @(Import-Csv -LiteralPath $download) } else { @() }
    }
    catch {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            'The Microsoft 365 groups activity report is unavailable to this sign-in or cloud ({0}). It needs Reports.Read.All and, for a delegated sign-in, a role such as Reports Reader. Writing the header only.' -f $_.Exception.Message)
        Export-AppendCsv -Path $csvPath -Column $columns
        return
    }
}
finally {
    Remove-Item -LiteralPath $download -Force -ErrorAction SilentlyContinue
}

function Read-Column { param($Row, [string[]]$Header) Get-ReportColumnValue -Row $Row -Header $Header }

$rows = foreach ($row in $report) {
    [pscustomobject]@{
        RunDate                       = $runDate
        ReportRefreshDate             = Read-Column $row 'Report Refresh Date'
        ReportPeriod                  = Read-Column $row 'Report Period'
        GroupId                       = Read-Column $row 'Group Id'
        GroupDisplayName              = Read-Column $row 'Group Display Name'
        IsDeleted                     = Read-Column $row 'Is Deleted'
        OwnerPrincipalName            = Read-Column $row 'Owner Principal Name'
        LastActivityDate              = Read-Column $row 'Last Activity Date'
        GroupType                     = Read-Column $row 'Group Type'
        MemberCount                   = Read-Column $row 'Member Count'
        ExternalMemberCount           = Read-Column $row $aliases.ExternalMemberCount
        ExchangeReceivedEmailCount    = Read-Column $row 'Exchange Received Email Count'
        SharePointActiveFileCount     = Read-Column $row 'SharePoint Active File Count'
        YammerPostedMessageCount      = Read-Column $row 'Yammer Posted Message Count'
        YammerReadMessageCount        = Read-Column $row 'Yammer Read Message Count'
        YammerLikedMessageCount       = Read-Column $row 'Yammer Liked Message Count'
        ExchangeMailboxTotalItemCount = Read-Column $row 'Exchange Mailbox Total Item Count'
        ExchangeMailboxStorageUsedByte = Read-Column $row 'Exchange Mailbox Storage Used (Byte)'
        SharePointTotalFileCount      = Read-Column $row 'SharePoint Total File Count'
        SharePointSiteStorageUsedByte = Read-Column $row 'SharePoint Site Storage Used (Byte)'
    }
}

$rows = @($rows | Where-Object { -not [string]::IsNullOrWhiteSpace($_.GroupId) })

$result = Export-AppendCsv -Path $csvPath -Rows $rows -Column $columns -KeyColumn @('RunDate', 'GroupId') -PassThru
Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'group-activity.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
