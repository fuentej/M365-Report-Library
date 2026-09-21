#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes guest-signins.csv: sign-ins by the tenant's guest users.

    .DESCRIPTION
        Source: Microsoft Graph list signIns (https://learn.microsoft.com/graph/api/signin-list),
        available in all clouds. Sign-in logs need Entra ID P1 or P2; without that licence
        the request is refused and the collector writes the header only.

        The v1.0 signIn resource has no supported $filter on guest status, so the query is
        scoped by createdDateTime and the results are narrowed to the user Ids in
        guests.csv. Run Get-Guests.ps1 first.

        The range is walked in windows (-WindowHours) so no single request asks the service
        for an unbounded span.

    .PARAMETER StartDate
        Overrides the start of the query window. By default the collector resumes from the
        latest CreatedDateTime already in guest-signins.csv, or -LookbackDays ago when the
        file is new.

    .EXAMPLE
        ./Get-GuestSignIns.ps1 -OutputPath ./out -LookbackDays 7
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

    [datetime]$StartDate,
    [datetime]$EndDate,

    [ValidateRange(1, 30)]
    [int]$LookbackDays = 30,

    [ValidateRange(1, 24)]
    [int]$WindowHours = 24,

    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Imported by path without -Force: a forced re-import builds a fresh module session
# state, which would discard anything the caller has arranged around the module - and
# that is exactly how the tests put mocks in front of the tenant cmdlets.
Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'GuestAccessSchema.psd1')
$columns = $schema.GuestSignIns
$source = 'guest-signins'
$csvPath = Join-Path $OutputPath 'guest-signins.csv'
$guestsPath = Join-Path $OutputPath 'guests.csv'

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

if (-not $SkipConnect) {
    Connect-M365Service -Service Graph -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization
}

$guestIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach ($guest in Get-CsvLatestSnapshot -Path $guestsPath) {
    if (-not [string]::IsNullOrWhiteSpace($guest.Id)) { [void]$guestIds.Add($guest.Id) }
}

if ($guestIds.Count -eq 0) {
    Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
        'guests.csv holds no guests, so there is nothing to narrow the sign-in log to. Run Get-Guests.ps1 first. Writing the header only.')
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}

$watermark = Get-CsvWatermark -Path $csvPath -Column 'CreatedDateTime'

$start = if ($PSBoundParameters.ContainsKey('StartDate')) { $StartDate.ToUniversalTime() }
elseif ($null -ne $watermark) { $watermark }
else { [datetime]::UtcNow.AddDays(-$LookbackDays) }

$end = if ($PSBoundParameters.ContainsKey('EndDate')) { $EndDate.ToUniversalTime() } else { [datetime]::UtcNow }

Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'Querying sign-ins for {0} guests from {1} to {2} in {3}-hour windows (watermark: {4}).' -f
    $guestIds.Count, (ConvertTo-CsvTimestamp $start), (ConvertTo-CsvTimestamp $end), $WindowHours,
    $(if ($null -eq $watermark) { 'none' } else { ConvertTo-CsvTimestamp $watermark }))

$rows = [System.Collections.Generic.List[object]]::new()

foreach ($window in Split-DateRange -Start $start -End $end -WindowMinutes ($WindowHours * 60)) {
    $filter = 'createdDateTime ge {0} and createdDateTime lt {1}' -f
        (ConvertTo-CsvTimestamp $window.Start), (ConvertTo-CsvTimestamp $window.End)

    try {
        $signIns = @(Get-MgAuditLogSignIn -All -Filter $filter -ErrorAction Stop)
    }
    catch {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            'The Entra ID sign-in log is unavailable to this sign-in ({0}). It needs Entra ID P1 or P2, the AuditLog.Read.All permission and the Reports Reader role. Writing the header only.' -f $_.Exception.Message)
        Export-AppendCsv -Path $csvPath -Column $columns
        return
    }

    foreach ($signIn in $signIns) {
        if (-not $guestIds.Contains([string]$signIn.UserId)) { continue }

        $city = ''
        $country = ''
        $location = $signIn.PSObject.Properties['Location']
        if ($location -and $null -ne $location.Value) {
            $city = [string](Get-GraphAdditionalProperty -Object $location.Value -Name 'city')
            $country = [string](Get-GraphAdditionalProperty -Object $location.Value -Name 'countryOrRegion')
        }

        $errorCode = ''
        $status = $signIn.PSObject.Properties['Status']
        if ($status -and $null -ne $status.Value) {
            $errorCode = [string](Get-GraphAdditionalProperty -Object $status.Value -Name 'errorCode')
        }

        $rows.Add([pscustomobject]@{
                CreatedDateTime         = ConvertTo-CsvTimestamp $signIn.CreatedDateTime
                Id                      = $signIn.Id
                UserId                  = $signIn.UserId
                UserPrincipalName       = $signIn.UserPrincipalName
                AppDisplayName          = $signIn.AppDisplayName
                ResourceDisplayName     = $signIn.ResourceDisplayName
                IpAddress               = $signIn.IPAddress
                City                    = $city
                CountryOrRegion         = $country
                ClientAppUsed           = $signIn.ClientAppUsed
                IsInteractive           = $signIn.IsInteractive
                ErrorCode               = $errorCode
                ConditionalAccessStatus = [string]$signIn.ConditionalAccessStatus
            })
    }
}

$result = Export-AppendCsv -Path $csvPath -Rows $rows.ToArray() -Column $columns -KeyColumn 'Id' -PassThru
Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'guest-signins.csv: {0} rows written, {1} skipped as already collected.' -f $result.Written, $result.Skipped)
