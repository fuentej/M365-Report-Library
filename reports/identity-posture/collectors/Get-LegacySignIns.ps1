#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes signins.csv: sign-ins that used a legacy authentication client, from the
        Entra ID sign-in log. Appended from the last exported timestamp.

    .DESCRIPTION
        Source: Microsoft Graph list signIns (https://learn.microsoft.com/graph/api/signin-list),
        v1.0. Each window is queried once per legacy clientAppUsed value (Exchange ActiveSync,
        IMAP, MAPI, SMTP, POP, other clients); the values are from
        https://learn.microsoft.com/graph/api/resources/signin#properties. -All follows
        @odata.nextLink; a page holds at most 1,000 sign-ins.

        CAVEAT: v1.0 returns sign-ins that are interactive in nature, plus successful
        federated sign-ins. Legacy protocols often authenticate non-interactively, so this
        file can under-report them. Non-interactive sign-ins need the beta signInEventTypes
        filter, which this library does not call (Microsoft does not support beta APIs in
        production); see the README.

        Retention is 7 days (free) or 30 days (P1 and P2)
        (https://learn.microsoft.com/entra/identity/monitoring-health/reference-reports-data-retention).
        A daily run builds a longer history than the source keeps. Rows are keyed by sign-in
        Id, so the overlap at the watermark is not repeated.

        Needs Microsoft Entra ID P1 or P2, AuditLog.Read.All and the Reports Reader role.
        conditionalAccessStatus is a property of the sign-in. Policy.Read.All is required
        only to read appliedConditionalAccessPolicies, which this collector does not.

    .EXAMPLE
        ./Get-LegacySignIns.ps1 -OutputPath ./out -LookbackDays 7
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

    # An alternate schema file. The tests use it to exercise the NotAvailable path.
    [string]$SchemaPath = (Join-Path $PSScriptRoot 'IdentityPostureSchema.psd1'),

    [datetime]$StartDate,
    [datetime]$EndDate,

    [ValidateRange(1, 30)]
    [int]$LookbackDays = 30,

    [ValidateRange(1, 24)]
    [int]$WindowHours = 24,

    # The clientAppUsed values to collect. Defaults to the legacy values in the schema.
    [string[]]$ClientAppUsed,

    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Imported by path without -Force: a forced re-import builds a fresh module session
# state, which would discard anything the caller has arranged around the module - and
# that is exactly how the tests put mocks in front of the tenant cmdlets.
Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

. (Join-Path $PSScriptRoot 'IdentityPostureHelpers.ps1')


$schema = Import-PowerShellDataFile -LiteralPath $SchemaPath
$columns = [string[]]$schema.SignIns
$source = 'signins'
$csvPath = Join-Path $OutputPath 'signins.csv'

if (-not $ClientAppUsed) { $ClientAppUsed = [string[]]$schema.LegacyClientAppValues }

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-IdentityPostureSourceAvailability -Source 'SignIns' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping signins.csv. $($availability.Reason) $($availability.Reference)")
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}
if ($availability.Status -eq 'Unverified') {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message $availability.Reason
}

if (-not $SkipConnect) {
    Connect-M365Service -Service Graph -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization `
        -Scopes (Get-IdentityPostureScope -Schema $schema -Source 'SignIns')
}

$watermark = Get-CsvWatermark -Path $csvPath -Column 'CreatedDateTime'

$start = if ($PSBoundParameters.ContainsKey('StartDate')) { $StartDate.ToUniversalTime() }
elseif ($null -ne $watermark) { $watermark }
else { [datetime]::UtcNow.AddDays(-$LookbackDays) }

$end = if ($PSBoundParameters.ContainsKey('EndDate')) { $EndDate.ToUniversalTime() } else { [datetime]::UtcNow }

if ($end -le $start) {
    if ($PSBoundParameters.ContainsKey('StartDate') -or $PSBoundParameters.ContainsKey('EndDate')) {
        # An inverted or empty range asked for explicitly is a mistake, and collecting
        # nothing while reporting success would hide it.
        throw ('The requested range is empty: the end ({0}) is not later than the start ({1}).' -f
            (ConvertTo-CsvTimestamp $end), (ConvertTo-CsvTimestamp $start))
    }

    Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
        'Nothing new to collect: the watermark ({0}) is already at or after the end of the range.' -f (ConvertTo-CsvTimestamp $start))
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}

Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'Querying sign-ins for client apps [{0}] from {1} to {2} in {3}-hour windows (watermark: {4}).' -f
    ($ClientAppUsed -join ', '), (ConvertTo-CsvTimestamp $start), (ConvertTo-CsvTimestamp $end), $WindowHours,
    $(if ($null -eq $watermark) { 'none' } else { ConvertTo-CsvTimestamp $watermark }))

$rows = [System.Collections.Generic.List[object]]::new()
$failure = $null

# A window is committed only when every client app in it was read. A failure part-way
# through would otherwise leave a partial window in the file, and the next run would
# resume from the watermark that window moved.
foreach ($window in Split-DateRange -Start $start -End $end -WindowMinutes ($WindowHours * 60)) {
    $windowRows = [System.Collections.Generic.List[object]]::new()

    foreach ($clientApp in $ClientAppUsed) {
        $filter = "createdDateTime ge {0} and createdDateTime lt {1} and clientAppUsed eq '{2}'" -f
            (ConvertTo-CsvTimestamp $window.Start), (ConvertTo-CsvTimestamp $window.End), ($clientApp -replace "'", "''")

        try {
            # -All follows @odata.nextLink; a page holds at most 1,000 sign-ins.
            $signIns = @(Get-MgAuditLogSignIn -All -Filter $filter -ErrorAction Stop)
        }
        catch {
            $failure = $_.Exception.Message
            break
        }

        foreach ($signIn in $signIns) {
            $status = $signIn.PSObject.Properties['Status']
            $errorCode = if ($status -and $null -ne $status.Value) {
                [string](Get-GraphAdditionalProperty -Object $status.Value -Name 'errorCode')
            }
            else { '' }

            $windowRows.Add([pscustomobject]@{
                    CreatedDateTime         = ConvertTo-CsvTimestamp $signIn.CreatedDateTime
                    Id                      = $signIn.Id
                    UserId                  = $signIn.UserId
                    UserPrincipalName       = $signIn.UserPrincipalName
                    AppDisplayName          = $signIn.AppDisplayName
                    ResourceDisplayName     = $signIn.ResourceDisplayName
                    IpAddress               = $signIn.IPAddress
                    ClientAppUsed           = $signIn.ClientAppUsed
                    IsInteractive           = $signIn.IsInteractive
                    ErrorCode               = $errorCode
                    ConditionalAccessStatus = [string]$signIn.ConditionalAccessStatus
                })
        }
    }

    if ($null -ne $failure) { break }
    foreach ($row in $windowRows) { $rows.Add($row) }
}

$result = Export-AppendCsv -Path $csvPath -Rows $rows.ToArray() -Column $columns -KeyColumn 'Id' -PassThru
Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'signins.csv: {0} rows written, {1} skipped as already collected.' -f $result.Written, $result.Skipped)


if ($null -ne $failure) {
    if ($result.Written -eq 0 -and (Test-LicenseError -Message $failure)) {
        Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
            "Skipping signins.csv: the sign-in log is not licensed in this tenant ($failure). It needs Microsoft Entra ID P1 or P2. Writing the header only.")
        return
    }

    $keptNote = if ($result.Written -gt 0) { " The windows read before the failure were kept." } else { "" }
    Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
        "The Entra ID sign-in log is unavailable to this sign-in ($failure). It needs Microsoft Entra ID P1 or P2, the AuditLog.Read.All permission and the Reports Reader role.$keptNote")
    if ($result.Written -gt 0) {
        throw "Reading the sign-in log stopped part-way ($failure). The complete windows were kept; re-run to resume."
    }
    throw "Reading the sign-in log failed ($failure)."
}

