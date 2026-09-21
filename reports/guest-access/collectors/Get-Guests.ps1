#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes guests.csv: a snapshot of every guest in the tenant, appended each run.

    .DESCRIPTION
        Source: Microsoft Graph list users filtered to userType eq 'Guest'
        (https://learn.microsoft.com/graph/api/user-list).

        signInActivity needs Entra ID P1 or P2 and the AuditLog.Read.All permission
        (https://learn.microsoft.com/entra/identity/monitoring-health/howto-manage-inactive-user-accounts).
        Some tenants also reject signInActivity combined with a $filter on userType. The
        collector tries the filtered query first, then the same query unfiltered with the
        userType test moved into the script, and finally the filtered query without
        signInActivity - in which case the three sign-in columns are left empty and the
        reason is written to run.log.

    .PARAMETER SkipConnect
        Use an existing Microsoft Graph session instead of signing in.

    .EXAMPLE
        ./Get-Guests.ps1 -OutputPath ./out -Environment GCCHigh
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

    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Imported by path without -Force: a forced re-import builds a fresh module session
# state, which would discard anything the caller has arranged around the module - and
# that is exactly how the tests put mocks in front of the tenant cmdlets.
Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'GuestAccessSchema.psd1')
$columns = $schema.Guests
$source = 'guests'
$csvPath = Join-Path $OutputPath 'guests.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

if (-not $SkipConnect) {
    Connect-M365Service -Service Graph -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization
}

function Get-ExternalDomain {
    <#
        .SYNOPSIS
            The guest's home domain.

        .DESCRIPTION
            Preferred source is the mail address. Failing that, a B2B guest's UPN has
            the form alice_partner1.example.com#EXT#@contoso.onmicrosoft.com, so the
            home domain is what sits between the last underscore and #EXT#.
    #>
    param(
        [AllowNull()][string]$Mail,
        [AllowNull()][string]$UserPrincipalName
    )

    if (-not [string]::IsNullOrWhiteSpace($Mail) -and $Mail.Contains('@')) {
        return $Mail.Substring($Mail.LastIndexOf('@') + 1)
    }

    if ([string]::IsNullOrWhiteSpace($UserPrincipalName)) { return '' }

    $extIndex = $UserPrincipalName.IndexOf('#EXT#', [System.StringComparison]::OrdinalIgnoreCase)
    if ($extIndex -gt 0) {
        $mangled = $UserPrincipalName.Substring(0, $extIndex)
        $underscore = $mangled.LastIndexOf('_')
        if ($underscore -ge 0 -and $underscore -lt $mangled.Length - 1) {
            return $mangled.Substring($underscore + 1)
        }
    }

    if ($UserPrincipalName.Contains('@')) {
        return $UserPrincipalName.Substring($UserPrincipalName.LastIndexOf('@') + 1)
    }

    return ''
}

$baseSelect = @(
    'id', 'displayName', 'mail', 'userPrincipalName', 'userType', 'createdDateTime'
    'creationType', 'externalUserState', 'externalUserStateChangeDateTime', 'accountEnabled'
)
$guestFilter = "userType eq 'Guest'"

$attempts = @(
    @{ Description = "the userType filter with signInActivity"; UseFilter = $true;  WithSignInActivity = $true }
    @{ Description = "signInActivity on all users, filtering userType in the script"; UseFilter = $false; WithSignInActivity = $true }
    @{ Description = "the userType filter without signInActivity"; UseFilter = $true;  WithSignInActivity = $false }
)

$guests = $null
$haveSignInActivity = $false

foreach ($attempt in $attempts) {
    $select = if ($attempt.WithSignInActivity) { $baseSelect + 'signInActivity' } else { $baseSelect }
    $params = @{ All = $true; Property = $select; ErrorAction = 'Stop' }
    if ($attempt.UseFilter) { $params['Filter'] = $guestFilter }

    try {
        $result = @(Get-MgUser @params)
        if (-not $attempt.UseFilter) {
            $result = @($result | Where-Object { $_.UserType -eq 'Guest' })
        }
        $guests = $result
        $haveSignInActivity = $attempt.WithSignInActivity
        break
    }
    catch {
        Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
            "Listing guests using {0} failed: {1}" -f $attempt.Description, $_.Exception.Message)
    }
}

if ($null -eq $guests) {
    Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
        'Microsoft Graph list users is unavailable in this tenant or cloud. Writing the header only.')
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}

if (-not $haveSignInActivity) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        'signInActivity was not returned: it needs Entra ID P1 or P2 and the AuditLog.Read.All permission. ' +
        'LastSignInDateTime, LastNonInteractiveSignInDateTime and LastSuccessfulSignInDateTime are left empty.')
}

$rows = foreach ($guest in $guests) {
    $lastSignIn = ''
    $lastNonInteractive = ''
    $lastSuccessful = ''

    if ($haveSignInActivity) {
        $activity = Get-GraphAdditionalProperty -Object $guest -Name 'signInActivity'
        if ($null -ne $activity) {
            $lastSignIn = ConvertTo-CsvTimestamp (Get-GraphAdditionalProperty -Object $activity -Name 'lastSignInDateTime')
            $lastNonInteractive = ConvertTo-CsvTimestamp (Get-GraphAdditionalProperty -Object $activity -Name 'lastNonInteractiveSignInDateTime')
            $lastSuccessful = ConvertTo-CsvTimestamp (Get-GraphAdditionalProperty -Object $activity -Name 'lastSuccessfulSignInDateTime')
        }
    }

    [pscustomobject]@{
        RunDate                          = $runDate
        Id                               = $guest.Id
        DisplayName                      = $guest.DisplayName
        Mail                             = $guest.Mail
        UserPrincipalName                = $guest.UserPrincipalName
        ExternalDomain                   = Get-ExternalDomain -Mail $guest.Mail -UserPrincipalName $guest.UserPrincipalName
        CreatedDateTime                  = ConvertTo-CsvTimestamp $guest.CreatedDateTime
        CreationType                     = $guest.CreationType
        ExternalUserState                = $guest.ExternalUserState
        ExternalUserStateChangeDateTime  = ConvertTo-CsvTimestamp $guest.ExternalUserStateChangeDateTime
        AccountEnabled                   = $guest.AccountEnabled
        LastSignInDateTime               = $lastSignIn
        LastNonInteractiveSignInDateTime = $lastNonInteractive
        LastSuccessfulSignInDateTime     = $lastSuccessful
    }
}

$result = Export-AppendCsv -Path $csvPath -Rows @($rows) -Column $columns -KeyColumn @('RunDate', 'Id') -PassThru
Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'guests.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
