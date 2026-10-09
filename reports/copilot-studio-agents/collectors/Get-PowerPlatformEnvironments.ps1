#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes environments.csv: a snapshot of every Power Platform environment, appended
        each run.

    .DESCRIPTION
        Source: the Power Platform inventory API, type microsoft.powerplatform/environments
        (https://learn.microsoft.com/power-platform/admin/inventory-api). Fields:
        https://learn.microsoft.com/power-platform/admin/inventory-schema#environments.
        The query follows SkipToken until no page returns one.

        SIGN-IN IS DELEGATED AND INTERACTIVE. The inventory API does not support service
        principals or managed identities and returns HTTP 403 to them, so this collector
        opens a browser sign-in as a user with the AI Reader role (Global Reader also
        works and is broader than needed). It does not take -AppId or -CertificateThumbprint.

        The Power Platform API base URL is documented for Commercial only. For GCC and
        GCC High pass -ApiHost; no page names it, so it is UNVERIFIED and never guessed.

        Inventory changes appear after about 15 to 20 minutes.

    .PARAMETER ApiHost
        The Power Platform API base URL. Required for GCC and GCC High.

    .PARAMETER SkipConnect
        Skip the sign-in. Used by Run-All.ps1 and the tests; the token still comes from
        Get-AzAccessToken, so an existing Az session is needed.

    .EXAMPLE
        ./Get-PowerPlatformEnvironments.ps1 -OutputPath ./out
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$OutputPath,

    [ValidateSet('Commercial', 'GCC', 'GCCHigh')]
    [string]$Environment = 'Commercial',

    [string]$TenantId,
    [string]$ApiHost,

    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Imported by path without -Force: a forced re-import builds a fresh module session
# state, which would discard anything the caller has arranged around the module - and
# that is exactly how the tests put mocks in front of the tenant cmdlets.
Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

. (Join-Path $PSScriptRoot 'CopilotStudioHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'CopilotStudioSchema.psd1')
$columns = $schema.Environments
$source = 'environments'
$csvPath = Join-Path $OutputPath 'environments.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-CopilotStudioSourceAvailability -Source 'Environments' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping environments.csv. $($availability.Reason) $($availability.Reference)")
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}
if ($availability.Status -eq 'Unverified') {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message $availability.Reason
}

$apiBase = Get-PowerPlatformApiHost -Environment $Environment -ApiHost $ApiHost
if ($null -eq $apiBase) {
    Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
        "No Power Platform API base URL is documented for $Environment (UNVERIFIED). Pass -ApiHost. Writing the header only.")
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}

try {
    $token = Get-DelegatedAccessToken -ResourceUrl $apiBase -Environment $Environment -TenantId $TenantId -SkipConnect:$SkipConnect
    $read = Invoke-InventoryRead -ApiHost $apiBase -Token $token -OutputPath $OutputPath -Source $source `
        -Clauses (New-InventoryTypeClause -Type 'microsoft.powerplatform/environments')
}
catch {
    Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
        'The Power Platform inventory is unavailable to this sign-in ({0}). It needs a signed-in user with the AI Reader role; service principals get HTTP 403. Writing the header only.' -f $_.Exception.Message)
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}

$items = @($read.Items)

$rows = foreach ($item in $items) {
    [pscustomobject]@{
        RunDate          = $runDate
        EnvironmentId    = [string]$item.name
        DisplayName      = [string](Get-JsonValue $item 'properties.displayName')
        Location         = [string]$item.location
        EnvironmentType  = [string](Get-JsonValue $item 'properties.environmentType')
        IsManaged        = Get-CsvBoolean (Get-JsonValue $item 'properties.isManaged')
        EnvironmentGroup = [string](Get-JsonValue $item 'properties.environmentGroup')
        LastModifiedAt   = ConvertTo-CsvTimestamp (Get-JsonValue $item 'properties.lastModifiedAt')
    }
}

$result = Export-AppendCsv -Path $csvPath -Rows @($rows) -Column $columns -KeyColumn @('RunDate', 'EnvironmentId') -PassThru
if (-not [string]::IsNullOrWhiteSpace($read.PagingError)) {
    Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
        'environments.csv: wrote {0} row(s) already read, then stopped. The snapshot is incomplete. {1}' -f $result.Written, $read.PagingError)
    throw $read.PagingError
}
Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'environments.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
