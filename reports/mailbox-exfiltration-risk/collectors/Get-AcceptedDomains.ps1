#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes accepted-domains.csv: the accepted domains of the organization, appended
        each run.

    .DESCRIPTION
        Source: Get-AcceptedDomain -ResultSize Unlimited
        (https://learn.microsoft.com/powershell/module/exchangepowershell/get-accepteddomain).
        ResultSize defaults to 1000, so Unlimited is required. The other collectors
        compare a forwarding destination's domain with these names to decide whether it is
        external.

        The cmdlet page defers to the permissions page for the role; the least privileged
        role is UNVERIFIED. Availability in GCC and GCC High is UNVERIFIED, so those clouds
        are attempted with a warning.

    .EXAMPLE
        ./Get-AcceptedDomains.ps1 -OutputPath ./out
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

. (Join-Path $PSScriptRoot 'MailboxExfiltrationHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'MailboxExfiltrationSchema.psd1')
$columns = $schema.AcceptedDomains
$source = 'accepted-domains'
$csvPath = Join-Path $OutputPath 'accepted-domains.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-MailboxSourceAvailability -Source 'AcceptedDomains' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping accepted-domains.csv. $($availability.Reason) $($availability.Reference)")
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}
if ($availability.Status -eq 'Unverified') {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message $availability.Reason
}

$connectedHere = -not $SkipConnect
if ($connectedHere) {
    Connect-M365Service -Service ExchangeOnline -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization
}

try {
    try {
        $domains = @(Get-AcceptedDomain -ResultSize Unlimited -ErrorAction Stop)
        $rows = foreach ($domain in $domains) {
            [pscustomobject]@{
                RunDate    = $runDate
                Name       = [string]$domain.Name
                DomainName = [string]$domain.DomainName
                DomainType = [string]$domain.DomainType
                IsDefault  = Get-CsvBoolean $domain.Default
            }
        }
    
        $result = Export-AppendCsv -Path $csvPath -Rows @($rows) -Column $columns -KeyColumn @('RunDate', 'DomainName') -PassThru
        Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
            'accepted-domains.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
    }
    catch {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            'Get-AcceptedDomain is unavailable to this sign-in ({0}). It needs a role that can view recipients and domains (the least privileged role is UNVERIFIED). Writing the header only.' -f $_.Exception.Message)
        Export-AppendCsv -Path $csvPath -Column $columns
        return
    }
}
finally {
    if ($connectedHere) {
        Disconnect-ExchangeOnline -Confirm:$false
    }
}
