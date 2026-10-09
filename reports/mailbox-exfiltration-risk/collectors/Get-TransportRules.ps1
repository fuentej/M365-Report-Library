#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes transport-rules.csv: the mail flow rules that redirect or blind-copy mail,
        appended each run.

    .DESCRIPTION
        Source: Get-TransportRule -ResultSize Unlimited
        (https://learn.microsoft.com/powershell/module/exchangepowershell/get-transportrule).
        ResultSize defaults to 1000. -ExcludeConditionActionDetails is left at its default
        of $false, because $true blanks Description, Conditions and Actions. Redirect is
        RedirectMessageTo and blind-copy is BlindCopyTo
        (https://learn.microsoft.com/exchange/security-and-compliance/mail-flow-rules/mail-flow-rule-actions).
        CopyTo and AddToRecipients add visible Cc and To recipients and are carried for
        context. Only rules with one of the four set are written.

        The cmdlet page defers to the permissions page for the role; the matching role,
        View-Only Configuration, is an inference and is UNVERIFIED.

    .EXAMPLE
        ./Get-TransportRules.ps1 -OutputPath ./out
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
$columns = $schema.TransportRules
$source = 'transport-rules'
$csvPath = Join-Path $OutputPath 'transport-rules.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-MailboxSourceAvailability -Source 'TransportRules' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping transport-rules.csv. $($availability.Reason) $($availability.Reference)")
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
        $rules = @(Get-TransportRule -ResultSize Unlimited -ExcludeConditionActionDetails $false -ErrorAction Stop)
        $accepted = Get-AcceptedDomainName -OutputPath $OutputPath -Source $source
    
        $rows = foreach ($rule in $rules) {
            $redirect = Join-ListValue $rule.RedirectMessageTo
            $blindCopy = Join-ListValue $rule.BlindCopyTo
            $copyTo = Join-ListValue $rule.CopyTo
            $addTo = Join-ListValue $rule.AddToRecipients
            if (-not $redirect -and -not $blindCopy -and -not $copyTo -and -not $addTo) { continue }
    
            $targets = Get-RecipientTargetSummary -Recipient @($rule.RedirectMessageTo, $rule.BlindCopyTo, $rule.CopyTo, $rule.AddToRecipients) -AcceptedDomain $accepted
    
            [pscustomobject]@{
                RunDate           = $runDate
                Name              = [string]$rule.Name
                Guid              = [string]$rule.Guid
                State             = [string]$rule.State
                Priority          = [string]$rule.Priority
                RedirectMessageTo = $redirect
                BlindCopyTo       = $blindCopy
                CopyTo            = $copyTo
                AddToRecipients   = $addTo
                TargetDomains     = $targets.Domains
                HasExternalTarget = $targets.HasExternal
            }
        }
    
        $result = Export-AppendCsv -Path $csvPath -Rows @($rows) -Column $columns -KeyColumn @('RunDate', 'Guid') -PassThru
        Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
            'transport-rules.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
    }
    catch {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            'Get-TransportRule is unavailable to this sign-in ({0}). It needs a role that can view mail flow settings (View-Only Configuration matches; the least privileged role is UNVERIFIED). Writing the header only.' -f $_.Exception.Message)
        Export-AppendCsv -Path $csvPath -Column $columns
        return
    }
}
finally {
    if ($connectedHere) {
        Disconnect-ExchangeOnline -Confirm:$false
    }
}
