#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes outbound-spam-policies.csv: each outbound spam filter policy with its automatic
        external forwarding setting and who it applies to, appended each run.

    .DESCRIPTION
        Sources: Get-HostedOutboundSpamFilterPolicy
        (https://learn.microsoft.com/powershell/module/exchangepowershell/get-hostedoutboundspamfilterpolicy)
        and Get-HostedOutboundSpamFilterRule
        (https://learn.microsoft.com/powershell/module/exchangepowershell/get-hostedoutboundspamfilterrule).
        AutoForwardingMode is Automatic, On or Off
        (https://learn.microsoft.com/powershell/module/exchangepowershell/set-hostedoutboundspamfilterpolicy).
        Off blocks automatic forwarding set by inbox rules and by mailbox forwarding; it does not
        block forwarding done by mail flow (transport) rules
        (https://learn.microsoft.com/defender-office-365/outbound-spam-policies-external-email-forwarding).
        Automatic behaves like On or Off depending on the organization, so it is recorded as returned.

        The default policy applies to every sender not covered by an enabled custom policy; it has
        the lowest priority. A custom policy applies to what its rule names with From, FromMemberOf
        and SenderDomainIs (different condition types must all match), less its ExceptIfFrom,
        ExceptIfFromMemberOf and ExceptIfSenderDomainIs exceptions
        (https://learn.microsoft.com/powershell/module/exchangepowershell/new-hostedoutboundspamfilterrule,
        https://learn.microsoft.com/defender-office-365/outbound-spam-policies-configure).
        A custom policy with no rule, or whose rule is disabled, applies to no one.

        The cmdlet pages defer to the permissions page for the role; the matching role,
        View-Only Configuration, is an inference and is UNVERIFIED.

    .EXAMPLE
        ./Get-OutboundSpamPolicies.ps1 -OutputPath ./out
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
$columns = $schema.OutboundSpamPolicies
$source = 'outbound-spam-policies'
$csvPath = Join-Path $OutputPath 'outbound-spam-policies.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-MailboxSourceAvailability -Source 'OutboundSpamPolicies' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping outbound-spam-policies.csv. $($availability.Reason) $($availability.Reference)")
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
        $policies = @(Get-HostedOutboundSpamFilterPolicy -ErrorAction Stop)
        $rules = @(Get-HostedOutboundSpamFilterRule -ErrorAction Stop)

        $rows = foreach ($policy in $policies) {
            $isDefault = [bool]$policy.IsDefault
            $rule = $rules | Where-Object { [string]$_.HostedOutboundSpamFilterPolicy -eq [string]$policy.Name } | Select-Object -First 1

            $appliesTo = if ($isDefault) {
                'All senders not covered by an enabled custom policy'
            }
            elseif ($null -eq $rule) {
                'No one (no rule)'
            }
            elseif ([string]$rule.State -eq 'Disabled') {
                'No one (rule disabled)'
            }
            else {
                $conditions = @()
                $exceptions = @()
                $users = Join-ListValue $rule.From
                $groups = Join-ListValue $rule.FromMemberOf
                $domains = Join-ListValue $rule.SenderDomainIs
                if ($users) { $conditions += "Users: $users" }
                if ($groups) { $conditions += "Groups: $groups" }
                if ($domains) { $conditions += "Domains: $domains" }
                $exceptUsers = Join-ListValue $rule.ExceptIfFrom
                $exceptGroups = Join-ListValue $rule.ExceptIfFromMemberOf
                $exceptDomains = Join-ListValue $rule.ExceptIfSenderDomainIs
                if ($exceptUsers) { $exceptions += "Users: $exceptUsers" }
                if ($exceptGroups) { $exceptions += "Groups: $exceptGroups" }
                if ($exceptDomains) { $exceptions += "Domains: $exceptDomains" }

                $text = if ($conditions.Count -eq 0) { 'No conditions' } else { $conditions -join ' AND ' }
                if ($exceptions.Count -gt 0) { $text += ' EXCEPT ' + ($exceptions -join ' AND ') }
                $text
            }

            [pscustomobject]@{
                RunDate            = $runDate
                Name               = [string]$policy.Name
                Guid               = [string]$policy.Guid
                IsDefault          = $isDefault.ToString()
                AutoForwardingMode = [string]$policy.AutoForwardingMode
                RuleName           = if ($null -eq $rule) { '' } else { [string]$rule.Name }
                RuleState          = if ($null -eq $rule) { '' } else { [string]$rule.State }
                AppliesTo          = $appliesTo
            }
        }

        $result = Export-AppendCsv -Path $csvPath -Rows @($rows) -Column $columns -KeyColumn @('RunDate', 'Guid') -PassThru
        Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
            'outbound-spam-policies.csv: {0} rows written, {1} skipped.' -f $result.Written, $result.Skipped)
    }
    catch {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            'Get-HostedOutboundSpamFilterPolicy is unavailable to this sign-in ({0}). It needs a role that can view organization settings (View-Only Configuration matches; the least privileged role is UNVERIFIED). Writing the header only.' -f $_.Exception.Message)
        Export-AppendCsv -Path $csvPath -Column $columns
        return
    }
}
finally {
    if ($connectedHere) {
        Disconnect-ExchangeOnline -Confirm:$false
    }
}
