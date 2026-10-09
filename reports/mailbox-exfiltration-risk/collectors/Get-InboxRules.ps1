#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes inbox-rules.csv: every inbox rule that forwards, redirects or deletes mail,
        by mailbox, appended each run.

    .DESCRIPTION
        Source: Get-InboxRule -Mailbox <id> -IncludeHidden -ResultSize Unlimited, one call
        per mailbox
        (https://learn.microsoft.com/powershell/module/exchangepowershell/get-inboxrule).
        -IncludeHidden is required to see hidden rules. ResultSize defaults to 1000. The
        forward, redirect and delete properties are ForwardTo, ForwardAsAttachmentTo,
        RedirectTo and DeleteMessage
        (https://learn.microsoft.com/powershell/module/exchangepowershell/new-inboxrule).
        Only rules that set one of them are written. The mailbox list comes from
        Get-EXOMailbox -ResultSize Unlimited.

        The cmdlet page says it does NOT work for View-Only Organization Management or
        the Entra Global Reader role, and defers to the permissions page for the role that
        does; the least privileged role is UNVERIFIED. A mailbox the sign-in cannot read is
        logged and skipped.

        The cmdlet returns no documented marker for a hidden rule, so hidden rules are
        read but not flagged.

    .PARAMETER MailboxLimit
        Read at most this many mailboxes. For a trial run; 0 means all.

    .EXAMPLE
        ./Get-InboxRules.ps1 -OutputPath ./out
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

    [ValidateRange(0, 1000000)]
    [int]$MailboxLimit = 0,

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
$columns = $schema.InboxRules
$source = 'inbox-rules'
$csvPath = Join-Path $OutputPath 'inbox-rules.csv'
$runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-MailboxSourceAvailability -Source 'InboxRules' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping inbox-rules.csv. $($availability.Reason) $($availability.Reference)")
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
        $mailboxes = @(Get-EXOMailbox -ResultSize Unlimited -ErrorAction Stop)
        if ($MailboxLimit -gt 0) { $mailboxes = @($mailboxes | Select-Object -First $MailboxLimit) }
        $accepted = Get-AcceptedDomainName -OutputPath $OutputPath -Source $source
    
        $rows = [System.Collections.Generic.List[object]]::new()
        $failed = 0
        foreach ($mailbox in $mailboxes) {
            $rules = $null
            try {
                $rules = @(Get-InboxRule -Mailbox ([string]$mailbox.UserPrincipalName) -IncludeHidden -ResultSize Unlimited -ErrorAction Stop)
            }
            catch {
                $failed++
                Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
                    'Get-InboxRule failed for {0} ({1}); skipping the mailbox.' -f $mailbox.UserPrincipalName, $_.Exception.Message)
                continue
            }
    
            foreach ($rule in $rules) {
                $forwardTo = Join-ListValue $rule.ForwardTo
                $forwardAsAttachmentTo = Join-ListValue $rule.ForwardAsAttachmentTo
                $redirectTo = Join-ListValue $rule.RedirectTo
                $delete = [bool]$rule.DeleteMessage
                if (-not $forwardTo -and -not $forwardAsAttachmentTo -and -not $redirectTo -and -not $delete) { continue }
    
                $targets = Get-RecipientTargetSummary -Recipient @($rule.ForwardTo, $rule.ForwardAsAttachmentTo, $rule.RedirectTo) -AcceptedDomain $accepted
    
                $rows.Add([pscustomobject]@{
                        RunDate                          = $runDate
                        MailboxUserPrincipalName         = [string]$mailbox.UserPrincipalName
                        MailboxExternalDirectoryObjectId = [string]$mailbox.ExternalDirectoryObjectId
                        RuleIdentity                     = [string]$rule.RuleIdentity
                        RuleName                         = [string]$rule.Name
                        Enabled                          = Get-CsvBoolean $rule.Enabled
                        Priority                         = [string]$rule.Priority
                        ForwardTo                        = $forwardTo
                        ForwardAsAttachmentTo            = $forwardAsAttachmentTo
                        RedirectTo                       = $redirectTo
                        DeleteMessage                    = $delete.ToString()
                        TargetDomains                    = $targets.Domains
                        HasExternalTarget                = $targets.HasExternal
                    })
            }
        }
    
        if ($failed -gt 0 -and $failed -eq $mailboxes.Count) {
            Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
                'Get-InboxRule failed for every mailbox. It needs a role the cmdlet page defers to the permissions page for; Global Reader and View-Only Organization Management do not work. Writing the header only.')
            Export-AppendCsv -Path $csvPath -Column $columns
            return
        }
    
        $result = Export-AppendCsv -Path $csvPath -Rows $rows.ToArray() -Column $columns `
            -KeyColumn @('RunDate', 'MailboxExternalDirectoryObjectId', 'RuleIdentity') -PassThru
        Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
            'inbox-rules.csv: {0} rows written, {1} skipped; {2} mailbox(es) could not be read.' -f $result.Written, $result.Skipped, $failed)
    }
    catch {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            'Get-EXOMailbox is unavailable to this sign-in ({0}). It needs a role that can view recipients (the least privileged role is UNVERIFIED). Writing the header only.' -f $_.Exception.Message)
        Export-AppendCsv -Path $csvPath -Column $columns
        return
    }
}
finally {
    if ($connectedHere) {
        Disconnect-ExchangeOnline -Confirm:$false
    }
}
