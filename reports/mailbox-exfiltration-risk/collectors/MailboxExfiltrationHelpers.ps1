#Requires -Version 7.0

# Dot-sourced by the mailbox-exfiltration-risk collectors.

function Get-MailboxSourceAvailability {
    <#
        .SYNOPSIS
            Says whether a source is available in a cloud, and why.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][ValidateSet('Commercial', 'GCC', 'GCCHigh')][string]$Environment,
        [Parameter(Mandatory)][hashtable]$Schema
    )

    $availability = $Schema.SourceAvailability
    if (-not $availability.ContainsKey($Source)) {
        throw "Unknown source '$Source'. Known sources: $(($availability.Keys | Sort-Object) -join ', ')."
    }

    $entry = $availability[$Source][$Environment]

    $reason = switch ($entry.Status) {
        'Available' { "$Source is documented as available in $Environment." }
        'NotAvailable' { "$Source is documented as unavailable in $Environment." }
        default { "Availability of $Source in $Environment is UNVERIFIED; the collector will attempt it anyway." }
    }

    [pscustomobject]@{
        Source      = $Source
        Environment = $Environment
        Status      = $entry.Status
        # Only NotAvailable skips. An unverified source is attempted so a guess
        # never drops data the tenant would have returned.
        ShouldSkip  = ($entry.Status -eq 'NotAvailable')
        Reason      = $reason
        Reference   = $entry.Reference
    }
}

function Join-ListValue {
    <#
        .SYNOPSIS
            Flattens a list into one CSV cell, separated by semicolons.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Position = 0)][AllowNull()]$Value)

    if ($null -eq $Value) { return '' }
    if ($Value -is [string]) { return $Value }
    return (@($Value) | Where-Object { $null -ne $_ } | ForEach-Object { [string]$_ }) -join ';'
}

function Get-CsvBoolean {
    <#
        .SYNOPSIS
            'True' or 'False' for a value that is set, an empty string for one that is not.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Position = 0)][AllowNull()]$Value)

    if ($null -eq $Value -or ($Value -is [string] -and [string]::IsNullOrWhiteSpace($Value))) { return '' }
    if ($Value -is [string]) {
        $parsed = $false
        if ([bool]::TryParse($Value, [ref]$parsed)) { return $parsed.ToString() }
        return $Value
    }
    return ([bool]$Value).ToString()
}

function Get-SmtpAddress {
    <#
        .SYNOPSIS
            The SMTP addresses found in a recipient value or a list of them.

        .DESCRIPTION
            Exchange renders recipients several ways: a bare address, smtp:user@domain,
            SMTP:user@domain, or a display name followed by the address in brackets.
            This finds every user@domain in the text, so each form works.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Position = 0)][AllowNull()]$Value)

    foreach ($item in @($Value)) {
        if ($null -eq $item) { continue }
        foreach ($match in [regex]::Matches([string]$item, "[A-Za-z0-9._%+'\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")) {
            $match.Value
        }
    }
}

function Get-AddressDomain {
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][string]$Address)

    return $Address.Substring($Address.LastIndexOf('@') + 1).ToLowerInvariant()
}

function Get-AcceptedDomainName {
    <#
        .SYNOPSIS
            The accepted domain names of the organization, or $null when they cannot be
            read.

        .DESCRIPTION
            Get-AcceptedDomain -ResultSize Unlimited
            (https://learn.microsoft.com/powershell/module/exchangepowershell/get-accepteddomain).
            ResultSize defaults to 1000, so Unlimited is required. A refusal returns $null
            and the caller leaves the external decision empty rather than guessing.
    #>
    [CmdletBinding()]
    param([string]$OutputPath, [string]$Source)

    try {
        $domains = @(Get-AcceptedDomain -ResultSize Unlimited -ErrorAction Stop)
        return , [string[]]@($domains | ForEach-Object { ([string]$_.DomainName).ToLowerInvariant() } | Where-Object { $_ })
    }
    catch {
        if ($OutputPath) {
            Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $Source -Message (
                'Get-AcceptedDomain is unavailable to this sign-in ({0}). External destinations cannot be decided, so IsExternal and HasExternalTarget stay empty.' -f $_.Exception.Message)
        }
        return $null
    }
}

function Test-ExternalDomain {
    <#
        .SYNOPSIS
            True when a domain is not one of the accepted domains. $null when the accepted
            domains are unknown.

        .DESCRIPTION
            A name matches an accepted domain exactly, or a wildcard accepted domain such
            as *.contoso.com. A subdomain of an accepted domain is external unless a
            wildcard covers it, because Exchange accepts only the domains it lists.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Domain,
        [AllowNull()][string[]]$AcceptedDomain
    )

    if ($null -eq $AcceptedDomain) { return $null }

    $name = $Domain.ToLowerInvariant()
    foreach ($accepted in $AcceptedDomain) {
        if ($name -eq $accepted) { return $false }
        if ($accepted.StartsWith('*.') -and ($name -eq $accepted.Substring(2) -or $name.EndsWith($accepted.Substring(1)))) { return $false }
    }
    return $true
}

function Get-RecipientTargetSummary {
    <#
        .SYNOPSIS
            The domains and the external flag for a set of recipient lists.

        .OUTPUTS
            An object with Domains (semicolon-joined) and HasExternal ('True', 'False', or
            '' when the accepted domains are unknown or no address was found).
    #>
    [CmdletBinding()]
    param(
        [AllowNull()]$Recipient,
        [AllowNull()][string[]]$AcceptedDomain
    )

    $domains = @(Get-SmtpAddress $Recipient | ForEach-Object { Get-AddressDomain $_ } | Select-Object -Unique)
    $external = ''
    if ($domains.Count -gt 0 -and $null -ne $AcceptedDomain) {
        $external = (@($domains | Where-Object { Test-ExternalDomain -Domain $_ -AcceptedDomain $AcceptedDomain }).Count -gt 0).ToString()
    }

    [pscustomobject]@{ Domains = ($domains -join ';'); HasExternal = $external }
}

function Test-FullAccessRow {
    <#
        .SYNOPSIS
            True for a Get-MailboxPermission row that is a real, explicit Full Access grant.

        .DESCRIPTION
            Keeps AccessRights like 'Full*', Deny false and IsInherited false, and drops
            NT AUTHORITY\SELF, to which the cmdlet page says FullAccess is assigned by
            default. Inherited entries for Administrator and Organization Management
            appear to allow FullAccess, and a Deny entry removes it.
            https://learn.microsoft.com/powershell/module/exchangepowershell/get-mailboxpermission
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param([Parameter(Mandatory)]$Row)

    $rights = Join-ListValue $Row.AccessRights
    if ($rights -notlike '*Full*') { return $false }
    if ([bool]$Row.Deny) { return $false }
    if ([bool]$Row.IsInherited) { return $false }
    if ([string]$Row.User -ieq 'NT AUTHORITY\SELF') { return $false }
    return $true
}

#region Unified audit log

function Test-AuditSearchHasMoreRecords {
    <#
        .SYNOPSIS
            True when a Search-UnifiedAuditLog record says another page is expected.
    #>
    [CmdletBinding()]
    param($Record)

    if ($null -eq $Record) { return $false }

    $metaProperty = $Record.PSObject.Properties['AuditSearchRequestMetadata']
    if (-not $metaProperty -or $null -eq $metaProperty.Value) { return $false }

    $meta = $metaProperty.Value
    $flag = $null
    if ($meta -is [System.Collections.IDictionary]) {
        foreach ($key in @('moreRecordsAvailable', 'MoreRecordsAvailable')) {
            if ($meta.Contains($key)) { $flag = $meta[$key]; break }
        }
    }
    else {
        foreach ($name in @('moreRecordsAvailable', 'MoreRecordsAvailable')) {
            $property = $meta.PSObject.Properties[$name]
            if ($property) { $flag = $property.Value; break }
        }
    }

    if ($null -eq $flag) { return $false }
    if ($flag -is [bool]) { return $flag }

    $parsed = $false
    if ([bool]::TryParse([string]$flag, [ref]$parsed)) { return $parsed }
    return $false
}

function Invoke-AuditSearch {
    <#
        .SYNOPSIS
            Searches the unified audit log over a range, in windows, paging each window
            with a ReturnLargeSet session.

        .DESCRIPTION
            Search-UnifiedAuditLog returns at most 100 records unless the same -SessionId
            is repeated with -SessionCommand ReturnLargeSet, which pages up to 50,000
            records a session (-ResultSize up to 5,000). A window that reaches 50,000 is
            incomplete and unsorted, so it is not returned: the caller gets the window in
            TruncatedWindow instead.
            https://learn.microsoft.com/powershell/module/exchangepowershell/search-unifiedauditlog

        .OUTPUTS
            An object with Records (one object per record: Audit, the parsed AuditData, and
            RecordType) and TruncatedWindow (a Start/End window, or $null).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][datetime]$Start,
        [Parameter(Mandatory)][datetime]$End,
        [Parameter(Mandatory)][int]$WindowHours,
        [string[]]$Operation,
        [string]$RecordType,
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)][string]$Source
    )

    $pageSize = 5000
    $sessionCap = 50000
    $nullPageRetries = 3
    $nullPageRetryDelayMs = 200

    $all = [System.Collections.Generic.List[object]]::new()
    $truncated = $null

    foreach ($window in Split-DateRange -Start $Start -End $End -WindowMinutes ($WindowHours * 60)) {
        $sessionId = 'mailbox-{0:yyyyMMddHHmmss}-{1}' -f $window.Start, [guid]::NewGuid().ToString('N').Substring(0, 8)
        $collected = 0
        $nullTries = 0
        $windowRecords = [System.Collections.Generic.List[object]]::new()
        $windowTruncated = $false

        while ($collected -lt $sessionCap) {
            $search = @{
                StartDate      = $window.Start
                EndDate        = $window.End
                SessionId      = $sessionId
                SessionCommand = 'ReturnLargeSet'
                ResultSize     = $pageSize
                ErrorAction    = 'Stop'
            }
            if ($Operation) { $search['Operations'] = $Operation }
            if ($RecordType) { $search['RecordType'] = $RecordType }

            $raw = Search-UnifiedAuditLog @search

            # $null and an empty collection both mean "nothing this call". Retry briefly
            # at the start of a window, because the service often returns nothing while
            # the search is prepared, then treat the window as empty.
            $records = @($raw | Where-Object { $null -ne $_ })
            if ($records.Count -eq 0) {
                if ($collected -eq 0 -and $nullTries -lt $nullPageRetries) {
                    $nullTries++
                    Start-Sleep -Milliseconds $nullPageRetryDelayMs
                    continue
                }
                break
            }

            # ResultCount is the hit count across every iteration of the session, not the
            # size of this page.
            $matched = 0
            $hasResultCount = $false
            $resultCountProperty = $records[0].PSObject.Properties['ResultCount']
            if ($resultCountProperty -and $null -ne $resultCountProperty.Value) {
                $hasResultCount = [int]::TryParse([string]$resultCountProperty.Value, [ref]$matched)
            }
            if ($hasResultCount -and $matched -gt $sessionCap) {
                $windowTruncated = $true
                break
            }

            $collected += $records.Count

            foreach ($record in $records) {
                $auditDataProperty = $record.PSObject.Properties['AuditData']
                if (-not $auditDataProperty -or [string]::IsNullOrWhiteSpace([string]$auditDataProperty.Value)) { continue }
                try {
                    $audit = [string]$auditDataProperty.Value | ConvertFrom-Json -ErrorAction Stop
                }
                catch {
                    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $Source -Message (
                        'Skipping a record whose AuditData is not valid JSON: {0}' -f $_.Exception.Message)
                    continue
                }
                $recordTypeProperty = $record.PSObject.Properties['RecordType']
                $windowRecords.Add([pscustomobject]@{
                        Audit      = $audit
                        RecordType = $(if ($recordTypeProperty) { [string]$recordTypeProperty.Value } else { '' })
                    })
            }

            if ($collected -ge $sessionCap) {
                $windowTruncated = $true
                break
            }

            # Repeat until the cmdlet returns nothing or the cap is hit. moreRecordsAvailable
            # says another iteration is expected even when a page is short.
            if (@($records | Where-Object { Test-AuditSearchHasMoreRecords -Record $_ }).Count -gt 0) { continue }
            if ($hasResultCount -and $matched -gt 0 -and $collected -ge $matched) { break }
            if ((-not $hasResultCount -or $matched -le 0) -and $records.Count -lt $pageSize) { break }
        }

        if ($windowTruncated) {
            $truncated = $window
            break
        }

        foreach ($item in $windowRecords) { $all.Add($item) }
    }

    [pscustomobject]@{ Records = $all.ToArray(); TruncatedWindow = $truncated }
}

function ConvertTo-AuditRow {
    <#
        .SYNOPSIS
            One CSV row for an audit record from Invoke-AuditSearch.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Item)

    $audit = $Item.Audit

    $pairs = foreach ($parameter in @(Get-GraphAdditionalProperty -Object $audit -Name 'Parameters')) {
        if ($null -eq $parameter) { continue }
        $name = Get-GraphAdditionalProperty -Object $parameter -Name 'Name'
        $value = Get-GraphAdditionalProperty -Object $parameter -Name 'Value'
        if ($name) { '{0}={1}' -f $name, $value }
    }

    $recordType = $Item.RecordType
    if (-not $recordType) { $recordType = [string](Get-GraphAdditionalProperty -Object $audit -Name 'RecordType') }

    [pscustomobject]@{
        CreationTime    = ConvertTo-CsvTimestamp (Get-GraphAdditionalProperty -Object $audit -Name 'CreationTime')
        Id              = [string](Get-GraphAdditionalProperty -Object $audit -Name 'Id')
        RecordType      = $recordType
        Operation       = [string](Get-GraphAdditionalProperty -Object $audit -Name 'Operation')
        UserId          = [string](Get-GraphAdditionalProperty -Object $audit -Name 'UserId')
        Workload        = [string](Get-GraphAdditionalProperty -Object $audit -Name 'Workload')
        ObjectId        = [string](Get-GraphAdditionalProperty -Object $audit -Name 'ObjectId')
        MailboxOwnerUPN = [string](Get-GraphAdditionalProperty -Object $audit -Name 'MailboxOwnerUPN')
        ClientIP        = [string](Get-GraphAdditionalProperty -Object $audit -Name 'ClientIP')
        ResultStatus    = [string](Get-GraphAdditionalProperty -Object $audit -Name 'ResultStatus')
        Parameters      = ($pairs -join ';')
    }
}

#endregion
