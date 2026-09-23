#Requires -Version 7.0

<#
    .SYNOPSIS
        Helpers specific to the Purview information protection collectors.

    .DESCRIPTION
        Dot-source this file from every collector in this report:

            . (Join-Path $PSScriptRoot 'PurviewIpHelpers.ps1')

        Dot-sourcing (not Import-Module) puts these functions in the caller's
        scope, the same reason PurviewIpSchema.psd1 is read with
        Import-PowerShellDataFile rather than imported as a module.

        Nothing here connects to a service. Ported from fuentej/purview-ip-report
        (scripts/PurviewCommon.ps1), trimmed to what this report's collectors use
        once Export-AppendCsv/Get-CsvWatermark from the shared module took over
        CSV writing and incremental resume.
#>

Set-StrictMode -Version Latest

function Get-PurviewProperty {
    <#
        .SYNOPSIS
            Reads a property that may not be there.

        .DESCRIPTION
            Not every Activity Explorer column is present on every activity, and
            an audit record only carries the properties its scenario produced.
            Under Set-StrictMode reaching for a missing property is an error, so
            go through this instead.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)][AllowNull()]$InputObject,
        [Parameter(Mandatory, Position = 1)][string]$Name,
        [Parameter(Position = 2)][AllowNull()]$Default = $null
    )

    if ($null -eq $InputObject) { return $Default }

    if ($InputObject -is [hashtable] -or $InputObject -is [System.Collections.Specialized.OrderedDictionary]) {
        if ($InputObject.Contains($Name)) { return $InputObject[$Name] }
        return $Default
    }

    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $property) { return $Default }
    if ($null -eq $property.Value) { return $Default }
    $property.Value
}

function ConvertTo-PurviewListValue {
    <#
        .SYNOPSIS
            Joins an array (or a scalar) into the ';'-separated string a CSV
            cell holds, for columns like Locations, LabelIds and PolicyRules.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Position = 0)][AllowNull()]$Value)

    if ($null -eq $Value) { return '' }
    if ($Value -is [string]) { return $Value }
    if ($Value -is [System.Collections.IEnumerable]) {
        return (@($Value) | ForEach-Object { ConvertTo-PurviewListValue $_ }) -join ';'
    }
    [string]$Value
}

function Get-PurviewSourceAvailability {
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

    [PSCustomObject]@{
        Source      = $Source
        Environment = $Environment
        Status      = $entry.Status
        # Collect unless the source is documented as unavailable. An
        # unverified source is attempted rather than skipped, so a guess never
        # silently drops data that the tenant would have returned.
        ShouldSkip  = ($entry.Status -eq 'NotAvailable')
        Reason      = $reason
        Reference   = $entry.Reference
    }
}

function Test-PurviewActivityName {
    <#
        .SYNOPSIS
            True when the value is one the schema's ActivityCategories recognises.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Position = 0)][AllowNull()][string]$Activity,
        [Parameter(Mandatory)][hashtable]$Schema
    )

    if (-not $Activity) { return $false }
    foreach ($category in $Schema.ActivityCategories.Keys) {
        if ($Activity -in $Schema.ActivityCategories[$category]) { return $true }
    }
    $false
}

function Resolve-PurviewActivityName {
    <#
        .SYNOPSIS
            Normalises an activity value to the filter-enum name.

        .DESCRIPTION
            Export-ActivityExplorerData filters on enum names (LabelApplied),
            but the records it returns have been observed carrying the portal's
            display name on Activity ("Label applied") with the enum on
            ActivityId. Microsoft documents the filter values and lists Activity
            as an exported column, but does not state which form the column
            holds, and this repository has no tenant to check against.

            So take ActivityId when it is there, and otherwise map a display
            name back to its enum by ignoring case and spacing. A value that
            matches nothing is passed through unchanged rather than guessed at.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Position = 0)][AllowNull()][string]$Activity,
        [Parameter(Position = 1)][AllowNull()][string]$ActivityId,
        [Parameter(Mandatory)][hashtable]$Schema
    )

    $byKey = @{}
    foreach ($category in $Schema.ActivityCategories.Keys) {
        foreach ($name in $Schema.ActivityCategories[$category]) {
            $byKey[$name] = $name
            $byKey[($name -replace '[\s_-]', '')] = $name
        }
    }

    foreach ($candidate in $ActivityId, $Activity) {
        if (-not $candidate) { continue }
        if ($byKey.Contains($candidate)) { return $byKey[$candidate] }

        $squashed = ($candidate -replace '[\s_-]', '')
        if ($byKey.Contains($squashed)) { return $byKey[$squashed] }
    }

    if ($Activity) { return $Activity }
    if ($ActivityId) { return $ActivityId }
    ''
}

function Get-PurviewActivityCategory {
    <#
        .SYNOPSIS
            Groups an activity value for reporting. Unrecognised values map to
            'Other'.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Position = 0)][AllowNull()][string]$Activity,
        [Parameter(Mandatory)][hashtable]$Schema
    )

    if (-not $Activity) { return 'Other' }
    foreach ($category in $Schema.ActivityCategories.Keys) {
        if ($Activity -in $Schema.ActivityCategories[$category]) { return $category }
    }
    'Other'
}

function Get-PurviewSensitiveInfoTypeSummary {
    <#
        .SYNOPSIS
            Flattens an Activity Explorer SensitiveInfoTypeData array into
            three cells.

        .DESCRIPTION
            A record can match several sensitive information types. Names are
            joined with ';', Count is their total, and Confidence is the
            highest of them.
    #>
    [CmdletBinding()]
    param([Parameter(Position = 0)][AllowNull()]$SensitiveInfoTypeData)

    $names = [System.Collections.Generic.List[string]]::new()
    $total = 0
    $confidence = $null

    foreach ($entry in @($SensitiveInfoTypeData)) {
        if ($null -eq $entry) { continue }
        $properties = $entry.PSObject.Properties.Name

        if ('SensitiveInfoTypeName' -in $properties -and $entry.SensitiveInfoTypeName) {
            $names.Add([string]$entry.SensitiveInfoTypeName)
        }
        elseif ('SensitiveInformationTypeName' -in $properties -and $entry.SensitiveInformationTypeName) {
            $names.Add([string]$entry.SensitiveInformationTypeName)
        }
        elseif ('Name' -in $properties -and $entry.Name) {
            $names.Add([string]$entry.Name)
        }

        foreach ($countProperty in 'SensitiveInfoTypeCount', 'Count') {
            if ($countProperty -in $properties -and $null -ne $entry.$countProperty) {
                $total += [int]$entry.$countProperty
                break
            }
        }

        foreach ($confidenceProperty in 'SensitiveInfoTypeConfidence', 'Confidence') {
            if ($confidenceProperty -in $properties -and $null -ne $entry.$confidenceProperty) {
                $value = [int]$entry.$confidenceProperty
                if ($null -eq $confidence -or $value -gt $confidence) { $confidence = $value }
                break
            }
        }
    }

    [PSCustomObject]@{
        Name       = ($names -join ';')
        Count      = $total
        Confidence = $confidence
    }
}

function Get-PurviewSensitiveInformationTypeName {
    <#
        .SYNOPSIS
            Pulls the classifier names out of a DLP rule's condition.

        .DESCRIPTION
            ContentContainsSensitiveInformation comes back in either of two
            shapes:

                @( @{ name = 'Credit Card Number' } )                 flat
                @{ operator; groups = @( @{ sensitivetypes = @(...)    grouped
                                            labels         = @(...) } ) }

            The grouped form nests the classifier under groups ->
            sensitivetypes, so reading `name` off the outer object yields the
            group's name ("Default") or nothing. Handle both, and pick up
            sensitivity labels used as a condition too - a Copilot DLP rule
            often matches on a label and names no sensitive information type
            at all.
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param([Parameter(Position = 0)][AllowNull()]$Condition)

    $names = [System.Collections.Generic.List[string]]::new()

    function Add-PurviewName {
        param([AllowNull()]$Entry)
        if ($null -eq $Entry) { return }
        if ($Entry -is [string]) { $names.Add($Entry); return }
        $value = Get-PurviewProperty $Entry 'name' (Get-PurviewProperty $Entry 'Name')
        if ($value) { $names.Add([string]$value) }
    }

    foreach ($item in @($Condition)) {
        if ($null -eq $item) { continue }

        $groups = Get-PurviewProperty $item 'groups' (Get-PurviewProperty $item 'Groups')
        if ($groups) {
            foreach ($group in @($groups)) {
                foreach ($property in 'sensitivetypes', 'SensitiveTypes', 'labels', 'Labels') {
                    foreach ($entry in @(Get-PurviewProperty $group $property @())) { Add-PurviewName $entry }
                }
            }
            continue
        }

        Add-PurviewName $item
    }

    [string[]]@($names | Where-Object { $_ } | Select-Object -Unique)
}

function Get-PurviewPolicyDetail {
    <#
        .SYNOPSIS
            Normalises an audit record's PolicyDetails into policy id, name
            and rules.

        .DESCRIPTION
            PolicyDetails is present only when a policy blocked or restricted
            access, and is documented as carrying "PolicyId, PolicyName, list
            of rules". It has been seen both as a single object and as an
            array of them, and an empty array means no policy at all - reading
            fields off the array itself yields nothing, and treating a
            non-null array as "blocked" marks a resource blocked with no
            policy to show for it.
    #>
    [CmdletBinding()]
    param([Parameter(Position = 0)][AllowNull()]$PolicyDetails)

    $ids = [System.Collections.Generic.List[string]]::new()
    $policyNames = [System.Collections.Generic.List[string]]::new()
    $ruleNames = [System.Collections.Generic.List[string]]::new()

    foreach ($detail in @($PolicyDetails)) {
        if ($null -eq $detail) { continue }

        $id = Get-PurviewProperty $detail 'PolicyId' (Get-PurviewProperty $detail 'policyId')
        $name = Get-PurviewProperty $detail 'PolicyName' (Get-PurviewProperty $detail 'policyName')
        if ($id) { $ids.Add([string]$id) }
        if ($name) { $policyNames.Add([string]$name) }

        foreach ($rule in @(Get-PurviewProperty $detail 'Rules' (Get-PurviewProperty $detail 'rules' @()))) {
            if ($null -eq $rule) { continue }
            if ($rule -is [string]) { $ruleNames.Add($rule); continue }
            $ruleName = Get-PurviewProperty $rule 'RuleName' (Get-PurviewProperty $rule 'Name' (Get-PurviewProperty $rule 'name'))
            if ($ruleName) { $ruleNames.Add([string]$ruleName) }
        }
    }

    [PSCustomObject]@{
        PolicyId   = ($ids | Select-Object -Unique) -join ';'
        PolicyName = ($policyNames | Select-Object -Unique) -join ';'
        Rules      = ($ruleNames | Select-Object -Unique) -join ';'
        # An empty PolicyDetails is not a policy.
        HasPolicy  = ($ids.Count -gt 0 -or $policyNames.Count -gt 0 -or $ruleNames.Count -gt 0)
    }
}

function Get-PurviewCopilotScope {
    <#
        .SYNOPSIS
            Says whether a DLP policy covers Microsoft 365 Copilot / Copilot
            Chat.

        .DESCRIPTION
            A Copilot DLP policy is scoped either through the
            CopilotExperiences enforcement plane or through the Microsoft 365
            Copilot location GUID in its Locations JSON.
            https://learn.microsoft.com/powershell/module/exchangepowershell/new-dlpcompliancepolicy
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [AllowNull()]$EnforcementPlanes,
        [AllowNull()]$Locations,
        [Parameter(Mandatory)][hashtable]$Schema
    )

    foreach ($plane in @($EnforcementPlanes)) {
        if ([string]$plane -eq $Schema.CopilotEnforcementPlane) { return $true }
    }

    $locationText = ConvertTo-PurviewListValue $Locations
    if ($locationText -and $locationText -match [regex]::Escape($Schema.CopilotLocationId)) {
        return $true
    }

    $false
}

# Slicing a time range for Export-ActivityExplorerData and Search-UnifiedAuditLog
# (both cap what one query can return) uses the shared module's Split-DateRange
# rather than a second copy of the same windowing logic.
