#Requires -Version 7.0

# Dot-sourced by the copilot-studio-agents collectors.

function Get-CopilotStudioSourceAvailability {
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

function Get-JsonValue {
    <#
        .SYNOPSIS
            Reads a nested property from a parsed JSON object by dotted path, or $null.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowNull()]$Object,
        [Parameter(Mandatory)][string]$Path
    )

    $current = $Object
    foreach ($part in $Path -split '\.') {
        if ($null -eq $current) { return $null }
        if ($current -is [System.Collections.IDictionary]) {
            if (-not $current.Contains($part)) { return $null }
            $current = $current[$part]
        }
        else {
            $property = $current.PSObject.Properties[$part]
            if (-not $property) { return $null }
            $current = $property.Value
        }
    }
    return $current
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

#region Delegated sign-in

function Get-AzureCloudName {
    <#
        .SYNOPSIS
            The Azure PowerShell environment name for a library cloud. GCC signs in to
            the commercial Azure cloud, as the library's Graph mapping does.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][ValidateSet('Commercial', 'GCC', 'GCCHigh')][string]$Environment)

    if ($Environment -eq 'GCCHigh') { return 'AzureUSGovernment' }
    return 'AzureCloud'
}

function Get-DelegatedAccessToken {
    <#
        .SYNOPSIS
            Signs in as a user, interactively, and returns a bearer token for a resource.

        .DESCRIPTION
            Uses Connect-AzAccount and Get-AzAccessToken -ResourceUrl, the sign-in the
            Dataverse Web API quickstart documents
            (https://learn.microsoft.com/power-apps/developer/data-platform/webapi/quick-start-ps).
            Get-AzAccessToken returns a SecureString token by default; both that and a
            plain string are accepted.

            This is delegated on purpose. The Power Platform inventory API returns HTTP 403
            to service principals and managed identities
            (https://learn.microsoft.com/power-platform/admin/inventory-api#authentication).
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][string]$ResourceUrl,
        [Parameter(Mandatory)][ValidateSet('Commercial', 'GCC', 'GCCHigh')][string]$Environment,
        [string]$TenantId,
        [switch]$SkipConnect
    )

    if (-not $SkipConnect) {
        $connect = @{ Environment = (Get-AzureCloudName -Environment $Environment) }
        if (-not [string]::IsNullOrWhiteSpace($TenantId)) { $connect['Tenant'] = $TenantId }
        Connect-AzAccount @connect | Out-Null
    }

    $token = (Get-AzAccessToken -ResourceUrl $ResourceUrl -AsSecureString).Token
    if ($token -is [securestring]) {
        return ConvertFrom-SecureString -SecureString $token -AsPlainText
    }
    return [string]$token
}

#endregion

#region Power Platform inventory

function Get-PowerPlatformApiHost {
    <#
        .SYNOPSIS
            The Power Platform API base URL for a cloud, or $null when none is documented.

        .DESCRIPTION
            Commercial is documented: POST https://api.powerplatform.com/resourcequery/...
            (https://learn.microsoft.com/rest/api/power-platform/resourcequery/resource-query/query-resources).
            No page names the base URL for GCC or GCC High, so the caller must pass
            -ApiHost there; a value is never guessed.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateSet('Commercial', 'GCC', 'GCCHigh')][string]$Environment,
        [string]$ApiHost
    )

    if (-not [string]::IsNullOrWhiteSpace($ApiHost)) { return $ApiHost.TrimEnd('/') }
    if ($Environment -eq 'Commercial') { return 'https://api.powerplatform.com' }
    return $null
}

function Invoke-InventoryQuery {
    <#
        .SYNOPSIS
            Runs a PowerPlatformResources query and follows SkipToken to the last page.

        .DESCRIPTION
            POST {host}/resourcequery/resources/query?api-version=2024-10-01
            (https://learn.microsoft.com/power-platform/admin/inventory-api). The POST
            carries a query specification and changes nothing in the tenant.

            A page is followed whenever it returns a skipToken. The inventory API page
            shows resultTruncated 1 beside a skipToken, while the REST reference
            describes 0 as truncated, so the flag is not trusted to decide: stopping on a
            misread flag would drop agents silently. A repeated token stops the loop.

        .OUTPUTS
            One object per ResourceItem row.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ApiHost,
        [Parameter(Mandatory)][string]$Token,
        [Parameter(Mandatory)][object[]]$Clauses,
        [int]$Top = 1000,
        [int]$MaxPages = 1000,
        [string]$OutputPath,
        [string]$Source
    )

    $uri = '{0}/resourcequery/resources/query?api-version=2024-10-01' -f $ApiHost.TrimEnd('/')
    $headers = @{ Authorization = "Bearer $Token" }
    $skipToken = $null
    $seenTokens = [System.Collections.Generic.HashSet[string]]::new()
    $pages = 0

    do {
        $options = @{ Top = $Top }
        if (-not [string]::IsNullOrWhiteSpace($skipToken)) { $options['SkipToken'] = $skipToken }

        $body = @{ TableName = 'PowerPlatformResources'; Clauses = $Clauses; Options = $options } |
            ConvertTo-Json -Depth 10

        $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $headers -Body $body `
            -ContentType 'application/json' -ErrorAction Stop
        $pages++

        foreach ($row in @($response.data)) {
            if ($null -ne $row) { $row }
        }

        $next = $null
        $tokenProperty = $response.PSObject.Properties['skipToken']
        if ($tokenProperty -and -not [string]::IsNullOrWhiteSpace([string]$tokenProperty.Value)) {
            $next = [string]$tokenProperty.Value
        }

        if ($null -ne $next) {
            $flag = $response.PSObject.Properties['resultTruncated']
            if ($flag -and [string]$flag.Value -in @('0', 'false', 'False') -and $OutputPath) {
                Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $Source -Message (
                    'The inventory returned a skipToken with resultTruncated {0}; following it anyway because Learn describes the flag two ways.' -f $flag.Value)
            }
            if (-not $seenTokens.Add($next)) {
                throw 'The inventory returned the same skipToken twice; stopping so the loop cannot run forever.'
            }
            if ($pages -ge $MaxPages) {
                throw "The inventory query reached $MaxPages pages and still returned a skipToken."
            }
        }
        $skipToken = $next
    } while ($null -ne $skipToken)
}

function New-InventoryTypeClause {
    <#
        .SYNOPSIS
            The where clause that selects one resource type. Values are KQL string
            literals, quoted as the inventory API page shows.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Type)

    return @(
        @{ '$type' = 'where'; FieldName = 'type'; Operator = '=='; Values = @("'$Type'") }
    )
}

#endregion

#region Dataverse

function Invoke-DataverseQuery {
    <#
        .SYNOPSIS
            GETs a Dataverse Web API collection and follows @odata.nextLink.

        .DESCRIPTION
            https://learn.microsoft.com/power-apps/developer/data-platform/webapi/query/page-results
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DataverseUrl,
        [Parameter(Mandatory)][string]$Token,
        [Parameter(Mandatory)][string]$Path
    )

    $uri = '{0}/api/data/v9.1/{1}' -f $DataverseUrl.TrimEnd('/'), $Path.TrimStart('/')
    $headers = @{
        Authorization      = "Bearer $Token"
        'OData-MaxVersion' = '4.0'
        'OData-Version'    = '4.0'
        Accept             = 'application/json'
    }
    $seen = [System.Collections.Generic.HashSet[string]]::new()

    while ($uri) {
        if (-not $seen.Add($uri)) { throw 'Dataverse returned the same @odata.nextLink twice; stopping.' }

        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $headers -ErrorAction Stop

        foreach ($row in @($response.value)) {
            if ($null -ne $row) { $row }
        }

        $uri = $null
        $link = $response.PSObject.Properties['@odata.nextLink']
        if ($link -and -not [string]::IsNullOrWhiteSpace([string]$link.Value)) { $uri = [string]$link.Value }
    }
}

#endregion
