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

function Get-HttpStatusCode {
    <#
        .SYNOPSIS
            The HTTP status on a failed Invoke-RestMethod error, or $null.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $ErrorRecord
    )

    $candidates = @($ErrorRecord.Exception, $ErrorRecord.Exception.InnerException)
    foreach ($exception in $candidates) {
        if ($null -eq $exception) { continue }
        $response = $exception.PSObject.Properties['Response']
        if (-not $response -or $null -eq $response.Value) { continue }
        $status = $response.Value.PSObject.Properties['StatusCode']
        if ($status -and $null -ne $status.Value) { return [int]$status.Value }
    }

    # PowerShell 7: "Response status code does not indicate success: 429 (Too Many Requests)."
    if ($ErrorRecord.Exception.Message -match ':\s*(\d{3})\b') {
        return [int]$Matches[1]
    }
    return $null
}

function Get-RetryAfterSeconds {
    <#
        .SYNOPSIS
            Seconds to wait after HTTP 429, from Retry-After when the response has one.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $ErrorRecord
    )

    $delay = 1
    $candidates = @($ErrorRecord.Exception, $ErrorRecord.Exception.InnerException)
    foreach ($exception in $candidates) {
        if ($null -eq $exception) { continue }
        $response = $exception.PSObject.Properties['Response']
        if (-not $response -or $null -eq $response.Value) { continue }
        $headers = $response.Value.PSObject.Properties['Headers']
        if (-not $headers -or $null -eq $headers.Value) { continue }

        $retryAfter = $headers.Value.PSObject.Properties['RetryAfter']
        if ($retryAfter -and $retryAfter.Value) {
            $delta = $retryAfter.Value.PSObject.Properties['Delta']
            if ($delta -and $delta.Value) {
                $delay = [int][math]::Ceiling($delta.Value.TotalSeconds)
                break
            }
        }

        $named = $null
        if ($headers.Value -is [System.Collections.IDictionary] -and $headers.Value.Contains('Retry-After')) {
            $named = [string]$headers.Value['Retry-After']
        }
        if (-not [string]::IsNullOrWhiteSpace($named)) {
            $seconds = 0
            if ([int]::TryParse($named, [ref]$seconds)) { $delay = $seconds }
        }
    }

    if ($delay -lt 1) { return 1 }
    # A multi-minute Retry-After would stall an interactive run. Cap the wait and
    # let a later attempt fail if the service is still throttling.
    if ($delay -gt 60) { return 60 }
    return $delay
}

function Test-HttpRetry {
    <#
        .SYNOPSIS
            Whether a failed read should be retried, and for how many seconds.

        .DESCRIPTION
            The inventory query returns 429 when Azure Resource Graph throttles
            (https://learn.microsoft.com/rest/api/power-platform/resourcequery/resource-query/query-resources).
            Dataverse service protection does the same and sends Retry-After
            (https://learn.microsoft.com/power-apps/developer/data-platform/api-limits).
            401 and 403 are not retried: those are an auth or role refusal, and
            Invoke-RestMethod's -MaximumRetryCount would retry every 400-599.

            The call itself stays in the caller so a test mock of Invoke-RestMethod
            still applies. A script block invoked from here would run outside that mock.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $ErrorRecord,

        [Parameter(Mandatory)]
        [int]$Attempt,

        [int]$MaxAttempts = 4
    )

    $status = Get-HttpStatusCode -ErrorRecord $ErrorRecord
    if ($status -ne 429 -or $Attempt -ge $MaxAttempts) {
        return [pscustomobject]@{ Retry = $false; DelaySeconds = 0 }
    }

    [pscustomobject]@{
        Retry        = $true
        DelaySeconds = (Get-RetryAfterSeconds -ErrorRecord $ErrorRecord)
    }
}

function Get-InventoryAgentId {
    <#
        .SYNOPSIS
            The agent's Dataverse bot id.

        .DESCRIPTION
            properties.name is the CDS bot id on every agent
            (https://learn.microsoft.com/microsoft-copilot-studio/admin-agent-inventory#core-properties).
            properties.botId is the same id when the Entra identity block is populated;
            that block is empty for Microsoft Copilot Agent Builder agents. The ARM
            resource name is only a last resort.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        $Item
    )

    foreach ($path in @('properties.name', 'properties.botId')) {
        $value = Get-JsonValue -Object $Item -Path $path
        if (-not [string]::IsNullOrWhiteSpace([string]$value)) { return [string]$value }
    }
    return [string]$Item.name
}

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

        $response = $null
        for ($attempt = 1; ; $attempt++) {
            try {
                $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $headers -Body $body `
                    -ContentType 'application/json' -ErrorAction Stop
                break
            }
            catch {
                $decision = Test-HttpRetry -ErrorRecord $_ -Attempt $attempt
                if (-not $decision.Retry) { throw }
                if ($OutputPath) {
                    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $Source -Message (
                        'HTTP 429. Waiting {0} seconds before attempt {1}.' -f $decision.DelaySeconds, ($attempt + 1))
                }
                Start-Sleep -Seconds $decision.DelaySeconds
            }
        }
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

function Invoke-InventoryRead {
    <#
        .SYNOPSIS
            Collects every inventory row, keeping pages already read if a later page fails.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ApiHost,
        [Parameter(Mandatory)][string]$Token,
        [Parameter(Mandatory)][object[]]$Clauses,
        [string]$OutputPath,
        [string]$Source
    )

    $items = [System.Collections.Generic.List[object]]::new()
    $pagingError = ''
    try {
        Invoke-InventoryQuery -ApiHost $ApiHost -Token $Token -Clauses $Clauses -OutputPath $OutputPath -Source $Source |
            ForEach-Object { $null = $items.Add($_) }
    }
    catch {
        # No row yet means the call never succeeded (401, 403, network). The
        # collector turns that into a header-only CSV. A failure after rows were
        # read must not discard those rows or look like a successful snapshot.
        if ($items.Count -eq 0) { throw }
        $pagingError = $_.Exception.Message
    }

    [pscustomobject]@{
        Items       = $items.ToArray()
        PagingError = $pagingError
    }
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
    $origin = $null
    if (-not [uri]::TryCreate($DataverseUrl, [UriKind]::Absolute, [ref]$origin)) {
        throw "DataverseUrl '$DataverseUrl' is not an absolute URL."
    }

    while ($uri) {
        $nextUri = $null
        if (-not [uri]::TryCreate($uri, [UriKind]::Absolute, [ref]$nextUri) -or
            -not $origin.Host.Equals($nextUri.Host, [StringComparison]::OrdinalIgnoreCase)) {
            throw "Dataverse returned an @odata.nextLink on a different host ($uri). The bearer token is not sent there."
        }
        if (-not $seen.Add($uri)) { throw 'Dataverse returned the same @odata.nextLink twice; stopping.' }

        $response = $null
        for ($attempt = 1; ; $attempt++) {
            try {
                $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $headers -ErrorAction Stop
                break
            }
            catch {
                $decision = Test-HttpRetry -ErrorRecord $_ -Attempt $attempt
                if (-not $decision.Retry) { throw }
                Start-Sleep -Seconds $decision.DelaySeconds
            }
        }

        foreach ($row in @($response.value)) {
            if ($null -ne $row) { $row }
        }

        $uri = $null
        $link = $response.PSObject.Properties['@odata.nextLink']
        if ($link -and -not [string]::IsNullOrWhiteSpace([string]$link.Value)) { $uri = [string]$link.Value }
    }
}

#endregion
