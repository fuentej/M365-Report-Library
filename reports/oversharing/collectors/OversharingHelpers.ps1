#Requires -Version 7.0

# Dot-sourced by the oversharing collectors.

function Get-OversharingSourceAvailability {
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

function Get-JsonProperty {
    <#
        .SYNOPSIS
            Reads one property from a parsed JSON object or a hashtable, or $null.

        .DESCRIPTION
            Invoke-MgGraphRequest returns hashtables; ConvertFrom-Json returns objects. The
            name is one key, not a path, because Graph names such as @odata.nextLink contain
            dots.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowNull()]$Object,
        [Parameter(Mandatory)][string]$Name
    )

    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name)) { return $Object[$Name] }
        return $null
    }
    $property = $Object.PSObject.Properties[$Name]
    if ($property) { return $property.Value }
    return $null
}

function Get-JsonValue {
    <#
        .SYNOPSIS
            Reads a nested property by dotted path, or $null.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowNull()]$Object,
        [Parameter(Mandatory)][string]$Path
    )

    $current = $Object
    foreach ($part in $Path -split '\.') {
        if ($null -eq $current) { return $null }
        $current = Get-JsonProperty -Object $current -Name $part
    }
    return $current
}

function Get-CsvField {
    <#
        .SYNOPSIS
            A field of an imported CSV row, found by any of several column names.

        .DESCRIPTION
            Column names are compared with case, spaces and punctuation ignored, so "Site URL",
            "SiteUrl" and "Site_URL" are one name. Returns an empty string when none matches.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]$Row,
        [Parameter(Mandatory)][string[]]$Name
    )

    $wanted = @($Name | ForEach-Object { ($_ -replace '[^A-Za-z0-9]', '').ToLowerInvariant() })
    foreach ($property in $Row.PSObject.Properties) {
        if (($property.Name -replace '[^A-Za-z0-9]', '').ToLowerInvariant() -in $wanted) {
            return [string]$property.Value
        }
    }
    return ''
}

#region Microsoft Graph

function Invoke-GraphGet {
    <#
        .SYNOPSIS
            One GET against Microsoft Graph, through the session Connect-MgGraph opened.

        .DESCRIPTION
            A relative path is sent to the cloud the session is signed in to:
            https://graph.microsoft.com for Commercial and GCC, https://graph.microsoft.us
            for GCC High (https://learn.microsoft.com/graph/deployments). An absolute URL,
            such as a returned @odata.nextLink, is sent as it is. This is the only place the
            collectors make a raw Graph request, and it only reads.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Uri)

    Invoke-MgGraphRequest -Method GET -Uri $Uri -ErrorAction Stop
}

function Get-GraphPagedValue {
    <#
        .SYNOPSIS
            Every item of a Graph collection, following @odata.nextLink until it is absent.

        .DESCRIPTION
            https://learn.microsoft.com/graph/paging. The returned URL is requested as it is.
            For getAllSites the sample nextLink changes the path to oneDrive.getAllSites, so
            rebuilding /sites/getAllSites from a skiptoken would ask for the wrong page.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Uri)

    $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
    $next = $Uri
    while (-not [string]::IsNullOrEmpty($next)) {
        if (-not $seen.Add($next)) {
            throw "Microsoft Graph returned a nextLink it had already returned: $next"
        }
        $page = Invoke-GraphGet -Uri $next
        foreach ($item in @(Get-JsonProperty -Object $page -Name 'value')) {
            if ($null -ne $item) { $item }
        }
        $next = [string](Get-JsonProperty -Object $page -Name '@odata.nextLink')
    }
}

function Get-AllSiteRow {
    <#
        .SYNOPSIS
            One row per site from GET /sites/getAllSites, for sites.csv.

        .DESCRIPTION
            https://learn.microsoft.com/graph/api/site-getallsites. Application permission
            Sites.Read.All; delegated is not supported. Learn points to
            https://learn.microsoft.com/onedrive/developer/rest-api/concepts/scan-guidance
            for large tenants.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$RunDate)

    foreach ($site in Get-GraphPagedValue -Uri '/v1.0/sites/getAllSites') {
        [pscustomobject]@{
            RunDate          = $RunDate
            SiteId           = [string](Get-JsonProperty -Object $site -Name 'id')
            Name             = [string](Get-JsonProperty -Object $site -Name 'name')
            WebUrl           = [string](Get-JsonProperty -Object $site -Name 'webUrl')
            IsPersonalSite   = Get-CsvBoolean (Get-JsonProperty -Object $site -Name 'isPersonalSite')
            HostName         = [string](Get-JsonValue -Object $site -Path 'siteCollection.hostName')
            DataLocationCode = [string](Get-JsonValue -Object $site -Path 'siteCollection.dataLocationCode')
        }
    }
}

function Get-DriveItemWalk {
    <#
        .SYNOPSIS
            Every file and folder in a drive, folder by folder.

        .DESCRIPTION
            GET /drives/{drive-id}/root/children, then GET /drives/{drive-id}/items/{item-id}/children
            for each folder (https://learn.microsoft.com/graph/api/driveitem-list-children),
            every collection paged. -MaxItems stops the walk after that many items; 0 walks
            the whole drive.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DriveId,
        [int]$MaxItems = 0
    )

    $select = '$select=id,name,webUrl,folder'
    $queue = [System.Collections.Generic.Queue[string]]::new()
    $queue.Enqueue("/v1.0/drives/$DriveId/root/children?$select")
    $count = 0

    while ($queue.Count -gt 0) {
        foreach ($item in Get-GraphPagedValue -Uri $queue.Dequeue()) {
            $id = [string](Get-JsonProperty -Object $item -Name 'id')
            $isFolder = $null -ne (Get-JsonProperty -Object $item -Name 'folder')

            [pscustomobject]@{
                Id       = $id
                Name     = [string](Get-JsonProperty -Object $item -Name 'name')
                WebUrl   = [string](Get-JsonProperty -Object $item -Name 'webUrl')
                IsFolder = $isFolder
            }

            $count++
            if ($MaxItems -gt 0 -and $count -ge $MaxItems) { return }
            if ($isFolder) { $queue.Enqueue("/v1.0/drives/$DriveId/items/$id/children?$select") }
        }
    }
}

function ConvertTo-ItemPermissionRow {
    <#
        .SYNOPSIS
            One CSV row for a permission returned by GET .../items/{item-id}/permissions.

        .DESCRIPTION
            https://learn.microsoft.com/graph/api/driveitem-list-permissions. link.webUrl and
            shareId are secrets the page says are returned only to callers who can create the
            permission; they are not copied. inheritedFrom names the ancestor a permission is
            inherited from.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$RunDate,
        [Parameter(Mandatory)][string]$SiteId,
        [Parameter(Mandatory)][string]$DriveId,
        [Parameter(Mandatory)]$Item,
        [Parameter(Mandatory)]$Permission
    )

    $inheritedFrom = Get-JsonProperty -Object $Permission -Name 'inheritedFrom'

    $names = [System.Collections.Generic.List[string]]::new()
    $granted = @(Get-JsonProperty -Object $Permission -Name 'grantedToV2')
    $granted += @(Get-JsonProperty -Object $Permission -Name 'grantedToIdentitiesV2')
    foreach ($entry in $granted) {
        if ($null -eq $entry) { continue }
        foreach ($kind in 'user', 'group', 'siteUser', 'siteGroup') {
            $name = [string](Get-JsonValue -Object $entry -Path "$kind.displayName")
            if ($name -and -not $names.Contains($name)) { $names.Add($name) }
        }
    }

    [pscustomobject]@{
        RunDate              = $RunDate
        SiteId               = $SiteId
        DriveId              = $DriveId
        ItemId               = $Item.Id
        ItemName             = $Item.Name
        ItemWebUrl           = $Item.WebUrl
        PermissionId         = [string](Get-JsonProperty -Object $Permission -Name 'id')
        Roles                = Join-ListValue (Get-JsonProperty -Object $Permission -Name 'roles')
        LinkScope            = [string](Get-JsonValue -Object $Permission -Path 'link.scope')
        LinkType             = [string](Get-JsonValue -Object $Permission -Path 'link.type')
        LinkPreventsDownload = Get-CsvBoolean (Get-JsonValue -Object $Permission -Path 'link.preventsDownload')
        HasPassword          = Get-CsvBoolean (Get-JsonProperty -Object $Permission -Name 'hasPassword')
        ExpirationDateTime   = ConvertTo-CsvTimestamp (Get-JsonProperty -Object $Permission -Name 'expirationDateTime')
        IsInherited          = ($null -ne $inheritedFrom).ToString()
        InheritedFromItemId  = [string](Get-JsonProperty -Object $inheritedFrom -Name 'id')
        GrantedTo            = ($names -join ';')
    }
}

#endregion

#region SharePoint Online

function Connect-SharePointAdmin {
    <#
        .SYNOPSIS
            Signs in to the SharePoint admin center for the Data access governance and
            site cmdlets.

        .DESCRIPTION
            Connect-SPOService
            (https://learn.microsoft.com/powershell/module/microsoft.online.sharepoint.powershell/connect-sposervice).
            The Data access governance page requires it without -Credential
            (https://learn.microsoft.com/sharepoint/powershell-for-data-access-governance), so
            interactive sign-in passes only -Url. -AppId, -CertificateThumbprint and -TenantId
            sign in with the certificate parameter set instead. -Region ITAR is passed for GCC
            High, the value the cmdlet page says applies only to GCC High and DoD tenants.

            No page found says the module connects in GCC or GCC High (an open item in
            docs/candidates/oversharing.md), so those clouds log a warning and try anyway.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminUrl,
        [ValidateSet('Commercial', 'GCC', 'GCCHigh')][string]$Environment = 'Commercial',
        [string]$AppId,
        [string]$CertificateThumbprint,
        [string]$TenantId,
        [string]$OutputPath,
        [string]$Source
    )

    $hasAppId = -not [string]::IsNullOrWhiteSpace($AppId)
    $hasCertificate = -not [string]::IsNullOrWhiteSpace($CertificateThumbprint)
    if ($hasAppId -ne $hasCertificate) {
        throw 'App-only sign-in needs both -AppId and -CertificateThumbprint. Omit both to sign in interactively.'
    }

    $params = @{ Url = $AdminUrl }
    if ($hasAppId) {
        if ([string]::IsNullOrWhiteSpace($TenantId)) {
            throw 'App-only sign-in to SharePoint Online requires -TenantId.'
        }
        $params['ClientId'] = $AppId
        $params['CertificateThumbprint'] = $CertificateThumbprint
        $params['TenantId'] = $TenantId
    }
    if ($Environment -eq 'GCCHigh') { $params['Region'] = 'ITAR' }

    if ($Environment -ne 'Commercial' -and $OutputPath) {
        Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $Source -Message (
            "Whether Connect-SPOService and the SharePoint Online cmdlets connect in $Environment is UNVERIFIED; the collector will attempt it anyway.")
    }

    Connect-SPOService @params
}

function Get-ReportTime {
    [CmdletBinding()]
    [OutputType([nullable[datetime]])]
    param([AllowNull()]$Value)

    if ($null -eq $Value) { return $null }
    if ($Value -is [datetime]) { return $Value.ToUniversalTime() }
    [datetime]$parsed = [datetime]::MinValue
    $styles = [System.Globalization.DateTimeStyles]::AdjustToUniversal -bor [System.Globalization.DateTimeStyles]::AssumeUniversal
    if ([datetime]::TryParse([string]$Value, [System.Globalization.CultureInfo]::InvariantCulture, $styles, [ref]$parsed)) {
        return $parsed
    }
    return $null
}

function Test-ActivityDataCollection {
    <#
        .SYNOPSIS
            Logs whether data collection for recent activity based reports is on.

        .DESCRIPTION
            Get-SPOAuditDataCollectionStatusForActivityInsights returns NotInitiated,
            InProgress or Paused, and a report can be generated only when it is InProgress
            (https://learn.microsoft.com/sharepoint/powershell-for-data-access-governance#check-the-data-collection-status-for-recent-activity-based-reports).
            Which -ReportEntity spelling the cmdlet accepts is UNVERIFIED: its accepted
            values have no underscore (SharingLinksAnyone) and its example uses
            SharingLinks_Anyone, so both are tried.

            This only reads and only logs. Start-SPOAuditDataCollectionForActivityInsights
            changes the tenant and is never called; with a SharePoint Advanced Management
            license the report runs without it.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ReportEntity,
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)][string]$Source
    )

    $spellings = @(($ReportEntity -replace '_', ''), $ReportEntity | Select-Object -Unique)
    $lastError = $null
    foreach ($spelling in $spellings) {
        try {
            $result = Get-SPOAuditDataCollectionStatusForActivityInsights -ReportEntity $spelling -ErrorAction Stop
            $statusProperty = if ($null -ne $result) { $result.PSObject.Properties['Status'] } else { $null }
            $status = if ($statusProperty) { [string]$statusProperty.Value } else { [string]$result }
            if ($status -eq 'InProgress') {
                Write-CollectorLog -OutputPath $OutputPath -Source $Source -Message (
                    "Data collection for $spelling is InProgress.")
            }
            else {
                Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $Source -Message (
                    "Data collection for $spelling is '$status'. Without a SharePoint Advanced Management license a report can be generated only when it is InProgress. This collector does not start data collection, because that changes the tenant.")
            }
            return
        }
        catch {
            $lastError = $_
        }
    }

    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $Source -Message (
        "The data collection status for $ReportEntity could not be read ({0}). Which -ReportEntity string Get-SPOAuditDataCollectionStatusForActivityInsights accepts is UNVERIFIED." -f $lastError.Exception.Message)
}

function Invoke-DagReport {
    <#
        .SYNOPSIS
            Starts or reuses a Data access governance report, waits for it, and returns the
            rows of its CSV.

        .DESCRIPTION
            Start-SPODataAccessGovernanceInsight, then Get-SPODataAccessGovernanceInsight
            for the status, then Export-SPODataAccessGovernanceInsight to download the CSV
            (https://learn.microsoft.com/sharepoint/powershell-for-data-access-governance).
            The reports are asynchronous: the first site permissions report takes up to five
            days, and a report can be run again only once every 30 days (24 hours for the
            activity and label reports). So a completed report no older than -MaxReportAgeHours
            is reused, and a report still running after -WaitMinutes is left running and
            reported as Pending, to be picked up by a later run.

            -MatchProperty narrows the reports that may be reused to those whose listing
            carries a property with that value; a report that lacks the property is not
            reused.

            The export is downloaded to a temporary folder, a zip is expanded, and the folder
            is removed afterwards.

        .OUTPUTS
            An object with State (Exported, Pending or Failed), ReportId, Report (the report
            listing) and Rows (the imported CSV rows).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ReportEntity,
        [Parameter(Mandatory)][ValidateSet('Snapshot', 'RecentActivity')][string]$ReportType,
        [ValidateSet('SharePoint', 'OneDriveForBusiness')][string]$Workload,
        [hashtable]$StartArgument = @{},
        [hashtable]$MatchProperty = @{},
        [int]$MaxReportAgeHours = 24,
        [int]$WaitMinutes = 30,
        [int]$PollSeconds = 60,
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)][string]$Source
    )

    $name = if ($Workload) { "$ReportEntity ($Workload)" } else { $ReportEntity }

    $listing = @{ ReportEntity = $ReportEntity; ReportType = $ReportType; ErrorAction = 'Stop' }
    if ($Workload) { $listing['Workload'] = $Workload }

    $reusable = $null
    try {
        $cutoff = [datetime]::UtcNow.AddHours(-$MaxReportAgeHours)
        $reusable = @(Get-SPODataAccessGovernanceInsight @listing | Where-Object { $null -ne $_ } | Where-Object {
                $report = $_
                if ([string]$report.Status -ne 'Completed') { return $false }
                $stamp = Get-ReportTime ($(if ($report.PSObject.Properties['TriggeredDateTime']) { $report.TriggeredDateTime } else { $report.CreatedDateTime }))
                if ($null -eq $stamp -or $stamp -lt $cutoff) { return $false }
                foreach ($key in $MatchProperty.Keys) {
                    $property = $report.PSObject.Properties[$key]
                    if (-not $property -or [string]$property.Value -ne [string]$MatchProperty[$key]) { return $false }
                }
                return $true
            } | Sort-Object { Get-ReportTime $_.CreatedDateTime } -Descending | Select-Object -First 1)
    }
    catch {
        Write-Verbose ("No earlier {0} report could be listed: {1}" -f $name, $_.Exception.Message)
    }

    if ($reusable.Count -gt 0) {
        $reportId = [string]$reusable[0].ReportId
        Write-CollectorLog -OutputPath $OutputPath -Source $Source -Message (
            "Reusing the completed $name report $reportId (no older than $MaxReportAgeHours hours).")
    }
    else {
        $start = @{ ReportEntity = $ReportEntity; ReportType = $ReportType; ErrorAction = 'Stop' }
        if ($Workload) { $start['Workload'] = $Workload }
        foreach ($key in $StartArgument.Keys) { $start[$key] = $StartArgument[$key] }

        $started = Start-SPODataAccessGovernanceInsight @start
        $reportId = [string](Get-JsonProperty -Object $started -Name 'ReportId')
        if ([string]::IsNullOrWhiteSpace($reportId)) {
            throw "Start-SPODataAccessGovernanceInsight returned no ReportId for $name."
        }
        Write-CollectorLog -OutputPath $OutputPath -Source $Source -Message "Started the $name report $reportId."
    }

    $deadline = [datetime]::UtcNow.AddMinutes($WaitMinutes)
    while ($true) {
        $report = @(Get-SPODataAccessGovernanceInsight -ReportID $reportId -ErrorAction Stop | Where-Object { $null -ne $_ } | Select-Object -First 1)[0]
        $status = [string](Get-JsonProperty -Object $report -Name 'Status')

        if ($status -eq 'Completed') { break }

        if ($status -match 'Fail|Error|Cancel') {
            Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $Source -Message (
                "The $name report $reportId ended with status '$status'.")
            return [pscustomobject]@{ State = 'Failed'; ReportId = $reportId; Report = $report; Rows = @() }
        }

        if ([datetime]::UtcNow -ge $deadline) {
            Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $Source -Message (
                "The $name report $reportId is '$status' after $WaitMinutes minutes. It is left running; a later run exports it.")
            return [pscustomobject]@{ State = 'Pending'; ReportId = $reportId; Report = $report; Rows = @() }
        }

        Start-Sleep -Seconds $PollSeconds
    }

    $folder = Join-Path ([System.IO.Path]::GetTempPath()) ('oversharing-dag-' + [guid]::NewGuid().ToString('N'))
    try {
        New-Item -Path $folder -ItemType Directory -Force | Out-Null
        Export-SPODataAccessGovernanceInsight -ReportID $reportId -DownloadPath $folder -ErrorAction Stop | Out-Null

        $zipIndex = 0
        foreach ($zip in @(Get-ChildItem -LiteralPath $folder -Recurse -File -Filter '*.zip')) {
            $zipIndex++
            Expand-Archive -LiteralPath $zip.FullName -DestinationPath (Join-Path $folder "expanded-$zipIndex") -Force
        }

        $files = @(Get-ChildItem -LiteralPath $folder -Recurse -File -Filter '*.csv')
        if ($files.Count -eq 0) {
            throw "Exporting the $name report $reportId produced no CSV file."
        }
        $rows = foreach ($file in $files) { Import-Csv -LiteralPath $file.FullName }
    }
    finally {
        Remove-Item -LiteralPath $folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    [pscustomobject]@{ State = 'Exported'; ReportId = $reportId; Report = $report; Rows = @($rows) }
}

function ConvertTo-ReportRowJson {
    <#
        .SYNOPSIS
            An exported CSV row as one compact JSON string, for CSVs whose columns Learn does
            not list.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)]$Row)

    $ordered = [ordered]@{}
    foreach ($property in $Row.PSObject.Properties) { $ordered[$property.Name] = $property.Value }
    return ($ordered | ConvertTo-Json -Compress -Depth 3)
}

#endregion

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
    foreach ($key in @('moreRecordsAvailable', 'MoreRecordsAvailable')) {
        $value = Get-JsonProperty -Object $meta -Name $key
        if ($null -ne $value) { $flag = $value; break }
    }

    if ($null -eq $flag) { return $false }
    if ($flag -is [bool]) { return $flag }

    $parsed = $false
    if ([bool]::TryParse([string]$flag, [ref]$parsed)) { return $parsed }
    return $false
}

function ConvertTo-AuditQueryDate {
    <#
        .SYNOPSIS
            A Search-UnifiedAuditLog date in UTC.

        .DESCRIPTION
            -StartDate and -EndDate are stored in UTC. A value with no time zone is
            midnight UTC; converting it from the machine's local zone would shift the window.
            https://learn.microsoft.com/powershell/module/exchangepowershell/search-unifiedauditlog
    #>
    [CmdletBinding()]
    [OutputType([datetime])]
    param([Parameter(Mandatory)][datetime]$Value)

    if ($Value.Kind -eq [DateTimeKind]::Unspecified) {
        return [datetime]::SpecifyKind($Value, [DateTimeKind]::Utc)
    }
    return $Value.ToUniversalTime()
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
            TruncatedWindow instead. The sharing operations are filtered with -Operations;
            searching SharePoint sharing by -RecordType was not confirmed on a Learn page.
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
        [Parameter(Mandatory)][string[]]$Operation,
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
        $sessionId = 'oversharing-{0:yyyyMMddHHmmss}-{1}' -f $window.Start, [guid]::NewGuid().ToString('N').Substring(0, 8)
        $collected = 0
        $nullTries = 0
        $windowRecords = [System.Collections.Generic.List[object]]::new()
        $windowTruncated = $false

        while ($collected -lt $sessionCap) {
            $search = @{
                StartDate      = $window.Start
                EndDate        = $window.End
                Operations     = $Operation
                SessionId      = $sessionId
                SessionCommand = 'ReturnLargeSet'
                ResultSize     = $pageSize
                # Without -Formatted, RecordType is an integer. The activities page uses
                # the display name, such as SharePointSharingOperation.
                # https://learn.microsoft.com/powershell/module/exchangepowershell/search-unifiedauditlog
                Formatted      = $true
                ErrorAction    = 'Stop'
            }

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

function ConvertTo-SharingAuditRow {
    <#
        .SYNOPSIS
            One CSV row for an audit record from Invoke-AuditSearch.

        .DESCRIPTION
            The SharePoint, file and sharing fields are the AuditData properties of the
            SharePoint base, file operation and sharing schemas
            (https://learn.microsoft.com/office/office-365-management-api/office-365-management-activity-api-schema).
            A field a record does not carry is left empty.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Item)

    $audit = $Item.Audit

    $recordType = $Item.RecordType
    if (-not $recordType) { $recordType = [string](Get-GraphAdditionalProperty -Object $audit -Name 'RecordType') }

    [pscustomobject]@{
        CreationTime          = ConvertTo-CsvTimestamp (Get-GraphAdditionalProperty -Object $audit -Name 'CreationTime')
        Id                    = [string](Get-GraphAdditionalProperty -Object $audit -Name 'Id')
        RecordType            = $recordType
        Operation             = [string](Get-GraphAdditionalProperty -Object $audit -Name 'Operation')
        UserId                = [string](Get-GraphAdditionalProperty -Object $audit -Name 'UserId')
        Workload              = [string](Get-GraphAdditionalProperty -Object $audit -Name 'Workload')
        ObjectId              = [string](Get-GraphAdditionalProperty -Object $audit -Name 'ObjectId')
        ItemType              = [string](Get-GraphAdditionalProperty -Object $audit -Name 'ItemType')
        SiteUrl               = [string](Get-GraphAdditionalProperty -Object $audit -Name 'SiteUrl')
        SourceRelativeUrl     = [string](Get-GraphAdditionalProperty -Object $audit -Name 'SourceRelativeUrl')
        SourceFileName        = [string](Get-GraphAdditionalProperty -Object $audit -Name 'SourceFileName')
        TargetUserOrGroupName = [string](Get-GraphAdditionalProperty -Object $audit -Name 'TargetUserOrGroupName')
        TargetUserOrGroupType = [string](Get-GraphAdditionalProperty -Object $audit -Name 'TargetUserOrGroupType')
        ClientIP              = [string](Get-GraphAdditionalProperty -Object $audit -Name 'ClientIP')
    }
}

#endregion
