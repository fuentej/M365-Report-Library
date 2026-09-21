#Requires -Version 7.0

<#
    M365ReportLibrary - shared layer for every report in this library.

    Read-only by design: the only cmdlets this module runs against a tenant are
    Connect-*, Disconnect-*, Get-* and Search-*.
#>

Set-StrictMode -Version Latest

#region Cloud endpoints

# Microsoft Graph national cloud deployments:
#   https://learn.microsoft.com/graph/deployments
# Exchange Online app-only / environment names:
#   https://learn.microsoft.com/powershell/exchange/app-only-auth-powershell-v2
# Security & Compliance PowerShell connection URIs:
#   https://learn.microsoft.com/powershell/exchange/connect-to-scc-powershell
$script:M365Endpoints = @{
    Graph = @{
        Commercial = @{ GraphEnvironment = 'Global'; ResourceEndpoint = 'https://graph.microsoft.com' }
        GCC        = @{ GraphEnvironment = 'Global'; ResourceEndpoint = 'https://graph.microsoft.com' }
        GCCHigh    = @{ GraphEnvironment = 'USGov';  ResourceEndpoint = 'https://graph.microsoft.us' }
    }
    ExchangeOnline = @{
        Commercial = @{ ExchangeEnvironmentName = 'O365Default' }
        GCC        = @{ ExchangeEnvironmentName = 'O365Default' }
        GCCHigh    = @{ ExchangeEnvironmentName = 'O365USGovGCCHigh' }
    }
    SecurityCompliance = @{
        Commercial = @{ ConnectionUri = $null; AzureADAuthorizationEndpointUri = $null }
        GCC        = @{ ConnectionUri = $null; AzureADAuthorizationEndpointUri = $null }
        GCCHigh    = @{
            ConnectionUri                   = 'https://ps.compliance.protection.office365.us/powershell-liveid/'
            AzureADAuthorizationEndpointUri = 'https://login.microsoftonline.us/organizations'
        }
    }
}

# Least-privilege read scopes used when signing in interactively.
$script:DefaultGraphScopes = @(
    'User.Read.All'
    'Directory.Read.All'
    'GroupMember.Read.All'
    'AuditLog.Read.All'
)

$script:EntraUserCsvColumns = @(
    'RunDate'
    'Id'
    'DisplayName'
    'UserPrincipalName'
    'Mail'
    'UserType'
    'AccountEnabled'
    'CreatedDateTime'
    'Department'
    'JobTitle'
    'City'
    'Country'
    'ManagerId'
    'ManagerUserPrincipalName'
)

function Get-M365ServiceEndpoint {
    <#
        .SYNOPSIS
            The connection settings for a service in a given cloud.

        .DESCRIPTION
            Returns the parameters Connect-M365Service splats onto the underlying
            Connect-* cmdlet. Kept separate from Connect-M365Service so that the
            endpoint choice can be asserted without a tenant.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Graph', 'ExchangeOnline', 'SecurityCompliance')]
        [string]$Service,

        [ValidateSet('Commercial', 'GCC', 'GCCHigh')]
        [string]$Environment = 'Commercial'
    )

    $settings = $script:M365Endpoints[$Service][$Environment]

    $result = @{
        Service     = $Service
        Environment = $Environment
    }
    foreach ($key in $settings.Keys) {
        $result[$key] = $settings[$key]
    }
    return $result
}

function Get-EntraUserCsvColumn {
    <#
        .SYNOPSIS
            The column order of users.csv.
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param()
    return , [string[]]$script:EntraUserCsvColumns
}

function Get-DefaultGraphScope {
    <#
        .SYNOPSIS
            The read-only Graph scopes requested on an interactive sign-in.
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param()
    return , [string[]]$script:DefaultGraphScopes
}

#endregion

#region Connection

function Connect-M365Service {
    <#
        .SYNOPSIS
            Signs in to Microsoft Graph, Exchange Online or Security & Compliance
            PowerShell in the Commercial, GCC or GCC High cloud.

        .DESCRIPTION
            Interactive sign-in by default. Passing -AppId and -CertificateThumbprint
            (plus -TenantId, and -Organization for the Exchange-based services) signs in
            app-only instead.

        .PARAMETER Organization
            The tenant's *.onmicrosoft.com domain. Required for app-only sign-in to
            Exchange Online and Security & Compliance PowerShell.

        .EXAMPLE
            Connect-M365Service -Service Graph -Environment GCCHigh
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Graph', 'ExchangeOnline', 'SecurityCompliance')]
        [string]$Service,

        [ValidateSet('Commercial', 'GCC', 'GCCHigh')]
        [string]$Environment = 'Commercial',

        [string]$AppId,
        [string]$CertificateThumbprint,
        [string]$TenantId,
        [string]$Organization,

        [string[]]$Scopes
    )

    $endpoint = Get-M365ServiceEndpoint -Service $Service -Environment $Environment

    # Half a credential is a mistake worth stopping on. Treating it as "no credential"
    # would quietly fall back to an interactive prompt, which on an unattended run
    # either hangs or signs in as whoever is at the keyboard.
    $hasAppId = -not [string]::IsNullOrWhiteSpace($AppId)
    $hasCertificate = -not [string]::IsNullOrWhiteSpace($CertificateThumbprint)
    if ($hasAppId -ne $hasCertificate) {
        $supplied = if ($hasAppId) { '-AppId' } else { '-CertificateThumbprint' }
        $missing = if ($hasAppId) { '-CertificateThumbprint' } else { '-AppId' }
        throw "App-only sign-in needs both -AppId and -CertificateThumbprint; $supplied was given without $missing. Omit both to sign in interactively."
    }
    $appOnly = $hasAppId -and $hasCertificate

    switch ($Service) {
        'Graph' {
            $params = @{
                Environment = $endpoint.GraphEnvironment
                NoWelcome   = $true
            }
            if ($appOnly) {
                if ([string]::IsNullOrWhiteSpace($TenantId)) {
                    throw 'App-only sign-in to Microsoft Graph requires -TenantId.'
                }
                $params['ClientId'] = $AppId
                $params['CertificateThumbprint'] = $CertificateThumbprint
                $params['TenantId'] = $TenantId
            }
            else {
                $params['Scopes'] = if ($Scopes) { $Scopes } else { Get-DefaultGraphScope }
                if (-not [string]::IsNullOrWhiteSpace($TenantId)) {
                    $params['TenantId'] = $TenantId
                }
            }
            Connect-MgGraph @params
        }

        'ExchangeOnline' {
            $params = @{
                ExchangeEnvironmentName = $endpoint.ExchangeEnvironmentName
                ShowBanner              = $false
            }
            if ($appOnly) {
                if ([string]::IsNullOrWhiteSpace($Organization)) {
                    throw 'App-only sign-in to Exchange Online requires -Organization (the tenant''s *.onmicrosoft.com domain).'
                }
                $params['AppId'] = $AppId
                $params['CertificateThumbprint'] = $CertificateThumbprint
                $params['Organization'] = $Organization
            }
            Connect-ExchangeOnline @params
        }

        'SecurityCompliance' {
            $params = @{ ShowBanner = $false }
            if ($null -ne $endpoint.ConnectionUri) {
                $params['ConnectionUri'] = $endpoint.ConnectionUri
                $params['AzureADAuthorizationEndpointUri'] = $endpoint.AzureADAuthorizationEndpointUri
            }
            if ($appOnly) {
                if ([string]::IsNullOrWhiteSpace($Organization)) {
                    throw 'App-only sign-in to Security & Compliance PowerShell requires -Organization (the tenant''s *.onmicrosoft.com domain).'
                }
                $params['AppId'] = $AppId
                $params['CertificateThumbprint'] = $CertificateThumbprint
                $params['Organization'] = $Organization
            }
            Connect-IPPSSession @params
        }
    }
}

#endregion

#region CSV output

function Write-CollectorLog {
    <#
        .SYNOPSIS
            Appends a line to run.log in the output folder.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$OutputPath,

        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Message,

        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Level = 'Info',

        [string]$Source
    )

    if (-not (Test-Path -LiteralPath $OutputPath)) {
        New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
    }

    $stamp = [datetime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ')
    $prefix = if ([string]::IsNullOrWhiteSpace($Source)) { '' } else { "[$Source] " }
    $line = "$stamp [$Level] $prefix$Message"

    Add-Content -LiteralPath (Join-Path $OutputPath 'run.log') -Value $line -Encoding utf8

    switch ($Level) {
        'Warning' { Write-Warning $Message }
        'Error' { Write-Warning $Message }
        default { Write-Verbose $line }
    }
}

function Get-CsvHeaderColumn {
    <#
        .SYNOPSIS
            The column names of an existing CSV, in file order.
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return $null
    }

    $headerLine = Get-Content -LiteralPath $Path -TotalCount 1
    if ([string]::IsNullOrWhiteSpace($headerLine)) {
        return $null
    }

    # Parse with the real CSV reader rather than splitting on commas: feeding the
    # header twice yields one row whose property names are the header fields.
    $parsed = @($headerLine, $headerLine) | ConvertFrom-Csv
    return , [string[]]@($parsed[0].PSObject.Properties.Name)
}

function Export-AppendCsv {
    <#
        .SYNOPSIS
            Appends rows to a CSV, creating the header once and skipping keys the
            file already holds.

        .DESCRIPTION
            On the first write the header is created from -Column, or from the first
            row's properties when -Column is omitted. On later writes the incoming
            columns must match the file's header exactly, in the same order, or the
            function throws. Passing no rows creates a header-only file, which is how
            a collector records a source that is unavailable or unlicensed.

        .PARAMETER KeyColumn
            One or more columns whose combined value identifies a row. Rows whose key
            is already in the file, or repeated within this batch, are skipped.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter()]
        [AllowNull()]
        [AllowEmptyCollection()]
        [object[]]$Rows,

        [string[]]$KeyColumn,

        [string[]]$Column,

        [switch]$PassThru
    )

    $rowList = @($Rows | Where-Object { $null -ne $_ })

    $schema = if ($Column) {
        [string[]]$Column
    }
    elseif ($rowList.Count -gt 0) {
        [string[]]@($rowList[0].PSObject.Properties.Name)
    }
    else {
        $null
    }

    $existingHeader = Get-CsvHeaderColumn -Path $Path

    if ($null -eq $schema) {
        if ($null -ne $existingHeader) {
            # Nothing to write and the file already has its header.
            if ($PassThru) {
                return [pscustomobject]@{ Path = $Path; Written = 0; Skipped = 0; HeaderCreated = $false }
            }
            return
        }
        throw "Export-AppendCsv cannot determine the columns for '$Path': pass -Column or at least one row."
    }

    if ($null -ne $existingHeader -and (Compare-Object -ReferenceObject $existingHeader -DifferenceObject $schema -SyncWindow 0)) {
        throw ("Column mismatch for '$Path'. File header: {0}. Incoming columns: {1}." -f ($existingHeader -join ', '), ($schema -join ', '))
    }

    $parent = Split-Path -Path $Path -Parent
    if ($parent -and -not (Test-Path -LiteralPath $parent)) {
        New-Item -Path $parent -ItemType Directory -Force | Out-Null
    }

    $headerCreated = $false
    if ($null -eq $existingHeader) {
        $headerText = (($schema | ForEach-Object { '"' + ($_ -replace '"', '""') + '"' }) -join ',')
        Set-Content -LiteralPath $Path -Value $headerText -Encoding utf8
        $headerCreated = $true
    }

    $written = 0
    $skipped = 0

    if ($rowList.Count -gt 0) {
        $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

        if ($KeyColumn) {
            foreach ($missing in $KeyColumn | Where-Object { $schema -notcontains $_ }) {
                throw "Key column '$missing' is not one of the columns of '$Path'."
            }
            if (-not $headerCreated) {
                foreach ($existing in Import-Csv -LiteralPath $Path) {
                    [void]$seen.Add((Get-RowKey -Row $existing -KeyColumn $KeyColumn))
                }
            }
        }

        $toWrite = [System.Collections.Generic.List[object]]::new()
        foreach ($row in $rowList) {
            $shaped = $row | Select-Object -Property $schema
            if ($KeyColumn) {
                $key = Get-RowKey -Row $shaped -KeyColumn $KeyColumn
                if (-not $seen.Add($key)) {
                    $skipped++
                    continue
                }
            }
            $toWrite.Add($shaped)
        }

        if ($toWrite.Count -gt 0) {
            $lines = $toWrite | ConvertTo-Csv -NoTypeInformation
            Add-Content -LiteralPath $Path -Value ($lines | Select-Object -Skip 1) -Encoding utf8
            $written = $toWrite.Count
        }
    }

    if ($PassThru) {
        return [pscustomobject]@{
            Path          = $Path
            Written       = $written
            Skipped       = $skipped
            HeaderCreated = $headerCreated
        }
    }
}

function Get-RowKey {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [object]$Row,

        [Parameter(Mandatory)]
        [string[]]$KeyColumn
    )

    $parts = foreach ($name in $KeyColumn) {
        $value = $Row.PSObject.Properties[$name]
        if ($null -eq $value -or $null -eq $value.Value) { '' } else { [string]$value.Value }
    }
    # Unit separator: cannot appear in a CSV field value we produce.
    return ($parts -join [char]0x1F)
}

function Get-CsvWatermark {
    <#
        .SYNOPSIS
            The latest timestamp already collected into a CSV, or $null.

        .DESCRIPTION
            Event collectors start their query at this value so a re-run picks up
            where the last one stopped. Returns $null when the file is missing,
            holds only a header, or holds no parseable timestamp.
    #>
    [CmdletBinding()]
    [OutputType([nullable[datetime]])]
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$Column
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return $null
    }

    $header = Get-CsvHeaderColumn -Path $Path
    if ($null -eq $header -or $header -notcontains $Column) {
        return $null
    }

    $latest = $null
    foreach ($row in Import-Csv -LiteralPath $Path) {
        $raw = $row.$Column
        if ([string]::IsNullOrWhiteSpace($raw)) { continue }

        [datetime]$parsed = [datetime]::MinValue
        $styles = [System.Globalization.DateTimeStyles]::AdjustToUniversal -bor [System.Globalization.DateTimeStyles]::AssumeUniversal
        if (-not [datetime]::TryParse($raw, [System.Globalization.CultureInfo]::InvariantCulture, $styles, [ref]$parsed)) {
            continue
        }
        if ($null -eq $latest -or $parsed -gt $latest) {
            $latest = $parsed
        }
    }

    return $latest
}

function Get-CsvLatestSnapshot {
    <#
        .SYNOPSIS
            The rows of the most recent snapshot in an appended snapshot CSV.

        .DESCRIPTION
            Snapshot CSVs hold one block of rows per run, tagged with RunDate. A
            collector that needs "the guests as of the last run" reads them from here.
            Returns nothing when the file is missing or holds only a header.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [string]$SnapshotColumn = 'RunDate'
    )

    if (-not (Test-Path -LiteralPath $Path)) { return }

    $rows = @(Import-Csv -LiteralPath $Path)
    if ($rows.Count -eq 0) { return }

    $header = Get-CsvHeaderColumn -Path $Path
    if ($header -notcontains $SnapshotColumn) {
        return $rows
    }

    $latest = ($rows | ForEach-Object { [string]$_.$SnapshotColumn } | Sort-Object -Descending | Select-Object -First 1)
    return @($rows | Where-Object { [string]$_.$SnapshotColumn -eq $latest })
}

#endregion

#region Collectors

function Invoke-EntraUserCollector {
    <#
        .SYNOPSIS
            Writes users.csv: a snapshot of every Entra ID user, appended each run.

        .DESCRIPTION
            Source: Microsoft Graph list users (https://learn.microsoft.com/graph/api/user-list),
            with the manager expanded. If the tenant or cloud rejects the expand, the
            manager is read per user instead. If the source is unavailable altogether the
            file is created with its header only and the reason is logged.

        .PARAMETER SkipConnect
            Use an existing Graph session instead of signing in.
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

    $columns = Get-EntraUserCsvColumn
    $csvPath = Join-Path $OutputPath 'users.csv'
    $runDate = [datetime]::UtcNow.ToString('yyyy-MM-dd')
    $source = 'users'

    if (-not (Test-Path -LiteralPath $OutputPath)) {
        New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
    }

    if (-not $SkipConnect) {
        Connect-M365Service -Service Graph -Environment $Environment -AppId $AppId `
            -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization
    }

    $select = @(
        'id', 'displayName', 'userPrincipalName', 'mail', 'userType', 'accountEnabled'
        'createdDateTime', 'department', 'jobTitle', 'city', 'country'
    )

    $users = $null
    $expanded = $true
    try {
        $users = @(Get-MgUser -All -Property $select -ExpandProperty 'manager($select=id,userPrincipalName)' -ErrorAction Stop)
    }
    catch {
        Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
            "Listing users with the manager expanded failed ({0}). Retrying without the expand." -f $_.Exception.Message)
        $expanded = $false
        try {
            $users = @(Get-MgUser -All -Property $select -ErrorAction Stop)
        }
        catch {
            Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
                "Microsoft Graph list users is unavailable in this tenant or cloud ({0}). Writing the header only." -f $_.Exception.Message)
            Export-AppendCsv -Path $csvPath -Column $columns
            return
        }
    }

    $rows = foreach ($user in $users) {
        $managerId = $null
        $managerUpn = $null

        if ($expanded) {
            $managerProperty = $user.PSObject.Properties['Manager']
            if ($managerProperty -and $null -ne $managerProperty.Value) {
                $manager = $managerProperty.Value
                $managerId = $manager.Id
                $managerUpn = Get-GraphAdditionalProperty -Object $manager -Name 'userPrincipalName'
            }
        }
        else {
            try {
                $manager = Get-MgUserManager -UserId $user.Id -ErrorAction Stop
                if ($manager) {
                    $managerId = $manager.Id
                    $managerUpn = Get-GraphAdditionalProperty -Object $manager -Name 'userPrincipalName'
                }
            }
            catch {
                # A user with no manager returns 404; anything else is logged once per user.
                Write-Verbose ("No manager read for {0}: {1}" -f $user.Id, $_.Exception.Message)
            }
        }

        [pscustomobject]@{
            RunDate                  = $runDate
            Id                       = $user.Id
            DisplayName              = $user.DisplayName
            UserPrincipalName        = $user.UserPrincipalName
            Mail                     = $user.Mail
            UserType                 = $user.UserType
            AccountEnabled           = $user.AccountEnabled
            CreatedDateTime          = ConvertTo-CsvTimestamp $user.CreatedDateTime
            Department               = $user.Department
            JobTitle                 = $user.JobTitle
            City                     = $user.City
            Country                  = $user.Country
            ManagerId                = $managerId
            ManagerUserPrincipalName = $managerUpn
        }
    }

    $result = Export-AppendCsv -Path $csvPath -Rows @($rows) -Column $columns -KeyColumn @('RunDate', 'Id') -PassThru
    Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
        "users.csv: {0} rows written, {1} skipped." -f $result.Written, $result.Skipped)
}

#endregion

#region Helpers

function ConvertTo-CsvTimestamp {
    <#
        .SYNOPSIS
            Formats a timestamp as round-trippable UTC, or an empty string.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Position = 0)]
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value -or ($Value -is [string] -and [string]::IsNullOrWhiteSpace($Value))) {
        return ''
    }

    [datetime]$parsed = [datetime]::MinValue
    if ($Value -is [datetime]) {
        $parsed = $Value
    }
    elseif ($Value -is [datetimeoffset]) {
        $parsed = $Value.UtcDateTime
    }
    else {
        $styles = [System.Globalization.DateTimeStyles]::AdjustToUniversal -bor [System.Globalization.DateTimeStyles]::AssumeUniversal
        if (-not [datetime]::TryParse([string]$Value, [System.Globalization.CultureInfo]::InvariantCulture, $styles, [ref]$parsed)) {
            return [string]$Value
        }
    }

    return $parsed.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
}

function Get-GraphAdditionalProperty {
    <#
        .SYNOPSIS
            Reads a property that the Graph SDK may have placed in AdditionalProperties.

        .DESCRIPTION
            The SDK maps unselected or loosely typed fields into AdditionalProperties,
            keyed by their camelCase Graph name. This looks in both places.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowNull()]
        [object]$Object,

        [Parameter(Mandatory)]
        [string]$Name
    )

    if ($null -eq $Object) { return $null }

    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name)) { return $Object[$Name] }
        return $null
    }

    $pascal = $Name.Substring(0, 1).ToUpperInvariant() + $Name.Substring(1)
    foreach ($candidate in @($pascal, $Name)) {
        $property = $Object.PSObject.Properties[$candidate]
        if ($property -and $null -ne $property.Value) {
            return $property.Value
        }
    }

    $additional = $Object.PSObject.Properties['AdditionalProperties']
    if ($additional -and $additional.Value -is [System.Collections.IDictionary] -and $additional.Value.Contains($Name)) {
        return $additional.Value[$Name]
    }

    return $null
}

function Split-DateRange {
    <#
        .SYNOPSIS
            Splits a date range into windows, so no single query is asked for more
            than a service will return.

        .OUTPUTS
            One object per window, with Start and End.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [datetime]$Start,

        [Parameter(Mandatory)]
        [datetime]$End,

        [ValidateRange(1, 1440)]
        [int]$WindowMinutes = 1440
    )

    if ($End -le $Start) { return }

    $windowSpan = [timespan]::FromMinutes($WindowMinutes)
    $cursor = $Start
    while ($cursor -lt $End) {
        $stop = $cursor.Add($windowSpan)
        if ($stop -gt $End) { $stop = $End }
        [pscustomobject]@{ Start = $cursor; End = $stop }
        $cursor = $stop
    }
}

#endregion

Export-ModuleMember -Function @(
    'Connect-M365Service'
    'ConvertTo-CsvTimestamp'
    'Export-AppendCsv'
    'Get-CsvHeaderColumn'
    'Get-CsvLatestSnapshot'
    'Get-CsvWatermark'
    'Get-DefaultGraphScope'
    'Get-EntraUserCsvColumn'
    'Get-GraphAdditionalProperty'
    'Get-M365ServiceEndpoint'
    'Invoke-EntraUserCollector'
    'Split-DateRange'
    'Write-CollectorLog'
)
