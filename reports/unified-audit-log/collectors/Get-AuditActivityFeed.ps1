#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes audit-activity-feed.csv: unified audit log events from the Office 365 Management
        Activity API, appended from where the last run stopped.

    .DESCRIPTION
        Source 3. Per content type: GET {root}/api/v1.0/{tenant}/activity/feed/subscriptions/list,
        then, for a type with no enabled subscription, POST .../subscriptions/start?contentType=...
        (unless -NoStartSubscription), then GET .../subscriptions/content for each window of at most
        24 hours, paged through the NextPageUri header, then GET each contentUri blob
        (https://learn.microsoft.com/office/office-365-management-api/office-365-management-activity-api-reference).
        The start call is a POST; it does not change tenant data. Decision D-007 allows it. It is the
        only POST this collector sends, it is sent only for a type that is not enabled, and a
        subscription is never stopped (stopping and starting again loses the content from the gap).
        A second start within 15 minutes is throttled, and the first blobs of a new subscription can
        take up to 12 hours.

        This is a 7-day feed. startTime and endTime select on when a blob became available
        (contentCreated), not on the event time; startTime is inclusive and endTime exclusive, and
        startTime cannot be more than 7 days back (error AF20030). A run starts at the latest
        ContentCreated already in the file, or -LookbackDays back on the first run, clamped to the last
        7 days; a gap older than that is logged and cannot be recovered. A window is written only when
        every content type in it has been read, so a failure in one type does not move the watermark
        past it. A blob whose contentExpiration has passed is skipped and logged; the API will not
        return it, and failing the window on it would stall every later blob. Rows are keyed on
        the event Id.

        Sign-in: the collector does not acquire a token. Pass -AccessToken, a SecureString holding a
        token for the feed root (Commercial https://manage.office.com, GCC https://manage-gcc.office.com,
        GCC High https://manage.office365.us) whose tenant ID is -TenantId and which carries the
        ActivityFeed.Read claim, from an Entra app with the application permission "Read activity data
        for an organization". DLP.All also needs "Read DLP sensitive data". The token is read from the
        SecureString only to build the Authorization header and is never logged.
        -PublisherIdentifier is the GUID of the tenant of whoever runs this code (not the customer tenant
        and not the app ID); a request without it shares one quota with every caller.

        Available in all three clouds (a root URL is documented for each).

    .EXAMPLE
        $token = Read-Host -AsSecureString 'Token'
        ./Get-AuditActivityFeed.ps1 -OutputPath ./out -TenantId $tenantId -PublisherIdentifier $publisherId -AccessToken $token
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$OutputPath,

    [ValidateSet('Commercial', 'GCC', 'GCCHigh')]
    [string]$Environment = 'Commercial',

    [Parameter(Mandatory)]
    [string]$TenantId,

    [Parameter(Mandatory)]
    [string]$PublisherIdentifier,

    [Parameter(Mandatory)]
    [securestring]$AccessToken,

    # Audit.General holds every workload that is not Entra, Exchange or SharePoint. DLP.All needs an
    # extra permission and is not read by default.
    [ValidateSet('Audit.AzureActiveDirectory', 'Audit.Exchange', 'Audit.SharePoint', 'Audit.General', 'DLP.All')]
    [string[]]$ContentType = @('Audit.AzureActiveDirectory', 'Audit.Exchange', 'Audit.SharePoint', 'Audit.General'),

    # How far back the first run looks, in days (1 to 7: the feed holds no more).
    [ValidateRange(1, 7)]
    [int]$LookbackDays = 7,

    # Never send the subscription start. A type that is not enabled is skipped with a warning.
    [switch]$NoStartSubscription
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Imported by path without -Force: a forced re-import builds a fresh module session
# state, which would discard anything the caller has arranged around the module - and
# that is exactly how the tests put mocks in front of the tenant cmdlets.
Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

. (Join-Path $PSScriptRoot 'UnifiedAuditLogHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'UnifiedAuditLogSchema.psd1')
$columns = $schema.AuditActivityFeed
$source = 'audit-activity-feed'
$csvPath = Join-Path $OutputPath 'audit-activity-feed.csv'

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

if (Test-AuditSourceSkipped -Source 'AuditActivityFeed' -LogSource $source -CsvPath $csvPath -Column $columns `
        -OutputPath $OutputPath -Environment $Environment -Schema $schema) {
    return
}

$parsedPublisher = [guid]::Empty
if (-not [guid]::TryParse($PublisherIdentifier, [ref]$parsedPublisher)) {
    throw "PublisherIdentifier '$PublisherIdentifier' is not a GUID."
}

$base = Get-ActivityFeedBase -Environment $Environment -TenantId $TenantId -Schema $schema
$feed = @{ Base = $base; AccessToken = $AccessToken; PublisherIdentifier = $PublisherIdentifier }

Export-AppendCsv -Path $csvPath -Column $columns
$total = [pscustomobject]@{ Written = 0; Skipped = 0 }

try {
    $enabled = Get-AuditActivitySubscription @feed
    $types = foreach ($type in $ContentType) {
        if ($enabled -contains $type) { $type; continue }
        if ($NoStartSubscription) {
            Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
                "$type has no enabled subscription and -NoStartSubscription was passed. Skipping it.")
            continue
        }
        Start-AuditActivitySubscription -Base $base -ContentType $type -AccessToken $AccessToken -PublisherIdentifier $PublisherIdentifier | Out-Null
        Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
            "Started the $type subscription. The first blobs can take up to 12 hours to appear.")
        $type
    }
    $types = @($types)

    $end = [datetime]::UtcNow
    $earliest = $end.AddDays(-7).AddMinutes(5)
    $start = Get-AuditRunStart -CsvPath $csvPath -WatermarkColumn 'ContentCreated' -End $end -LookbackDays $LookbackDays
    if ($start -lt $earliest) {
        Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
            'The last export is older than the feed holds. Starting at {0:yyyy-MM-ddTHH:mm:ssZ}; content before it can no longer be retrieved (AF20051).' -f $earliest)
        $start = $earliest
    }

    foreach ($window in Split-DateRange -Start $start -End $end -WindowMinutes 1440) {
        $rows = [System.Collections.Generic.List[object]]::new()
        foreach ($type in $types) {
            foreach ($blob in Get-AuditActivityContent @feed -ContentType $type -Start $window.Start -End $window.End) {
                $expires = Get-ObjectValue -Object $blob -Name 'contentExpiration'
                if ($null -ne $expires -and "$expires" -ne '') {
                    $expiry = ConvertTo-AuditUtc $expires
                    if ($expiry -le [datetime]::UtcNow) {
                        Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
                            'Skipping content {0}: its contentExpiration {1:yyyy-MM-ddTHH:mm:ssZ} has passed, so the API will not return it.' -f (Get-ObjectText -Object $blob -Name 'contentId'), $expiry)
                        continue
                    }
                }
                foreach ($row in Get-AuditActivityEvent -Blob $blob -AccessToken $AccessToken -Base $base -PublisherIdentifier $PublisherIdentifier) {
                    $rows.Add($row)
                }
            }
        }
        $result = Export-AppendCsv -Path $csvPath -Rows $rows.ToArray() -Column $columns -KeyColumn @('RecordId') -PassThru
        $total.Written += $result.Written
        $total.Skipped += $result.Skipped
    }
}
catch {
    Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
        'The Management Activity API failed ({0}). It needs a token for the feed root with the ActivityFeed.Read claim and a tenant ID that matches -TenantId. Windows already written are kept; the next run resumes from the latest ContentCreated.' -f $_.Exception.Message)
    throw
}

Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'audit-activity-feed.csv: {0} rows written, {1} skipped. Window {2:yyyy-MM-ddTHH:mm:ssZ} to {3:yyyy-MM-ddTHH:mm:ssZ}.' -f $total.Written, $total.Skipped, $start, $end)
