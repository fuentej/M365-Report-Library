#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes activity-explorer-events.csv: sensitivity label, protection, DLP,
        retention and Copilot/AI activity, appended each run.

    .DESCRIPTION
        Source: Export-ActivityExplorerData in Security & Compliance PowerShell
        (https://learn.microsoft.com/powershell/module/exchangepowershell/export-activityexplorerdata),
        covering: sensitivity label applied/changed/removed, protection
        changes, DLP matches/enforcements/overrides with their justification,
        retention label changes, and Copilot/AI app interactions.

        Activity Explorer keeps 30 days, so each run resumes from the
        watermark already in activity-explorer-events.csv (or -RetentionDays
        ago on a first run) and appends; a run at least every 30 days builds a
        history longer than the source keeps. Records are keyed on
        RecordIdentity, so a window that overlaps the previous run's does not
        duplicate rows.

        The window is queried in slices (-SliceMinutes) at -PageSize per page,
        because a PageCookie expires after 120 seconds and Microsoft
        recommends narrower ranges over one wide one.

        Read-only: the only tenant cmdlet it calls is Export-ActivityExplorerData.

    .PARAMETER RetentionDays
        How far back to go when there is no watermark to resume from. Activity
        Explorer keeps 30 days.

    .PARAMETER PageSize
        Records per page. Export-ActivityExplorerData accepts 1-5000.

    .PARAMETER SkipConnect
        Use an existing Security & Compliance PowerShell session instead of
        signing in.

    .EXAMPLE
        ./Get-ActivityExplorerEvents.ps1 -OutputPath ./out -Environment GCCHigh

    .LINK
        https://learn.microsoft.com/powershell/module/exchangepowershell/export-activityexplorerdata
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

    [ValidateRange(1, 30)]
    [int]$RetentionDays = 30,

    # Microsoft suggests a smaller page size to avoid the PageCookie expiring
    # mid-export.
    [ValidateRange(1, 5000)]
    [int]$PageSize = 1000,

    # The window is queried in slices of this many minutes, for the same reason.
    [ValidateRange(1, 1440)]
    [int]$SliceMinutes = 1440,

    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

. (Join-Path $PSScriptRoot 'PurviewIpHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'PurviewIpSchema.psd1')
$columns = $schema.ActivityExplorerEvents
$source = 'activity-explorer-events'
$csvPath = Join-Path $OutputPath 'activity-explorer-events.csv'

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-PurviewSourceAvailability -Source 'ActivityExplorer' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping activity-explorer-events.csv. $($availability.Reason) $($availability.Reference)")
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}
if ($availability.Status -eq 'Unverified') {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message $availability.Reason
}

$connectedHere = -not $SkipConnect
if ($connectedHere) {
    Connect-M365Service -Service SecurityCompliance -Environment $Environment -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization
}

try {
    $now = [datetime]::UtcNow
    $watermark = Get-CsvWatermark -Path $csvPath -Column 'Happened'

    $start = if ($null -ne $watermark) { $watermark } else { $now.AddDays(-$RetentionDays) }
    $end = $now

    if ($end -le $start) {
        Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
            'Nothing new to collect: the watermark ({0}) is already at or after now.' -f (ConvertTo-CsvTimestamp $start))
        Export-AppendCsv -Path $csvPath -Column $columns
        return
    }

    Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
        'Querying Activity Explorer from {0} to {1} (watermark: {2}).' -f
        (ConvertTo-CsvTimestamp $start), (ConvertTo-CsvTimestamp $end),
        $(if ($null -eq $watermark) { 'none' } else { ConvertTo-CsvTimestamp $watermark }))

    $rows = [System.Collections.Generic.List[object]]::new()
    # Activity values the pipeline could not map to a filter-enum name.
    # Reported at the end of the run: they land in the Other category, and
    # silence about them would let a whole class of event quietly stop being
    # counted.
    $unmappedActivities = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    $slices = @(Split-DateRange -Start $start -End $end -WindowMinutes $SliceMinutes)
    Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
        "Split into $($slices.Count) slice(s) of up to $SliceMinutes minute(s).")

    foreach ($slice in $slices) {
        $pageCookie = $null
        $page = 0

        do {
            $page++

            $exportParameters = @{
                StartTime    = $slice.Start
                EndTime      = $slice.End
                OutputFormat = 'Json'
                PageSize     = $PageSize
            }
            if ($pageCookie) { $exportParameters['PageCookie'] = $pageCookie }

            $response = Export-ActivityExplorerData @exportParameters

            $resultData = Get-PurviewProperty $response 'ResultData'
            $records = @()
            if ($resultData) {
                # Not "$records = if (...) { @(...) } else { @(...) }": when the chosen
                # branch's pipeline emits zero objects, assigning an if/else statement's
                # captured output collapses to $null instead of an empty array, and
                # $null has no .Count. Parsing first and wrapping @() around a plain
                # variable afterwards keeps the empty case an empty array.
                $parsed = if ($resultData -is [string]) { $resultData | ConvertFrom-Json } else { $resultData }
                $records = @($parsed | Where-Object { $null -ne $_ })
            }

            Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
                "Page $page returned $($records.Count) record(s).")

            foreach ($record in $records) {
                $recordIdentity = [string](Get-PurviewProperty $record 'RecordIdentity' '')

                $happened = Get-PurviewProperty $record 'Happened'
                $happenedTimestamp = if ($happened) { ConvertTo-CsvTimestamp $happened } else { '' }

                $rawActivity = [string](Get-PurviewProperty $record 'Activity' '')
                $activityId = [string](Get-PurviewProperty $record 'ActivityId' '')
                $activity = Resolve-PurviewActivityName -Activity $rawActivity -ActivityId $activityId -Schema $schema
                if ($activity -and -not (Test-PurviewActivityName -Activity $activity -Schema $schema)) {
                    $unmappedActivities.Add($rawActivity) | Out-Null
                }

                $labelEventType = [string](Get-PurviewProperty $record 'LabelEventType' '')
                $sitSummary = Get-PurviewSensitiveInfoTypeSummary (Get-PurviewProperty $record 'SensitiveInfoTypeData')

                $rows.Add([pscustomobject]@{
                        RecordIdentity              = $recordIdentity
                        Happened                    = $happenedTimestamp
                        EventDate                   = if ($happenedTimestamp) { $happenedTimestamp.Substring(0, 10) } else { '' }
                        Activity                    = $activity
                        ActivityRaw                 = $rawActivity
                        ActivityCategory            = Get-PurviewActivityCategory -Activity $activity -Schema $schema
                        Workload                    = Get-PurviewProperty $record 'Workload'
                        Application                 = Get-PurviewProperty $record 'Application'
                        User                        = Get-PurviewProperty $record 'User'
                        UserType                    = Get-PurviewProperty $record 'UserType'
                        ItemName                    = Get-PurviewProperty $record 'ItemName'
                        FilePath                    = Get-PurviewProperty $record 'FilePath'
                        FullUrl                     = Get-PurviewProperty $record 'FullUrl'
                        FileExtension               = Get-PurviewProperty $record 'FileExtension'
                        DeviceName                  = Get-PurviewProperty $record 'DeviceName'
                        Platform                    = Get-PurviewProperty $record 'Platform'
                        SensitivityLabel            = Get-PurviewProperty $record 'SensitivityLabel'
                        OldSensitivityLabel         = Get-PurviewProperty $record 'OldSensitivityLabel'
                        SensitivityLabelPolicyId    = Get-PurviewProperty $record 'SensitivityLabelPolicyId'
                        LabelEventType              = $labelEventType
                        IsLabelDowngrade            = ($labelEventType -eq 'LabelDowngraded')
                        HowApplied                  = Get-PurviewProperty $record 'HowApplied'
                        Justification               = Get-PurviewProperty $record 'Justification'
                        IsProtected                 = Get-PurviewProperty $record 'IsProtected'
                        IsProtectedBefore           = Get-PurviewProperty $record 'IsProtectedBefore'
                        ProtectionEventType         = Get-PurviewProperty $record 'ProtectionEventType'
                        ProtectionType              = Get-PurviewProperty $record 'ProtectionType'
                        ProtectionOwner             = Get-PurviewProperty $record 'ProtectionOwner'
                        RMSEncrypted                = Get-PurviewProperty $record 'RMSEncrypted'
                        RetentionLabel              = Get-PurviewProperty $record 'RetentionLabel'
                        OldRetentionLabel           = Get-PurviewProperty $record 'OldRetentionLabel'
                        SensitiveInfoTypeName       = $sitSummary.Name
                        SensitiveInfoTypeCount      = $sitSummary.Count
                        SensitiveInfoTypeConfidence = $sitSummary.Confidence
                        PolicyId                    = Get-PurviewProperty $record 'PolicyId'
                        PolicyName                  = Get-PurviewProperty $record 'PolicyName'
                        PolicyMode                  = Get-PurviewProperty $record 'PolicyMode'
                        RuleId                      = Get-PurviewProperty $record 'RuleId'
                        RuleName                    = Get-PurviewProperty $record 'RuleName'
                        RuleActions                 = Get-PurviewProperty $record 'RuleActions'
                        EnforcementMode             = Get-PurviewProperty $record 'EnforcementMode'
                        FalsePositive               = Get-PurviewProperty $record 'FalsePositive'
                        DlpPolicyMatchId            = Get-PurviewProperty $record 'DlpPolicyMatchId'
                    })
            }

            $lastPage = [bool](Get-PurviewProperty $response 'LastPage' $true)
            $pageCookie = if ($lastPage) { $null } else { Get-PurviewProperty $response 'Watermark' }
            if (-not $lastPage -and -not $pageCookie) {
                Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
                    'LastPage is false but no Watermark came back; stopping paging.')
                $lastPage = $true
            }
        } while (-not $lastPage)
    }

    if ($unmappedActivities.Count -gt 0) {
        Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
            "Could not map $($unmappedActivities.Count) activity value(s) to a known activity, so they " +
            "are categorised as Other: $(($unmappedActivities | Sort-Object) -join ', '). " +
            'Add them to ActivityCategories in PurviewIpSchema.psd1.')
    }
}
catch {
    Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
        "Activity Explorer is unavailable to this sign-in ({0}). Writing the header only." -f $_.Exception.Message)
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}
finally {
    if ($connectedHere) {
        Disconnect-ExchangeOnline -Confirm:$false
    }
}

$result = Export-AppendCsv -Path $csvPath -Rows $rows.ToArray() -Column $columns -KeyColumn 'RecordIdentity' -PassThru
Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'activity-explorer-events.csv: {0} rows written, {1} skipped as already collected.' -f $result.Written, $result.Skipped)
