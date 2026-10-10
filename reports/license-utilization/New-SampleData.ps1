#Requires -Version 7.0

<#
    .SYNOPSIS
        Generates the fake license utilization data in ./samples.

    .DESCRIPTION
        The sample set lets the later Power BI report, and anyone reading this repository,
        work against realistic shapes without a tenant. Nothing here comes from a real
        directory: every address is under example.com, which RFC 2606 reserves for
        documentation, and every id is made up.

        The generator is deterministic. The same -Seed and -EndDate always produce the
        same files.

        Every file is written through Export-AppendCsv with the column list the collectors
        use, so a sample file cannot drift from its collector's output. ./samples/gcchigh
        holds the header-only usage files that the collectors write there, where Microsoft
        documents the usage report APIs as not available.

    .PARAMETER EndDate
        The "now" the sample set is generated around. Fixed by default to keep the
        committed files stable.

    .EXAMPLE
        ./New-SampleData.ps1
#>
[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path $PSScriptRoot 'samples'),

    [datetime]$EndDate = [datetime]::new(2026, 9, 1, 0, 0, 0, [System.DateTimeKind]::Utc),

    [int]$Seed = 20260901
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot '../../shared/M365ReportLibrary.psm1')
. (Join-Path $PSScriptRoot 'collectors/LicenseUtilizationHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'collectors/LicenseUtilizationSchema.psd1')

$EndDate = $EndDate.ToUniversalTime()
$random = [System.Random]::new($Seed)

# Three weekly snapshots, the last one on EndDate.
$snapshotDates = 2..0 | ForEach-Object { $EndDate.AddDays(-7 * $_) }

function New-DeterministicGuid {
    $bytes = [byte[]]::new(16)
    $random.NextBytes($bytes)
    return [guid]::new($bytes).ToString()
}

function Format-Stamp {
    param([AllowNull()][object]$Value)
    return ConvertTo-CsvTimestamp $Value
}

if (Test-Path -LiteralPath $OutputPath) {
    Get-ChildItem -LiteralPath $OutputPath -Recurse -File -Filter '*.csv' | Remove-Item -Force
}
New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null

# Licences: a part number, how many are bought, and the service plans each holds. Intune
# is in both E3 and EMS, so a user holding both has an overlapping service plan.
$intunePlan = @{ Id = New-DeterministicGuid; Name = 'INTUNE_A' }
$skus = @(
    @{ Id = New-DeterministicGuid; Part = 'SPE_E3'; Enabled = 12; Plans = @(@{ Id = New-DeterministicGuid; Name = 'EXCHANGE_S_ENTERPRISE' }, @{ Id = New-DeterministicGuid; Name = 'TEAMS1' }, $intunePlan) }
    @{ Id = New-DeterministicGuid; Part = 'EMS'; Enabled = 4; Plans = @($intunePlan, @{ Id = New-DeterministicGuid; Name = 'AAD_PREMIUM' }) }
    @{ Id = New-DeterministicGuid; Part = 'POWER_BI_PRO'; Enabled = 6; Plans = @(@{ Id = New-DeterministicGuid; Name = 'BI_AZURE_P2' }) }
    @{ Id = New-DeterministicGuid; Part = 'Microsoft_365_Copilot'; Enabled = 3; Plans = @(@{ Id = New-DeterministicGuid; Name = 'M365_COPILOT_APPS' }, @{ Id = New-DeterministicGuid; Name = 'M365_COPILOT_TEAMS' }) }
)
$groupId = New-DeterministicGuid

$people = @(
    ,@('Avery Abara', 'Finance', 'Analyst', 'Tampa', 'US', 'Tampa HQ')
    ,@('Blake Brandt', 'Finance', 'Controller', 'Tampa', 'US', 'Tampa HQ')
    ,@('Casey Chaudhry', 'Sales', 'Account Executive', 'Austin', 'US', 'Austin Office')
    ,@('Devon Diaz', 'Sales', 'Sales Manager', 'Austin', 'US', 'Austin Office')
    ,@('Emery Eze', 'Engineering', 'Developer', 'Denver', 'US', 'Denver Office')
    ,@('Finley Fox', 'Engineering', 'Architect', 'Denver', 'US', 'Denver Office')
    ,@('Gray Gupta', 'IT', 'Administrator', 'Tampa', 'US', 'Tampa HQ')
    ,@('Harper Hale', 'HR', 'Recruiter', 'Toronto', 'CA', 'Toronto Office')
    ,@('Indigo Ito', 'HR', 'Director', 'Toronto', 'CA', 'Toronto Office')
    ,@('Jules Jung', 'Marketing', 'Designer', 'Austin', 'US', 'Austin Office')
    ,@('Kai Kovacs', 'Marketing', 'Manager', 'Austin', 'US', 'Austin Office')
    ,@('Lane Lopez', 'Operations', 'Coordinator', 'Tampa', 'US', 'Tampa HQ')
)
$users = foreach ($i in 0..($people.Count - 1)) {
    $person = $people[$i]
    [pscustomobject]@{
        Id       = New-DeterministicGuid
        Name     = $person[0]
        Upn      = (($person[0] -replace ' ', '.').ToLowerInvariant()) + '@example.com'
        Dept     = $person[1]
        Title    = $person[2]
        City     = $person[3]
        Country  = $person[4]
        Office   = $person[5]
        Enabled  = ($i -ne 11)
        Created  = $EndDate.AddDays(-$random.Next(120, 900))
        Index    = $i
    }
}
$users = @($users)

$runDates = $snapshotDates | ForEach-Object { $_.ToString('yyyy-MM-dd') }

# users.csv is the shared Entra users collector's file; the report joins to it for
# department, job title, city and country.
$userRows = foreach ($runDate in $runDates) {
    foreach ($user in $users) {
        [pscustomobject]@{
            RunDate                  = $runDate
            Id                       = $user.Id
            DisplayName              = $user.Name
            UserPrincipalName        = $user.Upn
            Mail                     = $user.Upn
            UserType                 = 'Member'
            AccountEnabled           = $user.Enabled
            CreatedDateTime          = Format-Stamp $user.Created
            Department               = $user.Dept
            JobTitle                 = $user.Title
            City                     = $user.City
            Country                  = $user.Country
            ManagerId                = if ($user.Index -gt 0) { $users[0].Id } else { '' }
            ManagerUserPrincipalName = if ($user.Index -gt 0) { $users[0].Upn } else { '' }
        }
    }
}
Export-AppendCsv -Path (Join-Path $OutputPath 'users.csv') -Rows @($userRows) -Column (Get-EntraUserCsvColumn) -KeyColumn @('RunDate', 'Id')

# Which SKUs each user holds. Everyone has E3 (user 11 gets it from the second snapshot on);
# the first four also hold EMS, so they overlap on Intune; user 4 holds E3 through a group
# as well as directly; Power BI Pro and Copilot go to a few. The last Copilot assignment
# fails.
function Get-Assignment {
    param([int]$Snapshot)

    foreach ($user in $users) {
        if ($user.Index -lt 11 -or $Snapshot -ge 1) {
            @{ User = $user; Sku = $skus[0]; Group = ''; State = 'Active'; Error = '' }
        }
        if ($user.Index -lt 4) { @{ User = $user; Sku = $skus[1]; Group = $groupId; State = 'Active'; Error = '' } }
        if ($user.Index -eq 4) { @{ User = $user; Sku = $skus[0]; Group = $groupId; State = 'Active'; Error = '' } }
        if ($user.Index % 3 -eq 0) { @{ User = $user; Sku = $skus[2]; Group = ''; State = 'Active'; Error = '' } }
        if ($user.Index -in 0, 2) { @{ User = $user; Sku = $skus[3]; Group = ''; State = 'Active'; Error = '' } }
        if ($user.Index -eq 5 -and $Snapshot -ge 1) { @{ User = $user; Sku = $skus[3]; Group = ''; State = 'Error'; Error = 'CountViolation' } }
    }
}

$skuRows = [System.Collections.Generic.List[object]]::new()
$planRows = [System.Collections.Generic.List[object]]::new()
$licenseRows = [System.Collections.Generic.List[object]]::new()
$detailRows = [System.Collections.Generic.List[object]]::new()
$signInRows = [System.Collections.Generic.List[object]]::new()
$settingRows = [System.Collections.Generic.List[object]]::new()

foreach ($snapshot in 0..($snapshotDates.Count - 1)) {
    $runDate = $runDates[$snapshot]
    $assignments = @(Get-Assignment -Snapshot $snapshot)

    foreach ($sku in $skus) {
        $consumed = @($assignments | Where-Object { $_.Sku.Id -eq $sku.Id }).Count
        $skuRows.Add([pscustomobject]@{
                RunDate          = $runDate
                SkuId            = $sku.Id
                SkuPartNumber    = $sku.Part
                AppliesTo        = 'User'
                CapabilityStatus = 'Enabled'
                ConsumedUnits    = $consumed
                PrepaidEnabled   = $sku.Enabled
                PrepaidSuspended = 0
                PrepaidWarning   = 0
                PrepaidLockedOut = 0
                FreeUnits        = $sku.Enabled - $consumed
            })
        foreach ($plan in $sku.Plans) {
            $planRows.Add([pscustomobject]@{
                    RunDate            = $runDate
                    SkuId              = $sku.Id
                    ServicePlanId      = $plan.Id
                    ServicePlanName    = $plan.Name
                    ProvisioningStatus = 'Success'
                    AppliesTo          = 'User'
                })
        }
    }

    foreach ($assignment in $assignments) {
        $user = $assignment.User
        $licenseRows.Add([pscustomobject]@{
                RunDate           = $runDate
                UserId            = $user.Id
                UserPrincipalName = $user.Upn
                AccountEnabled    = $user.Enabled
                UsageLocation     = $user.Country
                OfficeLocation    = $user.Office
                SkuId             = $assignment.Sku.Id
                AssignedByGroup   = $assignment.Group
                AssignmentType    = if ($assignment.Group) { 'Group' } else { 'Direct' }
                State             = $assignment.State
                Error             = $assignment.Error
                DisabledPlans     = ''
            })
    }

    $signInDate = $snapshotDates[$snapshot]
    foreach ($user in $users) {
        # Users 9 to 11 have never signed in; the rest signed in within the last weeks.
        $never = $user.Index -ge 9
        $last = $signInDate.AddDays(-$random.Next(0, 40))
        $signInRows.Add([pscustomobject]@{
                RunDate                          = $runDate
                UserId                           = $user.Id
                UserPrincipalName                = $user.Upn
                LastSignInDateTime               = if ($never) { '' } else { Format-Stamp $last }
                LastNonInteractiveSignInDateTime = if ($never) { '' } else { Format-Stamp $last.AddHours(3) }
                LastSuccessfulSignInDateTime     = if ($never) { '' } else { Format-Stamp $last }
            })
    }

    $settingRows.Add([pscustomobject]@{ RunDate = $runDate; DisplayConcealedNames = 'False' })
}

# Per-user licence details for the latest snapshot only (source 3 is optional).
$latest = $runDates[-1]
foreach ($assignment in @(Get-Assignment -Snapshot ($snapshotDates.Count - 1) | Where-Object { $_.State -eq 'Active' })) {
    foreach ($plan in $assignment.Sku.Plans) {
        $detailRows.Add([pscustomobject]@{
                RunDate            = $latest
                UserId             = $assignment.User.Id
                SkuId              = $assignment.Sku.Id
                SkuPartNumber      = $assignment.Sku.Part
                ServicePlanId      = $plan.Id
                ServicePlanName    = $plan.Name
                ProvisioningStatus = 'Success'
            })
    }
}

Export-AppendCsv -Path (Join-Path $OutputPath 'subscribed-skus.csv') -Rows $skuRows.ToArray() `
    -Column $schema.SubscribedSkus -KeyColumn @('RunDate', 'SkuId')
Export-AppendCsv -Path (Join-Path $OutputPath 'sku-service-plans.csv') -Rows $planRows.ToArray() `
    -Column $schema.SkuServicePlans -KeyColumn @('RunDate', 'SkuId', 'ServicePlanId')
Export-AppendCsv -Path (Join-Path $OutputPath 'user-licenses.csv') -Rows $licenseRows.ToArray() `
    -Column $schema.UserLicenses -KeyColumn @('RunDate', 'UserId', 'SkuId', 'AssignedByGroup')
Export-AppendCsv -Path (Join-Path $OutputPath 'license-details.csv') -Rows $detailRows.ToArray() `
    -Column $schema.LicenseDetails -KeyColumn @('RunDate', 'UserId', 'SkuId', 'ServicePlanId')
Export-AppendCsv -Path (Join-Path $OutputPath 'user-signin-activity.csv') -Rows $signInRows.ToArray() `
    -Column $schema.UserSignInActivity -KeyColumn @('RunDate', 'UserId')
Export-AppendCsv -Path (Join-Path $OutputPath 'report-settings.csv') -Rows $settingRows.ToArray() `
    -Column $schema.ReportSettings -KeyColumn @('RunDate')

# Usage reports: one row per licensed user per snapshot, shaped by header name. Values
# are plausible, not meaningful. A user who never signed in has no last activity.
function Get-SampleUsageValue {
    param([string]$Header, [object]$User, [datetime]$At, [string]$Period)

    $never = $User.Index -ge 9
    $day = { param($offset) if ($never) { '' } else { $At.AddDays(-$offset).ToString('yyyy-MM-dd') } }

    switch -Regex ($Header) {
        '^Report Refresh Date$' { return $At.ToString('yyyy-MM-dd') }
        '^User Principal Name$' { return $User.Upn }
        '^User Id$' { return $User.Id }
        '^Display Name$' { return $User.Name }
        '^Tenant Display Name$' { return 'Contoso Sample' }
        '^Shared Channel Tenant Display Names$' { return '' }
        '^Is Deleted$' { return 'False' }
        '^Deleted Date$' { return '' }
        '^Report Period$' { return $Period }
        '^Assigned Products$' { return 'MICROSOFT 365 E3+ENTERPRISE MOBILITY + SECURITY E3' }
        '^Has .* License$' { return 'True' }
        '^Is Licensed$' { return 'Yes' }
        '^Has Other Action$' { return 'No' }
        'License Assign Date$' { return $At.AddDays(-200).ToString('yyyy-MM-dd') }
        'Last Activation Date$' { return $At.AddDays(-150).ToString('yyyy-MM-dd') }
        'Last Activity Date$' { return (& $day ($User.Index * 3 + 1)) }
        'Duration In Seconds$' { return [string]($User.Index * 600) }
        'Duration$' { return ('{0:00}:{1:00}:00' -f $User.Index, ($User.Index * 5 % 60)) }
        'Count$|^Post Messages$|^Reply Messages$|^Urgent Messages$|^Prompts submitted' { return [string](($User.Index * 7 + 3) % 40) }
        '^Active Usage Days' { return [string]($User.Index % 20) }
        default { return [string]($User.Index % 2 -eq 0) }
    }
}

$usageFiles = [ordered]@{
    ActiveUserUsage         = @{ Csv = 'usage-active-users.csv'; Period = '30' }
    EmailActivityUsage      = @{ Csv = 'usage-email-activity.csv'; Period = '30' }
    TeamsActivityUsage      = @{ Csv = 'usage-teams-activity.csv'; Period = '30' }
    SharePointActivityUsage = @{ Csv = 'usage-sharepoint-activity.csv'; Period = '30' }
    OneDriveActivityUsage   = @{ Csv = 'usage-onedrive-activity.csv'; Period = '30' }
    M365AppUsage            = @{ Csv = 'usage-m365-apps.csv'; Period = '30' }
    CopilotUsage            = @{ Csv = 'usage-copilot.csv'; Period = '28' }
}
foreach ($key in $usageFiles.Keys) {
    $headers = $schema.UsageReports[$key]
    $columns = Get-UsageColumn -Header $headers
    $rows = foreach ($snapshot in 0..($snapshotDates.Count - 1)) {
        $at = $snapshotDates[$snapshot]
        # The Copilot report returns only users holding a Copilot licence.
        $reported = if ($key -eq 'CopilotUsage') { @($users | Where-Object { $_.Index -in 0, 2, 5 }) } else { $users }
        foreach ($user in $reported) {
            $row = [ordered]@{ RunDate = $runDates[$snapshot] }
            foreach ($header in $headers) {
                $row[(ConvertTo-UsageColumnName $header)] = Get-SampleUsageValue -Header $header -User $user -At $at -Period $usageFiles[$key].Period
            }
            [pscustomobject]$row
        }
    }
    Export-AppendCsv -Path (Join-Path $OutputPath $usageFiles[$key].Csv) -Rows @($rows) -Column $columns `
        -KeyColumn @('RunDate', 'UserPrincipalName', 'ReportPeriod')

    # GCC High: Microsoft documents the usage report APIs as not available, so the
    # collector writes the header only.
    Export-AppendCsv -Path (Join-Path $OutputPath "gcchigh/$($usageFiles[$key].Csv)") -Column $columns
}
