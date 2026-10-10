#Requires -Version 7.0

<#
    .SYNOPSIS
        Regenerates samples/: fake CSVs with exactly the columns the collectors write.

    .DESCRIPTION
        Every name is invented and every address is on example.com. Columns come from
        collectors/UnifiedAuditLogSchema.psd1, so a sample cannot drift from its collector.
        samples/gcchigh/ holds the header-only file a GCC High run leaves for the Graph Audit Search
        API, which is not available in that cloud.

    .EXAMPLE
        ./New-SampleData.ps1
#>
[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path $PSScriptRoot 'samples')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot '../../shared/M365ReportLibrary.psm1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'collectors/UnifiedAuditLogSchema.psd1')

function Write-Sample {
    param([string]$Folder, [string]$Name, [string[]]$Column, [object[]]$Row)
    $path = Join-Path $Folder $Name
    if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
    Export-AppendCsv -Path $path -Rows $Row -Column $Column
}

if (-not (Test-Path -LiteralPath $OutputPath)) { New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null }
$gccHigh = Join-Path $OutputPath 'gcchigh'
if (-not (Test-Path -LiteralPath $gccHigh)) { New-Item -Path $gccHigh -ItemType Directory -Force | Out-Null }

$org = '11111111-1111-1111-1111-111111111111'

# One invented event per row: Id, time, operation, workload, user, object, IP, result, record type.
$events = @(
    @{ Id = '00000000-0000-0000-0000-000000000001'; At = '2026-10-08T09:15:00Z'; Op = 'Add member to role.'; Work = 'AzureActiveDirectory'; User = 'avery.abara@example.com'; Obj = 'blake.bishop@example.com'; Ip = '203.0.113.10'; Result = 'success'; Name = 'azureActiveDirectory'; Num = '8'; Cmdlet = 'AzureActiveDirectory' }
    @{ Id = '00000000-0000-0000-0000-000000000002'; At = '2026-10-08T10:02:00Z'; Op = 'FileAccessed'; Work = 'SharePoint'; User = 'casey.cho@example.com'; Obj = 'https://example.com/sites/finance/Shared Documents/budget.xlsx'; Ip = '203.0.113.11'; Result = ''; Name = 'sharePointFileOperation'; Num = '6'; Cmdlet = 'SharePointFileOperation' }
    @{ Id = '00000000-0000-0000-0000-000000000003'; At = '2026-10-08T11:30:00Z'; Op = 'AnonymousLinkCreated'; Work = 'SharePoint'; User = 'devon.diaz@example.com'; Obj = 'https://example.com/sites/hr/Shared Documents/plan.docx'; Ip = '203.0.113.12'; Result = ''; Name = 'sharePointSharingOperation'; Num = '14'; Cmdlet = 'SharePointSharingOperation' }
    @{ Id = '00000000-0000-0000-0000-000000000004'; At = '2026-10-08T13:45:00Z'; Op = 'New-TransportRule'; Work = 'Exchange'; User = 'avery.abara@example.com'; Obj = 'Block external forwarding'; Ip = '203.0.113.10'; Result = 'True'; Name = 'exchangeAdmin'; Num = '1'; Cmdlet = 'ExchangeAdmin' }
    @{ Id = '00000000-0000-0000-0000-000000000005'; At = '2026-10-09T08:20:00Z'; Op = 'MailItemsAccessed'; Work = 'Exchange'; User = 'emery.ellis@example.com'; Obj = ''; Ip = '203.0.113.13'; Result = ''; Name = 'exchangeItemAggregated'; Num = '50'; Cmdlet = 'ExchangeItemAggregated' }
)

function Get-EventJson {
    param($E, [string]$Num)
    [ordered]@{
        Id = $E.Id; CreationTime = $E.At.TrimEnd('Z'); Operation = $E.Op; OrganizationId = $org
        RecordType = [int]$Num; ResultStatus = $E.Result; UserType = 0; Workload = $E.Work
        UserId = $E.User; ClientIP = $E.Ip; ObjectId = $E.Obj
    } | ConvertTo-Json -Compress
}

# Source 1
$rows = foreach ($e in $events) {
    [pscustomobject]@{
        CreationTime = $e.At; RecordId = $e.Id; RecordType = $e.Cmdlet; Operation = $e.Op; Workload = $e.Work
        UserId = $e.User; ObjectId = $e.Obj; ClientIP = $e.Ip; ResultStatus = $e.Result; OrganizationId = $org
        AuditData = Get-EventJson $e $e.Num
    }
}
Write-Sample $OutputPath 'audit-search-cmdlet.csv' $schema.AuditSearchCmdlet $rows

# Source 2
$queryId = '22222222-2222-2222-2222-222222222222'
$rows = foreach ($e in $events) {
    [pscustomobject]@{
        CreationTime = $e.At; RecordId = $e.Id; RecordType = $e.Name; Operation = $e.Op; Workload = $e.Work
        UserId = $e.User; ObjectId = $e.Obj; ClientIP = $e.Ip; ResultStatus = $e.Result; OrganizationId = $org
        QueryId = $queryId; AuditData = Get-EventJson $e $e.Num
    }
}
Write-Sample $OutputPath 'audit-graph-records.csv' $schema.AuditGraphRecords $rows
Write-Sample $gccHigh 'audit-graph-records.csv' $schema.AuditGraphRecords @()

# Source 3: two blobs, the second holds two events.
$blobs = @(
    @{ Id = 'blob-aad-1'; Type = 'Audit.AzureActiveDirectory'; At = '2026-10-08T09:40:00Z'; Events = @(0) }
    @{ Id = 'blob-sp-1'; Type = 'Audit.SharePoint'; At = '2026-10-08T12:00:00Z'; Events = @(1, 2) }
    @{ Id = 'blob-ex-1'; Type = 'Audit.Exchange'; At = '2026-10-09T09:00:00Z'; Events = @(3, 4) }
)
$rows = foreach ($b in $blobs) {
    foreach ($i in $b.Events) {
        $e = $events[$i]
        [pscustomobject]@{
            ContentCreated = $b.At; ContentType = $b.Type; ContentId = $b.Id; CreationTime = $e.At; RecordId = $e.Id
            RecordType = $e.Num; Operation = $e.Op; Workload = $e.Work; UserId = $e.User; ObjectId = $e.Obj
            ClientIP = $e.Ip; ResultStatus = $e.Result; OrganizationId = $org; AuditData = Get-EventJson $e $e.Num
        }
    }
}
Write-Sample $OutputPath 'audit-activity-feed.csv' $schema.AuditActivityFeed $rows

# Source 4
$rows = foreach ($run in '2026-10-01', '2026-10-08') {
    [pscustomobject]@{ RunDate = $run; UnifiedAuditLogIngestionEnabled = 'True' }
}
Write-Sample $OutputPath 'audit-ingestion.csv' $schema.AuditIngestion $rows

# Source 5
$rows = foreach ($run in '2026-10-01', '2026-10-08') {
    [pscustomobject]@{ RunDate = $run; Priority = 100; Name = 'Admin role changes, ten years'; RecordTypes = 'AzureActiveDirectory'; Operations = 'Add member to role.;Remove member from role.'; UserIds = ''; RetentionDuration = 'TenYears' }
    [pscustomobject]@{ RunDate = $run; Priority = 200; Name = 'Sharing, three months'; RecordTypes = 'SharePointSharingOperation'; Operations = ''; UserIds = 'devon.diaz@example.com'; RetentionDuration = 'ThreeMonths' }
}
Write-Sample $OutputPath 'audit-retention-policies.csv' $schema.AuditRetentionPolicies $rows
