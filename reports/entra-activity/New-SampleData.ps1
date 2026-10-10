#Requires -Version 7.0

<#
    .SYNOPSIS
        Regenerates samples/: fake CSVs with the exact columns each collector writes.

    .DESCRIPTION
        Every name, address and Id is invented; addresses use example.com and IP addresses
        are from the documentation range 203.0.113.0/24 (RFC 5737). The column order comes
        from collectors/EntraActivitySchema.psd1, so a sample cannot drift from its
        collector. No sign-in happens and no tenant is called.

        None of the five sources is marked NotAvailable in any cloud in the contract
        (docs/candidates/entra-activity.md), so there is no header-only sample.

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

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'collectors/EntraActivitySchema.psd1')

if (Test-Path -LiteralPath $OutputPath) {
    Get-ChildItem -LiteralPath $OutputPath -Filter '*.csv' | Remove-Item -Force
}

function New-SignInSample {
    param([string]$Id, [string]$At, [string]$UserId, [string]$Upn, [string]$App, [string]$Resource,
        [string]$Ip, [string]$City, [string]$Country, [string]$Client, [string]$Os, [string]$Browser,
        [bool]$Interactive, [string]$EventTypes, [string]$ErrorCode, [string]$Reason, [string]$Details,
        [string]$CaStatus)

    [pscustomobject]@{
        CreatedDateTime = $At; Id = $Id; UserId = $UserId; UserPrincipalName = $Upn
        AppId = '00000002-0000-0ff1-ce00-000000000000'; AppDisplayName = $App; ResourceDisplayName = $Resource
        IpAddress = $Ip; City = $City; State = ''; CountryOrRegion = $Country; ClientAppUsed = $Client
        DeviceOperatingSystem = $Os; DeviceBrowser = $Browser; DeviceIsCompliant = 'True'; DeviceIsManaged = 'True'
        IsInteractive = $Interactive; SignInEventTypes = $EventTypes; ErrorCode = $ErrorCode
        FailureReason = $Reason; AdditionalDetails = $Details; ConditionalAccessStatus = $CaStatus
        RiskDetail = 'none'; RiskLevelAggregated = 'none'; RiskLevelDuringSignIn = 'none'; RiskState = 'none'
    }
}

$interactive = @(
    New-SignInSample 'a1000000-0000-4000-8000-000000000001' '2026-09-10T06:05:00Z' 'aaaaaaaa-0000-4000-8000-000000000001' 'avery.abara@example.com' 'Office 365 Exchange Online' 'Office 365 Exchange Online' '203.0.113.10' 'Tampa' 'US' 'Browser' 'Windows 11' 'Edge 128.0' $true 'interactiveUser' '0' '' '' 'success'
    New-SignInSample 'a1000000-0000-4000-8000-000000000002' '2026-09-10T07:30:00Z' 'aaaaaaaa-0000-4000-8000-000000000002' 'casey.chaudhry@example.com' 'Microsoft Teams' 'Microsoft Teams' '203.0.113.11' 'Orlando' 'US' 'Mobile Apps and Desktop clients' 'iOS 17' 'Teams' $true 'interactiveUser' '50126' 'Invalid username or password or Invalid on-premise username or password.' 'Error validating credentials due to invalid username or password.' 'notApplied'
    New-SignInSample 'a1000000-0000-4000-8000-000000000003' '2026-09-11T08:15:00Z' '' '' 'Azure Portal' 'Windows Azure Service Management API' '203.0.113.12' '' 'US' 'Browser' 'macOS 14' 'Safari 17.6' $true 'interactiveUser' '50058' 'Session information is not sufficient for single-sign-on.' 'The user is not signed in; no user name was returned.' 'notApplied'
    New-SignInSample 'a1000000-0000-4000-8000-000000000004' '2026-09-11T09:40:00Z' 'aaaaaaaa-0000-4000-8000-000000000003' 'devon.dube@example.com' 'SharePoint Online Web Client Extensibility' 'Office 365 SharePoint Online' '203.0.113.13' 'Miami' 'US' 'Browser' 'Windows 11' 'Chrome 128.0' $true 'interactiveUser' '53003' 'Access has been blocked by Conditional Access policies.' 'Access has been blocked by Conditional Access policies. The access policy does not allow token issuance.' 'failure'
    New-SignInSample 'a1000000-0000-4000-8000-000000000005' '2026-09-12T10:00:00Z' 'aaaaaaaa-0000-4000-8000-000000000001' 'avery.abara@example.com' 'Office 365 Exchange Online' 'Office 365 Exchange Online' '203.0.113.10' 'Tampa' 'US' 'Browser' 'Windows 11' 'Edge 128.0' $true 'interactiveUser' '1024' 'Sample code that is not 5 or 6 digits.' '' 'success'
)
Export-AppendCsv -Path (Join-Path $OutputPath 'signins-interactive.csv') -Rows $interactive -Column $schema.SignIns -KeyColumn 'Id'

$nonInteractive = @(
    New-SignInSample 'b2000000-0000-4000-8000-000000000001' '2026-09-10T06:20:00Z' 'aaaaaaaa-0000-4000-8000-000000000001' 'avery.abara@example.com' 'Microsoft Office' 'Microsoft Graph' '203.0.113.10' 'Tampa' 'US' 'Mobile Apps and Desktop clients' 'Windows 11' '' $false 'nonInteractiveUser' '0' '' '' 'notApplied'
    New-SignInSample 'b2000000-0000-4000-8000-000000000002' '2026-09-11T11:00:00Z' 'aaaaaaaa-0000-4000-8000-000000000002' 'casey.chaudhry@example.com' 'Outlook Mobile' 'Office 365 Exchange Online' '203.0.113.11' 'Orlando' 'US' 'Mobile Apps and Desktop clients' 'iOS 17' '' $false 'nonInteractiveUser' '50173' 'The provided grant has expired due to it being revoked.' '' 'notApplied'
)
Export-AppendCsv -Path (Join-Path $OutputPath 'signins-noninteractive.csv') -Rows $nonInteractive -Column $schema.SignIns -KeyColumn 'Id'

function New-CaSample {
    param([string]$SignInId, [string]$At, [string]$UserId, [bool]$Interactive, [string]$Status,
        [string]$PolicyId, [string]$Name, [string]$Result, [string]$Grant, [string]$Session, [bool]$Readable)
    [pscustomobject]@{
        CreatedDateTime = $At; SignInId = $SignInId; UserId = $UserId; IsInteractive = $Interactive
        ConditionalAccessStatus = $Status; PolicyId = $PolicyId; PolicyDisplayName = $Name; PolicyResult = $Result
        EnforcedGrantControls = $Grant; EnforcedSessionControls = $Session; PolicyDetailReadable = $Readable
    }
}
$ca = @(
    New-CaSample 'a1000000-0000-4000-8000-000000000001' '2026-09-10T06:05:00Z' 'aaaaaaaa-0000-4000-8000-000000000001' $true 'success' 'cccccccc-0000-4000-8000-000000000001' 'Require MFA for all users' 'success' 'Mfa' '' $true
    New-CaSample 'a1000000-0000-4000-8000-000000000004' '2026-09-11T09:40:00Z' 'aaaaaaaa-0000-4000-8000-000000000003' $true 'failure' 'cccccccc-0000-4000-8000-000000000002' 'Block unmanaged devices' 'failure' 'Block' '' $true
    New-CaSample 'a1000000-0000-4000-8000-000000000004' '2026-09-11T09:40:00Z' 'aaaaaaaa-0000-4000-8000-000000000003' $true 'failure' 'cccccccc-0000-4000-8000-000000000003' 'Report-only: require compliant device' 'unknownFutureValue' '' '' $true
    New-CaSample 'a1000000-0000-4000-8000-000000000002' '2026-09-10T07:30:00Z' 'aaaaaaaa-0000-4000-8000-000000000002' $true 'notApplied' '' '' '' '' '' $false
)
Export-AppendCsv -Path (Join-Path $OutputPath 'signin-conditional-access.csv') -Rows $ca -Column $schema.SignInConditionalAccess -KeyColumn 'SignInId', 'PolicyId', 'PolicyDisplayName'

function New-AuditSample {
    param([string]$Id, [string]$At, [string]$Activity, [string]$Category, [string]$Operation, [string]$Result,
        [string]$Reason, [string]$Service, [string]$ByUserId, [string]$ByUpn, [string]$ByAppId, [string]$ByApp,
        [string]$Types, [string]$Ids, [string]$Names, [string]$Upns, [string]$Modified)
    [pscustomobject]@{
        ActivityDateTime = $At; Id = $Id; ActivityDisplayName = $Activity; Category = $Category
        OperationType = $Operation; Result = $Result; ResultReason = $Reason; LoggedByService = $Service
        CorrelationId = 'dddddddd-0000-4000-8000-0000000000' + $Id.Substring($Id.Length - 2)
        InitiatedByUserId = $ByUserId; InitiatedByUserPrincipalName = $ByUpn; InitiatedByAppId = $ByAppId
        InitiatedByAppDisplayName = $ByApp; TargetResourceTypes = $Types; TargetResourceIds = $Ids
        TargetResourceDisplayNames = $Names; TargetResourceUserPrincipalNames = $Upns; ModifiedProperties = $Modified
    }
}
$audits = @(
    New-AuditSample 'Directory_eeeeeeee-0000-4000-8000-000000000001_01' '2026-09-10T09:00:00Z' 'Add user' 'UserManagement' 'Add' 'success' '' 'Core Directory' 'aaaaaaaa-0000-4000-8000-000000000009' 'admin.alvarez@example.com' '' '' 'User' 'aaaaaaaa-0000-4000-8000-000000000010' 'Riley Rosen' 'riley.rosen@example.com' '[{"displayName":"AccountEnabled","oldValue":"","newValue":"[true]"}]'
    New-AuditSample 'Directory_eeeeeeee-0000-4000-8000-000000000002_02' '2026-09-10T09:30:00Z' 'Update Conditional Access policy' 'Policy' 'Update' 'success' '' 'Conditional Access' 'aaaaaaaa-0000-4000-8000-000000000009' 'admin.alvarez@example.com' '' '' 'Policy' 'cccccccc-0000-4000-8000-000000000001' 'Require MFA for all users' '' '[{"displayName":"State","oldValue":"[\"enabledForReportingButNotEnforced\"]","newValue":"[\"enabled\"]"}]'
    New-AuditSample 'SSGM_b662f17a-4e4d-4e1c-9248-cdec180024b2_MCDC4_88453290' '2026-09-11T12:00:00Z' 'Add member to group' 'GroupManagement' 'Assign' 'failure' 'Insufficient privileges to complete the operation.' 'Core Directory' '' '' 'ffffffff-0000-4000-8000-000000000001' 'Provisioning Sample App' 'Group;User' 'aaaaaaaa-0000-4000-8000-000000000020;aaaaaaaa-0000-4000-8000-000000000010' 'Sales team;Riley Rosen' ';riley.rosen@example.com' ''
    New-AuditSample 'Directory_eeeeeeee-0000-4000-8000-000000000004_04' '2026-09-12T08:45:00Z' 'Add member to role' 'RoleManagement' 'Assign' 'success' '' 'Core Directory' 'aaaaaaaa-0000-4000-8000-000000000009' 'admin.alvarez@example.com' '' '' 'Role;User' 'bbbbbbbb-0000-4000-8000-000000000001;aaaaaaaa-0000-4000-8000-000000000010' 'Helpdesk Administrator;Riley Rosen' ';riley.rosen@example.com' ''
)
Export-AppendCsv -Path (Join-Path $OutputPath 'directory-audits.csv') -Rows $audits -Column $schema.DirectoryAudits -KeyColumn 'Id'

$retention = foreach ($environment in 'Commercial', 'GCC', 'GCCHigh') {
    foreach ($level in $schema.RetentionLevels) {
        [pscustomobject]@{
            RunDate = '2026-09-30'; Environment = $environment; LicenseLevel = $level.LicenseLevel
            SignInRetentionDays = $level.SignInRetentionDays; AuditRetentionDays = $level.AuditRetentionDays
            RiskySignInRetentionDays = $level.RiskySignInRetentionDays; RetentionStatus = 'UNVERIFIED'
            Reference = $schema.SourceAvailability.RetentionReference[$environment].Reference
        }
    }
}
Export-AppendCsv -Path (Join-Path $OutputPath 'retention-reference.csv') -Rows $retention -Column $schema.RetentionReference -KeyColumn 'RunDate', 'Environment', 'LicenseLevel'
