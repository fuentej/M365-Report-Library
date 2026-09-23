#Requires -Version 7.0

<#
    .SYNOPSIS
        Declares the Microsoft Graph, Exchange Online and Security & Compliance cmdlets
        the collectors call, so the tests can mock them without those modules installed.

    .DESCRIPTION
        Every stub throws. A test that reaches a stub has failed to mock something, and
        the message says which cmdlet leaked.

        The stubs are defined in the global scope on purpose: a mock has to be resolvable
        both from a collector script and from inside M365ReportLibrary.psm1, which has its
        own session state.

        The parameters are only the ones this repository passes, plus the common ones
        Pester needs to match a -ParameterFilter on. They are not a complete copy of the
        real cmdlets' signatures.
#>

function global:Connect-MgGraph {
    [CmdletBinding()]
    param(
        [string]$Environment,
        [string[]]$Scopes,
        [string]$ClientId,
        [string]$CertificateThumbprint,
        [string]$TenantId,
        [switch]$NoWelcome
    )
    throw 'Connect-MgGraph was called for real. Mock it in the test.'
}

function global:Connect-ExchangeOnline {
    [CmdletBinding()]
    param(
        [string]$ExchangeEnvironmentName,
        [string]$AppId,
        [string]$CertificateThumbprint,
        [string]$Organization,
        [string]$UserPrincipalName,
        [switch]$ShowBanner
    )
    throw 'Connect-ExchangeOnline was called for real. Mock it in the test.'
}

function global:Connect-IPPSSession {
    [CmdletBinding()]
    param(
        [string]$ConnectionUri,
        [string]$AzureADAuthorizationEndpointUri,
        [string]$AppId,
        [string]$CertificateThumbprint,
        [string]$Organization,
        [string]$UserPrincipalName,
        [switch]$ShowBanner
    )
    throw 'Connect-IPPSSession was called for real. Mock it in the test.'
}

function global:Get-MgUser {
    [CmdletBinding()]
    param(
        [string]$UserId,
        [switch]$All,
        [string]$Filter,
        [string[]]$Property,
        [string[]]$ExpandProperty,
        [string]$ConsistencyLevel,
        [int]$Top
    )
    throw 'Get-MgUser was called for real. Mock it in the test.'
}

function global:Get-MgUserManager {
    [CmdletBinding()]
    param(
        [string]$UserId,
        [string[]]$Property
    )
    throw 'Get-MgUserManager was called for real. Mock it in the test.'
}

function global:Get-MgUserMemberOf {
    [CmdletBinding()]
    param(
        [string]$UserId,
        [switch]$All,
        [string]$Filter,
        [string[]]$Property
    )
    throw 'Get-MgUserMemberOf was called for real. Mock it in the test.'
}

function global:Get-MgUserMemberOfAsGroup {
    [CmdletBinding()]
    param(
        [string]$UserId,
        [switch]$All,
        [string]$Filter,
        [string[]]$Property
    )
    throw 'Get-MgUserMemberOfAsGroup was called for real. Mock it in the test.'
}

function global:Get-MgAuditLogDirectoryAudit {
    [CmdletBinding()]
    param(
        [switch]$All,
        [string]$Filter,
        [int]$Top
    )
    throw 'Get-MgAuditLogDirectoryAudit was called for real. Mock it in the test.'
}

function global:Get-MgAuditLogSignIn {
    [CmdletBinding()]
    param(
        [switch]$All,
        [string]$Filter,
        [int]$Top
    )
    throw 'Get-MgAuditLogSignIn was called for real. Mock it in the test.'
}

function global:Search-UnifiedAuditLog {
    [CmdletBinding()]
    param(
        [datetime]$StartDate,
        [datetime]$EndDate,
        [string[]]$Operations,
        [string[]]$RecordType,
        [string[]]$UserIds,
        [string]$SessionId,
        [string]$SessionCommand,
        [object]$ResultSize
    )
    throw 'Search-UnifiedAuditLog was called for real. Mock it in the test.'
}

function global:Disconnect-ExchangeOnline {
    [CmdletBinding()]
    param(
        [switch]$Confirm
    )
    # No-op: tests that care assert via a Mock. An unmocked call must not fail a
    # collector that opened a session, and must not reach the real cmdlet.
}

function global:Get-Label {
    [CmdletBinding()]
    param()
    throw 'Get-Label was called for real. Mock it in the test.'
}

function global:Get-LabelPolicy {
    [CmdletBinding()]
    param()
    throw 'Get-LabelPolicy was called for real. Mock it in the test.'
}

function global:Get-AutoSensitivityLabelPolicy {
    [CmdletBinding()]
    param()
    throw 'Get-AutoSensitivityLabelPolicy was called for real. Mock it in the test.'
}

function global:Get-DlpCompliancePolicy {
    [CmdletBinding()]
    param()
    throw 'Get-DlpCompliancePolicy was called for real. Mock it in the test.'
}

function global:Get-DlpComplianceRule {
    [CmdletBinding()]
    param()
    throw 'Get-DlpComplianceRule was called for real. Mock it in the test.'
}

function global:Get-RetentionCompliancePolicy {
    [CmdletBinding()]
    param()
    throw 'Get-RetentionCompliancePolicy was called for real. Mock it in the test.'
}

function global:Get-ComplianceTag {
    [CmdletBinding()]
    param()
    throw 'Get-ComplianceTag was called for real. Mock it in the test.'
}

function global:Get-DlpSensitiveInformationType {
    [CmdletBinding()]
    param()
    throw 'Get-DlpSensitiveInformationType was called for real. Mock it in the test.'
}

function global:Export-ActivityExplorerData {
    [CmdletBinding()]
    param(
        [datetime]$StartTime,
        [datetime]$EndTime,
        [string]$OutputFormat,
        [int]$PageSize,
        [string]$PageCookie
    )
    throw 'Export-ActivityExplorerData was called for real. Mock it in the test.'
}

function global:Export-ContentExplorerData {
    [CmdletBinding()]
    param(
        [string]$TagType,
        [string]$TagName,
        [string]$Workload,
        [int]$PageSize
    )
    throw 'Export-ContentExplorerData was called for real. Mock it in the test.'
}
