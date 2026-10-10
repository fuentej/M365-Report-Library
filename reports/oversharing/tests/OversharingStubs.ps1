#Requires -Version 7.0

<#
    .SYNOPSIS
        Declares the SharePoint Online, Exchange Online and Graph commands the oversharing
        collectors call that shared/tests/TenantCmdletStubs.ps1 does not, so the tests can
        mock them without those modules installed.

    .DESCRIPTION
        Every stub throws. A test that reaches one has failed to mock something. Defined in
        the global scope so a mock resolves from a collector script and from the
        dot-sourced helpers. Parameters are only the ones this report passes. Parameter
        binding happens before a Pester mock runs, so a parameter a collector passes must
        be declared here.
#>

function global:Connect-SPOService {
    [CmdletBinding()]
    param([string]$Url, [string]$ClientId, [string]$CertificateThumbprint, [string]$TenantId, [string]$Region)
    throw 'Connect-SPOService was called for real. Mock it in the test.'
}

function global:Disconnect-SPOService {
    [CmdletBinding()]
    param()
    throw 'Disconnect-SPOService was called for real. Mock it in the test.'
}

function global:Get-SPOSite {
    [CmdletBinding()]
    param([string]$Identity, [object]$Limit, [object]$IncludePersonalSite)
    throw 'Get-SPOSite was called for real. Mock it in the test.'
}

function global:Start-SPODataAccessGovernanceInsight {
    [CmdletBinding()]
    param([string]$ReportEntity, [string]$ReportType, [string]$Workload, [string]$Name, [int]$CountOfUsersMoreThan,
        [guid]$FileSensitivityLabelGUID, [string]$FileSensitivityLabelName)
    throw 'Start-SPODataAccessGovernanceInsight was called for real. Mock it in the test.'
}

function global:Get-SPODataAccessGovernanceInsight {
    [CmdletBinding()]
    param([string]$ReportEntity, [string]$ReportType, [string]$Workload, [string]$ReportID)
    throw 'Get-SPODataAccessGovernanceInsight was called for real. Mock it in the test.'
}

function global:Export-SPODataAccessGovernanceInsight {
    [CmdletBinding()]
    param([string]$ReportID, [string]$DownloadPath)
    throw 'Export-SPODataAccessGovernanceInsight was called for real. Mock it in the test.'
}

function global:Get-SPOAuditDataCollectionStatusForActivityInsights {
    [CmdletBinding()]
    param([string]$ReportEntity)
    throw 'Get-SPOAuditDataCollectionStatusForActivityInsights was called for real. Mock it in the test.'
}

function global:Start-SPOAuditDataCollectionForActivityInsights {
    [CmdletBinding()]
    param([string]$ReportEntity)
    throw 'Start-SPOAuditDataCollectionForActivityInsights was called for real. Tests must never reach it: it changes the tenant.'
}

function global:Get-AdminAuditLogConfig {
    [CmdletBinding()]
    param()
    throw 'Get-AdminAuditLogConfig was called for real. Mock it in the test.'
}

# Replaces the shared stub. Search-UnifiedAuditLog -Formatted is what turns RecordType
# into a display name.
# https://learn.microsoft.com/powershell/module/exchangepowershell/search-unifiedauditlog
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
        [object]$ResultSize,
        [switch]$Formatted
    )
    throw 'Search-UnifiedAuditLog was called for real. Mock it in the test.'
}

function global:Invoke-MgGraphRequest {
    [CmdletBinding()]
    param([string]$Method, [string]$Uri)
    throw 'Invoke-MgGraphRequest was called for real. Mock it in the test.'
}

function global:Get-SPOTenant {
    [CmdletBinding()]
    param()
    throw 'Get-SPOTenant was called for real. Mock it in the test.'
}
