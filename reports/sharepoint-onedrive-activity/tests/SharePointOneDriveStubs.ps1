#Requires -Version 7.0

<#
    .SYNOPSIS
        Declares the SharePoint Online, Exchange Online and Graph commands this report calls that
        shared/tests/TenantCmdletStubs.ps1 does not, so the tests can mock them without those
        modules installed.

    .DESCRIPTION
        Every stub throws. A test that reaches one has failed to mock something. Defined in the
        global scope so a mock resolves from a collector script and from the dot-sourced helpers.
        Parameters are only the ones this report passes; parameter binding happens before a Pester
        mock runs, so a parameter a collector passes must be declared here.
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

function global:Get-SPOTenant {
    [CmdletBinding()]
    param()
    throw 'Get-SPOTenant was called for real. Mock it in the test.'
}

# Replaces the shared stub. Search-UnifiedAuditLog -Formatted is what turns RecordType into a
# display name.
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

function global:Get-MgReportSharePointSiteUsageDetail {
    [CmdletBinding()]
    param([string]$Period, [datetime]$Date, [string]$OutFile)
    throw 'Get-MgReportSharePointSiteUsageDetail was called for real. Mock it in the test.'
}

function global:Get-MgReportOneDriveUsageAccountDetail {
    [CmdletBinding()]
    param([string]$Period, [datetime]$Date, [string]$OutFile)
    throw 'Get-MgReportOneDriveUsageAccountDetail was called for real. Mock it in the test.'
}

function global:Get-MgReportSharePointSiteUsageStorage {
    [CmdletBinding()]
    param([string]$Period, [string]$OutFile)
    throw 'Get-MgReportSharePointSiteUsageStorage was called for real. Mock it in the test.'
}

function global:Get-MgReportOneDriveUsageStorage {
    [CmdletBinding()]
    param([string]$Period, [string]$OutFile)
    throw 'Get-MgReportOneDriveUsageStorage was called for real. Mock it in the test.'
}

function global:Get-MgReportSharePointActivityUserDetail {
    [CmdletBinding()]
    param([string]$Period, [datetime]$Date, [string]$OutFile)
    throw 'Get-MgReportSharePointActivityUserDetail was called for real. Mock it in the test.'
}

function global:Get-MgReportOneDriveActivityUserDetail {
    [CmdletBinding()]
    param([string]$Period, [datetime]$Date, [string]$OutFile)
    throw 'Get-MgReportOneDriveActivityUserDetail was called for real. Mock it in the test.'
}

function global:Get-MgAdminReportSetting {
    [CmdletBinding()]
    param()
    throw 'Get-MgAdminReportSetting was called for real. Mock it in the test.'
}

function global:Invoke-MgGraphRequest {
    [CmdletBinding()]
    param([string]$Method, [string]$Uri)
    throw 'Invoke-MgGraphRequest was called for real. Mock it in the test.'
}
