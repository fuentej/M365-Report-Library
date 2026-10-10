#Requires -Version 7.0

<#
    .SYNOPSIS
        Declares the Graph and Exchange Online commands this report calls that
        shared/tests/TenantCmdletStubs.ps1 does not, so the tests can mock them without those
        modules installed.

    .DESCRIPTION
        Every stub throws. A test that reaches one has failed to mock something. Defined in the
        global scope so a mock resolves from a collector script and from the dot-sourced helpers.
        Parameters are only the ones this report passes; parameter binding happens before a Pester
        mock runs, so a parameter a collector passes must be declared here.
#>

function global:Get-MgReportTeamUserActivityUserDetail {
    [CmdletBinding()]
    param([string]$Period, [datetime]$Date, [string]$OutFile)
    throw 'Get-MgReportTeamUserActivityUserDetail was called for real. Mock it in the test.'
}

function global:Get-MgReportTeamUserActivityCount {
    [CmdletBinding()]
    param([string]$Period, [string]$OutFile)
    throw 'Get-MgReportTeamUserActivityCount was called for real. Mock it in the test.'
}

function global:Get-MgReportTeamDeviceUsageUserDetail {
    [CmdletBinding()]
    param([string]$Period, [datetime]$Date, [string]$OutFile)
    throw 'Get-MgReportTeamDeviceUsageUserDetail was called for real. Mock it in the test.'
}

function global:Get-MgAdminReportSetting {
    [CmdletBinding()]
    param()
    throw 'Get-MgAdminReportSetting was called for real. Mock it in the test.'
}

function global:Invoke-MgGraphRequest {
    [CmdletBinding()]
    param([string]$Method, [string]$Uri, [hashtable]$Headers)
    throw 'Invoke-MgGraphRequest was called for real. Mock it in the test.'
}

# Replaces the shared stub. -Formatted is what turns RecordType into a display name.
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
