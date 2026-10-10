#Requires -Version 7.0

<#
    .SYNOPSIS
        Declares the Microsoft Graph cmdlets the license utilization collectors call
        that shared/tests/TenantCmdletStubs.ps1 does not, so the tests can mock them
        without the Graph modules installed.

    .DESCRIPTION
        Every stub throws. A test that reaches one has failed to mock something. Defined
        in the global scope so a mock resolves from a collector script and from the
        dot-sourced helpers. Parameters are only the ones this report passes.
#>

function global:Get-MgSubscribedSku {
    [CmdletBinding()]
    param([switch]$All, [string[]]$Property)
    throw 'Get-MgSubscribedSku was called for real. Mock it in the test.'
}

function global:Get-MgUserLicenseDetail {
    [CmdletBinding()]
    param([string]$UserId, [switch]$All, [string[]]$Property)
    throw 'Get-MgUserLicenseDetail was called for real. Mock it in the test.'
}

function global:Get-MgAdminReportSetting {
    [CmdletBinding()]
    param([string[]]$Property)
    throw 'Get-MgAdminReportSetting was called for real. Mock it in the test.'
}

function global:Invoke-MgGraphRequest {
    [CmdletBinding()]
    param([string]$Method, [string]$Uri, [string]$OutputFilePath)
    throw 'Invoke-MgGraphRequest was called for real. Mock it in the test.'
}
