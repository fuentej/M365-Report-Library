#Requires -Version 7.0

<#
    .SYNOPSIS
        Declares the cmdlets this report calls that the shared stubs do not, so the tests can mock
        them without the Exchange Online and Graph modules installed.

    .DESCRIPTION
        Every stub throws. A test that reaches one has failed to mock something. Defined in the
        global scope so a mock resolves from a collector script and from the dot-sourced helpers.
        Parameters are only the ones this report passes. Search-UnifiedAuditLog,
        Disconnect-ExchangeOnline and the Connect-* stubs come from shared/tests/TenantCmdletStubs.ps1.
#>

function global:Get-AdminAuditLogConfig {
    [CmdletBinding()]
    param()
    throw 'Get-AdminAuditLogConfig was called for real. Mock it in the test.'
}

function global:Get-UnifiedAuditLogRetentionPolicy {
    [CmdletBinding()]
    param()
    throw 'Get-UnifiedAuditLogRetentionPolicy was called for real. Mock it in the test.'
}

function global:Invoke-MgGraphRequest {
    [CmdletBinding()]
    param([string]$Method, [string]$Uri, [object]$Body, [string]$ContentType)
    throw 'Invoke-MgGraphRequest was called for real. Mock it in the test.'
}
