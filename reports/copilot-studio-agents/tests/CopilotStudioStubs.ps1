#Requires -Version 7.0

<#
    .SYNOPSIS
        Declares the Az sign-in cmdlets the Copilot Studio collectors call, so the tests
        can mock them without the Az modules installed.

    .DESCRIPTION
        Every stub throws. A test that reaches one has failed to mock something. Defined
        in the global scope so a mock resolves from a collector script and from the
        dot-sourced helpers. Parameters are only the ones this report passes.
#>

function global:Connect-AzAccount {
    [CmdletBinding()]
    param([string]$Environment, [string]$Tenant)
    throw 'Connect-AzAccount was called for real. Mock it in the test.'
}

function global:Get-AzAccessToken {
    [CmdletBinding()]
    param([string]$ResourceUrl, [switch]$AsSecureString)
    throw 'Get-AzAccessToken was called for real. Mock it in the test.'
}
