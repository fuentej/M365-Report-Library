#Requires -Version 7.0

<#
    .SYNOPSIS
        Declares the Exchange Online and Graph cmdlets this report calls that the shared
        stubs do not, so the tests can mock them without those modules installed.

    .DESCRIPTION
        Every stub throws. A test that reaches one has failed to mock something. Defined in
        the global scope so a mock resolves from a collector script and from the dot-sourced
        helpers. Parameters are only the ones this report passes.
#>

function global:Get-EXOMailbox {
    [CmdletBinding()]
    param([object]$ResultSize, [string[]]$PropertySets, [string]$Identity, [switch]$IncludeInactiveMailbox, [switch]$SoftDeletedMailbox)
    throw 'Get-EXOMailbox was called for real. Mock it in the test.'
}

function global:Get-EXOMailboxStatistics {
    [CmdletBinding()]
    param([string]$Identity, [string[]]$PropertySets, [switch]$Archive)
    throw 'Get-EXOMailboxStatistics was called for real. Mock it in the test.'
}

function global:Get-EXOMobileDeviceStatistics {
    [CmdletBinding()]
    param([string]$Mailbox, [switch]$ActiveSync, [switch]$RestApi, [switch]$OWAforDevices, [switch]$UniversalOutlook)
    throw 'Get-EXOMobileDeviceStatistics was called for real. Mock it in the test.'
}

function global:Get-MessageTraceV2 {
    [CmdletBinding()]
    param(
        [datetime]$StartDate, [datetime]$EndDate, [string[]]$SenderAddress, [string[]]$RecipientAddress,
        [int]$ResultSize, [string]$StartingRecipientAddress, [guid]$MessageTraceId
    )
    throw 'Get-MessageTraceV2 was called for real. Mock it in the test.'
}

function global:Get-MgReportMailboxUsageDetail {
    [CmdletBinding()]
    param([string]$Period, [string]$OutFile)
    throw 'Get-MgReportMailboxUsageDetail was called for real. Mock it in the test.'
}

function global:Get-MgReportMailboxUsageStorage {
    [CmdletBinding()]
    param([string]$Period, [string]$OutFile)
    throw 'Get-MgReportMailboxUsageStorage was called for real. Mock it in the test.'
}

function global:Get-MgReportEmailActivityUserDetail {
    [CmdletBinding()]
    param([string]$Period, [datetime]$Date, [string]$OutFile)
    throw 'Get-MgReportEmailActivityUserDetail was called for real. Mock it in the test.'
}

function global:Get-MgReportEmailAppUsageUserDetail {
    [CmdletBinding()]
    param([string]$Period, [datetime]$Date, [string]$OutFile)
    throw 'Get-MgReportEmailAppUsageUserDetail was called for real. Mock it in the test.'
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
