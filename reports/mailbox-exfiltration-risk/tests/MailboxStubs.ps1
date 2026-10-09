#Requires -Version 7.0

<#
    .SYNOPSIS
        Declares the Exchange Online and Graph cmdlets this report calls that the shared
        stubs do not, so the tests can mock them without those modules installed.

    .DESCRIPTION
        Every stub throws. A test that reaches one has failed to mock something. Defined in
        the global scope so a mock resolves from a collector script and from the
        dot-sourced helpers. Parameters are only the ones this report passes.
#>

function global:Get-EXOMailbox {
    [CmdletBinding()]
    param([object]$ResultSize, [string[]]$PropertySets, [string]$Identity)
    throw 'Get-EXOMailbox was called for real. Mock it in the test.'
}

function global:Get-AcceptedDomain {
    [CmdletBinding()]
    param([object]$ResultSize)
    throw 'Get-AcceptedDomain was called for real. Mock it in the test.'
}

function global:Get-InboxRule {
    [CmdletBinding()]
    param([string]$Mailbox, [switch]$IncludeHidden, [object]$ResultSize)
    throw 'Get-InboxRule was called for real. Mock it in the test.'
}

function global:Get-TransportRule {
    [CmdletBinding()]
    param([object]$ResultSize, [bool]$ExcludeConditionActionDetails)
    throw 'Get-TransportRule was called for real. Mock it in the test.'
}

function global:Get-MailboxPermission {
    [CmdletBinding()]
    param([string]$Identity, [object]$ResultSize)
    throw 'Get-MailboxPermission was called for real. Mock it in the test.'
}

function global:Get-RecipientPermission {
    [CmdletBinding()]
    param([string]$Identity, [object]$ResultSize)
    throw 'Get-RecipientPermission was called for real. Mock it in the test.'
}

function global:Get-OrganizationConfig {
    [CmdletBinding()]
    param()
    throw 'Get-OrganizationConfig was called for real. Mock it in the test.'
}

function global:Get-AdminAuditLogConfig {
    [CmdletBinding()]
    param()
    throw 'Get-AdminAuditLogConfig was called for real. Mock it in the test.'
}

function global:Get-MgOauth2PermissionGrant {
    [CmdletBinding()]
    param([switch]$All)
    throw 'Get-MgOauth2PermissionGrant was called for real. Mock it in the test.'
}

function global:Get-MgServicePrincipal {
    [CmdletBinding()]
    param([string]$Filter, [switch]$All)
    throw 'Get-MgServicePrincipal was called for real. Mock it in the test.'
}

function global:Get-MgServicePrincipalAppRoleAssignedTo {
    [CmdletBinding()]
    param([string]$ServicePrincipalId, [switch]$All)
    throw 'Get-MgServicePrincipalAppRoleAssignedTo was called for real. Mock it in the test.'
}
