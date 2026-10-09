@{
    # The column order of every CSV this report produces. The collectors, the
    # sample-data generator and the tests all read the order from here, so a
    # sample file can never drift from its collector's output.
    # Contract: docs/candidates/mailbox-exfiltration-risk.md
    # users.csv comes from Invoke-EntraUserCollector in the shared module.

    # Source 1b. Get-AcceptedDomain.
    AcceptedDomains = @(
        'RunDate'
        'Name'
        'DomainName'
        'DomainType'
        'IsDefault'
    )

    # Source 1. One row per mailbox that has ForwardingAddress or ForwardingSmtpAddress set.
    # ForwardingAddress is an internal recipient. ForwardingSmtpAddress is an SMTP forward:
    # IsExternal is True when its domain is not an accepted domain, False when it is, and
    # empty when the accepted domains could not be read. A mailbox forwarding only to a
    # ForwardingAddress is False.
    MailboxForwarding = @(
        'RunDate'
        'ExternalDirectoryObjectId'
        'UserPrincipalName'
        'PrimarySmtpAddress'
        'ForwardingAddress'
        'ForwardingSmtpAddress'
        'ForwardingSmtpDomain'
        'DeliverToMailboxAndForward'
        'IsExternal'
    )

    # Source 4c. One row per delegate in GrantSendOnBehalfTo.
    SendOnBehalf = @(
        'RunDate'
        'ExternalDirectoryObjectId'
        'UserPrincipalName'
        'PrimarySmtpAddress'
        'Delegate'
    )

    # Source 2. Get-InboxRule properties named on New-InboxRule: ForwardTo,
    # ForwardAsAttachmentTo, RedirectTo, DeleteMessage. TargetDomains is the SMTP domains
    # found in those three recipient lists; HasExternalTarget is True when any is not an
    # accepted domain. Only rules that forward, redirect or delete are written.
    InboxRules = @(
        'RunDate'
        'MailboxUserPrincipalName'
        'MailboxExternalDirectoryObjectId'
        'RuleIdentity'
        'RuleName'
        'Enabled'
        'Priority'
        'ForwardTo'
        'ForwardAsAttachmentTo'
        'RedirectTo'
        'DeleteMessage'
        'TargetDomains'
        'HasExternalTarget'
    )

    # Source 3. Only rules with a redirect or blind-copy action are written; CopyTo and
    # AddToRecipients add visible recipients and are carried for context.
    TransportRules = @(
        'RunDate'
        'Name'
        'Guid'
        'State'
        'Priority'
        'RedirectMessageTo'
        'BlindCopyTo'
        'CopyTo'
        'AddToRecipients'
        'TargetDomains'
        'HasExternalTarget'
    )

    # Source 4a. Full Access rows that survive the filter: AccessRights like 'Full*',
    # Deny false, IsInherited false, and not NT AUTHORITY\SELF.
    MailboxFullAccess = @(
        'RunDate'
        'MailboxExternalDirectoryObjectId'
        'MailboxUserPrincipalName'
        'User'
        'AccessRights'
    )

    # Source 4b.
    SendAsPermissions = @(
        'RunDate'
        'Identity'
        'Trustee'
        'AccessRights'
        'AccessControlType'
        'IsInherited'
    )

    # Source 5a. https://learn.microsoft.com/graph/api/resources/oauth2permissiongrant
    # HasMailScope is True when a scope in Scope starts with Mail. or MailboxSettings.
    DelegatedConsents = @(
        'RunDate'
        'Id'
        'ClientId'
        'ConsentType'
        'PrincipalId'
        'ResourceId'
        'Scope'
        'HasMailScope'
    )

    # Source 5b. https://learn.microsoft.com/graph/api/resources/approleassignment
    # Assignments of the Microsoft Graph service principal's app roles. IsMailRole is True
    # when AppRoleValue starts with Mail. or MailboxSettings.
    AppRoleAssignments = @(
        'RunDate'
        'AssignmentId'
        'AppRoleId'
        'AppRoleValue'
        'PrincipalId'
        'PrincipalDisplayName'
        'PrincipalType'
        'ResourceId'
        'ResourceDisplayName'
        'CreatedDateTime'
        'IsMailRole'
    )

    # Sources 6a and 6b share these columns. Parameters is the Name=Value pairs of an
    # Exchange admin record, joined with semicolons.
    AuditEvents = @(
        'CreationTime'
        'Id'
        'RecordType'
        'Operation'
        'UserId'
        'Workload'
        'ObjectId'
        'MailboxOwnerUPN'
        'ClientIP'
        'ResultStatus'
        'Parameters'
    )

    # Source 6c. Scope is Organization or Mailbox. AuditDisabled and
    # UnifiedAuditLogIngestionEnabled are set on the Organization row; the rest on Mailbox
    # rows. Get-Mailbox always shows AuditEnabled True, so it is not collected.
    AuditConfiguration = @(
        'RunDate'
        'Scope'
        'Identity'
        'AuditDisabled'
        'UnifiedAuditLogIngestionEnabled'
        'DefaultAuditSet'
        'AuditAdmin'
        'AuditDelegate'
        'AuditOwner'
    )

    # Source 6a. Mailbox operations, copied exactly from the contract doc.
    MailboxChangeOperations = @(
        'New-InboxRule'
        'Set-InboxRule'
        'UpdateInboxRules'
        'Set-Mailbox'
        'Add-MailboxPermission'
        'Remove-MailboxPermission'
    )

    # Source 6b.
    MailAccessOperations = @(
        'MailItemsAccessed'
        'Send'
        'SendAs'
        'SendOnBehalf'
    )

    # Available / NotAvailable / Unverified, from the source table in the contract doc.
    # Only NotAvailable skips a source. No source in the doc is NotAvailable in any cloud.
    SourceAvailability = @{
        AcceptedDomains = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/powershell/module/exchangepowershell/get-accepteddomain' }
            GCC        = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/powershell/module/exchangepowershell/get-accepteddomain' }
            GCCHigh    = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/powershell/module/exchangepowershell/get-accepteddomain' }
        }
        MailboxForwarding = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/powershell/module/exchangepowershell/get-mailbox' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/office-365-platform-service-description/office-365-us-government/exchange-online-for-us-government-environments#exchange-online-features' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/office-365-platform-service-description/office-365-us-government/exchange-online-for-us-government-environments#exchange-online-features' }
        }
        SendOnBehalf = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/powershell/module/exchangepowershell/get-mailbox' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/office-365-platform-service-description/office-365-us-government/exchange-online-for-us-government-environments#exchange-online-features' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/office-365-platform-service-description/office-365-us-government/exchange-online-for-us-government-environments#exchange-online-features' }
        }
        InboxRules = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/powershell/module/exchangepowershell/get-inboxrule' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/office-365-platform-service-description/office-365-us-government/exchange-online-for-us-government-environments#exchange-online-features' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/office-365-platform-service-description/office-365-us-government/exchange-online-for-us-government-environments#exchange-online-features' }
        }
        TransportRules = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/exchange/security-and-compliance/mail-flow-rules/mail-flow-rules' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/office-365-platform-service-description/office-365-us-government/exchange-online-for-us-government-environments#exchange-online-features' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/office-365-platform-service-description/office-365-us-government/exchange-online-for-us-government-environments#exchange-online-features' }
        }
        MailboxFullAccess = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/exchange/recipients-in-exchange-online/manage-permissions-for-recipients' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/office-365-platform-service-description/office-365-us-government/exchange-online-for-us-government-environments#exchange-online-features' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/office-365-platform-service-description/office-365-us-government/exchange-online-for-us-government-environments#exchange-online-features' }
        }
        SendAsPermissions = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/powershell/module/exchangepowershell/get-recipientpermission' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/office-365-platform-service-description/office-365-us-government/exchange-online-for-us-government-environments#exchange-online-features' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/office-365-platform-service-description/office-365-us-government/exchange-online-for-us-government-environments#exchange-online-features' }
        }
        DelegatedConsents = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/oauth2permissiongrant-list' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/oauth2permissiongrant-list' }
        }
        AppRoleAssignments = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/serviceprincipal-list-approleassignedto' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/deployments' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/graph/api/serviceprincipal-list-approleassignedto' }
        }
        MailboxChangeEvents = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/purview/audit-solutions-overview#comparison-of-key-capabilities' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-deployments' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-high-deployments' }
        }
        MailAccessEvents = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/purview/audit-log-investigate-accounts' }
            GCC        = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-deployments' }
            GCCHigh    = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-high-deployments' }
        }
        AuditConfiguration = @{
            Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/purview/audit-mailboxes#verify-mailbox-auditing-on-by-default-is-turned-on' }
            GCC        = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/purview/audit-mailboxes#verify-mailbox-auditing-on-by-default-is-turned-on' }
            GCCHigh    = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/purview/audit-mailboxes#verify-mailbox-auditing-on-by-default-is-turned-on' }
        }
    }
}
