@{
    # The column order of every CSV this report produces. The collectors, the
    # sample-data generator and the tests all read the order from here, so a
    # sample file can never drift from its collector's output.

    Guests = @(
        'RunDate'
        'Id'
        'DisplayName'
        'Mail'
        'UserPrincipalName'
        'ExternalDomain'
        'CreatedDateTime'
        'CreationType'
        'ExternalUserState'
        'ExternalUserStateChangeDateTime'
        'AccountEnabled'
        'LastSignInDateTime'
        'LastNonInteractiveSignInDateTime'
        'LastSuccessfulSignInDateTime'
    )

    GuestInvitations = @(
        'ActivityDateTime'
        'Id'
        'ActivityDisplayName'
        'Result'
        'InitiatedByUserPrincipalName'
        'InitiatedByAppDisplayName'
        'TargetUserId'
        'TargetUserPrincipalName'
    )

    GuestSignIns = @(
        'CreatedDateTime'
        'Id'
        'UserId'
        'UserPrincipalName'
        'AppDisplayName'
        'ResourceDisplayName'
        'IpAddress'
        'City'
        'CountryOrRegion'
        'ClientAppUsed'
        'IsInteractive'
        'ErrorCode'
        'ConditionalAccessStatus'
    )

    SharingEvents = @(
        'CreationTime'
        'Id'
        'Operation'
        'UserId'
        'Workload'
        'SiteUrl'
        'ObjectId'
        'SourceFileName'
        'TargetUserOrGroupName'
        'TargetUserOrGroupType'
    )

    GuestMemberships = @(
        'RunDate'
        'GuestId'
        'GroupId'
        'GroupDisplayName'
        'IsTeam'
        'Visibility'
    )

    # Entra audit activities that record a guest invitation or its redemption.
    # Verified against the "Invited users" and "B2B Auth" sections of
    # https://learn.microsoft.com/entra/identity/monitoring-health/reference-audit-activities
    # "Redeem extern user invite" is spelled that way in the service and in the
    # reference; it is not a typo in this file.
    InvitationActivities = @(
        'Invite external user'
        'Invite external user with reset invitation status'
        'Invite internal user to B2B collaboration'
        'Invitation Email'
        'Redeem external user invite'
        'Redeem extern user invite'
        'Bulk invite users - finished (bulk)'
    )

    # Sharing and access request activities in the unified audit log:
    # https://learn.microsoft.com/purview/audit-log-activities#sharing-and-access-request-activities
    SharingOperations = @(
        'SharingSet'
        'SharingRevoked'
        'SharingInvitationCreated'
        'SharingInvitationAccepted'
        'SharingInvitationRevoked'
        'AnonymousLinkCreated'
        'AnonymousLinkUsed'
        'AnonymousLinkRemoved'
        'SecureLinkCreated'
        'AddedToSecureLink'
        'SecureLinkUsed'
    )
}
