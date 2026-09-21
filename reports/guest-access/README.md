# Guest and external access

Who the guests in your tenant are, who invited them, whether they ever showed up, what
they can still reach, and what has been shared outside the organisation.

The collectors write one CSV per source into an output folder of your choosing, plus a
`run.log`. Snapshot files get a new block of rows on every run; event files are appended
from where the last run stopped. Nothing is ever rewritten or deleted, so the folder is a
history you can chart over time.

The Power BI report that reads these files is tracked separately. `samples/` holds a
generated set of fake data so the report can be built without a tenant.

## Contents

| Path | What it is |
| --- | --- |
| `collectors/Get-Guests.ps1` | `guests.csv` — a snapshot of every guest |
| `collectors/Get-GuestInvitations.ps1` | `guest-invitations.csv` — invitation and redemption events |
| `collectors/Get-GuestSignIns.ps1` | `guest-signins.csv` — sign-ins by guests |
| `collectors/Get-SharingEvents.ps1` | `sharing-events.csv` — SharePoint and OneDrive sharing activity |
| `collectors/Get-GuestMemberships.ps1` | `guest-memberships.csv` — the groups each guest is in |
| `collectors/Run-All.ps1` | Runs the users collector and all five of the above |
| `collectors/GuestAccessSchema.psd1` | The column order of every CSV, and the activity and operation names collected |
| `New-SampleData.ps1` | Regenerates `samples/` |
| `tests/` | Pester tests; every tenant call is mocked |

`users.csv` comes from `Invoke-EntraUserCollector` in the shared module, not from this
folder, because every report needs it.

## Before you run it

### Modules

```powershell
Install-Module Microsoft.Graph.Users, Microsoft.Graph.Groups, Microsoft.Graph.Reports, Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

### Licences

| Source | Licence needed |
| --- | --- |
| Guest list (without sign-in activity) | None beyond Microsoft 365 / Entra ID Free |
| `signInActivity` on a guest | [Entra ID P1 or P2](https://learn.microsoft.com/entra/identity/monitoring-health/howto-manage-inactive-user-accounts) |
| Directory audit log | Entra ID Free retains 7 days; P1/P2 retains 30 ([retention](https://learn.microsoft.com/entra/identity/monitoring-health/reference-reports-data-retention)) |
| Sign-in log | [Entra ID P1 or P2](https://learn.microsoft.com/entra/identity/monitoring-health/reference-reports-data-retention) |
| Unified audit log | [Audit (Standard)](https://learn.microsoft.com/purview/audit-solutions-overview#audit-standard), retaining 180 days ([retention](https://learn.microsoft.com/purview/audit-log-retention-policies)) |
| Group memberships | None beyond Microsoft 365 |

A source the tenant is not licensed for leaves a header-only CSV and a line in `run.log`
saying so. The run carries on.

### Microsoft Graph permissions

| Permission | Used for |
| --- | --- |
| `User.Read.All` | `users.csv`, `guests.csv` |
| `AuditLog.Read.All` | `signInActivity`, `guest-invitations.csv`, `guest-signins.csv` |
| `Directory.Read.All` | manager and group lookups |
| `GroupMember.Read.All` | `guest-memberships.csv` |

An interactive sign-in asks for exactly these. App-only needs them granted as application
permissions with admin consent.

### Directory and Exchange roles

| Role | Needed for | Reference |
| --- | --- | --- |
| Reports Reader | the sign-in log and the directory audit log | [directoryAudits](https://learn.microsoft.com/graph/api/directoryaudit-list), [signIns](https://learn.microsoft.com/graph/api/signin-list) |
| View-Only Audit Logs, or Audit Logs | `Search-UnifiedAuditLog` | [Audit get started](https://learn.microsoft.com/purview/audit-get-started) |

Auditing also has to be turned on for the organisation before the unified audit log
returns anything — see the same page.

## Cloud availability

Every collector takes `-Environment Commercial|GCC|GCCHigh`, defaulting to `Commercial`.
The shared module holds the endpoints; see the root README for the table.

| Source | Commercial | GCC | GCC High | Reference |
| --- | --- | --- | --- | --- |
| Graph list users (`guests.csv`, `users.csv`) | Available | Available | Available | [national cloud deployments](https://learn.microsoft.com/graph/deployments) |
| `signInActivity` on a user | Available (P1/P2) | Available (P1/P2) | **UNVERIFIED** | [inactive user accounts](https://learn.microsoft.com/entra/identity/monitoring-health/howto-manage-inactive-user-accounts) — no per-cloud statement found for US Gov L4 |
| Directory audit log | Available | Available | Available | [directoryAudits](https://learn.microsoft.com/graph/api/directoryaudit-list), [deployments](https://learn.microsoft.com/graph/deployments) |
| Sign-in log | Available (P1/P2) | Available (P1/P2) | Available (P1/P2) | [signIns](https://learn.microsoft.com/graph/api/signin-list), [deployments](https://learn.microsoft.com/graph/deployments) |
| Unified audit log (`Search-UnifiedAuditLog`) | Available | Available | Available | [Audit (Standard) in GCC](https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-deployments), [in GCC High](https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-high-deployments) |
| Group memberships (`memberOf`) | Available | Available | Available | [list memberOf](https://learn.microsoft.com/graph/api/user-list-memberof), [deployments](https://learn.microsoft.com/graph/deployments) |

Where a row says UNVERIFIED, the collector does not assume either way: it asks for the
data, and if the service refuses, it leaves those columns empty and records the reason in
`run.log`.

Individual audit *activities* can also be missing from a government cloud even when the
log itself is available — the [audit log activities
reference](https://learn.microsoft.com/purview/audit-log-activities) footnotes each one.

## Running it

```powershell
# Everything, interactively, against the commercial cloud.
./collectors/Run-All.ps1 -OutputPath ./out

# GCC High, app-only.
./collectors/Run-All.ps1 -OutputPath ./out -Environment GCCHigh `
    -AppId $appId -CertificateThumbprint $thumbprint `
    -TenantId $tenantId -Organization contoso.onmicrosoft.com

# One collector, over a window you choose.
./collectors/Get-SharingEvents.ps1 -OutputPath ./out `
    -StartDate (Get-Date).AddDays(-30) -EndDate (Get-Date) -WindowHours 6
```

`-Organization` is the tenant's `*.onmicrosoft.com` domain, and app-only sign-in to
Exchange Online needs it.

Order matters: `Get-GuestSignIns.ps1` and `Get-GuestMemberships.ps1` read the guest list
out of `guests.csv`, so run `Get-Guests.ps1` first. `Run-All.ps1` already does.

### Resuming

Each event collector starts at the latest timestamp already in its CSV, so running it
daily costs one short query. Pass `-StartDate` to override, or `-LookbackDays` to set how
far back a first run reaches. If a collector has not run for longer than the log's
retention window, the gap cannot be recovered — the service no longer holds it.

### Re-running on the same day

Snapshot files are keyed on `RunDate` plus the object's id, so running twice in one day
does not duplicate a snapshot. Event files are keyed on `Id`.

## The files

### `guests.csv` — snapshot

| Column | Meaning |
| --- | --- |
| `RunDate` | The UTC date of the run that wrote this block of rows |
| `Id` | The guest's object id in Entra ID |
| `DisplayName` | As shown in the directory |
| `Mail` | The address the invitation went to |
| `UserPrincipalName` | The B2B UPN, of the form `alice_partner.example.com#EXT#@contoso.onmicrosoft.com` |
| `ExternalDomain` | The guest's home domain, taken from `Mail` or parsed out of the UPN |
| `CreatedDateTime` | When the guest object was created, which is when they were invited |
| `CreationType` | How the account came about — `Invitation` for a normal B2B guest |
| `ExternalUserState` | `PendingAcceptance` or `Accepted`. Pending means the invitation was never redeemed |
| `ExternalUserStateChangeDateTime` | When that state last changed, i.e. when they redeemed |
| `AccountEnabled` | False for a guest that has been disabled but not removed |
| `LastSignInDateTime` | Last interactive sign-in. Empty without Entra ID P1/P2 |
| `LastNonInteractiveSignInDateTime` | Last non-interactive sign-in — a token refresh counts, so this stays recent longer |
| `LastSuccessfulSignInDateTime` | Last sign-in that actually succeeded |

### `guest-invitations.csv` — events, keyed on `Id`

| Column | Meaning |
| --- | --- |
| `ActivityDateTime` | When the activity was recorded |
| `Id` | The audit record's id |
| `ActivityDisplayName` | One of the activities listed below |
| `Result` | `success`, `failure`, `timeout` |
| `InitiatedByUserPrincipalName` | Who invited them, or the guest themself on a redemption |
| `InitiatedByAppDisplayName` | The app that did it, when a service rather than a person |
| `TargetUserId` | The guest's object id |
| `TargetUserPrincipalName` | The guest's UPN |

The activities collected, verified against the "Invited users" and "B2B Auth" sections of
the [Entra audit activity
reference](https://learn.microsoft.com/entra/identity/monitoring-health/reference-audit-activities):

- `Invite external user`
- `Invite external user with reset invitation status`
- `Invite internal user to B2B collaboration`
- `Invitation Email`
- `Redeem external user invite`
- `Redeem extern user invite` — spelled that way in the service and in the reference
- `Bulk invite users - finished (bulk)`

The same reference also lists `Delete external user` and `Email not sent, user
unsubscribed` under "Invited users". They are not invitation or redemption events, so
they are not collected; add them to `InvitationActivities` in
`collectors/GuestAccessSchema.psd1` if you want them.

### `guest-signins.csv` — events, keyed on `Id`

| Column | Meaning |
| --- | --- |
| `CreatedDateTime` | When the sign-in happened |
| `Id` | The sign-in record's id |
| `UserId` | The guest's object id |
| `UserPrincipalName` | The guest's UPN |
| `AppDisplayName` | The application signed in to |
| `ResourceDisplayName` | The resource that application asked for |
| `IpAddress` | The client's address as the service saw it |
| `City`, `CountryOrRegion` | Where the service placed that address |
| `ClientAppUsed` | Browser, mobile, desktop, or a legacy protocol |
| `IsInteractive` | False for token refreshes and other background sign-ins |
| `ErrorCode` | `0` on success; anything else is the failure reason |
| `ConditionalAccessStatus` | `success`, `failure`, `notApplied` |

The v1.0 `signIn` resource has no supported filter on guest status, so the collector
queries by date and keeps the rows whose `UserId` is in `guests.csv`.

### `sharing-events.csv` — events, keyed on `Id`

Taken from each record's `AuditData`.

| Column | Meaning |
| --- | --- |
| `CreationTime` | When the sharing action happened |
| `Id` | The audit record's id |
| `Operation` | One of the operations listed below |
| `UserId` | Who did it. An anonymous link use shows as `urn:spo:anon#...` |
| `Workload` | `SharePoint` or `OneDrive` |
| `SiteUrl` | The site the item lives in |
| `ObjectId` | The full path of the item |
| `SourceFileName` | The file or folder name |
| `TargetUserOrGroupName` | Who it was shared with. Empty for an anonymous link |
| `TargetUserOrGroupType` | `Guest`, `Member`, `Group`, `SecurityGroup`, `Partner` |

The operations collected are the [sharing and access request
activities](https://learn.microsoft.com/purview/audit-log-activities#sharing-and-access-request-activities):
`SharingSet`, `SharingRevoked`, `SharingInvitationCreated`, `SharingInvitationAccepted`,
`SharingInvitationRevoked`, `AnonymousLinkCreated`, `AnonymousLinkUsed`,
`AnonymousLinkRemoved`, `SecureLinkCreated`, `AddedToSecureLink`, `SecureLinkUsed`.

`Search-UnifiedAuditLog` returns at most 50,000 unsorted results per session, so the
collector walks the range in windows and starts a new session for each. A window that
reaches the cap is logged as possibly truncated; re-run it with a smaller `-WindowHours`.

### `guest-memberships.csv` — snapshot

| Column | Meaning |
| --- | --- |
| `RunDate` | The UTC date of the run that wrote this block of rows |
| `GuestId` | The guest's object id |
| `GroupId` | The group's object id |
| `GroupDisplayName` | The group's name |
| `IsTeam` | True when the group's `resourceProvisioningOptions` contains `Team` |
| `Visibility` | `Public`, `Private` or `HiddenMembership` |

### `users.csv` — snapshot

Written by the shared module so that every report can join to it. Columns: `RunDate`,
`Id`, `DisplayName`, `UserPrincipalName`, `Mail`, `UserType`, `AccountEnabled`,
`CreatedDateTime`, `Department`, `JobTitle`, `City`, `Country`, `ManagerId`,
`ManagerUserPrincipalName`. `UserType` is `Member` or `Guest`, so guests appear here too;
`ManagerId` is what lets a report roll guest sponsorship up a management chain.

### `run.log`

One line per event: `2026-09-21T20:51:00Z [Info] [guests] guests.csv: 60 rows written, 0
skipped.` Levels are `Info`, `Warning` and `Error`. An `Error` line always says which
source was skipped and why.

## The sample data

`samples/` is generated by `New-SampleData.ps1` and holds nothing real. Every address is
under `example.com` or a `partnerN.example.com` subdomain, and every sign-in address is in
`203.0.113.0/24`, both reserved for documentation.

It contains 200 members across six departments with a management chain three levels deep,
60 guests from 12 partner domains, six monthly snapshots and 240 days of events —
including guests who never accepted, guests quiet for more than 90 days, disabled guests,
anonymous and secure links, revoked sharing, invitations from 26 different members, and
guests in both Teams and plain groups.

The generator is deterministic: the same `-Seed` and `-EndDate` produce the same files,
so regenerating shows up in a diff only when the generator itself changes.

```powershell
./New-SampleData.ps1
```

## Tests

```powershell
pwsh -NoProfile -Command "Invoke-Pester -Path ./shared/tests, ./reports/guest-access/tests -CI"
```

Run from the root of the repository. No tenant is contacted: the Graph, Exchange Online
and Security & Compliance cmdlets are stubbed in `shared/tests/TenantCmdletStubs.ps1` and
mocked per test.
