# Teams and Groups lifecycle

Which Microsoft 365 groups and Teams have no owner or only one, which are inactive, which
are archived or soft-deleted and when they purge, whether an expiration policy exists and
which groups it covers, and who created what.

The collectors write one CSV per source into an output folder of your choosing, plus a
`run.log`. Snapshot files get a new block of rows tagged `RunDate` on every run; the
creation-event file is appended from where the last run stopped. Nothing is rewritten or
deleted, so the folder is a history you can chart over time.

The sources were verified against Microsoft Learn in
[`docs/candidates/teams-groups-lifecycle.md`](../../docs/candidates/teams-groups-lifecycle.md),
which is the contract for this folder. The Power BI project is in `report/`; saving it as
a `.pbit` needs Power BI Desktop and is a separate step. Guest membership in groups is out of scope: the [Guest and external
access](../guest-access/README.md) report covers it.

## Contents

| Path | What it is |
| --- | --- |
| `collectors/Get-Groups.ps1` | `groups.csv` — a snapshot of every Microsoft 365 group |
| `collectors/Get-GroupOwners.ps1` | `group-owners.csv` — the owners of each group |
| `collectors/Get-DeletedGroups.ps1` | `deleted-groups.csv` — soft-deleted groups still inside the restore window |
| `collectors/Get-GroupLifecyclePolicies.ps1` | `group-lifecycle-policies.csv` — the group expiration policy |
| `collectors/Get-GroupLifecycleCoverage.ps1` | `group-lifecycle-coverage.csv` — whether the policy covers each group |
| `collectors/Get-TeamActivity.ps1` | `team-activity.csv` — Teams usage by team |
| `collectors/Get-GroupActivity.ps1` | `group-activity.csv` — Microsoft 365 groups usage by group |
| `collectors/Get-ArchivedTeams.ps1` | `team-archive-status.csv` — whether each team is archived |
| `collectors/Get-GroupCreationEvents.ps1` | `group-creation-events.csv` — group and team creation, with the creator |
| `collectors/Run-All.ps1` | Runs the shared users collector and all nine of the above, in dependency order |
| `collectors/TeamsGroupsSchema.psd1` | The column order of every CSV, the per-cloud availability of every source |
| `collectors/TeamsGroupsHelpers.ps1` | Small helpers the collectors dot-source |
| `New-SampleData.ps1` | Regenerates `samples/` |
| `samples/` | Generated fake data; `samples/gcchigh/` holds the header-only files GCC High writes |
| `report/` | The Power BI project (PBIP): `TeamsGroupsLifecycle.SemanticModel` in TMDL and `TeamsGroupsLifecycle.Report` in PBIR |
| `tests/` | Pester tests; every tenant call is mocked. `ReportSchema`, `ReportModel` and `TmdlDataType` hold the PBIP to its schemas and to these CSVs |

`users.csv` comes from `Invoke-EntraUserCollector` in the shared module, not from this
folder, because every report needs it. Join it to `group-owners.csv` to find owners whose
account is disabled.

## Before you run it

### Modules

```powershell
Install-Module Microsoft.Graph.Authentication, Microsoft.Graph.Groups, Microsoft.Graph.Identity.DirectoryManagement, Microsoft.Graph.Teams, Microsoft.Graph.Users, Microsoft.Graph.Reports -Scope CurrentUser
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

### Licences

| Source | Licence needed |
| --- | --- |
| Groups, owners, soft-deleted groups, archived teams | None beyond Microsoft 365 |
| Group expiration policy | Microsoft Entra ID P1 or P2 for the members of every group the policy applies to ([expiration](https://learn.microsoft.com/entra/identity/users/groups-lifecycle)). A Selected scope holds at most 500 groups |
| Usage reports | None beyond Microsoft 365. The organization setting that conceals user, group and site names blanks the names in these reports ([show details](https://learn.microsoft.com/microsoft-365/admin/activity-reports/activity-reports#show-user-group-or-site-details-in-usage-reports)) |
| Creation events | [Audit (Standard)](https://learn.microsoft.com/purview/audit-solutions-overview#audit-standard), retaining 180 days; auditing must be turned on |

A source the tenant is not licensed for, or the sign-in cannot read, leaves a header-only
CSV and a line in `run.log` saying so. The run carries on.

### Microsoft Graph permissions and roles

| Collector | Permission | Role or condition |
| --- | --- | --- |
| `Get-Groups.ps1` | `Group.Read.All` | None named on [list groups](https://learn.microsoft.com/graph/api/group-list). The page's least-privileged `Group-NestingSupport.ReadWrite.All` only reads and writes `disableNesting`, so it is not used |
| `Get-GroupOwners.ps1` | `GroupMember.Read.All` | Delegated callers also need a role listed on [list owners](https://learn.microsoft.com/graph/api/group-list-owners): group owners, Member users, Guest users (limited), or Directory Readers |
| `Get-DeletedGroups.ps1` | `Group.Read.All` | |
| `Get-GroupLifecyclePolicies.ps1`, `Get-GroupLifecycleCoverage.ps1` | `Directory.Read.All` | Entra ID P1 or P2 licence, above |
| `Get-TeamActivity.ps1`, `Get-GroupActivity.ps1` | `Reports.Read.All` | Delegated callers need a limited admin role such as Reports Reader ([authorization](https://learn.microsoft.com/graph/reportroot-authorization)). **Global Reader and Usage Summary Reports Reader see tenant totals only, not the detail rows** |
| `Get-ArchivedTeams.ps1` | `Team.ReadBasic.All` (delegated) | If a token with only that permission omits `isArchived`, use `TeamSettings.Read.All`. Application: `TeamSettings.Read.Group` (resource-specific consent) or `Team.ReadBasic.All` for an organization-wide read |
| `Get-GroupCreationEvents.ps1` | none (Exchange Online) | View-Only Audit Logs or Audit Logs ([audit permissions](https://learn.microsoft.com/purview/audit-get-started#step-2-assign-permissions-to-search-the-audit-log)); app-only needs `Exchange.ManageAsApp` |

An interactive sign-in to Graph asks for the library's default read scopes
(`User.Read.All`, `Directory.Read.All`, `GroupMember.Read.All`, `AuditLog.Read.All`). Those
do not include `Group.Read.All`, `Reports.Read.All` or `Team.ReadBasic.All`; pass
`-Scopes` through `Connect-M365Service` or sign in app-only with them granted as
application permissions with admin consent.

## Availability by cloud

`-Environment` is `Commercial`, `GCC` or `GCCHigh` (default `Commercial`). Graph uses
`https://graph.microsoft.com` for Commercial and GCC and `Connect-MgGraph -Environment USGov`
(`https://graph.microsoft.us`) for GCC High
([national cloud deployments](https://learn.microsoft.com/graph/deployments)).

* **Available** — collected.
* **NotAvailable** — the collector skips the source: it writes a header-only CSV and logs why.
* **UNVERIFIED** — no Microsoft page found says either way. The collector asks for the data,
  logs a warning, and records any refusal in `run.log`.

| CSV | Commercial | GCC | GCC High |
| --- | --- | --- | --- |
| `groups.csv` | [Available](https://learn.microsoft.com/graph/api/group-list) | [Available](https://learn.microsoft.com/graph/deployments) (global service) | [Available](https://learn.microsoft.com/graph/api/group-list) |
| `group-owners.csv` | [Available](https://learn.microsoft.com/graph/api/group-list-owners) | [Available](https://learn.microsoft.com/graph/deployments) (global service) | [Available](https://learn.microsoft.com/graph/api/group-list-owners) |
| `deleted-groups.csv` | [Available](https://learn.microsoft.com/graph/api/directory-deleteditems-list) | [Available](https://learn.microsoft.com/graph/deployments) (global service) | [Available](https://learn.microsoft.com/graph/api/directory-deleteditems-list) |
| `group-lifecycle-policies.csv`, `group-lifecycle-coverage.csv` | [Available](https://learn.microsoft.com/graph/api/grouplifecyclepolicy-list) | [Available](https://learn.microsoft.com/graph/deployments) (global service) | [Available](https://learn.microsoft.com/graph/api/grouplifecyclepolicy-list) |
| `team-activity.csv` | [Available](https://learn.microsoft.com/graph/api/reportroot-getteamsteamactivitydetail) | UNVERIFIED | [NotAvailable](https://learn.microsoft.com/graph/api/reportroot-getteamsteamactivitydetail) |
| `group-activity.csv` | [Available](https://learn.microsoft.com/graph/api/reportroot-getoffice365groupsactivitydetail) | UNVERIFIED | [NotAvailable](https://learn.microsoft.com/graph/api/reportroot-getoffice365groupsactivitydetail) |
| `team-archive-status.csv` | [Available](https://learn.microsoft.com/graph/api/team-get) | [Available](https://learn.microsoft.com/graph/api/team-get) (global service) | [Available](https://learn.microsoft.com/graph/api/team-get) |
| `group-creation-events.csv` | [Available](https://learn.microsoft.com/purview/audit-solutions-overview#comparison-of-key-capabilities) (Audit (Standard)) | Audit (Standard) available; `AddGroup` and `TeamCreated` records UNVERIFIED | Audit (Standard) available; `AddGroup` and `TeamCreated` records UNVERIFIED |

Why the usage reports are `UNVERIFIED` in GCC: each API page marks the global service
available and US Government L4 (GCC High) not available, and the usage-report overview marks
"Microsoft Cloud for US Government" unavailable without separating GCC from GCC High. GCC
calls the global endpoint, so a live GCC tenant would decide. In GCC High the inactivity
page therefore has no API source, even though the same reports exist in the Microsoft 365
admin center. See [the contract](../../docs/candidates/teams-groups-lifecycle.md) for the
Learn pages behind each cell.

## Running it

```powershell
./collectors/Run-All.ps1 -OutputPath ./out

./collectors/Run-All.ps1 -OutputPath ./out -Environment GCCHigh -AppId $appId `
    -CertificateThumbprint $thumbprint -TenantId $tenantId -Organization contoso.onmicrosoft.com
```

Order matters: `groups.csv` first, because `Get-GroupOwners.ps1`,
`Get-GroupLifecycleCoverage.ps1` and `Get-ArchivedTeams.ps1` read the groups it lists, and
`Get-GroupLifecyclePolicies.ps1` before `Get-GroupLifecycleCoverage.ps1`. `Run-All.ps1`
follows that order and carries on if one collector stops.

Listing: groups, deleted groups and owners are paged. Graph returns 100 objects a page by
default and 999 at most, `$skip` is not supported on list groups, and every collector asks
with `-All`, which follows `@odata.nextLink` until it is absent
([paging](https://learn.microsoft.com/graph/paging)).

## The CSVs

Timestamps are UTC, `yyyy-MM-ddTHH:mm:ssZ`. Lists are joined with `;`.

### groups.csv — snapshot, key `RunDate` + `Id`

Microsoft 365 groups only (`groupTypes/any(c:c eq 'Unified')`). `createdDateTime` is here
for every group, so creation over time needs no audit record.

| Column | Meaning |
| --- | --- |
| `RunDate` | UTC date of the run |
| `Id`, `DisplayName`, `Mail` | The group |
| `GroupTypes` | `Unified` for a Microsoft 365 group |
| `SecurityEnabled`, `MailEnabled`, `Visibility` | As on the group resource |
| `CreatedDateTime`, `RenewedDateTime` | When the group was created and last renewed |
| `ExpirationDateTime` | When the expiration policy will expire the group; empty when no policy applies. The group is deleted one day after this, and the 30-day restore window then starts |
| `DeletedDateTime` | Empty for an active group |
| `OnPremisesSyncEnabled` | `True` for a group synchronized from on-premises |
| `ResourceProvisioningOptions` | Contains `Team` for most teams |
| `IsTeam` | `ResourceProvisioningOptions` contains `Team`. Certain unused old teams do not carry that value, so use `team-archive-status.csv` to find every team |

### group-owners.csv — snapshot, key `RunDate` + `GroupId` + `OwnerId`

| Column | Meaning |
| --- | --- |
| `GroupId` | The group |
| `OwnerListStatus` | `Listed` (one row per owner), `None` (the call succeeded and returned nobody), or `Unknown` (Graph cannot return owners for this group, or the call failed) |
| `OwnerId`, `OwnerType`, `OwnerDisplayName`, `OwnerUserPrincipalName` | The owner; empty unless `Listed` |

Owners are not available for groups created in Exchange, distribution groups, or groups
synchronized from on-premises. Report those as "owner unknown", not "no owner".

### deleted-groups.csv — snapshot, key `RunDate` + `Id`

| Column | Meaning |
| --- | --- |
| `Id`, `DisplayName`, `GroupTypes`, `SecurityEnabled`, `MailEnabled`, `CreatedDateTime` | The group. A soft-deleted security group returns `SecurityEnabled` false; `GroupTypes` (`Unified` or empty) tells Microsoft 365 groups from security groups |
| `DeletedDateTime` | When the group was deleted |
| `PurgeDateTime` | `DeletedDateTime` plus 30 days, when a Microsoft 365 or security group is permanently deleted. Distribution groups are deleted immediately and never appear |
| `ResourceProvisioningOptions`, `IsTeam` | As in `groups.csv` |

### group-lifecycle-policies.csv — snapshot, key `RunDate` + `Id`

| Column | Meaning |
| --- | --- |
| `Id` | The policy; a tenant has at most one |
| `GroupLifetimeInDays` | Days before a group expires |
| `ManagedGroupTypes` | `All`, `Selected` or `None` |
| `AlternateNotificationEmails` | Where notices go for groups without an owner |

A tenant with no policy writes the header only.

### group-lifecycle-coverage.csv — snapshot, key `RunDate` + `GroupId`

| Column | Meaning |
| --- | --- |
| `GroupId`, `PolicyId` | The group and the policy (empty when the tenant has none) |
| `CoverageStatus` | `Covered`, `NotCovered`, or `Unknown` (the per-group call failed). `All` and `None` are worked out from the policy; only a `Selected` policy costs one call per group |

### team-activity.csv — snapshot, key `RunDate` + `TeamId`

A snapshot of a rolling period, `D180` by default. `LastActivityDate` is the latest activity
whatever the period; the counts cover the period.

| Column | Meaning |
| --- | --- |
| `ReportRefreshDate`, `ReportPeriod` | When Microsoft refreshed the report, and its period in days |
| `TeamId`, `TeamName`, `TeamType` | The team (`Team type` and `Team Type` are both accepted). Names are blank when the conceal-names setting is on |
| `LastActivityDate` | Latest activity |
| `ActiveUsers`, `ActiveChannels`, `Guests`, `Reactions`, `MeetingsOrganized`, `PostMessages`, `ReplyMessages`, `ChannelMessages`, `UrgentMessages`, `Mentions`, `ActiveSharedChannels`, `ActiveExternalUsers` | Activity counts for the period |

### group-activity.csv — snapshot, key `RunDate` + `GroupId`

| Column | Meaning |
| --- | --- |
| `ReportRefreshDate`, `ReportPeriod` | As above |
| `GroupId`, `GroupDisplayName`, `GroupType`, `IsDeleted`, `OwnerPrincipalName` | The group. Names are blank when the conceal-names setting is on |
| `LastActivityDate` | Latest activity; not capped at the period |
| `MemberCount`, `ExternalMemberCount` | Members and external members. The Learn page names this column `External Member Count` in its header list and `Guest Count` in its schema example; whichever is present is read |
| `ExchangeReceivedEmailCount`, `ExchangeMailboxTotalItemCount`, `ExchangeMailboxStorageUsedByte` | Mailbox activity and size |
| `SharePointActiveFileCount`, `SharePointTotalFileCount`, `SharePointSiteStorageUsedByte` | Site activity and size |
| `YammerPostedMessageCount`, `YammerReadMessageCount`, `YammerLikedMessageCount` | Viva Engage activity |

### team-archive-status.csv — snapshot, key `RunDate` + `TeamId`

| Column | Meaning |
| --- | --- |
| `TeamId`, `DisplayName` | The team |
| `IsArchived` | `isArchived` from `GET /teams/{team-id}`. `GET /groups` does not return it |

Every Microsoft 365 group in the latest `groups.csv` snapshot is asked; a group that is not
a team answers 404 and has no row.

### group-creation-events.csv — events, key `Id`

Resumes from the latest `CreationTime` already in the file. Audit (Standard) keeps 180
days, so a group created earlier has no creator here; `groups.csv` still has its
`CreatedDateTime`.

| Column | Meaning |
| --- | --- |
| `CreationTime`, `Id` | When the record was written, and its audit id |
| `Operation` | `AddGroup` ("Added group") or `TeamCreated` ("Created team"), copied exactly. `AddGroup` covers Microsoft 365 groups and security groups created in the admin center or the Azure portal; it does not replace `TeamCreated` |
| `UserId` | The creator |
| `Workload` | The service that wrote the record |
| `ObjectId` | The object the record names |
| `TargetDisplayName`, `TargetGroupId` | `TeamName` and `TeamGuid` from the record; empty where the record has neither (Learn documents no group-name property for `AddGroup`) |

## Sample data

`samples/` is generated by `New-SampleData.ps1`: three monthly snapshots of 48 groups
(30 of them Teams), six soft-deleted groups, ownerless, single-owner and unknown-owner
groups, archived teams, a policy scoped to Selected groups, and creation events inside the
last 180 days. It is deterministic and uses only `example.com`. `samples/gcchigh/` holds
the header-only `team-activity.csv` and `group-activity.csv` that GCC High writes.

## The Power BI report

Open `report/TeamsGroupsLifecycle.pbip` in Power BI Desktop and set the `CsvFolder`
parameter to the folder the collectors wrote (it defaults to `C:\TeamsGroupsLifecycleData`;
point it at `samples/` to see the sample data). Six pages follow the contract: Overview,
Ownership, Inactivity, Expiration policy, Archived and deleted, and Creation. Every page has
a date range, a Group or Team, and a Group slicer, plus the Anonymize toggle, which swaps
names for stable pseudonyms (`Group 12345`, `User 12345`) derived from the id.

* Counts are for the latest snapshot inside the date range. An empty value is Unknown, never
  zero: owners Graph cannot list are "Owner unknown", an empty last activity date is
  "Unknown", and the activity cards are blank when a snapshot has no usage-report rows (the
  usage reports are not available in GCC High, and GCC is unverified).
* `users.csv` is joined to owners to find disabled owners; an owner missing from it is Unknown.
