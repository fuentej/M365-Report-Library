# Candidate: Teams and Groups lifecycle

Status: sources verified against Microsoft Learn; nothing built. Written for BRO-314 so the build issue can
be written without guessing. Nothing here was run against a tenant.

Cloud names follow the library: `Commercial`, `GCC`, `GCCHigh` (`-Environment`). Graph uses
`https://graph.microsoft.com` for Commercial and GCC and `Connect-MgGraph -Environment USGov` for GCC High
([national cloud deployments](https://learn.microsoft.com/graph/deployments)). Every Graph page cited below
carries a national-cloud table; "US Government L4" is the GCC High column, and GCC calls the global endpoint.
Guest membership in groups is out of scope: the Guest access report covers it.

## Finding that shapes the report

The Microsoft 365 usage report APIs (`getTeamsTeamActivityDetail`, `getOffice365GroupsActivityDetail`) that
answer "which Teams and groups are inactive" are **not available in GCC High**: their Graph pages list the
global service only, and the usage-report overview shows "Microsoft Cloud for US Government" as unavailable
([usage reports in Graph](https://learn.microsoft.com/graph/api/resources/report#cloud-deployments)). The same
reports do exist in the Microsoft 365 admin center for GCC and GCC High
([usage reports overview](https://learn.microsoft.com/microsoft-365/admin/activity-reports/activity-reports#available-usage-reports-in-the-microsoft-365-admin-center)),
but that is a portal view, not a collector source. In GCC High the inactivity page has no verified API source;
audit records are the fallback.

## How to read the table

* `Available` / `NotAvailable` appear only where a Microsoft page says so; the link is in the cell.
* `UNVERIFIED` means no Microsoft page found says either way. The collector should ask for the data and
  record a refusal in `run.log`, as the existing reports do.
* The library's rule for an unavailable source is to skip it and say why in the README. Sources 5 and 6 would
  be skipped in GCC High.

## Sources

| # | Source | Endpoint or cmdlet | Least privileged role or permission | Event / state | Retention | Commercial | GCC | GCC High |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 1 | Groups and Teams inventory (type, `createdDateTime`, `expirationDateTime`, `resourceProvisioningOptions` contains `Team`, `deletedDateTime`) | [`GET /groups`](https://learn.microsoft.com/graph/api/group-list) (`Get-MgGroup`); filter `groupTypes/any(c:c eq 'Unified')` for Microsoft 365 groups. Any group with a team has `resourceProvisioningOptions` containing `Team` ([Teams API overview](https://learn.microsoft.com/graph/api/resources/teams-api-overview#teams-and-groups)); `expirationDateTime` is set for Microsoft 365 groups by the lifecycle policy ([group properties](https://learn.microsoft.com/graph/templates/terraform/reference/v1.0/groups#property-values)) | `Group.Read.All` is listed among the permissions; the page's "least privileged" row (`Group-NestingSupport.ReadWrite.All`) is a write permission, so the least privileged read permission is UNVERIFIED. Entra role: not read | State | n/a | [Available](https://learn.microsoft.com/graph/api/group-list) | [Available](https://learn.microsoft.com/graph/deployments) (global service) | [Available](https://learn.microsoft.com/graph/api/group-list) (page lists US Government L4) |
| 2 | Group owners | [`GET /groups/{id}/owners`](https://learn.microsoft.com/graph/api/group-list-owners) (`Get-MgGroupOwner`). Owners are not available for groups created in Exchange, distribution groups, or groups synchronized from on-premises | `GroupMember.Read.All` is the first permission listed on the cmdlet page ([Get-MgGroupOwner](https://learn.microsoft.com/powershell/module/microsoft.graph.groups/get-mggroupowner)). Entra role: not read | State | n/a | [Available](https://learn.microsoft.com/graph/api/group-list-owners) | [Available](https://learn.microsoft.com/graph/deployments) (global service) | [Available](https://learn.microsoft.com/graph/api/group-list-owners) (US Government L4) |
| 3 | Soft-deleted groups | [`GET /directory/deletedItems/microsoft.graph.group`](https://learn.microsoft.com/graph/api/directory-deleteditems-list) (`Get-MgDirectoryDeletedItemAsGroup`); each item carries `deletedDateTime`. Deleted groups are retained 30 days, then permanently deleted ([Delete group](https://learn.microsoft.com/graph/api/group-delete)). Purge date is `deletedDateTime` plus 30 days, a report calculation | Application and delegated `Group.Read.All` ([permissions](https://learn.microsoft.com/graph/api/directory-deleteditems-list#permissions)) | State (a snapshot of what is still recoverable; the entry disappears at 30 days) | 30 days | [Available](https://learn.microsoft.com/graph/api/directory-deleteditems-list) | [Available](https://learn.microsoft.com/graph/deployments) (global service) | [Available](https://learn.microsoft.com/graph/api/directory-deleteditems-list) (US Government L4) |
| 4 | Group expiration policy (lifetime in days, `managedGroupTypes` All, Selected or None, notification email) and per-group coverage | [`GET /groupLifecyclePolicies`](https://learn.microsoft.com/graph/api/grouplifecyclepolicy-list) (one policy per tenant); [`GET /groups/{id}/groupLifecyclePolicies`](https://learn.microsoft.com/graph/api/group-list-grouplifecyclepolicies) for a group. Near-expiry is `expirationDateTime` from source 1 | `Directory.Read.All` (both pages) | State | n/a | [Available](https://learn.microsoft.com/graph/api/grouplifecyclepolicy-list) | [Available](https://learn.microsoft.com/graph/deployments) (global service) | [Available](https://learn.microsoft.com/graph/api/grouplifecyclepolicy-list) (US Government L4). License needed for the policy: UNVERIFIED (not stated on these pages) |
| 5 | Team activity (last activity date, active users, active channels, messages) | [`GET /reports/getTeamsTeamActivityDetail(period='D180')`](https://learn.microsoft.com/graph/api/reportroot-getteamsteamactivitydetail) returns a CSV via a 302 redirect; includes `Last Activity Date` and `Team Id` | `Reports.Read.All` (application and delegated). Delegated callers need a limited admin role ([authorization](https://learn.microsoft.com/graph/reportroot-authorization)) | Snapshot of a rolling period (D7, D30, D90 or D180) | Periods up to 180 days | [Available](https://learn.microsoft.com/graph/api/reportroot-getteamsteamactivitydetail) | UNVERIFIED (Graph pages list the global service; the overview table marks "Microsoft Cloud for US Government" unavailable, [usage reports](https://learn.microsoft.com/graph/api/resources/report#cloud-deployments), and no page separates GCC). The admin center report is available in GCC ([overview](https://learn.microsoft.com/microsoft-365/admin/activity-reports/activity-reports#available-usage-reports-in-the-microsoft-365-admin-center)) | [NotAvailable](https://learn.microsoft.com/graph/api/reportroot-getteamsteamactivitydetail) (page lists US Government L4 as not available) |
| 6 | Microsoft 365 group activity (last activity date, owner, member count, external member count, mailbox and site activity, `Is Deleted`) | [`GET /reports/getOffice365GroupsActivityDetail(period='D180')`](https://learn.microsoft.com/graph/api/reportroot-getoffice365groupsactivitydetail) returns a CSV via a 302 redirect | `Reports.Read.All`. Delegated callers need a limited admin role | Snapshot of a rolling period (D7 to D180) | Periods up to 180 days | [Available](https://learn.microsoft.com/graph/api/reportroot-getoffice365groupsactivitydetail) | UNVERIFIED (same reasoning as source 5) | [NotAvailable](https://learn.microsoft.com/graph/api/resources/report#cloud-deployments) ("Microsoft 365 groups activity": Microsoft Cloud for US Government not available) |
| 7 | Archived teams | `Get-Team -Archived $true` in the MicrosoftTeams PowerShell module ([Get-Team](https://learn.microsoft.com/powershell/module/microsoftteams/get-team)). The List groups page also names an `isArchived` property that `$select` does not return ([List groups](https://learn.microsoft.com/graph/api/group-list)); how to read it was not verified | Role: UNVERIFIED (not stated on the cmdlet page) | State | n/a | [Available](https://learn.microsoft.com/powershell/module/microsoftteams/get-team) | UNVERIFIED (no page found states cmdlet availability in GCC) | UNVERIFIED (same; Teams Graph APIs are listed as available in GCC High, [Plan for government clouds](https://learn.microsoft.com/microsoftteams/platform/concepts/cloud-overview#teams-app-capabilities), which does not name this property) |
| 8 | Group and Team creation events with the creator | Microsoft Purview audit: `Search-UnifiedAuditLog`, operation `AddGroup` ("Added group") ([group administration activities](https://learn.microsoft.com/purview/audit-log-activities#microsoft-entra-group-administration-activities)). Entra directory audit logs also record group changes ([audit logs](https://learn.microsoft.com/entra/identity/monitoring-health/concept-audit-logs)); `GET /auditLogs/directoryAudits` ([overview](https://learn.microsoft.com/graph/api/resources/azure-ad-auditlog-overview)) | Purview: View-Only Audit Logs or Audit Logs ([audit permissions](https://learn.microsoft.com/purview/audit-get-started#step-2-assign-permissions-to-search-the-audit-log)). Entra directory audits: not verified here | Event | Purview: 180 days (Audit Standard), one year for Microsoft Entra records of E5 users ([audit-search](https://learn.microsoft.com/purview/audit-search#before-you-search-the-audit-log)). Entra audit logs: 7 days Free, 30 days P1 and P2 ([Entra retention](https://learn.microsoft.com/entra/identity/monitoring-health/reference-reports-data-retention)) | [Available](https://learn.microsoft.com/purview/audit-solutions-overview#comparison-of-key-capabilities) | [Available](https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-deployments#step-4-understand-which-capabilities-are-currently-unavailable-or-disabled-by-default-in-microsoft-365-government-%E2%80%93-gcc%5E1%5E) (Audit (Standard) Available). That `AddGroup` is recorded in GCC: UNVERIFIED | [Available](https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-high-deployments#step-4-understand-which-capabilities-are-currently-unavailable-or-disabled-by-default-in-microsoft-365-government-%E2%80%93-gcc-high%5E1%5E) (Audit (Standard) Available). That `AddGroup` is recorded in GCC High: UNVERIFIED |

Consolidated notes:

* Source 1 gives `createdDateTime` for every group, so creation over time works without audit records. Only
  the creator needs source 8, and only for creations inside its retention window.
* Sources 5 and 6 return CSV through a redirect, not JSON. Preauthenticated download URLs are valid for a
  few minutes ([getTeamsTeamActivityDetail](https://learn.microsoft.com/graph/api/reportroot-getteamsteamactivitydetail#response)).
* The Teams reports in the Teams admin center are available in Public, GCC, GCC High and DoD
  ([Teams reporting reference](https://learn.microsoft.com/microsoftteams/teams-analytics-and-reports/teams-reporting-reference#teams-reporting-reference)).
  They are a UI source and are not used here.

## Proposed report pages

| # | Page | Reads | What it shows |
| --- | --- | --- | --- |
| 1 | Overview | 1, 2, 3, 4 | Counts of groups and Teams, ownerless, single-owner, expiring, soft-deleted; trend by snapshot date |
| 2 | Ownership | 1, 2 | Groups and Teams with zero or one owner; owners who are disabled (join to the library's `users.csv`) |
| 3 | Inactivity | 1, 5, 6 | Last activity date per Team and group, bucketed by days; not available in GCC High |
| 4 | Expiration policy | 1, 4 | Policy lifetime and scope, groups covered, groups expiring in the next 30, 60 and 90 days |
| 5 | Archived and deleted | 3, 7 | Archived Teams; soft-deleted groups with deletion date and purge date |
| 6 | Creation | 1, 8 | Groups and Teams created per month; by creator where audit records exist |

## Starting questions

| # | Question | Status |
| --- | --- | --- |
| 1 | Which groups and Teams have no owner, or only one owner | Covered by sources 1 and 2. Groups whose owners Graph cannot return (Exchange-created, distribution, on-premises synced) must be reported as "owner unknown", not "no owner". |
| 2 | Which Teams and groups are inactive | Covered by sources 5 and 6 in Commercial. GCC availability UNVERIFIED. **Dropped for GCC High**: both APIs are documented as not available there; no other API source was found. |
| 3 | Which Teams are archived | Source 7, with availability UNVERIFIED in both government clouds |
| 3 | ...which groups are soft-deleted and when they purge | Covered by source 3 |
| 4 | Whether an expiration policy exists, which groups it covers, which are near expiry | Covered by sources 4 and 1. License requirement UNVERIFIED. |
| 5 | Group and Team creation over time, by creator | Creation over time: source 1. Creator: source 8, limited to the audit retention window (180 days, or longer with Audit Premium). |

## Open items for the build issue

* Confirm whether the usage report APIs work in GCC, or treat GCC like GCC High and skip them.
* Find the least privileged read permission for `GET /groups` and confirm `Group.Read.All` as the choice.
* Confirm how to read archived state through Graph, or accept the Teams PowerShell module as a dependency.
* Confirm the license, if any, that the group expiration policy needs.
