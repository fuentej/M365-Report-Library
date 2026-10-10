# Candidate: Entra sign-in and audit activity

Status: sources verified against Microsoft Learn; nothing built. Written for BRO-383 so the build issue can
be written without guessing. Nothing here was run against a tenant. Pages read 2026-10-10.

Cloud names follow the library: `Commercial`, `GCC`, `GCCHigh` (`-Environment`). Graph uses
`https://graph.microsoft.com` for Commercial and GCC and `Connect-MgGraph -Environment USGov` for GCC High
([national cloud deployments](https://learn.microsoft.com/graph/deployments)). Every Graph page cited below
carries a national-cloud table; "US Government L4" is the GCC High column, and GCC calls the global endpoint.

## How to read the table

* `Available` / `NotAvailable` appear only where a Microsoft page says so; the link is in the cell.
* `UNVERIFIED` means no Microsoft page found says either way. The collector should ask for the data and
  record a refusal in `run.log`, as the existing reports do.
* Sources 1 to 4 are Graph calls through the `Microsoft.Graph` PowerShell SDK or REST. Source 5 is a public
  reference page, not a tenant call.
* The shared sign-in already requests `AuditLog.Read.All`, which is the least privileged permission for
  sources 1, 2 and 4. Source 3 needs a Conditional Access read permission the shared sign-in does not request
  (`Policy.Read.All`, `Policy.Read.ConditionalAccess`, or `Policy.ReadWrite.ConditionalAccess`; the page lists
  the third as a write permission, so the build should take one of the first two).
* Sources 1 and 2 are the **same endpoint**, `GET /auditLogs/signIns`, on v1.0 and beta. The v1.0 page returns
  interactive sign-ins and successful federated sign-ins; the beta page says non-interactive and service
  principal sign-ins come back only when `signInEventTypes` is filtered. Non-interactive history therefore
  depends on a beta API, which Microsoft says is not supported for production applications.
* **The headline finding for GCC High:** both Graph audit-log pages (sign-ins and directory audits) mark
  US Government L4 as available, so the API is `Available` in all three clouds. What no page read states is
  the log retention window in a national cloud, so retention is `UNVERIFIED` for GCC and GCC High. A GCC High
  run should not promise more history than it can read.

## Sources

| # | Source | Endpoint or cmdlet | Least privileged role or permission | License | Event / state | Retention | Commercial | GCC | GCC High |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 1 | Interactive sign-ins: user, app, resource, IP address, `location`, `clientAppUsed`, `deviceDetail`, `status.errorCode`, `status.failureReason`, `conditionalAccessStatus`, risk fields | [`GET /auditLogs/signIns`](https://learn.microsoft.com/graph/api/signin-list) (`Get-MgAuditLogSignIn`). v1.0 includes sign-ins "interactive in nature" and successful federated sign-ins. Maximum and default page size is 1,000, newest first; follow `@odata.nextLink`. Filter `createdDateTime ge ... and createdDateTime le ...`; the page says to always apply a time range to avoid timeouts. Property filters (`eq`) include `userPrincipalName`, `appDisplayName`, `clientAppUsed`, `conditionalAccessStatus`, `status/errorCode` ([signIn](https://learn.microsoft.com/graph/api/resources/signin#properties)) | Application and delegated `AuditLog.Read.All`. Signed in: Reports Reader (the least privileged role named on the [troubleshooting page](https://learn.microsoft.com/entra/identity/monitoring-health/howto-troubleshoot-sign-in-errors)), Global Reader, Security Administrator, Security Operator or Security Reader ([list signIns](https://learn.microsoft.com/graph/api/signin-list#permissions)) | The [signIn resource page](https://learn.microsoft.com/graph/api/resources/signin) says Entra ID P1 or P2 is needed to download sign-in logs through the Graph API. The [licensing page](https://learn.microsoft.com/entra/fundamentals/licensing#microsoft-entra-monitoring-and-health) lists sign-in logs as "Yes" for Entra ID Free. The two pages disagree; the build should treat P1 or P2 as required and record a license refusal as "not licensed". `riskDetail`, `riskLevelAggregated` and `riskLevelDuringSignIn` return `hidden` unless P2 | Event | 7 days Free, 30 days P1 and P2 ([data retention](https://learn.microsoft.com/entra/identity/monitoring-health/reference-reports-data-retention#how-long-does-microsoft-entra-id-store-the-data)). The list page says only events inside the default retention period are available. Appending from the last exported `createdDateTime` on a daily run builds a longer history than the source holds | [Available](https://learn.microsoft.com/graph/api/signin-list) | [Available](https://learn.microsoft.com/graph/deployments) (global service) | [Available](https://learn.microsoft.com/graph/api/signin-list) (US Government L4 ✅) |
| 2 | Non-interactive sign-ins (and, with another `signInEventTypes` value, service principal and managed identity sign-ins) | Beta only: [`GET /beta/auditLogs/signIns?$filter=(createdDateTime ge ... and createdDateTime le ...) and signInEventTypes/any(t: t eq 'nonInteractiveUser')`](https://learn.microsoft.com/graph/api/signin-list?view=graph-rest-beta) (`Get-MgBetaAuditLogSignIn`). The beta page says the API returns only interactive sign-ins unless `signInEventTypes` is filtered. Values shown: `interactiveUser`, `nonInteractiveUser`, `servicePrincipal`, `managedIdentity`. Same 1,000-row page and `@odata.nextLink`. Beta is not supported for production applications | Same as source 1 | Same as source 1 | Event | Same as source 1 | [Available](https://learn.microsoft.com/graph/api/signin-list?view=graph-rest-beta) | [Available](https://learn.microsoft.com/graph/deployments) (global service) | [Available](https://learn.microsoft.com/graph/api/signin-list?view=graph-rest-beta) (US Government L4 ✅ on the beta page) |
| 3 | Conditional Access result on each sign-in: `conditionalAccessStatus` (`success`, `failure`, `notApplied`) and per-policy `appliedConditionalAccessPolicies` (`displayName`, `enforcedGrantControls`, `enforcedSessionControls`, `result`) | Read from the source 1 and 2 responses; no separate call. `result` values include `success`, `failure`, `notApplied`, `notEnabled`, `unknown` and the report-only values `reportOnlySuccess`, `reportOnlyFailure`, `reportOnlyNotApplied`, `reportOnlyInterrupted`, the last four only when the request sends `Prefer: include-unknown-enum-members` ([appliedConditionalAccessPolicy](https://learn.microsoft.com/graph/api/resources/appliedconditionalaccesspolicy)) | `conditionalAccessStatus` needs only the source 1 permission. `appliedConditionalAccessPolicies` is omitted from the response unless the caller can read Conditional Access data: application permission `Policy.Read.All`, `Policy.Read.ConditionalAccess` or `Policy.ReadWrite.ConditionalAccess`; signed-in Global Reader, Security Administrator, Security Reader or Conditional Access Administrator ([list signIns](https://learn.microsoft.com/graph/api/signin-list#permissions)) | Sign-in logs license from source 1. Conditional Access itself needs Entra ID P1 ([licensing](https://learn.microsoft.com/entra/fundamentals/licensing#microsoft-entra-conditional-access)); a tenant without it has no policy results to report | Event | Same as source 1 | [Available](https://learn.microsoft.com/graph/api/signin-list) (documented in the list response and permissions of source 1) | [Available](https://learn.microsoft.com/graph/deployments) (global service) | [Available](https://learn.microsoft.com/graph/api/signin-list) (US Government L4 ✅ on the page that documents the property; the `appliedConditionalAccessPolicy` resource page itself has no cloud table) |
| 4 | Directory audit events: who changed users, groups, roles, applications and policies, when, and the result. Fields `activityDateTime`, `activityDisplayName`, `category` (`UserManagement`, `GroupManagement`, `ApplicationManagement`, `RoleManagement`), `operationType` (`Add`, `Assign`, `Update`, `Unassign`, `Delete`), `result`, `resultReason`, `loggedByService`, `initiatedBy` (user or app), `targetResources` (`User`, `Device`, `Directory`, `App`, `Role`, `Group`, `Policy`, `Other`) with `modifiedProperties` | [`GET /auditLogs/directoryaudits`](https://learn.microsoft.com/graph/api/directoryaudit-list) (`Get-MgAuditLogDirectoryAudit`). Supports `$filter` (`eq`, `ge`, `le`, `startswith`), `$top`, `$orderby`, `skiptoken`. Filter `activityDateTime ge ...` to append from the last run; filter `category`-style questions client-side or on `loggedByService`, `activityDisplayName`, `initiatedBy/user/id` ([directoryAudit](https://learn.microsoft.com/graph/api/resources/directoryaudit#properties)). Includes PIM, access review and password management activity | Application and delegated `AuditLog.Read.All` (`Directory.Read.All` is the higher privileged alternative). Signed in: Reports Reader, Security Administrator or Security Reader | The [licensing page](https://learn.microsoft.com/entra/fundamentals/licensing#microsoft-entra-monitoring-and-health) lists audit logs as "Yes" for Entra ID Free and for P1 or P2. The list page names no license | Event | 7 days Free, 30 days P1 and P2 ([data retention](https://learn.microsoft.com/entra/identity/monitoring-health/reference-reports-data-retention#how-long-does-microsoft-entra-id-store-the-data)) | [Available](https://learn.microsoft.com/graph/api/directoryaudit-list) | [Available](https://learn.microsoft.com/graph/deployments) (global service) | [Available](https://learn.microsoft.com/graph/api/directoryaudit-list) (US Government L4 ✅) |
| 5 | How far back a run can read, by license level (answers starting question 5; not a tenant call) | [Microsoft Entra data retention](https://learn.microsoft.com/entra/identity/monitoring-health/reference-reports-data-retention): audit logs and sign-ins are 7 days on Free and 30 days on P1 and P2. Risky sign-ins (not read here) are 7, 30 and 90 days. Collection starts at sign-up on P1 and P2, and the first time the portal or reporting APIs are used on Free. After an upgrade from Free only the data still inside the 7-day window is kept, so history cannot be recovered. Longer retention means archiving to a storage account or Azure Monitor, or, for organizations with Microsoft 365 E5, Office 365 E5, Microsoft Purview Suite or the E5 eDiscovery and Audit add-on, Purview Audit (Premium) | None; a public page | None | State (reference; the run records which license window applied) | n/a | UNVERIFIED (the page names no cloud) | UNVERIFIED (the page names no cloud) | UNVERIFIED (the page names no cloud) |

Consolidated notes:

* **Sign-in log events are not complete history.** Appending from the last exported timestamp on a daily run
  builds more history than the source retains (30 days at best, 7 on Free). A first run reads at most the
  retention window, so a report cannot show a trend longer than the days the collector has been running plus
  the retention window it found on day one. Each page holds at most 1,000 rows; stopping after the first page
  drops the rest.
* **Interactive vs non-interactive.** A report built on source 1 alone leaves out non-interactive sign-ins.
  Source 2 is beta. The build issue has to accept the beta dependency or the report states that it covers
  interactive sign-ins only.
* **Failed sign-ins and error codes.** `status.errorCode` is a 5-6 digit integer and `status.failureReason` is
  the message ([signInStatus](https://learn.microsoft.com/graph/api/resources/signinstatus)). The v1.0 example shows `0` with a null `failureReason` for a
  successful sign-in. Microsoft lists 50058, 90025, 500121 and 70046 as common but says the list is not
  exhaustive and points to its
  [error code reference](https://learn.microsoft.com/entra/identity/monitoring-health/howto-troubleshoot-sign-in-errors#sign-in-error-codes).
  The report should show the code and the Microsoft `failureReason` text, not a library-written translation.
* **License failures must show "not licensed" instead of zero.** The two Learn pages disagree on whether the
  sign-in logs API works on Free (see source 1).
* **Conditional Access detail is permission-dependent.** A token with only `AuditLog.Read.All` still gets
  `conditionalAccessStatus` but loses `appliedConditionalAccessPolicies` without error, so an empty list is not
  proof that no policy applied. The collector should record which permission set it ran with.
* **Directory audit logs do not carry every directory change.** The page lists what is included (user, app,
  device and group management, PIM, access reviews, terms of use, Identity Protection, password management,
  self-service group management). It does not list Conditional Access policy edits by name and this file did
  not research that.
* **Unified audit log is separate.** Microsoft states that Entra audit and sign-in logs are separate from the
  Microsoft 365 unified audit log and that its retention is not affected by Entra licensing
  ([data retention](https://learn.microsoft.com/entra/identity/monitoring-health/reference-reports-data-retention)).
  This candidate does not read the unified audit log.

## Proposed report pages

| # | Page | Reads | What it shows |
| --- | --- | --- | --- |
| 1 | Overview | 1, 2, 4, 5 | Sign-in count, failure rate, directory changes per day; the retention window the run could read |
| 2 | Sign-ins over time | 1, 2 | Counts by day split interactive vs non-interactive, by user, application (`appDisplayName`), client (`clientAppUsed`) and location (`location.countryOrRegion`, `city`) |
| 3 | Failed sign-ins | 1, 2 | `status.errorCode` and `failureReason` ranked by count, by user and application; success vs failure over time |
| 4 | Conditional Access results | 1, 2, 3 | `conditionalAccessStatus` over time; `appliedConditionalAccessPolicies` result per policy, split enforced and report-only; "no policy detail" when the permission was missing |
| 5 | Directory changes | 4 | Events by `category` (users, groups, roles, applications) and `operationType`; target resource type; success vs failure |
| 6 | Who changed what | 4 | `initiatedBy` user or app, `activityDisplayName`, `targetResources`, `modifiedProperties`, time |
| 7 | Coverage and retention | 5, `run.log` | Retention window by license, first and last event timestamp per source, sources skipped or refused in this cloud, and whether policy detail was readable |

## Starting questions

| # | Question | Status |
| --- | --- | --- |
| 1 | Interactive and non-interactive sign-ins over time, by user, application, location, client and result | Covered by source 1 (interactive) and source 2 (non-interactive, beta). User, application, location, client and result are fields on the signIn object. Available in all three clouds for the API. Retention in GCC and GCC High is UNVERIFIED. |
| 2 | Failed sign-ins and their error codes | Covered by sources 1 and 2: `status.errorCode` and `status.failureReason`. The code list is Microsoft's and not exhaustive. |
| 3 | The Conditional Access result recorded on each sign-in | Covered by source 3: `conditionalAccessStatus` always, `appliedConditionalAccessPolicies` only with a Conditional Access read permission. Available in all three clouds. |
| 4 | Directory audit events: who changed users, groups, roles and applications, and when | Covered by source 4. Available in all three clouds. |
| 5 | How long sign-in and audit logs are kept for each licence level, and how far back a run can read | Covered by source 5 (reference) and the retention column: 7 days on Free, 30 days on P1 and P2. Cloud-specific retention is UNVERIFIED for GCC and GCC High. A run can read at most the retention window, and no further. |

No starting question is dropped. Two limits sit inside question 1: the non-interactive half depends on a beta
API, and a report cannot show more than the retention window of history on its first run.

## Overlap with other candidates

* Identity posture ([identity-posture.md](identity-posture.md), BRO-315, BRO-321, BRO-326): its sources 6 and 6b
  read the same `GET /auditLogs/signIns` endpoint, filtered to legacy `clientAppUsed` values, and its source 5
  reads per-user `signInActivity`. This report reads the full event stream, the Conditional Access result and
  directory audits, and does not rebuild MFA, role or risky-user content. Whether the two reports share one
  sign-in collector is a decision for the build issue.

## Open items for the build issue

* `signInActivity` (last sign-in per user) is not a source here. If a page ever needs it, its GCC High
  availability is UNVERIFIED and unsettled; see BRO-243.
* Source 2 is beta. Accept it, or ship interactive sign-ins only.
* Settle whether to ship a shared sign-in collector with identity posture, or two collectors reading the same
  endpoint.
* Decide how a first run is described when the retention window is shorter than the report period.
* Whether the Learn pages' disagreement on Free sign-in log access matters: a Free tenant may be refused by the
  API even though the licensing page lists "Yes".
