# Candidate: license utilization

Status: sources verified against Microsoft Learn; nothing built. Written for BRO-364 so the build issue can
be written without guessing. Nothing here was run against a tenant. Pages read 2026-10-10.

Cloud names follow the library: `Commercial`, `GCC`, `GCCHigh` (`-Environment`). Graph uses
`https://graph.microsoft.com` for Commercial and GCC and `Connect-MgGraph -Environment USGov` for GCC High
([national cloud deployments](https://learn.microsoft.com/graph/deployments)). Every Graph page cited below
carries a national-cloud table; "US Government L4" is the GCC High column, and GCC calls the global endpoint.

## How to read the table

* `Available` / `NotAvailable` appear only where a Microsoft page says so; the link is in the cell.
* `UNVERIFIED` means no Microsoft page found says either way. The collector should ask for the data and
  record a refusal in `run.log`, as the existing reports do.
* Source 2 reuses `Invoke-EntraUserCollector` in `shared/M365ReportLibrary.psm1` for the user attributes. That
  collector writes `Department`, `JobTitle`, `City` and `Country` and does **not** collect `assignedLicenses`,
  `licenseAssignmentStates`, `officeLocation`, `usageLocation` or `signInActivity`, so this report needs its own
  `$select` for them.
* The shared sign-in requests `User.Read.All`, `Directory.Read.All`, `GroupMember.Read.All` and
  `AuditLog.Read.All`. The build still has to add `LicenseAssignment.Read.All` (or rely on `Directory.Read.All`,
  which the pages list as a higher privileged alternative), `Reports.Read.All` and `ReportSettings.Read.All`.
* **The headline finding:** every Microsoft 365 usage report API the usage questions need (sources 5a to 5g and
  6) is marked `❌` for US Government L4 on its Learn page, so those sources are `NotAvailable` through Graph in
  GCC High. The same reports do exist in the Microsoft 365 admin center for GCC High (source 8). This report
  is therefore a full report in Commercial and GCC and a licence inventory with no usage data in GCC High
  unless the admin center path is used by hand.

## Sources

| # | Source | Endpoint or cmdlet | Least privileged role or permission | License | Event / state | Retention or period | Commercial | GCC | GCC High |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 1 | Licences the tenant owns: units bought, assigned, suspended, and the service plans each SKU holds | [`GET /subscribedSkus`](https://learn.microsoft.com/graph/api/subscribedsku-list) (`Get-MgSubscribedSku`). Per SKU: `skuId`, `skuPartNumber`, `capabilityStatus`, `consumedUnits` (licences assigned), `prepaidUnits` (`enabled`, `suspended`, `warning`, `lockedOut`) and `servicePlans` ([subscribedSku](https://learn.microsoft.com/graph/api/resources/subscribedsku)). Only SKUs with `appliesTo` `User` are assignable. Does not support `$filter`. Free units are not a field; compute `prepaidUnits.enabled` minus `consumedUnits` | `LicenseAssignment.Read.All`. Higher: `Directory.Read.All`, `Organization.Read.All`. Signed in: Global Reader or Directory Readers (also Dynamics 365 Business Central Administrator, standard properties only) | None named on the page | State | n/a | [Available](https://learn.microsoft.com/graph/api/subscribedsku-list) | [Available](https://learn.microsoft.com/graph/deployments) (global service) | [Available](https://learn.microsoft.com/graph/api/subscribedsku-list) (US Government L4) |
| 2 | Users and the licences they hold, with department, job title and location | [`GET /users?$select=id,userPrincipalName,assignedLicenses,licenseAssignmentStates,usageLocation,department,jobTitle,officeLocation,city,country,accountEnabled`](https://learn.microsoft.com/graph/api/user-list) (`Get-MgUser -Property ...`). `assignedLicenses` includes inherited (group-based) licences and does not say which; `licenseAssignmentStates` does: `assignedByGroup` is `null` for a direct assignment and the group id otherwise, `state` is `Active`, `ActiveWithError`, `Disabled` or `Error`, `disabledPlans` lists plans switched off, `error` holds the failure ([user](https://learn.microsoft.com/graph/api/resources/user), [licenseAssignmentState](https://learn.microsoft.com/graph/api/resources/licenseassignmentstate)). Those properties return only with `$select`. Default page 100, maximum 999 | The list page names `User.ReadBasic.All` as least privileged delegated and `User.Read.All` as least privileged application. No page found says whether `User.ReadBasic.All` returns `assignedLicenses`, so use `User.Read.All`, which the shared sign-in already requests | None named on the page | State | n/a | [Available](https://learn.microsoft.com/graph/api/user-list) | [Available](https://learn.microsoft.com/graph/deployments) (global service) | [Available](https://learn.microsoft.com/graph/api/user-list) (US Government L4) |
| 3 | Per-user licence detail with the service plans in each licence (optional; needed only if the join of 1 and 2 is not enough) | [`GET /users/{id}/licenseDetails`](https://learn.microsoft.com/graph/api/user-list-licensedetails) (`Get-MgUserLicenseDetail`). Returns directly assigned and group-assigned licences, each with `skuId`, `skuPartNumber` and `servicePlans` (`servicePlanId`, `servicePlanName`, `provisioningStatus`). One call per user | Delegated `LicenseAssignment.Read.All`; the page lists **Application: Not supported**. Signed in: Guest Inviter, Directory Readers, Directory Writers, License Administrator or User Administrator | None named on the page | State | n/a | [Available](https://learn.microsoft.com/graph/api/user-list-licensedetails) | [Available](https://learn.microsoft.com/graph/deployments) (global service) | [Available](https://learn.microsoft.com/graph/api/user-list-licensedetails) (US Government L4) |
| 4 | Last sign-in per user (`signInActivity`: `lastSignInDateTime`, `lastNonInteractiveSignInDateTime`, `lastSuccessfulSignInDateTime`) | `GET /users?$select=signInActivity`, as specified in source 5 of [identity posture](identity-posture.md) and not re-specified here. Page size drops to 500 when selected; a blank or `0001-01-01T00:00:00Z` value is not a sign-in | `AuditLog.Read.All` plus `User.Read.All`; signed in: Reports Reader ([inactive accounts](https://learn.microsoft.com/entra/identity/monitoring-health/howto-manage-inactive-user-accounts)) | Microsoft Entra ID P1 or P2 ([user](https://learn.microsoft.com/graph/api/resources/user)) | State | n/a | [Available](https://learn.microsoft.com/graph/api/resources/user) | [Available](https://learn.microsoft.com/graph/deployments) (global service) | UNVERIFIED (list users is marked for US Government L4, no page states the property itself in that cloud; open question BRO-243) |
| 5a | Active users per workload: licence flags and last activity per service (Exchange, OneDrive, SharePoint, Teams, Yammer, Skype for Business) | [`GET /reports/getOffice365ActiveUserDetail(period='D30')`](https://learn.microsoft.com/graph/api/reportroot-getoffice365activeuserdetail) (`Get-MgReportOffice365ActiveUserDetail`). CSV via 302 redirect. Columns include `Has <service> License`, `<service> Last Activity Date`, `<service> License Assign Date` and `Assigned Products` | `Reports.Read.All`. Signed in: an Entra limited admin role such as Reports Reader ([authorization](https://learn.microsoft.com/graph/reportroot-authorization)) | None named on the page | State | `period` D7, D30, D90 or D180; `date` form only the past 30 days | [Available](https://learn.microsoft.com/graph/api/reportroot-getoffice365activeuserdetail) | [Available](https://learn.microsoft.com/graph/deployments) (global service) | [NotAvailable](https://learn.microsoft.com/graph/api/reportroot-getoffice365activeuserdetail) (US Government L4 `❌`) |
| 5b | Email activity per user (send, receive, read counts and last activity) | [`GET /reports/getEmailActivityUserDetail(period='D30')`](https://learn.microsoft.com/graph/api/reportroot-getemailactivityuserdetail) | `Reports.Read.All`, same roles as 5a | None named on the page | State | D7, D30, D90, D180; `date` form only the past 28 days | [Available](https://learn.microsoft.com/graph/api/reportroot-getemailactivityuserdetail) | [Available](https://learn.microsoft.com/graph/deployments) (global service) | [NotAvailable](https://learn.microsoft.com/graph/api/reportroot-getemailactivityuserdetail) (US Government L4 `❌`) |
| 5c | Teams user activity (chat, call and meeting counts, last activity, `Is Licensed`) | [`GET /reports/getTeamsUserActivityUserDetail(period='D30')`](https://learn.microsoft.com/graph/api/reportroot-getteamsuseractivityuserdetail) | `Reports.Read.All`, same roles as 5a | None named on the page | State | D7, D30, D90, D180; `date` form only the past 30 days | [Available](https://learn.microsoft.com/graph/api/reportroot-getteamsuseractivityuserdetail) | [Available](https://learn.microsoft.com/graph/deployments) (global service) | [NotAvailable](https://learn.microsoft.com/graph/api/reportroot-getteamsuseractivityuserdetail) (US Government L4 `❌`) |
| 5d | SharePoint activity per user (files viewed or edited, synced, shared internally and externally, pages visited, last activity) | [`GET /reports/getSharePointActivityUserDetail(period='D30')`](https://learn.microsoft.com/graph/api/reportroot-getsharepointactivityuserdetail) | `Reports.Read.All`, same roles as 5a | None named on the page | State | D7, D30, D90, D180; `date` form only the past 30 days | [Available](https://learn.microsoft.com/graph/api/reportroot-getsharepointactivityuserdetail) | [Available](https://learn.microsoft.com/graph/deployments) (global service) | [NotAvailable](https://learn.microsoft.com/graph/api/reportroot-getsharepointactivityuserdetail) (US Government L4 `❌`) |
| 5e | OneDrive activity per user (same file counts and last activity) | [`GET /reports/getOneDriveActivityUserDetail(period='D30')`](https://learn.microsoft.com/graph/api/reportroot-getonedriveactivityuserdetail) | `Reports.Read.All`, same roles as 5a | None named on the page | State | D7, D30, D90, D180; `date` form only the past 30 days | [Available](https://learn.microsoft.com/graph/api/reportroot-getonedriveactivityuserdetail) | [Available](https://learn.microsoft.com/graph/deployments) (global service) | [NotAvailable](https://learn.microsoft.com/graph/api/reportroot-getonedriveactivityuserdetail) (US Government L4 `❌`) |
| 5f | Microsoft 365 apps used per user (Windows, Mac, mobile, web; Outlook, Word, Excel, PowerPoint, OneNote, Teams; last activation and activity dates) | [`GET /reports/getM365AppUserDetail(period='D30')`](https://learn.microsoft.com/graph/api/reportroot-getm365appuserdetail). CSV or JSON with `$format`; JSON pages 200 rows and carries `@odata.nextLink` | `Reports.Read.All`, same roles as 5a | None named on the page | State | D7, D30, D90, D180; `date` form only the past 30 days | [Available](https://learn.microsoft.com/graph/api/reportroot-getm365appuserdetail) | [Available](https://learn.microsoft.com/graph/deployments) (global service) | [NotAvailable](https://learn.microsoft.com/graph/api/reportroot-getm365appuserdetail) (US Government L4 `❌`) |
| 5g | Microsoft 365 Copilot last activity per user and app (Teams, Word, Excel, PowerPoint, Outlook, OneNote, Loop, Copilot Chat) | The beta [`/reports`](https://learn.microsoft.com/graph/api/reportroot-getmicrosoft365copilotusageuserdetail?view=graph-rest-beta) page says to use [`/copilot`](https://learn.microsoft.com/microsoft-365-copilot/extensibility/api/admin-settings/reports/copilotreportroot-getmicrosoft365copilotusageuserdetail) going forward. v1.0 call: `GET /copilot/reports/getMicrosoft365CopilotUsageUserDetail(period='D7', version='v2')`. `version` defaults to `v2`, and `v2` is the supported version value. v2 periods are `D7`, `D28`, `D90`, `D180`, `ALL`. Do not send `D30` on that default; `D30` is only in the v1 period list, and v1 is not a supported `version` value. `ALL` is the four periods in one response, not lifetime. v1.0 returns `200 OK` with the CSV in the body, not a redirect. Beta `/copilot` returns JSON. Beta `/reports` defaults to `application/json` and its example carries `@odata.nextLink` with `$skiptoken`; follow that link, because one page is not the full set. v2 adds prompts submitted, active usage days, Copilot Chat (work), Copilot Chat (web), Microsoft 365 Copilot, `Edge Last Activity Date` and Copilot Agent, on top of Teams, Word, Excel, PowerPoint, Outlook, OneNote, Loop and Copilot Chat. Beta is not supported for production applications. Returns only users with a Microsoft 365 Copilot licence. Unlicensed Copilot Chat is not in this API (admin center Copilot Chat Usage, Purview audit, or `Search-UnifiedAuditLog`) | `Reports.Read.All`. Signed in on the `/copilot` page: Company Administrator, AI Administrator, Exchange Administrator, SharePoint Administrator, Lync Administrator, Teams Service Administrator, Teams Communications Administrator or Reports Reader. That page does not list Global Reader or Usage Summary Reports Reader. The beta `/reports` page lists both, and [authorization](https://learn.microsoft.com/graph/reportroot-authorization) says they have no visibility into detailed metrics | A Microsoft 365 Copilot licence is the population; unlicensed Copilot Chat use is not in this API ([page](https://learn.microsoft.com/microsoft-365-copilot/extensibility/api/admin-settings/reports/copilotreportroot-getmicrosoft365copilotusageuserdetail)) | State | `version` defaults to `v2`: `D7`, `D28`, `D90`, `D180`, `ALL` (not `D30`) | [Available](https://learn.microsoft.com/graph/api/reportroot-getmicrosoft365copilotusageuserdetail?view=graph-rest-beta) | [Available](https://learn.microsoft.com/graph/deployments) (global service) | [NotAvailable](https://learn.microsoft.com/microsoft-365-copilot/extensibility/api/admin-settings/reports/copilotreportroot-getmicrosoft365copilotusageuserdetail) (US Government L4 `❌` on both the `/reports` and `/copilot` pages) |
| 6 | Whether usage reports show names or concealed identifiers: `displayConcealedNames` | [`GET /admin/reportSettings`](https://learn.microsoft.com/graph/api/adminreportsettings-get) (`Get-MgAdminReportSetting`). `true` means reports conceal "usernames, groups, and sites"; the property represents a setting in the Microsoft 365 admin center ([adminReportSettings](https://learn.microsoft.com/graph/api/resources/adminreportsettings)). The admin center menu path is not stated on the pages read | `ReportSettings.Read.All`. Signed in: an Entra limited admin role ([authorization](https://learn.microsoft.com/graph/reportroot-authorization)). Changing it needs `ReportSettings.ReadWrite.All` and is out of scope: the library is read-only | None named on the page | State | n/a | [Available](https://learn.microsoft.com/graph/api/adminreportsettings-get) | [Available](https://learn.microsoft.com/graph/deployments) (global service) | [NotAvailable](https://learn.microsoft.com/graph/api/adminreportsettings-get) (US Government L4 `❌`) |
| 7 | Product names for SKU ids and service plan ids | Not a tenant call. [Product names and service plan identifiers for licensing](https://learn.microsoft.com/entra/identity/users/licensing-service-plan-reference) lists `skuPartNumber`, GUID and the service plans in each product, with a CSV download. It says it is accurate only as of its last update, which it gives as 2026-08-19 | None; a public page | None | State (reference file; a build step must ship a dated copy) | n/a | n/a (public reference) | n/a (public reference) | n/a (public reference) |
| 8 | Fallback for GCC High: the same usage reports in the Microsoft 365 admin center | Admin center **Reports > Usage**, not a Graph call. The [usage reports overview](https://learn.microsoft.com/microsoft-365/admin/activity-reports/activity-reports) marks `Yes` for GCC-High on Active Users, Email activity, OneDrive user activity, SharePoint activity, Microsoft Teams user activity and Microsoft Copilot usage, and `N/A` (not yet released) for Microsoft 365 Apps usage. Whether the data can be exported or read by a script in GCC High is not stated on the pages read | Global Administrator, Exchange, SharePoint, Teams, Reports Reader and others per the overview; Usage Summary Reports Reader sees no user details | None named | State | 7, 30, 90 and 180 days | [Available](https://learn.microsoft.com/microsoft-365/admin/activity-reports/activity-reports) | [Available](https://learn.microsoft.com/microsoft-365/admin/activity-reports/activity-reports) | Partly: [Available](https://learn.microsoft.com/microsoft-365/admin/activity-reports/activity-reports) in the admin center for the reports listed; scripted read UNVERIFIED |

Consolidated notes:

* **GCC High usage data.** Sources 5a to 5g and 6 return `NotAvailable` through Graph. Under the library rule
  ("a source not available in a cloud is skipped, and the README says why"), a GCC High run writes sources 1 to 4
  and skips 5a to 5g and 6. Source 8 shows the data exists in the admin center; turning that into a collector is
  an open item, not a claim.
* **Overlapping licences (starting question 4) has no source of its own.** No Microsoft page found offers an
  overlap report. It is derived: take each user's `assignedLicenses` or `licenseAssignmentStates` (source 2), map
  `skuId` to `servicePlans` from source 1, and flag a `servicePlanId` that appears under two SKUs for the same
  user. Source 3 returns the same plans per user and is the fallback. The derivation is the library's logic, not
  a Microsoft-documented result.
* **Free licences.** `consumedUnits` is "the number of licenses that have been assigned". Free units are
  derived (`prepaidUnits.enabled` minus `consumedUnits`). Suspended, warning and locked-out units are separate
  fields; do not fold them into free.
* **Direct versus group-assigned.** `assignedLicenses` does not distinguish; `licenseAssignmentStates` does
  (`assignedByGroup`). A user with a licence assignment in `Error` or `ActiveWithError` holds a licence that is
  not working; show it apart from clean assignments.
* **Concealed names.** When `displayConcealedNames` is `true`, the usage reports (5a to 5g) hide usernames, so
  they cannot be joined to source 2 on `userPrincipalName`. Source 6 must be read first, and the report must show
  "names concealed, usage not joined" instead of zero-usage users.
* **Usage reports are periods, not events.** The count columns aggregate 7, 30, 90 or 180 days. A daily run
  appends a snapshot stamped with the run date; it is not an event stream. The last activity date is the most
  recent intentional activity in that app, regardless of the selected time period, so a snapshot can show
  inactivity older than 180 days ([activity reports](https://learn.microsoft.com/microsoft-365/admin/activity-reports/activity-reports)).
  Compare that date with the period start. Reports typically become available within 24 to 72 hours and sometimes
  take several days, so a gap of a few days is not proof of no use. Usage reports don't include perpetual license
  models. When a user account is deleted, that user's usage data is removed within 30 days and they leave the
  user-detail table; a missing row is not a zero-activity row.
* **Download mechanics.** The CSV reports for sources 5a to 5f return `302 Found` to a pre-authenticated URL that
  "is only valid for a short period of time (a few minutes)". The collector must follow the redirect at once.
  Source 5g v1.0 returns `200 OK` with the CSV in the body, not a redirect.
* **Source 4 depends on an open question.** `signInActivity` in GCC High is UNVERIFIED (BRO-243). If the property
  is refused, "not signed in" cannot be answered for GCC High from this source.

## Proposed report pages

| # | Page | Reads | What it shows |
| --- | --- | --- | --- |
| 1 | Overview | 1, 2, 4, 5a, 5g | Headline counts: licences bought, assigned, free, assigned to users who never signed in, assigned with last activity before the period start; trend by snapshot date |
| 2 | SKU inventory | 1, 7 | Per SKU: bought, assigned, free, suspended, warning, `capabilityStatus`; friendly product name from the reference copy |
| 3 | Assignments by organisation | 2, shared users | Who holds each SKU by department, job title, office location, city and country; direct versus group-assigned |
| 4 | Assigned but inactive | 2, 4 | Licensed users with no sign-in in 30, 90 or 180 days, or no sign-in recorded; GCC High marked UNVERIFIED |
| 5 | Workload usage | 2, 5a to 5f, 6 | Per workload (Exchange, Teams, SharePoint, OneDrive, Microsoft 365 apps): licensed users whose last activity date is before the period start, or who have no date. Count columns are the period aggregate. Hidden when names are concealed; absent in GCC High |
| 6 | Copilot adoption | 1, 2, 5g | Microsoft 365 Copilot licences assigned (service plan found through source 1) against users with a last activity date per Copilot app; absent in GCC High |
| 7 | Overlapping licences | 1, 2, 3 | Users holding two SKUs that include the same service plan, with the SKU pair; derived, see notes |
| 8 | Coverage and licence errors | 2, 6, `run.log` | Assignments in `Error` or `ActiveWithError` with the `error` value; which sources ran, which were skipped in this cloud, and whether names were concealed |

## Starting questions

| # | Question | Status |
| --- | --- | --- |
| 1 | Which licences the tenant owns, how many are assigned and how many are free | Covered by source 1. Free is derived from `prepaidUnits.enabled` and `consumedUnits`. Available in all three clouds. |
| 2 | Which users hold each licence, by department, job title and location | Covered by source 2 and the shared users collector. Location is `officeLocation`, `city`, `country` or `usageLocation`; the page does not say which one the report should treat as "location", so the build issue has to choose. Available in all three clouds. |
| 3 | Which assigned users have not signed in, or have not used the licensed workloads, over a period | Sign-in: source 4 (UNVERIFIED in GCC High). Workloads: sources 5a to 5g for Exchange (5a, 5b), Teams (5a, 5c), SharePoint (5a, 5d), OneDrive (5a, 5e), Microsoft 365 apps (5f) and Copilot (5g). **Dropped for GCC High through Graph:** every usage source is `NotAvailable` there (source 8 is the manual route). Period for sources 5a to 5f is 7, 30, 90 or 180 days. Copilot `version` defaults to `v2`, whose periods use 28 days instead of 30. |
| 4 | Which users hold overlapping licences | Derived from sources 1 and 2 (source 3 as fallback). Not a Microsoft-provided report; see the note above. |
| 5 | Whether usage reports return real names or concealed identifiers, and the tenant setting that controls it | Covered by source 6 (`displayConcealedNames`). `NotAvailable` in GCC High through Graph; the admin center setting itself is not confirmed for GCC High on a page read, so it is UNVERIFIED there. |

No starting question is dropped outright. Question 3's usage half is dropped for GCC High, for the reason above.

## Overlap with other candidates

* Resolve: the project catalog lists "license utilization (overlaps Resolve)"; nothing here was derived from
  Resolve code, which was not read, so the build issue should compare the two before building.
* Identity posture ([identity-posture.md](identity-posture.md)): shares the users collector and source 4
  (`signInActivity`). It does not rebuild MFA, Conditional Access, role or risky-user content.

## Open items for the build issue

* GCC High usage data: decide between skipping 5a to 5g (the library rule) and a documented manual export from
  source 8. No page read says the admin center reports can be exported by script in GCC High.
* Source 5g: call `/copilot` v1.0. `version` defaults to `v2` (the supported version value); do not send `D30`.
  Beta `/reports` is not supported for production applications.
* Source 3 lists Application as not supported. If the report ever runs unattended with an app token, use the
  source 1 and 2 join instead.
* Source 7 is a dated copy of a public page. Decide where the copy lives and how it is refreshed.
