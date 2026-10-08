# Candidate: identity posture

Status: sources verified against Microsoft Learn; nothing built. Written for BRO-315 so the build issue can
be written without guessing. Nothing here was run against a tenant.

Cloud names follow the library: `Commercial`, `GCC`, `GCCHigh` (`-Environment`). Graph uses
`https://graph.microsoft.com` for Commercial and GCC and `Connect-MgGraph -Environment USGov` for GCC High
([national cloud deployments](https://learn.microsoft.com/graph/deployments)). Every Graph page cited below
carries a national-cloud table; "US Government L4" is the GCC High column, and GCC calls the global endpoint.

## How to read the table

* `Available` / `NotAvailable` appear only where a Microsoft page says so; the link is in the cell.
* `UNVERIFIED` means no Microsoft page found says either way. The collector should ask for the data and
  record a refusal in `run.log`, as the existing reports do.
* Source 1 reuses `Invoke-EntraUserCollector` in `shared/M365ReportLibrary.psm1`; it is not re-specified.
  That collector writes `users.csv` (`Id`, `DisplayName`, `UserPrincipalName`, `Mail`, `UserType`,
  `AccountEnabled`, `CreatedDateTime`, `Department`, `JobTitle`, `City`, `Country`, `ManagerId`,
  `ManagerUserPrincipalName`) and does **not** collect `signInActivity`, so stale and never-signed-in
  accounts need source 5.
* All Graph sources below are read through the `Microsoft.Graph` PowerShell SDK or REST with the permissions
  named; the library already requests `AuditLog.Read.All` for the shared collector.

## Sources

| # | Source | Endpoint or cmdlet | Least privileged role or permission | License | Event / state | Retention | Commercial | GCC | GCC High |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 1 | Users (shared collector) | `Get-MgUser` via `Invoke-EntraUserCollector` | Already specified in the shared module | Already specified | State | n/a | [Available](https://learn.microsoft.com/graph/deployments) | [Available](https://learn.microsoft.com/graph/deployments) | [Available](https://learn.microsoft.com/graph/deployments) |
| 2 | Authentication method registration per user (methods registered, MFA registered and capable, passwordless capable, SSPR, `isAdmin`) | [`GET /reports/authenticationMethods/userRegistrationDetails`](https://learn.microsoft.com/graph/api/authenticationmethodsroot-list-userregistrationdetails) (`Get-MgReportAuthenticationMethodUserRegistrationDetail`). Does not return disabled users | Application and delegated `AuditLog.Read.All`. Signed in: Reports Reader, Security Reader, Security Administrator or Global Reader | Microsoft Entra ID P1 or P2: a call without it returns "Neither tenant is B2C or tenant doesn't have premium license" ([troubleshooting](https://learn.microsoft.com/troubleshoot/entra/entra-id/users-groups-entra-apis/b2c-or-tenant-premium-license-sign-in-activities)). That page also lists `Directory.Read.All` for tenant license lookups | State | n/a | [Available](https://learn.microsoft.com/graph/api/authenticationmethodsroot-list-userregistrationdetails) | [Available](https://learn.microsoft.com/graph/deployments) (global service) | [Available](https://learn.microsoft.com/graph/api/authenticationmethodsroot-list-userregistrationdetails) (page lists US Government L4) |
| 3 | Conditional Access policies (state, include and exclude users, groups, roles and apps, grant controls) | [`GET /identity/conditionalAccess/policies`](https://learn.microsoft.com/graph/api/conditionalaccessroot-list-policies) (`Get-MgIdentityConditionalAccessPolicy`) | Application and delegated `Policy.Read.All`. Signed in: Security Reader, Global Reader, Security Administrator or Conditional Access Administrator | None named on the page | State | n/a | [Available](https://learn.microsoft.com/graph/api/conditionalaccessroot-list-policies) | [Available](https://learn.microsoft.com/graph/deployments) (global service) | [Available](https://learn.microsoft.com/graph/api/conditionalaccessroot-list-policies) (US Government L4) |
| 4a | Active privileged role assignments | [`GET /roleManagement/directory/roleAssignmentScheduleInstances`](https://learn.microsoft.com/graph/api/rbacapplication-list-roleassignmentscheduleinstances) | `RoleAssignmentSchedule.Read.Directory`. Read roles: Global Reader, Security Operator, Security Reader, Security Administrator or Privileged Role Administrator | None named on the page; PIM license requirement not found on this page. UNVERIFIED | State | n/a | [Available](https://learn.microsoft.com/graph/api/rbacapplication-list-roleassignmentscheduleinstances) | [Available](https://learn.microsoft.com/graph/deployments) (global service) | [Available](https://learn.microsoft.com/graph/api/rbacapplication-list-roleassignmentscheduleinstances) (US Government L4) |
| 4b | Eligible privileged role assignments | [`GET /roleManagement/directory/roleEligibilityScheduleInstances`](https://learn.microsoft.com/graph/api/rbacapplication-list-roleeligibilityscheduleinstances) | `RoleEligibilitySchedule.Read.Directory`. Same read roles as 4a | As 4a | State | n/a | [Available](https://learn.microsoft.com/graph/api/rbacapplication-list-roleeligibilityscheduleinstances) | [Available](https://learn.microsoft.com/graph/deployments) (global service) | [Available](https://learn.microsoft.com/graph/api/rbacapplication-list-roleeligibilityscheduleinstances) (US Government L4) |
| 5 | Last sign-in per user (`signInActivity`: `lastSignInDateTime`, `lastNonInteractiveSignInDateTime`, `lastSuccessfulSignInDateTime`) | [`GET /users?$select=signInActivity`](https://learn.microsoft.com/graph/api/user-list#example-11-get-users-including-their-last-sign-in-time) ([resource](https://learn.microsoft.com/graph/api/resources/signinactivity)) | `AuditLog.Read.All` (page note). Role: UNVERIFIED | Microsoft Entra ID P1 or P2 ([List users note](https://learn.microsoft.com/graph/api/user-list#example-11-get-users-including-their-last-sign-in-time)) | State (kept as long as the user object exists; `lastSuccessfulSignInDateTime` is not backfilled before 1 December 2023) ([resource](https://learn.microsoft.com/graph/api/resources/signinactivity)) | n/a | [Available](https://learn.microsoft.com/graph/api/user-list#example-11-get-users-including-their-last-sign-in-time) | UNVERIFIED (the page excerpt read does not show the national-cloud table for this property) | UNVERIFIED (same) |
| 6 | Sign-in events, including `clientAppUsed` for legacy authentication | [`GET /auditLogs/signIns`](https://learn.microsoft.com/graph/api/signin-list) (`Get-MgAuditLogSignIn`); `clientAppUsed` values `Exchange ActiveSync`, `IMAP`, `MAPI`, `SMTP`, `POP`, `other clients` are legacy ([signIn resource](https://learn.microsoft.com/graph/api/resources/signin#properties)). Filter by `createdDateTime`; filter `clientAppUsed eq` is supported | `AuditLog.Read.All`. Signed in: Reports Reader (least privileged for sign-in logs per [least privileged roles](https://learn.microsoft.com/entra/identity/role-based-access-control/delegate-by-task#monitoring-and-health---audit-and-sign-in-logs-least-privileged-roles)), Global Reader, Security Reader, Security Operator or Security Administrator. Reading `appliedConditionalAccessPolicies` needs a Conditional Access read role or `Policy.Read.All` | Microsoft Entra ID P1 or P2 for the sign-in activity reports ([troubleshooting](https://learn.microsoft.com/troubleshoot/entra/entra-id/users-groups-entra-apis/b2c-or-tenant-premium-license-sign-in-activities)) | Event | Free: 7 days; P1 and P2: 30 days ([Entra data retention](https://learn.microsoft.com/entra/identity/monitoring-health/reference-reports-data-retention)). Returns interactive sign-ins and successful federated sign-ins; the page says only interactive-in-nature sign-ins are included ([List signIns](https://learn.microsoft.com/graph/api/signin-list)) | [Available](https://learn.microsoft.com/graph/api/signin-list) | [Available](https://learn.microsoft.com/graph/deployments) (global service) | [Available](https://learn.microsoft.com/graph/api/signin-list) (US Government L4) |
| 7 | Users Entra flags as risky (`riskLevel`, `riskState`, `riskLastUpdatedDateTime`) | [`GET /identityProtection/riskyUsers`](https://learn.microsoft.com/graph/api/riskyuser-list) (`Get-MgRiskyUser`) | `IdentityRiskyUser.Read.All`. Signed in: Global Reader, Security Operator, Security Reader or Security Administrator | **Microsoft Entra ID P2** ([riskyUser resource](https://learn.microsoft.com/graph/api/resources/riskyuser)) | State | Risky users: no limit, kept until the risk is remediated ([Entra data retention](https://learn.microsoft.com/entra/identity/monitoring-health/reference-reports-data-retention)) | [Available](https://learn.microsoft.com/graph/api/riskyuser-list) | [Available](https://learn.microsoft.com/graph/deployments) (global service) | [Available](https://learn.microsoft.com/graph/api/riskyuser-list) (US Government L4) |

Consolidated notes:

* Sources 2, 5, 6 and 7 need Entra ID P1 or P2 and fail without it. A report built on them must show "not
  licensed" instead of zero when the tenant lacks the license.
* Source 6 retention is 30 days at best. Appending from the last exported timestamp on a daily run builds a
  longer history than the source retains, as the Activity Explorer collector does.
* Directory audit logs (`GET /auditLogs/directoryAudits`) are an event source for changes to roles and
  policies ([audit logs overview](https://learn.microsoft.com/graph/api/resources/azure-ad-auditlog-overview));
  they were not in the starting questions and were not researched further.
* Source 3 is a snapshot, so a Conditional Access change shows up as a difference between two runs, with no
  record of who made it.

## Proposed report pages

| # | Page | Reads | What it shows |
| --- | --- | --- | --- |
| 1 | Overview | 1, 2, 3, 4a, 4b, 5, 7 | Headline counts: users without MFA, enabled policies, privileged accounts, stale accounts, risky users; trend by snapshot date |
| 2 | Authentication methods | 1, 2 | Methods registered per user, MFA registered vs capable, passwordless capable; users with no multifactor method, split member and guest |
| 3 | Conditional Access | 3 | Policies by state, targets (users, groups, roles, apps), grant controls; policies in report-only |
| 4 | Privileged roles | 1, 2, 4a, 4b | Active vs eligible holders per role; privileged accounts without MFA (`isAdmin` on source 2) |
| 5 | Account hygiene | 1, 5 | Disabled, stale (by `lastSuccessfulSignInDateTime`) and never-signed-in accounts |
| 6 | Legacy authentication | 6 | Legacy-client sign-ins over time by protocol and user; blocked vs succeeded |
| 7 | Risky users | 7 | Risk level and state, last updated; overlap with privileged roles |

## Starting questions

| # | Question | Status |
| --- | --- | --- |
| 1 | Which users have which authentication methods registered, and which have no multifactor method | Covered by source 2. Disabled users are not returned, so a "no MFA" count covers enabled users only; join to source 1 to list disabled users separately. |
| 2 | Which Conditional Access policies exist, their state, and what they target | Covered by source 3 |
| 3 | Who holds privileged directory roles, split active and eligible | Covered by sources 4a and 4b. "Privileged" needs a defined role list; Learn marks some built-in roles with a privileged label ([privileged roles](https://learn.microsoft.com/entra/identity/role-based-access-control/privileged-roles-permissions)), which was not read in full here. |
| 4 | Which accounts are disabled, stale or have never signed in | Disabled: source 1 (`AccountEnabled`). Stale and never signed in: source 5 (needs P1 or P2; GCC and GCC High availability UNVERIFIED). |
| 5 | Legacy authentication sign-ins over time | Covered by source 6, limited to 30 days of history per run and to the sign-ins that source returns (interactive) |
| 6 | Which users Entra flags as risky, and the license | Covered by source 7; requires Entra ID P2 |

## Open items for the build issue

* Confirm whether the `signInActivity` property is available in GCC and GCC High; if not, derive staleness
  from source 6 for those clouds.
* Confirm PIM licensing for sources 4a and 4b.
* Confirm the non-interactive sign-in gap in source 6 against the legacy-authentication question: legacy
  protocols often authenticate non-interactively, and the List signIns page says only interactive-in-nature
  sign-ins are included.
* Decide the privileged-role definition for page 4.
