# Entra sign-in and audit activity

Who signed in to Entra ID, from where, with what result and under which Conditional Access
policies, and who changed users, groups, roles, applications and policies. Built from the
source contract in [`docs/candidates/entra-activity.md`](../../docs/candidates/entra-activity.md)
(BRO-383); the Power BI project is a separate, later issue and is not here.

The collectors write one CSV per source into an output folder of your choosing, plus a
`run.log`. Event files are appended from the last exported timestamp; the snapshot file gets
a block of rows stamped with the run date. Nothing is rewritten or deleted. `samples/`
holds fake data (`example.com`, documentation IP addresses) with the exact columns each
collector writes.

Nothing here has been run against a tenant. Every tenant call is mocked in `tests/`.

## Contents

| Path | What it is |
| --- | --- |
| `collectors/Get-InteractiveSignIns.ps1` | `signins-interactive.csv` — source 1, v1.0 sign-in log |
| `collectors/Get-NonInteractiveSignIns.ps1` | `signins-noninteractive.csv` — source 2, **beta** sign-in log filtered to `nonInteractiveUser` |
| `collectors/Get-SignInConditionalAccess.ps1` | `signin-conditional-access.csv` — source 3, the Conditional Access result on each sign-in, one row per applied policy |
| `collectors/Get-DirectoryAudits.ps1` | `directory-audits.csv` — source 4, directory audit events |
| `collectors/Get-RetentionReference.ps1` | `retention-reference.csv` — source 5, log retention by licence level (no tenant call) |
| `collectors/Run-All.ps1` | Runs the shared users collector and all five of the above |
| `collectors/EntraActivitySchema.psd1` | The column order of every CSV, and the availability of every source per cloud |
| `collectors/EntraActivityHelpers.ps1` | The report-specific helper: the window loop, availability and sign-in row shaping |
| `New-SampleData.ps1` | Regenerates `samples/` |
| `tests/` | Pester tests; every tenant call is mocked |

`users.csv` comes from `Invoke-EntraUserCollector` in the shared module, not from this
folder. The report joins sign-ins to user attributes through it on `UserId`; the identity
posture collectors (BRO-321) are not rebuilt here. Sources 1, 2 and 3 read the same sign-in
endpoint as identity posture's legacy sign-in collector but keep the full event stream; this
report ships its own collectors rather than sharing one.

## Before you run it

```powershell
Install-Module Microsoft.Graph.Authentication, Microsoft.Graph.Users -Scope CurrentUser
```

Non-interactive sign-ins call `GET /beta/auditLogs/signIns` on the signed-in Graph
host. That does not load `Microsoft.Graph.Beta.Reports`.

### Role and licence each collector needs

| Collector | Graph permission | Least privileged role | Licence |
| --- | --- | --- | --- |
| `Get-InteractiveSignIns.ps1` | `AuditLog.Read.All` | Reports Reader ([troubleshooting page](https://learn.microsoft.com/entra/identity/monitoring-health/howto-troubleshoot-sign-in-errors)); Global Reader, Security Administrator, Security Operator or Security Reader also work ([list signIns](https://learn.microsoft.com/graph/api/signin-list#permissions)) | Entra ID P1 or P2. The [signIn resource page](https://learn.microsoft.com/graph/api/resources/signin) says P1 or P2 is needed to download sign-in logs through the API; the [licensing page](https://learn.microsoft.com/entra/fundamentals/licensing#microsoft-entra-monitoring-and-health) lists them as available on Free. The pages disagree, so a refusal is logged as "not licensed". `riskDetail` and the risk levels read `hidden` without P2 |
| `Get-NonInteractiveSignIns.ps1` | `AuditLog.Read.All` | As above | As above |
| `Get-SignInConditionalAccess.ps1` | `AuditLog.Read.All` **and** `Policy.Read.All` (or `Policy.Read.ConditionalAccess`) for the per-policy detail | As above, plus Conditional Access Administrator, Global Reader, Security Administrator or Security Reader for signed-in use | Sign-in log licence as above; Conditional Access itself needs [Entra ID P1](https://learn.microsoft.com/entra/fundamentals/licensing#microsoft-entra-conditional-access) |
| `Get-DirectoryAudits.ps1` | `AuditLog.Read.All` (`Directory.Read.All` is the higher privileged alternative) | Reports Reader, Security Administrator or Security Reader ([list directoryAudits](https://learn.microsoft.com/graph/api/directoryaudit-list#permissions)) | None named on the list page; the licensing page lists audit logs as available on Free and P1 or P2 |
| `Get-RetentionReference.ps1` | None | None | None; a public page |

An interactive sign-in asks for the shared read scopes, plus `Policy.Read.All` for
`Get-SignInConditionalAccess.ps1`. App-only needs the permissions granted as application
permissions with admin consent. `Policy.ReadWrite.ConditionalAccess` also unlocks the
policy detail but is a write permission, so it is not requested.

## Cloud availability

Every collector takes `-Environment Commercial|GCC|GCCHigh`, defaulting to `Commercial`.
Graph is `https://graph.microsoft.com` for Commercial and GCC and `Connect-MgGraph
-Environment USGov` for GCC High ([national cloud deployments](https://learn.microsoft.com/graph/deployments)).
Copied from the contract; GCC calls the global service.

| # | Source | Commercial | GCC | GCC High |
| --- | --- | --- | --- | --- |
| 1 | Interactive sign-ins | [Available](https://learn.microsoft.com/graph/api/signin-list) | [Available](https://learn.microsoft.com/graph/deployments) | [Available](https://learn.microsoft.com/graph/api/signin-list) (US Government L4) |
| 2 | Non-interactive sign-ins (beta) | [Available](https://learn.microsoft.com/graph/api/signin-list?view=graph-rest-beta) | [Available](https://learn.microsoft.com/graph/deployments) | [Available](https://learn.microsoft.com/graph/api/signin-list?view=graph-rest-beta) (US Government L4) |
| 3 | Conditional Access result on each sign-in | [Available](https://learn.microsoft.com/graph/api/signin-list) | [Available](https://learn.microsoft.com/graph/deployments) | [Available](https://learn.microsoft.com/graph/api/signin-list) (documented on the sign-in list page; the `appliedConditionalAccessPolicy` resource page has no cloud table) |
| 4 | Directory audit events | [Available](https://learn.microsoft.com/graph/api/directoryaudit-list) | [Available](https://learn.microsoft.com/graph/deployments) | [Available](https://learn.microsoft.com/graph/api/directoryaudit-list) (US Government L4) |
| 5 | Log retention by licence | UNVERIFIED | UNVERIFIED | UNVERIFIED |

Source 5's [retention page](https://learn.microsoft.com/entra/identity/monitoring-health/reference-reports-data-retention)
names no cloud, so retention is UNVERIFIED in all three. That also bounds sources 1 to 4: a
GCC or GCC High run should not promise more history than it can read. No source is
`NotAvailable` in any cloud, so there is no header-only sample; the `NotAvailable` path is
tested against a modified copy of the schema. An `UNVERIFIED` source is attempted and logs a
warning.

`signInActivity` (last sign-in per user) is **not** a source here, so its unsettled GCC High
availability ([BRO-243](https://linear.app/broekncode/issue/BRO-243/settle-whether-signinactivity-is-available-in-gcc-high))
does not affect this report. If a page later reads it, treat GCC High as `UNVERIFIED`.

## Limits to know about

* **Not complete history.** Sign-in and audit logs hold 7 days on Free and 30 on P1 and P2.
  A first run reads at most that window; a daily run then builds more history than the
  source holds. After an upgrade from Free only the data still in the 7-day window is kept.
  If the free licence had no activity data, it can take up to three days to show after the
  upgrade. Purview Audit (Premium) retains Entra audit logs only and does not extend
  sign-in retention. Routing both logs to an Azure Monitor storage account is what lengthens
  them
  ([data retention](https://learn.microsoft.com/entra/identity/monitoring-health/reference-reports-data-retention)).
* **Paging.** A sign-in page holds at most 1,000 rows (the maximum and the default), so the
  request does not send `$top`. The list pages do not take `$skip` or `$select`. The audit
  page size is not stated. Every page follows the full `@odata.nextLink` until it is absent
  ([paging](https://learn.microsoft.com/graph/paging)). A repeated nextLink, or a nextLink
  that is not on `graph.microsoft.com` or `graph.microsoft.us`, fails the window. Each run
  is walked in `-WindowHours` windows (default 24) filtered on `createdDateTime` or
  `activityDateTime`, as the list pages advise. Every page sends
  `Prefer: include-unknown-enum-members`, including each nextLink page
  ([SDK paging](https://learn.microsoft.com/graph/sdks/paging) does not forward that header).
* **A 429 is a throttle, not an empty log.** Identity and access reports allow 122 requests
  per 10 seconds per app and five per 10 seconds per app per tenant
  ([limits](https://learn.microsoft.com/graph/throttling-limits#identity-and-access-reports-service-limits)).
  The collector waits the `Retry-After` seconds and retries that same request
  ([throttling](https://learn.microsoft.com/graph/throttling)). It does not follow a
  nextLink from the error, and a `DirectoryPageTokenNotFoundException` is not retried with
  a different link. If the 429 continues, the window is halved down to one hour
  (the reports guidance starts at three days and then shortens the span). A window that
  still fails is not written; earlier windows are kept and the next run resumes at the
  watermark. HTTP 401 and 403 are not retried.
* **Source 2 is beta**, which Microsoft does not support for production applications. Ship
  `signins-interactive.csv` alone and say the report covers interactive sign-ins only if
  that is not acceptable. The filter is `nonInteractiveUser`; `ne 'interactiveUser'` is not
  used because it also returns service principal and managed identity sign-ins.
* **Conditional Access detail depends on a permission.** With `AuditLog.Read.All` alone
  `conditionalAccessStatus` is returned but `appliedConditionalAccessPolicies` is dropped
  without an error. `PolicyDetailReadable` in `signin-conditional-access.csv` says whether the
  run held a Conditional Access read permission, so an empty policy list is not read as "no
  policy applied". Report-only results (`reportOnlySuccess`, `reportOnlyFailure`,
  `reportOnlyNotApplied`, `reportOnlyInterrupted`) are returned only when the request sends
  `Prefer: include-unknown-enum-members`
  ([appliedConditionalAccessPolicy](https://learn.microsoft.com/graph/api/resources/appliedconditionalaccesspolicy)).
  Every sign-in page sends that header. `unknownFutureValue` is kept when the service sends it.
* **Error codes are kept as returned.** `ErrorCode` `0` is success; `1024` and other codes
  that are not 5 or 6 digits are kept. A `50058` may have no user, so a row with an empty
  `UserId` is a sign-in, not a gap. `UserPrincipalName` is lowercase and a guest is the home
  UPN, not the `#EXT#` form; group by `UserId`.
* **Directory audit values are not a closed set.** Every `Category` (Conditional Access
  changes are `Policy`), `ActivityDisplayName`, `OperationType` and target `type` is kept as
  returned, and ids that are not GUIDs are kept.
* **The unified audit log is a different source** and is not read here.

## The CSVs

Timestamps are UTC. List columns are joined with `;`. All files are keyed so a re-run never
repeats a row.

### `signins-interactive.csv` and `signins-noninteractive.csv`

Same columns; key `Id`.

| Column | Meaning |
| --- | --- |
| `CreatedDateTime` | UTC time of the sign-in |
| `Id` | Sign-in id |
| `UserId`, `UserPrincipalName` | Directory id and lowercase UPN (empty for some failures such as 50058) |
| `AppId`, `AppDisplayName`, `ResourceDisplayName` | Client application and the resource it asked for |
| `IpAddress`, `City`, `State`, `CountryOrRegion` | Where the sign-in came from |
| `ClientAppUsed` | Client type, such as `Browser` |
| `DeviceOperatingSystem`, `DeviceBrowser`, `DeviceIsCompliant`, `DeviceIsManaged` | From `deviceDetail` |
| `IsInteractive` | Whether the sign-in was interactive |
| `SignInEventTypes` | Event types on the sign-in, such as `interactiveUser` or `nonInteractiveUser` |
| `ErrorCode`, `FailureReason`, `AdditionalDetails` | `status`: `0` is success; the text is Microsoft's, not translated |
| `ConditionalAccessStatus` | `success`, `failure`, `notApplied` or `unknownFutureValue` |
| `RiskDetail`, `RiskLevelAggregated`, `RiskLevelDuringSignIn`, `RiskState` | Risk fields; `hidden` without P2 |

### `signin-conditional-access.csv`

Key `SignInId`, `PolicyId`, `PolicyDisplayName`.

| Column | Meaning |
| --- | --- |
| `CreatedDateTime`, `SignInId`, `UserId`, `IsInteractive` | The sign-in the row belongs to |
| `ConditionalAccessStatus` | The sign-in's overall result |
| `PolicyId`, `PolicyDisplayName`, `PolicyResult` | An applied policy and its result; empty when the sign-in carried no policy detail |
| `EnforcedGrantControls`, `EnforcedSessionControls` | What the policy enforced |
| `PolicyDetailReadable` | Whether the run held a Conditional Access read permission |

### `directory-audits.csv`

Key `Id`.

| Column | Meaning |
| --- | --- |
| `ActivityDateTime` | UTC time of the event |
| `Id`, `CorrelationId` | Event id (not always a GUID) and correlation id |
| `ActivityDisplayName`, `Category`, `OperationType`, `LoggedByService` | What happened and which service logged it |
| `Result`, `ResultReason` | `success`, `failure`, `timeout` or `unknownFutureValue`, and the reason for a failure or timeout |
| `InitiatedByUserId`, `InitiatedByUserPrincipalName`, `InitiatedByAppId`, `InitiatedByAppDisplayName` | Who did it: a user or an app |
| `TargetResourceTypes`, `TargetResourceIds`, `TargetResourceDisplayNames`, `TargetResourceUserPrincipalNames` | The targets, in the same order |
| `ModifiedProperties` | Targets' `modifiedProperties` as compact JSON |

### `retention-reference.csv`

Key `RunDate`, `Environment`, `LicenseLevel`.

| Column | Meaning |
| --- | --- |
| `RunDate` | UTC date of the run |
| `Environment` | The `-Environment` of the run |
| `LicenseLevel` | `Free`, `P1` or `P2` |
| `SignInRetentionDays`, `AuditRetentionDays`, `RiskySignInRetentionDays` | Documented windows (risky sign-ins are not collected) |
| `RetentionStatus` | `UNVERIFIED` in every cloud: the Microsoft page names no cloud |
| `Reference` | The Learn page |
