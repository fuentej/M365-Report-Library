# Identity posture

Who is registered for multifactor authentication, which Conditional Access policies exist
and what state they are in, who holds directory roles (active and eligible), which accounts
have stale or no sign-in, which users Entra ID Protection flags as risky, and which sign-ins
used a legacy authentication client.

The collectors write one CSV per source into an output folder of your choosing, plus a
`run.log`. Snapshot files get a new block of rows tagged `RunDate` on every run; the sign-in
file is appended from where the last run stopped. Nothing is rewritten or deleted, so the
folder is a history you can chart over time. Every collector is read-only.

The sources were verified against Microsoft Learn in
[`docs/candidates/identity-posture.md`](../../docs/candidates/identity-posture.md), which is
the contract for this folder. The Power BI project is a separate, later piece of work.
Nothing here was run against a tenant: the tests mock every Graph call and the sample files
are fake.

## Contents

| Path | What it is |
| --- | --- |
| `collectors/Get-AuthenticationMethods.ps1` | `authentication-methods.csv` — methods registered, MFA and passwordless capability, per user |
| `collectors/Get-ConditionalAccessPolicies.ps1` | `conditional-access-policies.csv` — every policy, its state, targets and grant controls |
| `collectors/Get-ActiveRoleAssignments.ps1` | `role-assignments-active.csv` — active role assignments, including PIM activations |
| `collectors/Get-EligibleRoleAssignments.ps1` | `role-assignments-eligible.csv` — eligible (PIM) role assignments |
| `collectors/Get-RoleAssignments.ps1` | `role-assignments.csv` — active role assignments without PIM |
| `collectors/Get-UserSignInActivity.ps1` | `user-signin-activity.csv` — last sign-in times per user |
| `collectors/Get-RiskyUsers.ps1` | `risky-users.csv` — users Entra ID Protection flags as risky |
| `collectors/Get-LegacySignIns.ps1` | `signins.csv` — sign-ins that used a legacy authentication client |
| `collectors/Run-All.ps1` | Runs the shared users collector and all eight above |
| `collectors/IdentityPostureSchema.psd1` | The column order of every CSV, the per-cloud availability and the extra Graph scopes |
| `collectors/IdentityPostureHelpers.ps1` | Code the collectors share |
| `samples/` | Fake data in the exact shape the collectors write |
| `tests/` | Pester 5 tests; every tenant call is mocked |

`users.csv` comes from the shared Entra users collector (`Invoke-EntraUserCollector` in
`shared/M365ReportLibrary.psm1`); this report does not carry a second one.

## Before you run it

### Modules

PowerShell 7 and `Microsoft.Graph.Authentication`, `Microsoft.Graph.Users`,
`Microsoft.Graph.Reports`, `Microsoft.Graph.Identity.SignIns`,
`Microsoft.Graph.Identity.Governance` and `Microsoft.Graph.Identity.DirectoryManagement`.

### Licences

A source whose licence is missing is a logged skip: the collector writes the header only,
puts a warning in `run.log` and carries on. A header-only file means "not collected", never
"zero". See [`samples/unlicensed/`](samples/unlicensed).

| CSV | Licence |
| --- | --- |
| `authentication-methods.csv` | Microsoft Entra ID P1 or P2 |
| `conditional-access-policies.csv` | Microsoft Entra ID P1 (Microsoft 365 Business Premium also includes Conditional Access) |
| `role-assignments-active.csv`, `role-assignments-eligible.csv` | Microsoft Entra ID P2 or Microsoft Entra ID Governance. The PIM licence requirement is `UNVERIFIED`: the contract cites the PIM getting-started and licensing pages, not the API pages |
| `role-assignments.csv` | None named on the API page |
| `user-signin-activity.csv` | Microsoft Entra ID P1 or P2 |
| `risky-users.csv` | Microsoft Entra ID P2. P1 is not enough |
| `signins.csv` | Microsoft Entra ID P1 or P2. Retention is 7 days (free) or 30 days (P1, P2) |

### Microsoft Graph permissions and roles

The shared sign-in already requests `User.Read.All`, `Directory.Read.All`,
`GroupMember.Read.All` and `AuditLog.Read.All`. Each collector adds what it needs on top for
an interactive sign-in; app-only sign-in uses the permissions granted to the app.

| CSV | Permission | Least privileged signed-in role |
| --- | --- | --- |
| `authentication-methods.csv` | `AuditLog.Read.All` (plus `Directory.Read.All` to read the tenant licence) | Reports Reader |
| `conditional-access-policies.csv` | `Policy.Read.All` | Security Reader (or Global Reader, Conditional Access Administrator) |
| `role-assignments-active.csv` | `RoleAssignmentSchedule.Read.Directory` | Global Reader (or Security Reader, Privileged Role Administrator) |
| `role-assignments-eligible.csv` | `RoleEligibilitySchedule.Read.Directory` | Same as above |
| `role-assignments.csv` | `RoleManagement.Read.Directory` | Directory Readers (or Global Reader, Privileged Role Administrator) |
| `user-signin-activity.csv` | `User.Read.All` and `AuditLog.Read.All` | Reports Reader |
| `risky-users.csv` | `IdentityRiskyUser.Read.All` | Security Reader (or Global Reader, Security Operator) |
| `signins.csv` | `AuditLog.Read.All`, and `Policy.Read.All` to read the Conditional Access result | Reports Reader |

A token that has only `AuditLog.Read.All` can intermittently fail with
`Authentication_RequestFromNonPremiumTenantOrB2CTenant`, because reading the tenant licence
needs `Directory.Read.All` when it is not cached. Grant both
([troubleshooting](https://learn.microsoft.com/troubleshoot/entra/entra-id/users-groups-entra-apis/b2c-or-tenant-premium-license-sign-in-activities)).

## Availability by cloud

`-Environment` is `Commercial`, `GCC` or `GCCHigh` (default `Commercial`). Graph is
`https://graph.microsoft.com` for Commercial and GCC, and `Connect-MgGraph -Environment USGov`
for GCC High ([national cloud deployments](https://learn.microsoft.com/graph/deployments)).

`Available` and `NotAvailable` appear only where a Microsoft page says so. `UNVERIFIED` means
no page found says either way: the collector attempts the call, logs a warning, and records
a refusal in `run.log`. Only `NotAvailable` skips a source and writes the header only. No
source in the contract is `NotAvailable` in any cloud today; the skip path is tested and is
one word in `IdentityPostureSchema.psd1`.

| CSV | Commercial | GCC | GCC High |
| --- | --- | --- | --- |
| `authentication-methods.csv` | [Available](https://learn.microsoft.com/graph/api/authenticationmethodsroot-list-userregistrationdetails) | [Available](https://learn.microsoft.com/graph/deployments) | [Available](https://learn.microsoft.com/graph/api/authenticationmethodsroot-list-userregistrationdetails) |
| `conditional-access-policies.csv` | [Available](https://learn.microsoft.com/graph/api/conditionalaccessroot-list-policies) | [Available](https://learn.microsoft.com/graph/deployments) | [Available](https://learn.microsoft.com/graph/api/conditionalaccessroot-list-policies) |
| `role-assignments-active.csv` | [Available](https://learn.microsoft.com/graph/api/rbacapplication-list-roleassignmentscheduleinstances) | [Available](https://learn.microsoft.com/graph/deployments) | [Available](https://learn.microsoft.com/graph/api/rbacapplication-list-roleassignmentscheduleinstances) |
| `role-assignments-eligible.csv` | [Available](https://learn.microsoft.com/graph/api/rbacapplication-list-roleeligibilityscheduleinstances) | [Available](https://learn.microsoft.com/graph/deployments) | [Available](https://learn.microsoft.com/graph/api/rbacapplication-list-roleeligibilityscheduleinstances) |
| `role-assignments.csv` | [Available](https://learn.microsoft.com/graph/api/rbacapplication-list-roleassignments) | [Available](https://learn.microsoft.com/graph/deployments) | [Available](https://learn.microsoft.com/graph/api/rbacapplication-list-roleassignments) |
| `user-signin-activity.csv` | [Available](https://learn.microsoft.com/graph/api/user-list#example-11-get-users-including-their-last-sign-in-time) | [Available](https://learn.microsoft.com/graph/deployments) | **UNVERIFIED** ([list users](https://learn.microsoft.com/graph/api/user-list) is marked for US Government L4; no page states the `signInActivity` property in that cloud) |
| `risky-users.csv` | [Available](https://learn.microsoft.com/graph/api/riskyuser-list) | [Available](https://learn.microsoft.com/graph/deployments) | [Available](https://learn.microsoft.com/graph/api/riskyuser-list) |
| `signins.csv` | [Available](https://learn.microsoft.com/graph/api/signin-list) | [Available](https://learn.microsoft.com/graph/deployments) | [Available](https://learn.microsoft.com/graph/api/signin-list) |

GCC links to the national cloud deployments page because GCC calls the global service.

## Documented, not collected, beta only

These two sources are in the contract and are not implemented. Microsoft does not support
beta APIs in production, and this library is public and shipped to other tenants.

| Source | Endpoint | What it would add |
| --- | --- | --- |
| 4d: which role definitions are privileged | `GET /beta/roleManagement/directory/roleDefinitions?$filter=isPrivileged eq true` ([privileged roles](https://learn.microsoft.com/entra/identity/role-based-access-control/privileged-roles-permissions)) | The `isPrivileged` flag to join `RoleDefinitionId` against. Without it the report can still group by role but cannot say which roles are privileged |
| 6b: non-interactive sign-ins | `GET /beta/auditLogs/signIns?$filter=signInEventTypes/any(t: t eq 'nonInteractiveUser')` ([signIns, beta](https://learn.microsoft.com/graph/api/signin-list?view=graph-rest-beta)) | Non-interactive sign-ins, which is how legacy protocols often authenticate |

## The sign-in caveat

`signins.csv` comes from the v1.0 list signIns API, which returns sign-ins that are
interactive in nature plus successful federated sign-ins
([List signIns](https://learn.microsoft.com/graph/api/signin-list)). Legacy protocols often
authenticate non-interactively, so **this file can under-report legacy authentication**.
Treat a low count as "at least", not "none". A page holds at most 1,000 sign-ins and the
collector follows `@odata.nextLink`; the log keeps at most 30 days (P1 and P2), so a daily
run builds a longer history than the source retains.

Each window is queried once per legacy `clientAppUsed` value (`Exchange ActiveSync`, `IMAP`,
`MAPI`, `SMTP`, `POP`, `other clients`), so the file holds legacy sign-ins only.

## Running it

```powershell
./collectors/Run-All.ps1 -OutputPath ./out
./collectors/Run-All.ps1 -OutputPath ./out -Environment GCCHigh -AppId $appId `
    -CertificateThumbprint $thumbprint -TenantId $tenantId
./collectors/Get-LegacySignIns.ps1 -OutputPath ./out -LookbackDays 7
```

`Run-All.ps1` runs the users collector first and then each collector in turn; a collector
that fails outright is logged and the rest still run, and the script exits with an error at
the end. Each collector signs in for itself and accepts `-SkipConnect` to reuse a session.

## The CSVs

Booleans are `True`/`False`. Timestamps are UTC, `yyyy-MM-ddTHH:mm:ssZ`. List values are
separated by semicolons. An empty cell means Graph returned no value.

### users.csv — snapshot, key `RunDate` + `Id`

From the shared collector; see `shared/M365ReportLibrary.psm1`. It does not hold
`signInActivity`; use `user-signin-activity.csv`.

### authentication-methods.csv — snapshot, key `RunDate` + `UserId`

| Column | Meaning |
| --- | --- |
| `UserId`, `UserPrincipalName`, `UserDisplayName`, `UserType` | The user; `UserType` is `member` or `guest` |
| `IsAdmin` | The user holds an admin role |
| `IsMfaRegistered`, `IsMfaCapable` | A multifactor method is registered / the user can use one |
| `IsPasswordlessCapable` | The user can sign in without a password |
| `IsSsprEnabled`, `IsSsprRegistered`, `IsSsprCapable` | Self-service password reset state |
| `IsSystemPreferredAuthenticationMethodEnabled` | System-preferred MFA is on for the user |
| `DefaultMfaMethod`, `UserPreferredMethodForSecondaryAuthentication` | As returned by Graph |
| `MethodsRegistered`, `SystemPreferredAuthenticationMethods` | Lists of method names |
| `LastUpdatedDateTime` | When the registration last changed |

Disabled users are not returned by the API, so a "no MFA" count covers enabled users only.

### conditional-access-policies.csv — snapshot, key `RunDate` + `Id`

| Column | Meaning |
| --- | --- |
| `Id`, `DisplayName`, `CreatedDateTime`, `ModifiedDateTime` | The policy |
| `State` | `enabled`, `disabled` or `enabledForReportingButNotEnforced` (report-only) |
| `IncludeUsers`, `ExcludeUsers`, `IncludeGroups`, `ExcludeGroups`, `IncludeRoles`, `ExcludeRoles` | Who it targets (ids, `All`, `GuestsOrExternalUsers`) |
| `IncludeApplications`, `ExcludeApplications` | The apps it targets |
| `ClientAppTypes`, `SignInRiskLevels`, `UserRiskLevels` | Conditions |
| `GrantOperator`, `BuiltInControls` | The grant controls (`OR`/`AND`; `mfa`, `block`, ...) |

A change shows as a difference between two `RunDate` blocks. The API records no one who made it.

### role-assignments-active.csv — snapshot, key `RunDate` + `Id`

| Column | Meaning |
| --- | --- |
| `Id`, `PrincipalId`, `RoleDefinitionId`, `DirectoryScopeId`, `AppScopeId` | The assignment |
| `AssignmentType` | `Assigned` or `Activated` |
| `MemberType` | `Inherited`, `Direct` or `Group` |
| `StartDateTime`, `EndDateTime` | Empty `EndDateTime` is a permanent assignment |
| `RoleAssignmentOriginId`, `RoleAssignmentScheduleId` | Where it came from |

### role-assignments-eligible.csv — snapshot, key `RunDate` + `Id`

Same as above without `AssignmentType`, `RoleAssignmentOriginId` and
`RoleAssignmentScheduleId`, plus `RoleEligibilityScheduleId`.

### role-assignments.csv — snapshot, key `RunDate` + `Id`

`Id`, `PrincipalId`, `RoleDefinitionId`, `DirectoryScopeId`, `AppScopeId`. Active holders only;
use it when `role-assignments-active.csv` is header-only because PIM is not licensed.

### user-signin-activity.csv — snapshot, key `RunDate` + `UserId`

| Column | Meaning |
| --- | --- |
| `LastSignInDateTime` | Last interactive sign-in. Blank means the account never signed in or its last attempt was before April 2020 |
| `LastNonInteractiveSignInDateTime` | Last non-interactive sign-in |
| `LastSuccessfulSignInDateTime` | Last successful sign-in. Not backfilled, so blank is not proof the account never signed in |

Graph returns `0001-01-01T00:00:00Z` or `""` for "no value"; both are written as an empty
cell, never as a date. Selecting `signInActivity` caps a page at 500 users.

### risky-users.csv — snapshot, key `RunDate` + `Id`

`Id`, `UserPrincipalName`, `UserDisplayName`, `RiskLevel` (`low`, `medium`, `high`, `hidden`,
`none`, `unknownFutureValue`), `RiskState` (`none`, `confirmedSafe`, `remediated`,
`dismissed`, `atRisk`, `confirmedCompromised`, `unknownFutureValue`), `RiskDetail`,
`RiskLastUpdatedDateTime`, `IsDeleted`, `IsProcessing`.

### signins.csv — events, key `Id`

`CreatedDateTime`, `Id`, `UserId`, `UserPrincipalName`, `AppDisplayName`,
`ResourceDisplayName`, `IpAddress`, `ClientAppUsed` (a legacy value), `IsInteractive`,
`ErrorCode` (`0` is success), `ConditionalAccessStatus`. Appended from the latest
`CreatedDateTime` already in the file, or `-LookbackDays` (default 30) ago.

## Sample data

`samples/` holds fake data in the exact shape the collectors write: ten users (two guests, one
disabled), two snapshots a month apart, a Conditional Access policy that moves from report-only
to enabled, an administrator who registers for MFA between snapshots, and legacy sign-ins from
documentation IP addresses. Every address is under `example.com`. `samples/unlicensed/` holds
the header-only files a tenant without Entra ID P2 or Governance gets.

## Tests

```powershell
Invoke-Pester -Path ./reports/identity-posture/tests -CI
```

Every tenant call is mocked from the stubs in `shared/tests/TenantCmdletStubs.ps1`. The tests
check that each sample's columns match its collector's output, each connection targets the right
endpoints per `-Environment`, paging is followed, empty `signInActivity` values stay empty, and
a `NotAvailable` source writes a header only.
