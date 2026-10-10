# License utilization

Which licences the tenant owns and how many are free, which users hold which licences (directly or
through a group, and which assignments are in error), when each licensed user last signed in, and what
each licensed user actually does in Exchange, Teams, SharePoint, OneDrive, the Microsoft 365 apps and
Microsoft 365 Copilot.

The collectors write one CSV per source into an output folder of your choosing, plus a `run.log`.
Every source here is a **state** source: each run appends a block of rows tagged with the run date
(`RunDate`), and nothing is rewritten or deleted, so the folder is a history you can chart over time.
There is no event stream in this report, so there is no watermark.

The sources were verified against Microsoft Learn in
[`docs/candidates/license-utilization.md`](../../docs/candidates/license-utilization.md), which is the
contract for this folder. The Power BI project is a separate, later piece of work and is not here.

## Run it

```powershell
./reports/license-utilization/collectors/Run-All.ps1 -OutputPath ./out -Environment Commercial
```

`-Environment` is `Commercial`, `GCC` or `GCCHigh` (default `Commercial`). Commercial and GCC sign in to
`https://graph.microsoft.com`; GCC High signs in with `Connect-MgGraph -Environment USGov`
(`https://graph.microsoft.us`) ([national cloud deployments](https://learn.microsoft.com/graph/deployments)).
Without `-AppId` and `-CertificateThumbprint` the sign-in is interactive and delegated. `Run-All.ps1` runs the
shared Entra users collector first (`users.csv`), then report settings, then the rest; one collector failing
does not stop the others, and the run throws at the end if any failed.
`-IncludeLicenseDetails` adds the per-user licence detail collector, which makes one call per user.
Each collector also runs on its own, for example `./collectors/Get-SubscribedSkus.ps1 -OutputPath ./out`.
The usage reports take `-Period` (`D7`, `D30`, `D90`, `D180`; the Copilot report takes the v2 periods).

Sample data with fake `example.com` users is in `samples/`; `New-SampleData.ps1` regenerates it.
`samples/gcchigh/` holds the header-only files a GCC High run writes for the usage reports.

## Reused from the shared layer

* **User attributes** (department, job title, city, country) come from the shared Entra users collector
  (`Invoke-EntraUserCollector`, `users.csv`). `user-licenses.csv` does not repeat them: join on `UserId`.
* The sign-in, CSV append and run log are the shared layer's (`shared/M365ReportLibrary.psm1`).
* Report-specific helper: `collectors/LicenseUtilizationHelpers.ps1` (availability check, the shared state
  runner, usage report download and paging). Column lists and per-cloud availability are in
  `collectors/LicenseUtilizationSchema.psd1`; the collectors, samples and tests all read the columns from there.

## Roles and licences

| Collector | Permission | Signed-in role | Licence |
| --- | --- | --- | --- |
| Subscribed SKUs | `LicenseAssignment.Read.All` (`Directory.Read.All`, `Organization.Read.All` are higher) | Global Reader or Directory Readers | None named |
| User licences | `User.Read.All` | Directory read | None named |
| Licence details (optional) | Delegated `LicenseAssignment.Read.All`; **application permissions are not supported** | Guest Inviter, Directory Readers, Directory Writers, License Administrator or User Administrator | None named |
| Sign-in activity | `AuditLog.Read.All` + `User.Read.All` | Reports Reader | Microsoft Entra ID P1 or P2 |
| Report settings | `ReportSettings.Read.All` | An Entra limited admin role | None named |
| Usage reports (active users, email, Teams, SharePoint, OneDrive, apps) | `Reports.Read.All` | Reports Reader and similar; Global Reader and Usage Summary Reports Reader do not receive user detail rows | None named |
| Copilot usage | `Reports.Read.All` | As the usage reports | Only users with a Microsoft 365 Copilot licence are returned |

The collectors request only read scopes. `Reports.Read.All`, `LicenseAssignment.Read.All` and
`ReportSettings.Read.All` are added to the shared sign-in's scopes per collector. A tenant without the Entra ID
licence a source needs gets a logged skip and a header-only file, not an error.

## Availability by cloud

`Available` and `NotAvailable` appear only where a Microsoft page says so. **UNVERIFIED** means no page found
says either way: the collector asks for the data anyway, logs a warning, and records a refusal in `run.log`.
Only `NotAvailable` skips a source and writes the header only.

| Source | Collector | Commercial | GCC | GCC High |
| --- | --- | --- | --- | --- |
| 1 Subscribed SKUs | `Get-SubscribedSkus` | [Available](https://learn.microsoft.com/graph/api/subscribedsku-list) | [Available](https://learn.microsoft.com/graph/deployments) | [Available](https://learn.microsoft.com/graph/api/subscribedsku-list) |
| 2 User licences | `Get-UserLicenses` | [Available](https://learn.microsoft.com/graph/api/user-list) | [Available](https://learn.microsoft.com/graph/deployments) | [Available](https://learn.microsoft.com/graph/api/user-list) |
| 3 Licence details | `Get-LicenseDetails` | [Available](https://learn.microsoft.com/graph/api/user-list-licensedetails) | [Available](https://learn.microsoft.com/graph/deployments) | [Available](https://learn.microsoft.com/graph/api/user-list-licensedetails) |
| 4 Sign-in activity | `Get-UserSignInActivity` | [Available](https://learn.microsoft.com/graph/api/resources/user) | [Available](https://learn.microsoft.com/graph/deployments) | **UNVERIFIED** ([list users](https://learn.microsoft.com/graph/api/user-list); open question [BRO-243](https://linear.app/broekncode/issue/BRO-243/settle-whether-signinactivity-is-available-in-gcc-high)) |
| 5a Active users | `Get-ActiveUserUsage` | [Available](https://learn.microsoft.com/graph/api/reportroot-getoffice365activeuserdetail) | [Available](https://learn.microsoft.com/graph/deployments) | [NotAvailable](https://learn.microsoft.com/graph/api/reportroot-getoffice365activeuserdetail) |
| 5b Email activity | `Get-EmailActivityUsage` | [Available](https://learn.microsoft.com/graph/api/reportroot-getemailactivityuserdetail) | [Available](https://learn.microsoft.com/graph/deployments) | [NotAvailable](https://learn.microsoft.com/graph/api/reportroot-getemailactivityuserdetail) |
| 5c Teams activity | `Get-TeamsActivityUsage` | [Available](https://learn.microsoft.com/graph/api/reportroot-getteamsuseractivityuserdetail) | [Available](https://learn.microsoft.com/graph/deployments) | [NotAvailable](https://learn.microsoft.com/graph/api/reportroot-getteamsuseractivityuserdetail) |
| 5d SharePoint activity | `Get-SharePointActivityUsage` | [Available](https://learn.microsoft.com/graph/api/reportroot-getsharepointactivityuserdetail) | [Available](https://learn.microsoft.com/graph/deployments) | [NotAvailable](https://learn.microsoft.com/graph/api/reportroot-getsharepointactivityuserdetail) |
| 5e OneDrive activity | `Get-OneDriveActivityUsage` | [Available](https://learn.microsoft.com/graph/api/reportroot-getonedriveactivityuserdetail) | [Available](https://learn.microsoft.com/graph/deployments) | [NotAvailable](https://learn.microsoft.com/graph/api/reportroot-getonedriveactivityuserdetail) |
| 5f Microsoft 365 apps | `Get-M365AppUsage` | [Available](https://learn.microsoft.com/graph/api/reportroot-getm365appuserdetail) | [Available](https://learn.microsoft.com/graph/deployments) | [NotAvailable](https://learn.microsoft.com/graph/api/reportroot-getm365appuserdetail) |
| 5g Copilot usage | `Get-CopilotUsage` | [Available](https://learn.microsoft.com/microsoft-365-copilot/extensibility/api/admin-settings/reports/copilotreportroot-getmicrosoft365copilotusageuserdetail) | [Available](https://learn.microsoft.com/graph/deployments) | [NotAvailable](https://learn.microsoft.com/microsoft-365-copilot/extensibility/api/admin-settings/reports/copilotreportroot-getmicrosoft365copilotusageuserdetail) |
| 6 Report settings | `Get-ReportSettings` | [Available](https://learn.microsoft.com/graph/api/adminreportsettings-get) | [Available](https://learn.microsoft.com/graph/deployments) | **UNVERIFIED** (the [Graph page](https://learn.microsoft.com/graph/api/adminreportsettings-get) marks US Government L4 unavailable; [activity reports](https://learn.microsoft.com/microsoft-365/admin/activity-reports/activity-reports) says the API works in all environments) |

**Rows 7 and 8 of the contract have no collector.** Row 7 (product names for SKU and service plan ids) is a
public reference page, not a tenant call, and the contract leaves where a dated copy lives as an open item.
Row 8 (the GCC High fallback) is the Microsoft 365 admin center Reports > Usage page, not a Graph call; whether
it can be read by a script in GCC High is not stated on the pages read. In GCC High the usage CSVs are
header-only.

## Names in usage reports

By default Microsoft 365 usage reports show concealed identifiers instead of names. The setting is
**Settings, Org settings, Services, Reports, "Conceal user, group, and site names in all reports"** in the
Microsoft 365 admin center, and `displayConcealedNames` in Graph
([activity reports](https://learn.microsoft.com/microsoft-365/admin/activity-reports/activity-reports)). When it
is on, `UserPrincipalName` in the usage CSVs is a hash that cannot be joined to `user-licenses.csv`. The report
settings collector (`report-settings.csv`) records it, and each usage collector logs a warning when the latest
snapshot says names are concealed. The library only reads the setting; changing it needs
`ReportSettings.ReadWrite.All` and is out of scope.

## What the usage reports do and do not show

* A missing row is not a zero-activity row: reports usually appear within 24 to 72 hours, and a deleted
  user's row leaves the report within 30 days.
* Last Activity Date is the most recent intentional activity whatever the period, so a snapshot can show
  inactivity older than the period. The counts aggregate the period.
* Active users (5a) is the usage CSV whose [documented header](https://learn.microsoft.com/graph/api/reportroot-getoffice365activeuserdetail)
  stops at Assigned Products. When that download has no Report Period column, `ReportPeriod` is the day
  count from `-Period` (`D30` is stored as `30`), the same shape the other usage CSVs use.
* Copilot usage (5g) uses the v1.0 `/copilot` function with `version='v2'` and returns only users with a
  Microsoft 365 Copilot licence. The version 2 additions are named in prose on the Learn page, not printed as
  a CSV header, so their column spellings are **UNVERIFIED**: a header that does not match leaves the cell empty.
* Microsoft 365 apps usage (5f) is read as JSON and `@odata.nextLink` is followed until it is absent; a next
  link on another host is refused.
* The usage reports have no SDK cmdlet in this build, so they are read with `Invoke-MgGraphRequest -Method GET`
  in two helper functions (`Get-GraphReportCsv`, `Get-GraphReportJsonPage`). `tests/ReadOnly.Tests.ps1` holds
  the helper to that.

## CSV files

All files start with `RunDate`, the UTC date of the run.

| File | Source | One row per | Columns |
| --- | --- | --- | --- |
| `subscribed-skus.csv` | 1 | SKU | `SkuId`, `SkuPartNumber`, `AppliesTo`, `CapabilityStatus`, `ConsumedUnits` (assigned), `PrepaidEnabled`, `PrepaidSuspended`, `PrepaidWarning`, `PrepaidLockedOut`, `FreeUnits` (derived: enabled minus consumed) |
| `sku-service-plans.csv` | 1 | SKU and service plan | `SkuId`, `ServicePlanId`, `ServicePlanName`, `ProvisioningStatus`, `AppliesTo` |
| `user-licenses.csv` | 2 | User and licence assignment state | `UserId`, `UserPrincipalName`, `AccountEnabled`, `UsageLocation`, `OfficeLocation`, `SkuId`, `AssignedByGroup` (empty when direct), `AssignmentType` (`Direct` or `Group`), `State` (`Active`, `ActiveWithError`, `Disabled`, `Error`), `Error`, `DisabledPlans` (semicolon list) |
| `license-details.csv` | 3 (optional) | User, SKU and service plan | `UserId`, `SkuId`, `SkuPartNumber`, `ServicePlanId`, `ServicePlanName`, `ProvisioningStatus` |
| `user-signin-activity.csv` | 4 | User | `UserId`, `UserPrincipalName`, `LastSignInDateTime`, `LastNonInteractiveSignInDateTime`, `LastSuccessfulSignInDateTime`. An empty cell means Graph returned no value; it never holds `0001-01-01` |
| `report-settings.csv` | 6 | Run | `DisplayConcealedNames` |
| `usage-active-users.csv` | 5a | User | The report's columns: licence flags, last activity and licence assign date per service, `AssignedProducts`, `ReportPeriod` |
| `usage-email-activity.csv` | 5b | User | Send, receive, read and meeting counts, last activity |
| `usage-teams-activity.csv` | 5c | User | Chat, call and meeting counts, durations, last activity, `IsLicensed` |
| `usage-sharepoint-activity.csv` | 5d | User | Files viewed or edited, synced, shared internally and externally, pages visited, last activity |
| `usage-onedrive-activity.csv` | 5e | User | The same file counts and last activity |
| `usage-m365-apps.csv` | 5f | User | Platform and app flags (Windows, Mac, Mobile, Web; Outlook, Word, Excel, PowerPoint, OneNote, Teams, and each per platform), last activation and activity |
| `usage-copilot.csv` | 5g | User | Last activity date per Copilot app, plus the v2 additions (**UNVERIFIED** spellings) |

Usage CSV columns are the report's own headers with spaces and punctuation removed (`Outlook (Windows)` becomes
`OutlookWindows`), after `RunDate`; the exact lists are in `collectors/LicenseUtilizationSchema.psd1`.

## Tests

```powershell
Invoke-Pester -Path ./reports/license-utilization/tests -CI
```

Every tenant call is mocked; nothing connects to a tenant. The tests check that each sample matches its
collector's output, that each connection targets the endpoint of its `-Environment`, that paged reads are
followed, that state sources stamp the run date, that a `NotAvailable` source writes a header only, and that
the scripts call only read-only commands.
