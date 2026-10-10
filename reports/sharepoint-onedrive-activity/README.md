# SharePoint and OneDrive usage and activity

How SharePoint sites and OneDrive accounts are used: storage and file counts per site and account,
storage trend, who views, syncs and shares files, which sites and accounts have no activity, and
the tenant's storage quota.

This folder holds the collectors, fake sample CSVs and tests. **There is no Power BI project
yet**; it is a separate, later issue. The sources and their per-cloud availability were verified in
[`docs/candidates/sharepoint-onedrive-activity.md`](../../docs/candidates/sharepoint-onedrive-activity.md)
(BRO-381); the tables below are copied from it. Nothing here was run against a tenant.

The collectors write one CSV per source into an output folder, plus a `run.log`. State sources
append a snapshot stamped with `RunDate`; the event source (`file-events.csv`) appends from the day
after the latest `Date` already in the file. Nothing is rewritten or deleted. Everything is
read-only against the tenant.

## Contents

| Path | What it is |
| --- | --- |
| `collectors/Get-SharePointSiteUsageDetail.ps1` | `sharepoint-site-usage-detail.csv`: storage, file counts and last activity per site (source 1) |
| `collectors/Get-OneDriveUsageAccountDetail.ps1` | `onedrive-usage-account-detail.csv`: the same per OneDrive account (source 2) |
| `collectors/Get-SharePointSiteUsageStorage.ps1` | `sharepoint-site-usage-storage.csv`: SharePoint storage, one row per day (source 3) |
| `collectors/Get-OneDriveUsageStorage.ps1` | `onedrive-usage-storage.csv`: OneDrive storage, one row per day (source 4) |
| `collectors/Get-SharePointActivityUserDetail.ps1` | `sharepoint-activity-user-detail.csv`: files viewed or edited, synced, shared, pages visited per user (source 5) |
| `collectors/Get-OneDriveActivityUserDetail.ps1` | `onedrive-activity-user-detail.csv`: the same without pages visited (source 6) |
| `collectors/Get-ReportSettings.ps1` | `report-settings.csv`: whether usage reports conceal names (source 7) |
| `collectors/Get-TenantStorage.ps1` | `tenant-storage.csv`: tenant storage and resource quota from `Get-SPOTenant` (source 9, route b) |
| `collectors/Get-SpoSites.ps1` | `spo-sites.csv`: every site collection and OneDrive site from `Get-SPOSite` (source 10) |
| `collectors/Get-DriveQuota.ps1` | `drive-quota.csv`: used and total bytes per drive through Graph (source 11) |
| `collectors/Get-FileEvents.ps1` | `file-events.csv`: audit-log file and sharing operations counted per day and user (source 12) |
| `collectors/Get-SiteActivity.ps1` | `site-activity.csv`: access, create, edit, delete and move counts per site per day (source 13) |
| `collectors/Run-All.ps1` | Signs in once each and runs all twelve |
| `collectors/SharePointOneDriveSchema.psd1` | Column order of every CSV and the per-cloud availability of every source |
| `collectors/SharePointOneDriveHelpers.ps1` | Report-specific helpers (below) |
| `New-SampleData.ps1` | Regenerates `samples/` |
| `samples/` | Fake CSVs (everything is on `example.com`); `samples/gcchigh/` holds the header-only files a GCC High run leaves for the six Graph usage reports |
| `tests/` | Pester 5 tests; every tenant call is mocked |

Two contract rows are **not** collectors:

* Source 8, the same reports in the Microsoft 365 admin center (**Reports > Usage**), is a manual CSV
  export. The pages read do not say a script can read it in GCC High, so it stays an open item in
  the contract.
* Source 9 route (a), the **SharePoint Storage** report in the admin center, is manual too.
  Route (b), `Get-SPOTenant`, is `tenant-storage.csv`.

`GET /users/{id}/drives` (named in source 11) is not called. A personal site that `getAllSites`
returns is read through `GET /sites/{id}/drives` like any other site, and `drive-quota.csv` marks it
`IsPersonalSite`.

`SharePointOneDriveHelpers.ps1` is report-specific because the shared layer has no usage-report,
audit-log counting, SharePoint Online sign-in or Graph paging logic: the availability rule, the
usage-report download (a CSV behind a `302` to a short-lived URL), the Graph read that follows
`@odata.nextLink` as returned and only to a Graph host, the `Connect-SPOService` sign-in (with
`-Region ITAR` for GCC High), the `Search-UnifiedAuditLog` `ReturnLargeSet` paging, and the daily
count. The shared connection (`Connect-M365Service`) and CSV append functions are used as they are.

## Before you run it

```powershell
Install-Module Microsoft.Graph, ExchangeOnlineManagement, Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser
```

```powershell
./collectors/Run-All.ps1 -OutputPath ./out -AdminUrl https://contoso-admin.sharepoint.com
./collectors/Run-All.ps1 -OutputPath ./out -Environment GCCHigh -AdminUrl https://contoso-admin.sharepoint.us `
    -AppId $appId -CertificateThumbprint $thumbprint -TenantId $tenantId -Organization contoso.onmicrosoft.com
```

`-Environment` is `Commercial` (default), `GCC` or `GCCHigh`. Graph signs in to
`https://graph.microsoft.com` for Commercial and GCC and with `Connect-MgGraph -Environment USGov`
(`https://graph.microsoft.us`) for GCC High
([national cloud deployments](https://learn.microsoft.com/graph/deployments)). Exchange Online connects
with `O365Default` or `O365USGovGCCHigh`. SharePoint Online connects with `Connect-SPOService -Url
<admin url>`, and `-Region ITAR` in GCC High, the value the cmdlet page says applies only to GCC High
and DoD tenants.

`getAllSites` supports application permissions only, so `drive-quota.csv` and `site-activity.csv`
return data only on an app-only sign-in (`-AppId` and `-CertificateThumbprint`). Without `-AdminUrl`,
`tenant-storage.csv` and `spo-sites.csv` are header-only. A collector whose source is refused or
unavailable leaves a header-only CSV and a line in `run.log`; the others still run. Pass `-Period`
(`D7`, `D30`, `D90`, `D180`) to the usage reports, `-Date` to the two activity reports for a single
day within the past 30 days, `-LookbackDays` (up to 365 for the audit log; site activity
stays under 90 days) and `-SiteLimit` for a trial run.

## What each collector needs

| Collector | Role or permission | Licence |
| --- | --- | --- |
| Sources 1 to 6 | `Reports.Read.All` (delegated and application). Signed in: Company Administrator, Exchange Administrator, SharePoint Administrator, Lync Administrator, Teams Service Administrator, Teams Communications Administrator or Reports Reader. Global Reader and Usage Summary Reports Reader "only have access to tenant-level data, without visibility into detailed metrics" ([authorization](https://learn.microsoft.com/graph/reportroot-authorization)) | None named. The admin center OneDrive usage page limits the report to users with a valid OneDrive licence |
| Source 7 | `ReportSettings.Read.All`; an Entra limited admin role. Changing the setting needs `ReportSettings.ReadWrite.All` and is out of scope | None named |
| Source 9 (b), 10 | SharePoint Online administrator (and site collection administrator for `Get-SPOSite`) | None named |
| Source 11 | Application `Sites.Read.All` for `getAllSites` (delegated not supported); application `Files.Read.All` (higher: `Sites.Read.All`) for the drive list | None named |
| Source 12 | View-Only Audit Logs (least privileged read) or Audit Logs ([search the audit log](https://learn.microsoft.com/purview/audit-search)) | Audit (Standard). Auditing is on by default for enterprise organizations, and not for Microsoft 365 Business Basic, Business Standard, Business Premium, or unmanaged trial tenants ([turn auditing on or off](https://learn.microsoft.com/purview/audit-log-enable-disable)) |
| Source 13 | Application `Files.Read.All` (delegated least privileged `Files.Read`) | None named |

## Availability per cloud

`Available` and `NotAvailable` appear only where a Microsoft page says so; the link is in the cell.
`UNVERIFIED` means no Microsoft page read says either way: the collector asks for the data and
records a refusal in `run.log`. A `NotAvailable` source is skipped: the CSV gets a header only and
`run.log` says why.

| # | CSV | Commercial | GCC | GCC High |
| --- | --- | --- | --- | --- |
| 1 | `sharepoint-site-usage-detail.csv` | [Available](https://learn.microsoft.com/graph/api/reportroot-getsharepointsiteusagedetail) | [Available](https://learn.microsoft.com/graph/deployments) (global service) | [NotAvailable](https://learn.microsoft.com/graph/api/reportroot-getsharepointsiteusagedetail) (US Government L4 `❌`) |
| 2 | `onedrive-usage-account-detail.csv` | [Available](https://learn.microsoft.com/graph/api/reportroot-getonedriveusageaccountdetail) | [Available](https://learn.microsoft.com/graph/deployments) (global service) | [NotAvailable](https://learn.microsoft.com/graph/api/reportroot-getonedriveusageaccountdetail) (US Government L4 `❌`) |
| 3 | `sharepoint-site-usage-storage.csv` | [Available](https://learn.microsoft.com/graph/api/reportroot-getsharepointsiteusagestorage) | [Available](https://learn.microsoft.com/graph/deployments) (global service) | [NotAvailable](https://learn.microsoft.com/graph/api/reportroot-getsharepointsiteusagestorage) (US Government L4 `❌`) |
| 4 | `onedrive-usage-storage.csv` | [Available](https://learn.microsoft.com/graph/api/reportroot-getonedriveusagestorage) | [Available](https://learn.microsoft.com/graph/deployments) (global service) | [NotAvailable](https://learn.microsoft.com/graph/api/reportroot-getonedriveusagestorage) (US Government L4 `❌`) |
| 5 | `sharepoint-activity-user-detail.csv` | [Available](https://learn.microsoft.com/graph/api/reportroot-getsharepointactivityuserdetail) | [Available](https://learn.microsoft.com/graph/deployments) (global service) | [NotAvailable](https://learn.microsoft.com/graph/api/reportroot-getsharepointactivityuserdetail) (US Government L4 `❌`) |
| 6 | `onedrive-activity-user-detail.csv` | [Available](https://learn.microsoft.com/graph/api/reportroot-getonedriveactivityuserdetail) | [Available](https://learn.microsoft.com/graph/deployments) (global service) | [NotAvailable](https://learn.microsoft.com/graph/api/reportroot-getonedriveactivityuserdetail) (US Government L4 `❌`) |
| 7 | `report-settings.csv` | [Available](https://learn.microsoft.com/graph/api/adminreportsettings-get) | [Available](https://learn.microsoft.com/graph/deployments) (global service) | UNVERIFIED: [the Graph page](https://learn.microsoft.com/graph/api/adminreportsettings-get) marks US Government L4 `❌`, and the [usage reports overview](https://learn.microsoft.com/microsoft-365/admin/activity-reports/activity-reports) says "an API in all environments" can change the setting |
| 9 (b) | `tenant-storage.csv` | UNVERIFIED (no page states this cmdlet per cloud) | UNVERIFIED (no cloud statement on the [cmdlet page](https://learn.microsoft.com/powershell/module/microsoft.online.sharepoint.powershell/get-spotenant)) | UNVERIFIED (no cloud statement on the cmdlet page) |
| 10 | `spo-sites.csv` | UNVERIFIED (no page states this cmdlet per cloud) | UNVERIFIED (no cloud statement on the [cmdlet page](https://learn.microsoft.com/powershell/module/microsoft.online.sharepoint.powershell/get-sposite)) | UNVERIFIED: [`-Region ITAR`](https://learn.microsoft.com/powershell/module/microsoft.online.sharepoint.powershell/connect-sposervice) shows the module connects to GCC High; no page read says `Get-SPOSite` returns these properties there |
| 11 | `drive-quota.csv` | [Available](https://learn.microsoft.com/graph/api/site-getallsites); [Available](https://learn.microsoft.com/graph/api/drive-list) | [Available](https://learn.microsoft.com/graph/deployments) (global service) | [Available](https://learn.microsoft.com/graph/api/site-getallsites) (US Government L4 `✅`); [Available](https://learn.microsoft.com/graph/api/drive-list) (US Government L4 `✅`) |
| 12 | `file-events.csv` | [Available](https://learn.microsoft.com/purview/audit-solutions-overview) (Audit (Standard) and `Search-UnifiedAuditLog` listed) | [Available](https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-deployments) (Audit (Standard): Available) | [Available](https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-high-deployments) (Audit (Standard) listed `Available` for GCC High; the cmdlet itself is not named there) |
| 13 | `site-activity.csv` | [Available](https://learn.microsoft.com/graph/api/itemactivitystat-getactivitybyinterval) (global service `✅`) | [Available](https://learn.microsoft.com/graph/deployments) (global service) | UNVERIFIED: [the API page](https://learn.microsoft.com/graph/api/itemactivitystat-getactivitybyinterval) marks US Government L4 `✅` and says itemAnalytics is not yet available in all national deployments |

**GCC High.** Every Graph usage report this report needs (sources 1 to 6) is `NotAvailable` through
Graph there, although the same reports are listed `Yes` for GCC-High in the Microsoft 365 admin
center (source 8). So the report is full in Commercial and GCC; in GCC High a script gets the site
list and per-drive quota (`drive-quota.csv`), audit-log event counts (`file-events.csv`) and, if the
service answers, site action counts (`site-activity.csv`). File counts, active file counts and per-user
viewed, edited, synced and shared counts have no scripted source there.

## CSVs

Every CSV is written with the columns in `SharePointOneDriveSchema.psd1`; `samples/` holds one
of each.

| CSV | Kind | One row per | Columns that need a note |
| --- | --- | --- | --- |
| `sharepoint-site-usage-detail.csv` | State, period aggregate | site per run and period | `StorageUsedByte` and `StorageAllocatedByte` are bytes; the admin center shows MB, so never add the two. `ActiveFileCount` counts files saved, synced, modified or shared in the period and can exceed `FileCount`. `LastActivityDate` is driven by audit events, which for sites include page views. `SiteId` and `SiteUrl` are concealed or empty (see below) |
| `onedrive-usage-account-detail.csv` | State, period aggregate | account per run and period | No `SiteId` column in the Graph CSV. `LastActivityDate` is blank when the OneDrive has no file activity. `IsDeleted` is set after at least seven days |
| `sharepoint-site-usage-storage.csv` | State, daily points in a period | report date per run and period | `SiteType`. No allocated column |
| `onedrive-usage-storage.csv` | State, daily points in a period | report date per run and period | As above |
| `sharepoint-activity-user-detail.csv` | State, period aggregate or one day | user per run, query date and period | `QueryDate` is the day asked for with `-Date`, empty for a period. Counts are files, not links or permissions. `LastActivityDate` is for the selected range, not lifetime. `IsDeleted` means the licence was removed |
| `onedrive-activity-user-detail.csv` | State | user per run, query date and period | As above, no `VisitedPageCount` |
| `report-settings.csv` | State | run | `DisplayConcealedNames` |
| `tenant-storage.csv` | State | run | The properties `Get-SPOTenant` returns; one it does not return is empty and `run.log` names it |
| `spo-sites.csv` | State | site per run | `StorageUsageCurrent`, `ResourceUsageCurrent` and `WebsCount` are empty unless the cmdlet returns them without the deprecated `-Detailed` |
| `drive-quota.csv` | State | drive per run | Bytes. `QuotaState` is `normal`, `nearing` (under 10% left), `critical` (under 1%) or `exceeded`. No file count. `LastModifiedDateTime` is when the drive was modified, not the usage report's last activity. Drives with the system facet are included because the request `$select`s `system`; without that they are hidden |
| `file-events.csv` | Event | UTC day, workload, user and operation | `EventCount` is counted by the library from the audit log and will not match sources 5 and 6. Operations: `FileAccessed`, `FileModified`, `FileDownloaded`, `FileUploaded`, `FileSyncDownloadedFull`, `FileSyncUploadedFull`, `PageViewed`, `SharingSet`, `AnonymousLinkCreated`, `SecureLinkCreated`. `FileAccessed` and `FileModified` are not logged again for the same user and file for five minutes. Only whole UTC days before today are written |
| `site-activity.csv` | State, interval aggregate | site and day per run | Site action counts, not file counts and not per-user counts. An action an interval does not carry is empty, not zero. `IncompleteData` is `True` when Graph says the interval is based on incomplete data; a zero there is not "no activity" |

Retention at the source: the `date` form of the activity reports reaches back 30 days on the Graph
page (the admin center selected-day table, 28); the audit log keeps Audit (Standard) records for 180
days, and one year for E5 users of SharePoint and OneDrive, so `-LookbackDays` accepts 365.
Usage reports are typically available
within 24 to 72 hours and sometimes take several days, so the newest days are missing, and a missing
day is not zero activity.

## Concealed names

The tenant setting that controls this is **Settings, Org settings, Services, Reports, "Conceal user,
group, and site names in all reports"** in the Microsoft 365 admin center, read through
`GET /admin/reportSettings` (`displayConcealedNames`). By default usage reports hide user names,
display names, groups and sites; the setting covers "Site IDs and Site URLs" in the OneDrive and
SharePoint site usage reports. When it is `true`, sources 1, 2, 5 and 6 hide site ids, site URLs and
user principal names, so they cannot be joined to `drive-quota.csv`, `spo-sites.csv` or `users.csv`.
Read `report-settings.csv` first (`Run-All.ps1` does): the report must show "names concealed, usage
not joined" instead of zero-activity sites. A change takes a few minutes and also applies to the
Graph usage reports. This report only reads the setting.

## Other reports that cover the same ground

* **Oversharing** ([`reports/oversharing`](../oversharing/)) owns sharing links, permissions and
  sharing audit events. Sources 5 and 6 here count files shared per user; they are not a link or
  permission inventory.
* **Teams and Groups lifecycle** ([`reports/teams-groups-lifecycle`](../teams-groups-lifecycle/)) owns
  group and Team ownership and lifecycle. This report shows site inactivity and storage and does not
  decide which groups to expire.
* **License utilization** (candidate: [`docs/candidates/license-utilization.md`](../../docs/candidates/license-utilization.md))
  shares the user activity reports (its sources 5d and 5e, here sources 5 and 6) and
  `displayConcealedNames` (its source 6, here source 7). It is not built yet; when it is, call each
  report once and share the result.

## Dropped starting questions

Nothing is dropped outright. In GCC High the usage reports are skipped, so file counts, active file
counts, per-user viewed, edited, synced and shared counts, and `Last Activity Date`-based inactivity
are not collected there; `file-events.csv` is a different measure and must not be compared with them.
