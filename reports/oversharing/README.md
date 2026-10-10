# Oversharing

Which sites, libraries and files are open more widely than they should be: sites reachable by many users, items
shared with "Everyone" or "Everyone except external users" (EEEU), items with Anyone, organization-wide or
specific-people sharing links, the sharing settings that allow it, and the sharing events over time.

The collectors write one CSV per source into an output folder of your choosing, plus a `run.log`. State files
get a new block of rows tagged with `RunDate` on every run; the two audit event files are appended from the
last `CreationTime` already in the file. Nothing is rewritten or deleted, so the folder is a history you can chart.

The sources were verified against Microsoft Learn in
[`docs/candidates/oversharing.md`](../../docs/candidates/oversharing.md), which is the contract for this folder.
Its "dropped" list is not built: who created a link older than the 180-day audit window, Graph-level
EEEU/Everyone site group membership, and sensitive information types. The Power BI project is a separate, later piece
of work and is not here. Nothing in this folder has been run against a tenant.

## Run it

```powershell
./reports/oversharing/collectors/Run-All.ps1 -OutputPath ./out -AdminUrl https://contoso-admin.sharepoint.com `
    -AppId <app> -CertificateThumbprint <thumbprint> -TenantId <tenant> -Environment Commercial
```

`-Environment` is `Commercial`, `GCC` or `GCCHigh` (default `Commercial`). Commercial and GCC sign in to
`https://graph.microsoft.com`; GCC High signs in with `Connect-MgGraph -Environment USGov`
(`https://graph.microsoft.us`) ([national cloud deployments](https://learn.microsoft.com/graph/deployments)).
The SharePoint Online cmdlets need `-AdminUrl` (the admin center URL); GCC High adds `-Region ITAR`.
`Run-All.ps1` signs in once per service, runs the collectors with `-SkipConnect`, keeps going if one fails, and
throws at the end if any did. A collector that is refused logs the reason and leaves a header-only file.

Options worth knowing:

* `-MaxSites`, `-MaxItemsPerDrive`, `-LinksOnly`, `-SkipItemPermissions` bound the item permission walk, which is
  one Graph call per item. Read Learn's [scan guidance](https://learn.microsoft.com/onedrive/developer/rest-api/concepts/scan-guidance) before running it against a large tenant.
* `-StartDate`, `-EndDate`, `-LookbackDays` (default 90), `-WindowHours` (default 24) set the audit search range.
* `-LabelGuid` names the sensitivity labels for the labelled-files report; without it the collector lists file labels with `Get-Label`.
* `-WaitMinutes` is how long a Data access governance report is waited for. A report still running is left
  running and a later run exports it. A completed report newer than `-MaxReportAgeHours` is reused, because these
  reports can be re-run only once every 30 days.

Sample data with fake `example.com` users is in `samples/`; `New-SampleData.ps1` regenerates it.

## Reused and added

* The sign-in, CSV append, watermark and run log are the shared layer's (`shared/M365ReportLibrary.psm1`).
* **Guest and external access** and **sensitivity label coverage** are other reports' jobs and are not rebuilt.
  The contract has no join to users, so this report writes no `users.csv`; run the shared `Invoke-EntraUserCollector`
  if you want one beside it. Specific-people links appear only as counts by site.
* Report-specific helper: `collectors/OversharingHelpers.ps1` (availability check, Graph GET and paging, the
  drive walk, SharePoint admin sign-in, the Data access governance start/poll/export helper, and the audit search).
  Column lists and per-cloud availability are in `collectors/OversharingSchema.psd1`.
* **Data collection for activity reports is never started.** Without a SharePoint Advanced Management licence, the
  contract suggests `Start-SPOAuditDataCollectionForActivityInsights`. That turns collection on in the tenant, which a
  read-only library cannot do, so 3c and 3d only read `Get-SPOAuditDataCollectionStatusForActivityInsights` and log a
  warning when it is not `InProgress`.

## Roles and licences

| Collector | Least privileged role | Licence |
| --- | --- | --- |
| Sites, item permissions | Application `Sites.Read.All`; `Files.Read.All` (delegated: `Files.Read`). `getAllSites` does not support delegated | None named |
| Site permission breadth, Everyone item exposure, sharing link and EEEU activity, labelled-file sites | SharePoint Administrator or SharePoint Advanced Management Administrator (item-level Everyone report: the latter, assigned by a Global Administrator); `Connect-SPOService` without `-Credential` | SharePoint Advanced Management (a Microsoft 365 Copilot licence, or the SAM Plan 1 add-on). E5 without SAM gets no snapshot reports; activity reports then return at most 10,000 sites |
| Site sharing settings | SharePoint Online administrator and site collection administrator | None named |
| Anonymous link and sharing events | Audit Reader, and the Exchange View-Only Audit Logs or Audit Logs role | Audit (Standard); 180 days retention (one year with E5) |
| Audit log status | Audit roles as above; the role for `Get-AdminAuditLogConfig` is **UNVERIFIED** | None named |

The labelled-files collector also runs `Get-Label` when no `-LabelGuid` is given; its role is **UNVERIFIED**.
Data access governance reports might not work when "Display concealed user, group, and site names in all
reports" is cleared in the Microsoft 365 admin center, which only a Global Administrator can change
([limitations](https://learn.microsoft.com/sharepoint/data-access-governance-reports#limitations-or-known-issues)).
They are unavailable for Microsoft 365 operated by 21Vianet.

## Availability by cloud

`Available` and `NotAvailable` appear only where a Microsoft page says so. **UNVERIFIED** means no page found says
either way: the collector asks for the data anyway, logs a warning, and records a refusal in `run.log`. Only
`NotAvailable` skips a source and writes the header only; no source is `NotAvailable` in any cloud.

| Source | Collector | Commercial | GCC | GCC High |
| --- | --- | --- | --- | --- |
| 1 Site list | `Get-Sites` | [Available](https://learn.microsoft.com/graph/api/site-getallsites) | [Available](https://learn.microsoft.com/graph/deployments#microsoft-graph-and-graph-explorer-service-root-endpoints) | [Available](https://learn.microsoft.com/graph/api/site-getallsites) |
| 2 Item sharing permissions | `Get-ItemSharingPermissions` | [Available](https://learn.microsoft.com/graph/api/driveitem-list-permissions) | [Available](https://learn.microsoft.com/graph/deployments#microsoft-graph-and-graph-explorer-service-root-endpoints) | [Available](https://learn.microsoft.com/graph/api/driveitem-list-permissions) |
| 3a Site permission breadth | `Get-SitePermissionBreadth` | [Available](https://learn.microsoft.com/sharepoint/sharepoint-advanced-management-features-copilot-license#sam-features-in-copilot-environments) | [Available](https://learn.microsoft.com/sharepoint/sharepoint-advanced-management-features-copilot-license#sam-features-in-copilot-environments) (feature row, cmdlet not named) | [Available](https://learn.microsoft.com/sharepoint/sharepoint-advanced-management-features-copilot-license#sam-features-in-copilot-environments) (feature row, cmdlet not named) |
| 3b Everyone / EEEU items | `Get-EveryoneItemExposure` | [Available](https://learn.microsoft.com/sharepoint/sharepoint-advanced-management-features-copilot-license#sam-features-in-copilot-environments) | **UNVERIFIED** (the feature row says Yes for GCC; no page says this item-level report is available there) | **UNVERIFIED** (same) |
| 3c Sharing link activity | `Get-SharingLinkActivity` | [Available](https://learn.microsoft.com/sharepoint/sharepoint-advanced-management-features-copilot-license#sam-features-in-copilot-environments) | [Available](https://learn.microsoft.com/sharepoint/sharepoint-advanced-management-features-copilot-license#sam-features-in-copilot-environments) (feature row, cmdlet not named) | [Available](https://learn.microsoft.com/sharepoint/sharepoint-advanced-management-features-copilot-license#sam-features-in-copilot-environments) (feature row, cmdlet not named) |
| 3d EEEU activity | `Get-EeeuActivity` | [Available](https://learn.microsoft.com/sharepoint/sharepoint-advanced-management-features-copilot-license#sam-features-in-copilot-environments) | [Available](https://learn.microsoft.com/sharepoint/sharepoint-advanced-management-features-copilot-license#sam-features-in-copilot-environments) (feature row, cmdlet not named) | [Available](https://learn.microsoft.com/sharepoint/sharepoint-advanced-management-features-copilot-license#sam-features-in-copilot-environments) (feature row, cmdlet not named) |
| 3e Labelled-file sites | `Get-LabeledFileSites` | [Available](https://learn.microsoft.com/sharepoint/sharepoint-advanced-management-features-copilot-license#sam-features-in-copilot-environments) | [Available](https://learn.microsoft.com/sharepoint/sharepoint-advanced-management-features-copilot-license#sam-features-in-copilot-environments) (feature row, cmdlet not named) | [Available](https://learn.microsoft.com/sharepoint/sharepoint-advanced-management-features-copilot-license#sam-features-in-copilot-environments) (feature row, cmdlet not named) |
| 4 Site and tenant sharing settings | `Get-SiteSharingSettings` | **UNVERIFIED** (no page states availability per cloud; [Get-SPOSite](https://learn.microsoft.com/powershell/module/microsoft.online.sharepoint.powershell/get-sposite)) | **UNVERIFIED** | **UNVERIFIED** |
| 5a Anonymous link events | `Get-AnonymousLinkEvents` | [Available](https://learn.microsoft.com/purview/audit-solutions-overview#audit-standard) | [Available](https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-deployments) (Audit (Standard)) | [Available](https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-high-deployments) (Audit (Standard)) |
| 5b Sharing events | `Get-SharingEvents` | [Available](https://learn.microsoft.com/purview/audit-solutions-overview#audit-standard) | [Available](https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-deployments) (the specific SharePoint operations are not listed on that page) | [Available](https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-high-deployments) (same note) |
| 5c Audit log on | `Get-AuditLogStatus` | [Available](https://learn.microsoft.com/purview/audit-log-enable-disable) | **UNVERIFIED** | **UNVERIFIED** |

Whether `Connect-SPOService` and the `*-SPODataAccessGovernanceInsight` cmdlets connect in GCC and GCC High is an
open item in the contract: those clouds log a warning and try anyway.

## What the sources do and do not show

* **Item permissions (2).** What an app-only caller sees from `GET .../permissions` is **UNVERIFIED**: Learn says a
  non-owner caller gets only the permissions that apply to it. Every run logs a warning, and the file is not proof
  that every link was seen. `link.webUrl` and `shareId` are secrets and are never written. Learn does not list the
  creator of a permission; who created a link comes from the audit events (5a, 5b).
* **Data access governance reports (3a to 3e)** are asynchronous: the first site permissions report takes up to five
  days, later ones within 24 hours, data can be 48 hours old, and a report can be re-run once every 30 days. Reports
  3c and 3d cover a rolling 28 days; a gap longer than 28 days between runs loses events, and the audit log is the
  longer record.
* **Column names.** The columns of the site permissions report (3a) and the Everyone/EEEU item report (3b) are the
  ones on their Learn pages. Learn does not list the CSV columns of 3c, 3d and 3e, so those files keep the whole
  exported row as JSON in `ReportRow` and copy `SiteId` and `SiteUrl` out when present: their inner column names are
  **UNVERIFIED**. Which `-ReportEntity` spelling the data collection status cmdlet accepts
  (`SharingLinksAnyone` or `SharingLinks_Anyone`) is **UNVERIFIED**, so both are tried.
* **Site sharing settings (4).** `Get-SPOSite -Limit All` does not populate the sharing properties, so each site is
  read again by URL. That `Get-SPOTenant` returns the tenant sharing level is **UNVERIFIED**; a refusal is logged.
* **Audit events (5a, 5b).** `Search-UnifiedAuditLog` is paged with a `ReturnLargeSet` session (5,000 per page, 50,000
  per session). A window that reaches 50,000 is incomplete and unsorted, so it is not written: the collector throws
  and tells you the window to re-run with a smaller `-WindowHours`. Audit (Standard) keeps 180 days.

## CSV files

| File | Source | Kind | One row per | Columns |
| --- | --- | --- | --- | --- |
| `sites.csv` | 1 | State | Site | `SiteId`, `Name`, `WebUrl`, `IsPersonalSite`, `HostName`, `DataLocationCode` |
| `item-permissions.csv` | 2 | State | Permission on an item | `SiteId`, `DriveId`, `ItemId`, `ItemName`, `ItemWebUrl`, `PermissionId`, `Roles`, `LinkScope` (`anonymous`, `organization`, `users`, `existingAccess`), `LinkType`, `LinkPreventsDownload`, `HasPassword`, `ExpirationDateTime`, `IsInherited`, `InheritedFromItemId`, `GrantedTo` |
| `site-permission-breadth.csv` | 3a | State | Site and workload | `Workload`, `ReportId`, `ReportDate`, `SiteId`, `SiteName`, `SiteUrl`, `SiteTemplate`, primary admin, `ExternalSharing`, `SitePrivacy`, `SiteSensitivity`, `UsersWithAccess`, guest / external / Entra group permission counts, `FileCount`, `ItemsWithUniquePermissions`, `PeopleInYourOrgLinks`, `AnyoneLinks`, `EeeuPermissions`, `EveryonePermissions` |
| `everyone-item-exposure.csv` | 3b | State | Item shared with EEEU or Everyone | `ReportEntity`, `ReportId`, `ReportDate`, the site to item identifiers, `ItemType`, `ItemUrl`, `RoleDefinition`, `LinkScope`, `Recipient`, the parent group columns, `TotalUserCount` |
| `sharing-link-activity.csv` | 3c | State (28-day window) | Exported report row | `ReportEntity`, `Workload`, `ReportId`, `ReportStartTime`, `ReportEndTime`, `SiteId`, `SiteUrl`, `ReportRow` (JSON, **UNVERIFIED** inner columns) |
| `eeeu-activity.csv` | 3d | State (28-day window) | Exported report row | As 3c |
| `labeled-file-sites.csv` | 3e | State | Exported report row, per label | `LabelGuid`, `LabelName`, `Workload`, `ReportId`, `ReportCreatedDateTime`, `SiteId`, `SiteUrl`, `ReportRow` |
| `site-sharing-settings.csv` | 4 | State | Tenant, then each site | `Scope` (`Tenant`, `Site`), `Url`, `Title`, `Template`, `SharingCapability`, `DefaultSharingLinkType`, `DisableCompanyWideSharingLinks`, `SensitivityLabel` |
| `anonymous-link-events.csv` | 5a | Event | Audit record | `CreationTime`, `Id`, `RecordType`, `Operation`, `UserId`, `Workload`, `ObjectId`, `ItemType`, `SiteUrl`, `SourceRelativeUrl`, `SourceFileName`, `TargetUserOrGroupName`, `TargetUserOrGroupType`, `ClientIP` |
| `sharing-events.csv` | 5b | Event | Audit record | As 5a |
| `audit-log-status.csv` | 5c | State | Run | `UnifiedAuditLogIngestionEnabled` |

State files start with `RunDate`, the UTC date of the run.

## Tests

```powershell
Invoke-Pester -Path ./reports/oversharing/tests -CI
```

Every tenant call is mocked; nothing connects to a tenant. The tests check that each sample matches its collector's
output, that each connection targets the endpoint of its `-Environment`, that paged reads are followed, that
event sources resume from the last exported timestamp, that state sources stamp the run date, that a source that
cannot be read writes a header only, and that the scripts call only read-only commands and never start data collection.
