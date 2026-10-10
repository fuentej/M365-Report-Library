# Unified audit log

Which admin and user operations happen, by workload, user and day; who changed admin roles and policies;
which file and sharing operations happened; which mailboxes were accessed by someone other than the owner;
how long records are kept. Five sources read the audit log or the settings that decide what it holds.

This folder holds the collectors, fake sample CSVs and tests. **There is no Power BI project yet**; it is a
separate, later issue, and the report pages in the contract are for that issue. The sources and their
per-cloud availability were verified in [`docs/candidates/unified-audit-log.md`](../../docs/candidates/unified-audit-log.md)
(BRO-379); the tables below are copied from it.

The collectors write one CSV per source into an output folder, plus a `run.log`. Event sources (1 to 3) append
from the latest timestamp already in the file; state sources (4 and 5) append a snapshot stamped with `RunDate`.
Nothing is rewritten or deleted.

## Read-only, and the two POSTs

Every tenant call is a `Connect-*`, `Disconnect-*`, `Get-*` or `Search-*` cmdlet or an HTTP GET, except two POSTs.
Both start a read and neither changes tenant data (the contract says so). Decision D-007 (Joshua, 2026-10-10:
"Do what enables the functionality") allows them, and `tests/ReadOnly.Tests.ps1` pins each to its path and to the
one function that sends it.

| Collector | POST | Why it is needed |
| --- | --- | --- |
| `Get-AuditGraphRecords.ps1` | `POST /v1.0/security/auditLog/queries` ([create auditLogQuery](https://learn.microsoft.com/graph/api/security-auditcoreroot-post-auditlogqueries)) | Records can only be listed from a query object that this call creates. |
| `Get-AuditActivityFeed.ps1` | `POST {root}/api/v1.0/{tenant}/activity/feed/subscriptions/start?contentType=...` ([start a subscription](https://learn.microsoft.com/office/office-365-management-api/office-365-management-activity-api-reference#start-a-subscription)) | Content can only be listed for a content type that has a subscription. |

The subscription start is sent only for a content type whose subscription is not already `enabled`, and
`-NoStartSubscription` withholds it entirely. A subscription is never stopped: stopping and starting it again does
not return the content from the gap, and a second start within 15 minutes is throttled.

## Contents

| Path | What it is |
| --- | --- |
| `collectors/Get-AuditSearchCmdlet.ps1` | `audit-search-cmdlet.csv`: records from `Search-UnifiedAuditLog` (source 1) |
| `collectors/Get-AuditGraphRecords.ps1` | `audit-graph-records.csv`: records from the Graph Audit Search API (source 2) |
| `collectors/Get-AuditActivityFeed.ps1` | `audit-activity-feed.csv`: events from the Office 365 Management Activity API (source 3) |
| `collectors/Get-AuditIngestion.ps1` | `audit-ingestion.csv`: whether audit ingestion is on (source 4) |
| `collectors/Get-AuditRetentionPolicies.ps1` | `audit-retention-policies.csv`: retention policies in force (source 5) |
| `collectors/Run-All.ps1` | Runs all five, signing in to each service in turn |
| `collectors/UnifiedAuditLogSchema.psd1` | Column order of every CSV and the per-cloud availability of every source |
| `collectors/UnifiedAuditLogHelpers.ps1` | Report-specific helpers (below) |
| `New-SampleData.ps1` | Regenerates `samples/` |
| `samples/` | Fake CSVs (everything is on `example.com`); `samples/gcchigh/` holds the header-only file a GCC High run leaves for the Graph Audit Search API |
| `tests/` | Pester 5 tests; every tenant call is mocked |

`UnifiedAuditLogHelpers.ps1` is report-specific because the shared layer has no audit-log logic: the availability
rule, the `ReturnLargeSet` session and its 50,000-record split, the Graph query create, poll and records paging, the
Management Activity subscription, listing and blob reads, 429 handling, and the mapping of a record to the common
columns. The shared connection (`Connect-M365Service`), CSV append (`Export-AppendCsv`), watermark
(`Get-CsvWatermark`) and range split (`Split-DateRange`) functions are used as they are. The one change outside this
folder is `-Formatted` added to the `Search-UnifiedAuditLog` test stub in `shared/tests/TenantCmdletStubs.ps1`, and
this folder added to `.github/workflows/tests.yml`.

## Before you run it

```powershell
Install-Module Microsoft.Graph.Authentication, ExchangeOnlineManagement -Scope CurrentUser
```

```powershell
./collectors/Run-All.ps1 -OutputPath ./out
./collectors/Run-All.ps1 -OutputPath ./out -Environment GCCHigh -AppId $appId `
    -CertificateThumbprint $thumbprint -TenantId $tenantId -Organization contoso.onmicrosoft.us `
    -PublisherIdentifier $publisherId -AccessToken $token
```

`-Environment` is `Commercial` (default), `GCC` or `GCCHigh`. Exchange Online connects with
`-ExchangeEnvironmentName O365USGovGCCHigh` in GCC High and with the default environment otherwise
([app-only authentication](https://learn.microsoft.com/powershell/exchange/app-only-auth-powershell-v2)).
Security & Compliance PowerShell uses the defaults in Commercial and GCC and, in GCC High,
`-ConnectionUri https://ps.compliance.protection.office365.us/powershell-liveid/ -AzureADAuthorizationEndpointUri https://login.microsoftonline.us/organizations`
([connect to SCC PowerShell](https://learn.microsoft.com/powershell/exchange/connect-to-scc-powershell)).
Graph uses `https://graph.microsoft.com` for Commercial and GCC and `Connect-MgGraph -Environment USGov` for GCC High
([national cloud deployments](https://learn.microsoft.com/graph/deployments)).

**The Management Activity API needs a token you supply.** The collector does not sign in to it. `-AccessToken` is a
SecureString holding a token for the feed root (Commercial `https://manage.office.com`, GCC `https://manage-gcc.office.com`,
GCC High `https://manage.office365.us`) for the tenant in `-TenantId`, carrying the `ActivityFeed.Read` claim.
`-PublisherIdentifier` is the GUID of the tenant of whoever runs the code, not the customer's tenant and not the app ID;
a request without it shares one quota with every caller. Without the three values `Run-All.ps1` skips source 3, writes
no file for it, and says so in `run.log`. The token is read from the SecureString only to build the `Authorization`
header and is never logged.

## Role and licence each collector needs

| Collector | Role or permission | Licence |
| --- | --- | --- |
| `Get-AuditSearchCmdlet.ps1` | Exchange Online role *View-Only Audit Logs* or *Audit Logs* (default: Compliance Management and Organization Management role groups) ([script article](https://learn.microsoft.com/purview/audit-log-search-script)) | Audit (Standard) covers the cmdlet; Audit (Premium) adds retention and some events ([overview](https://learn.microsoft.com/purview/audit-solutions-overview)) |
| `Get-AuditGraphRecords.ps1` | `AuditLogsQuery.Read.All` (all workloads) to create a query and list records; per-workload alternatives are `AuditLogsQuery-Entra.Read.All`, `-Exchange.`, `-SharePoint.`, `-OneDrive.`, `-CRM.` and `-Endpoint.Read.All` (`-Entra` does not return Exchange, SharePoint or OneDrive records). Reading one query lists `ThreatIntelligence.Read.All` as the least privileged permission, so the collector requests it too. The Purview role is not stated on the API pages | None named on the API pages; the overview lists the API under Audit (Standard) |
| `Get-AuditActivityFeed.ps1` | Entra app with the application permission *Read activity data for an organization* (claim `ActivityFeed.Read`); `DLP.All` also needs *Read DLP sensitive data* and is not read by default. No Purview or Exchange role is named | Audit (Standard) and Audit (Premium) both include access; Premium has higher bandwidth |
| `Get-AuditIngestion.ps1` | *View-Only Audit Logs* or *Audit Logs* in Exchange Online ([turn auditing on or off](https://learn.microsoft.com/purview/audit-log-enable-disable)) | None named |
| `Get-AuditRetentionPolicies.ps1` | The cmdlet page points to the Defender and Purview permission pages and names no role. Organization Configuration creates or changes a policy; that is not the read role | Audit (Premium) |

## Availability per cloud

Copied from the contract. `UNVERIFIED` means no Microsoft page read says either way; the collector attempts the source and
writes a warning to `run.log`. Only `NotAvailable` skips: the CSV gets its header only and the reason is logged.

| # | Source | Commercial | GCC | GCC High |
| --- | --- | --- | --- | --- |
| 1 | `Search-UnifiedAuditLog` | [Available](https://learn.microsoft.com/purview/audit-solutions-overview) | UNVERIFIED (no page names the cmdlet in GCC; Audit (Standard) is [Available](https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-deployments) in GCC) | UNVERIFIED (Audit (Standard) is [Available](https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-high-deployments) in GCC High; no page names the cmdlet) |
| 2 | Graph Audit Search API | [Available](https://learn.microsoft.com/graph/api/security-auditcoreroot-list-auditlogqueries) | [Available](https://learn.microsoft.com/graph/deployments) (global endpoint; the API pages do not name GCC) | [NotAvailable](https://learn.microsoft.com/graph/api/security-auditcoreroot-list-auditlogqueries) (US Government L4 is marked unsupported on the list, get, create and records pages) |
| 3 | Management Activity API | [Available](https://learn.microsoft.com/office/office-365-management-api/office-365-management-activity-api-reference#activity-api-operations) | [Available](https://learn.microsoft.com/office/office-365-management-api/office-365-management-activity-api-reference#activity-api-operations) | [Available](https://learn.microsoft.com/office/office-365-management-api/office-365-management-activity-api-reference#activity-api-operations) |
| 4 | `Get-AdminAuditLogConfig` | [Available](https://learn.microsoft.com/purview/audit-log-search-script) | UNVERIFIED | UNVERIFIED |
| 5 | `Get-UnifiedAuditLogRetentionPolicy` | [Available](https://learn.microsoft.com/powershell/module/exchangepowershell/get-unifiedauditlogretentionpolicy) | UNVERIFIED (the [GCC planning page](https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-deployments) lists Audit (Premium) log retention as Available; it does not name the cmdlet) | UNVERIFIED (the [GCC High planning page](https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-high-deployments) lists Audit (Premium) log retention as Available; it does not name the cmdlet) |

**Which read path to use.** Source 3 is the only one with a Learn page for all three clouds, and it is a 7-day feed.
Retained history comes from source 1, or from source 2 outside GCC High. Learn recommends the Management Activity API over
`Search-UnifiedAuditLog` for scripted, regular retrieval
([script article](https://learn.microsoft.com/purview/audit-log-search-script)).

## How each event source pages and resumes

| Source | Window | Paging | Resumes from | Limit handled |
| --- | --- | --- | --- | --- |
| 1 `Search-UnifiedAuditLog` | `-SliceMinutes` slices (60 by default), oldest first | One `SessionId`, `-SessionCommand ReturnLargeSet`, `-ResultSize 5000`, repeated until a call returns nothing or `AuditSearchRequestMetadata.moreRecordsAvailable` is false. The first empty call is retried three times | Latest `CreationTime` in the file, else `-LookbackDays` (7, up to 3653) back | A session whose `ResultCount` is 50,000, or that returns 50,000 records, is not the full slice: it is dropped and read as two halves. At 2 minutes the run stops and those rows are not written. `-HighCompleteness` is preview and is not used; without it results can be missing. `-StartDate` and `-EndDate` are UTC with a time |
| 2 Graph Audit Search | `-SliceMinutes` (1440) per query, oldest first | `GET .../queries/{id}` until `succeeded`, then `GET .../records` through `@odata.nextLink` | Latest `CreationTime` in the file, else `-LookbackDays` (7, up to 3653) back | `isRecordCountLimitExceeded` is read (the query still reports `succeeded` when over the limit) and the range is halved while it is longer than 60 minutes. At that floor the run stops and those rows are not written. A ten-year first run is many queries; a tenant allows at least 200 submissions per rolling 24 hours. `429` waits `Retry-After` (at least one second), else 30, 60, 120... seconds, up to 5 attempts, never at once |
| 3 Management Activity | 24-hour windows of `contentCreated`, oldest first, per content type | `NextPageUri` response header until it is absent; then `GET` each `contentUri` | Latest `ContentCreated` in the file, else `-LookbackDays` (7, the most the feed holds) back | `startTime` is clamped to 7 days (AF20030) and the gap is logged. A window is written only when every content type in it was read. A blob past `contentExpiration` is skipped and logged. `429` as above |

Two things to know about event timing. Audit events can take hours to become searchable
([turn auditing on or off](https://learn.microsoft.com/purview/audit-log-enable-disable)), so a record that appears after a
later one was exported is not picked up by a run that has moved past it. And source 3 selects on when a blob became
available, not on the event time, so `CreationTime` in `audit-activity-feed.csv` is not in order.

## CSVs

Common columns (sources 1 to 3), from the audit record's common schema
([Management Activity API schema](https://learn.microsoft.com/office/office-365-management-api/office-365-management-activity-api-schema)).
Workload-specific properties are not flattened; they stay in `AuditData`.

| Column | Meaning |
| --- | --- |
| `CreationTime` | When the event happened, UTC (`yyyy-MM-ddTHH:mm:ssZ`) |
| `RecordId` | The record's `Id`. The row key: a record already in the file is skipped |
| `RecordType` | Source 1: the name the cmdlet returns (for example `SharePointFileOperation`). Source 2: the Graph `auditLogRecordType` (for example `sharePointFileOperation`). Source 3: the number (for example `6`). The three are not the same text |
| `Operation` | The operation, for example `FileAccessed`, `Add member to role.`, `New-TransportRule` |
| `Workload` | The service. Source 2 reads the record's `service` property |
| `UserId` | The user that did it (source 2: the user principal name, else `userId`) |
| `ObjectId` | The object acted on: a file path, a mailbox, a rule name |
| `ClientIP` | The device's IP address |
| `ResultStatus` | The record's `ResultStatus` when it has one. Source 2 reads it from `auditData` |
| `OrganizationId` | The tenant GUID |
| `AuditData` | The record's JSON on one line, as returned. Source 3 keeps each event's own text |

| CSV | Extra columns | Key | Resumes on |
| --- | --- | --- | --- |
| `audit-search-cmdlet.csv` | none | `RecordId` | `CreationTime` |
| `audit-graph-records.csv` | `QueryId`: the `auditLogQuery` that returned the row (before `AuditData`) | `RecordId` | `CreationTime` |
| `audit-activity-feed.csv` | `ContentCreated`, `ContentType`, `ContentId`: the blob the event came in (first three columns) | `RecordId` | `ContentCreated` |
| `audit-ingestion.csv` | `RunDate`, `UnifiedAuditLogIngestionEnabled`. In Security & Compliance PowerShell the property is always `False` even when search is on, so this reads Exchange Online | `RunDate` | one row per run |
| `audit-retention-policies.csv` | `RunDate`, `Priority`, `Name`, `RecordTypes`, `Operations`, `UserIds`, `RetentionDuration`; lists are joined with `;`. `RetentionDuration` is written as returned: the portal offers durations (7 Days, 30 Days, 3, 5 and 7 Years) that are not among the five names the cmdlet page lists, and none is dropped | `RunDate`, `Priority`, `Name` | one snapshot per run |

`audit-retention-policies.csv` having no rows is not "no one-year retention": the cmdlet does not return the default policy.
Per-user licence for the one-year rule comes from the licence collectors in
[`license-utilization.md`](../../docs/candidates/license-utilization.md); the join is the library's logic, not a Microsoft report.
A search that returns nothing within 60 minutes of turning auditing on, or within hours of it, is not proof of no activity.

## Overlap with other reports

This report does not rebuild the mailbox exfiltration risk or oversharing collectors.

* [`mailbox-exfiltration-risk`](../mailbox-exfiltration-risk/) already reads mailbox audit and forwarding signals and scores
  them. This report would only show the raw non-owner access and permission-change events (`FolderBind`, `MessageBind`,
  `MailItemsAccessed`, `Add-MailboxPermission`) from `AuditData`; "not the owner" is `Logon_type` other than Owner, with
  `MailboxUPN` and `User`. `MailItemsAccessed` is logged for the owner as well, and its absence below an E3 licence is not
  proof of no access.
* [`oversharing`](../oversharing/) reads current sharing state. This report reads the history of sharing events
  (`SharingSet`, `AnonymousLinkCreated`, ...) and does not rebuild the state snapshot.

## Not built

The contract's "dropped" list is empty: no starting question was dropped. Starting question 5 (how far back a run can read per
licence level) is answered in part by the two state CSVs plus the licence collectors, as the contract notes. The seven report
pages in the contract belong to the later Power BI project.
