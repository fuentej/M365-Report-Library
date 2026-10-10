# Candidate: unified audit log

Status: sources verified against Microsoft Learn; nothing built. Written for BRO-379 so the build issue can
be written without guessing. Nothing here was run against a tenant. Pages read 2026-10-10.

Cloud names follow the library: `Commercial`, `GCC`, `GCCHigh` (`-Environment`). Graph uses
`https://graph.microsoft.com` for Commercial and GCC and `Connect-MgGraph -Environment USGov` for GCC High
([national cloud deployments](https://learn.microsoft.com/graph/deployments)). Security & Compliance and
Exchange Online PowerShell connect as in the library rules
([SCC](https://learn.microsoft.com/powershell/exchange/connect-to-scc-powershell),
[Exchange](https://learn.microsoft.com/powershell/exchange/app-only-auth-powershell-v2)).

## How to read the table

* `Available` / `NotAvailable` appear only where a Microsoft page says so; the link is in the cell.
* `UNVERIFIED` means no Microsoft page found says either way for that source in that cloud. The collector
  should ask for the data and record a refusal in `run.log`, as the existing reports do.
* **The headline finding:** there are three ways to read audit records, and they split by cloud. The Graph
  Audit Search API is marked `❌` for US Government L4 on every page read, so it is `NotAvailable` in GCC High.
  The Office 365 Management Activity API has a documented root URL for GCC and for GCC High, so it is the one
  read path with a Learn page for all three clouds. It lists content blobs from the last 7 days only, by when
  the blob was published, so it is not the retained-history search. `Search-UnifiedAuditLog` is documented, but no page read
  names GCC or GCC High for the cmdlet itself (the Purview planning pages list Audit (Standard) and Audit
  (Premium) features as available in both, which is the feature, not the cmdlet).
* Learn itself recommends the Management Activity API over `Search-UnifiedAuditLog` for scripted, regular
  retrieval ([script article](https://learn.microsoft.com/purview/audit-log-search-script),
  [cmdlet page](https://learn.microsoft.com/powershell/module/exchangepowershell/search-unifiedauditlog)).
* **Read-only caveat for the build issue.** Source 2 needs `POST /security/auditLog/queries` to create a query
  object before records can be listed, and source 3 needs a `POST .../subscriptions/start` once per content
  type. Neither changes tenant data, but neither is a pure GET. The library rule is "read-only against the
  tenant", so the build issue has to decide whether these two count. Source 1 has no such step.

## Sources

| # | Source | Endpoint or cmdlet | Least privileged role or permission | License | Event / state | Retention or limit | Commercial | GCC | GCC High |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 1 | Unified audit log through Exchange Online PowerShell | [`Search-UnifiedAuditLog`](https://learn.microsoft.com/powershell/module/exchangepowershell/search-unifiedauditlog) `-StartDate -EndDate [-RecordType] [-Operations] [-UserIds] [-ObjectIds] [-SiteIds] [-FreeText] -SessionId -SessionCommand ReturnLargeSet -ResultSize 5000`. Records carry `AuditData` JSON. Page by re-running with the same `SessionId` and the same `SessionCommand` until `AuditSearchRequestMetadata.moreRecordsAvailable` is false or nothing returns. Do not stop a session because `ResultIndex` equals `ResultCount` while `moreRecordsAvailable` is still true. Use `-Formatted` for readable operation names | Exchange Online role *View-Only Audit Logs* or *Audit Logs* (default: Compliance Management and Organization Management role groups) ([script article](https://learn.microsoft.com/purview/audit-log-search-script)). The Purview portal side uses the same roles through the *Audit Reader* and *Audit Manager* role groups ([get started](https://learn.microsoft.com/purview/audit-get-started)) | Audit (Standard) covers the cmdlet; Audit (Premium) adds retention and some events ([overview](https://learn.microsoft.com/purview/audit-solutions-overview)) | Event | Default 100 records. `-ResultSize` maximum 5,000. `ReturnLargeSet` pages up to 50,000 unsorted; `ReturnNextPreviewPage` is sorted but capped at 5,000. Do not switch commands within one `SessionId` (limit drops to 10,000). A session that reaches 50,000 is not the full window and must not be treated as complete: split the range ([cmdlet page](https://learn.microsoft.com/powershell/module/exchangepowershell/search-unifiedauditlog), [script article](https://learn.microsoft.com/purview/audit-log-search-script)). The script article slices the range into 60-minute intervals and lowers it on `maximum results limitation reached`. `-HighCompleteness` is preview and is not in every tenant; without it the cmdlet page says results can be missing. `-StartDate` and `-EndDate` are stored and interpreted as UTC: a value with no time zone is UTC, a date with no time is midnight UTC, and the same date for both with no time returns nothing. Not usable in Office 365 operated by 21Vianet (returns nothing) | [Available](https://learn.microsoft.com/purview/audit-solutions-overview) | UNVERIFIED (no page read names GCC for the cmdlet; Audit (Standard) is [Available](https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-deployments) in GCC) | UNVERIFIED (no page read names GCC High for the cmdlet; Audit (Standard) is [Available](https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-high-deployments) in GCC High) |
| 2 | Unified audit log through the Graph Audit Search API | [`POST /security/auditLog/queries`](https://learn.microsoft.com/graph/api/security-auditcoreroot-post-auditlogqueries) (`Get-MgSecurityAuditLogQuery` reads), then [`GET /security/auditLog/queries/{id}`](https://learn.microsoft.com/graph/api/security-auditlogquery-get) for `status` (`notStarted`, `running`, `succeeded`, `failed`, `cancelled`), then [`GET /security/auditLog/queries/{id}/records`](https://learn.microsoft.com/graph/api/security-auditlogquery-list-records). Body filters: `serviceFilter` (the workload), `operationFilters`, `userPrincipalNameFilters`, `ipAddressFilters`, `objectIdFilters`, `administrativeUnitIdFilters`, `keywordFilter`; date range properties are on the [auditLogQuery resource](https://learn.microsoft.com/graph/api/resources/security-auditlogquery) | Create and list records: `AuditLogsQuery-Entra.Read.All` is least privileged; higher are per-workload `AuditLogsQuery-Exchange.Read.All`, `-SharePoint.Read.All`, `-OneDrive.Read.All`, `-CRM.Read.All`, `-Endpoint.Read.All` and `AuditLogsQuery.Read.All`, delegated and application. The Get page lists `ThreatIntelligence.Read.All` as least privileged for reading a query; the build issue should request the create and records set and confirm the Get permission when it builds. Purview role needed is not stated on these pages; [get started](https://learn.microsoft.com/purview/audit-get-started) says the API "requires additional permissions to be configured in Microsoft Graph" | None named on the API pages. Overview lists "Audit Search Graph API" under Audit (Standard) ([overview](https://learn.microsoft.com/purview/audit-solutions-overview)) | Event | Tenant limits: at least 200 submissions per rolling 24 hours, at least 50 queued or running queries, at least 1,000,000 records per query; over the record limit the query still reports `succeeded`, so check `isRecordCountLimitExceeded`. Throttled submissions return `429` with `Retry-After` ([throttling limits](https://learn.microsoft.com/graph/throttling-limits#security-audit-log-query-service-limits)). Paging of `/records` follows `@odata.nextLink` (not separately confirmed on the page read) | [Available](https://learn.microsoft.com/graph/api/security-auditcoreroot-list-auditlogqueries) (Global service ✅) | [Available](https://learn.microsoft.com/graph/deployments) (global endpoint; the API pages do not name GCC) | [NotAvailable](https://learn.microsoft.com/graph/api/security-auditcoreroot-list-auditlogqueries) (US Government L4 `❌` on list, get, create and records pages) |
| 3 | Unified audit log through the Office 365 Management Activity API | Per tenant: `{root}/api/v1.0/{tenant_id}/activity/feed/subscriptions/start?contentType=Audit.<Workload>`, `.../subscriptions/content?contentType=...&startTime=...&endTime=...`, then GET each `contentUri` blob. Content types: `Audit.AzureActiveDirectory`, `Audit.Exchange`, `Audit.SharePoint`, `Audit.General` (every other workload), and `DLP.All` (DLP events only; needs the extra permission *Read DLP sensitive data*). Pass `PublisherIdentifier` (the vendor tenant GUID, not the customer tenant or the app id) on every request; requests without it share one quota. A second start within 15 minutes is throttled. The first blobs can take up to 12 hours. Stopping a subscription and starting it again does not return content from the gap. Roots: Commercial `https://manage.office.com`, GCC `https://manage-gcc.office.com`, GCC High `https://manage.office365.us` ([API reference](https://learn.microsoft.com/office/office-365-management-api/office-365-management-activity-api-reference#activity-api-operations)). A blob holds 1 to N events, so counts of blobs are not counts of events ([FAQ](https://learn.microsoft.com/office/office-365-management-api/troubleshooting-the-office-365-management-activity-api)) | Entra app with application permission *Read activity data for your organization*; the token needs the `ActivityFeed.Read` claim and its tenant ID must match the URL. The page read does not name a Purview or Exchange role | Audit (Standard) and Audit (Premium) both include access; Premium has higher bandwidth ([overview](https://learn.microsoft.com/purview/audit-solutions-overview)) | Event | Listing content: a range of at most 24 hours; `startTime` and `endTime` are used together or not at all; results are paged through the `NextPageUri` response header until it is absent ([FAQ](https://learn.microsoft.com/office/office-365-management-api/troubleshooting-the-office-365-management-activity-api), [API reference](https://learn.microsoft.com/office/office-365-management-api/office-365-management-activity-api-reference#list-available-content)). Baseline 2,000 requests per minute per tenant, about twice that for E5, A5 and G5 ([overview](https://learn.microsoft.com/purview/audit-solutions-overview)). `startTime` and `endTime` select blobs by `contentCreated` (when the blob became available), not by the time of the event: `startTime` is inclusive and `endTime` is exclusive. `startTime` is no more than 7 days in the past (error `AF20030`); content older than 7 days cannot be retrieved (`AF20051`). Each blob includes `contentExpiration`. This API is a 7-day feed, not the 180-day or one-year search | [Available](https://learn.microsoft.com/office/office-365-management-api/office-365-management-activity-api-reference#activity-api-operations) | [Available](https://learn.microsoft.com/office/office-365-management-api/office-365-management-activity-api-reference#activity-api-operations) | [Available](https://learn.microsoft.com/office/office-365-management-api/office-365-management-activity-api-reference#activity-api-operations) |
| 4 | Whether audit ingestion is on | `Get-AdminAuditLogConfig \| Format-List UnifiedAuditLogIngestionEnabled`, run in **Exchange Online** PowerShell. In Security & Compliance PowerShell the property is always `False` even when search is on ([search the audit log](https://learn.microsoft.com/purview/audit-search)) | Exchange Online role *Audit Logs* turns auditing on or off; this page does not say which role only reads the setting. Use the same read role as source 1 and confirm in the build | None named | State | n/a | [Available](https://learn.microsoft.com/purview/audit-log-search-script) (the page tells admins to run it) | UNVERIFIED | UNVERIFIED |
| 5 | Audit log retention policies in force (decides how far back a run can read, per record type, operation and user) | [`Get-UnifiedAuditLogRetentionPolicy`](https://learn.microsoft.com/powershell/module/exchangepowershell/get-unifiedauditlogretentionpolicy), Security & Compliance PowerShell only. Properties: `Priority`, `Name`, `RecordTypes`, `Operations`, `UserIds`, `RetentionDuration`. The cmdlet accepted values are `ThreeMonths`, `SixMonths`, `NineMonths`, `TwelveMonths`, `TenYears`. The portal duration list also includes 7 Days, 30 Days, 3 Years, 5 Years, and 7 Years, which are not in that accepted list; 7 Days and 30 Days need an E5 subscription, and 3, 5, and 7 years need the 10-year add-on in addition to E5. Do not drop a returned duration that is not one of the five cmdlet names. The cmdlet does not return the default audit log retention policy ([retention policies](https://learn.microsoft.com/purview/audit-log-retention-policies)) | The cmdlet page says permissions are needed and points to the Defender and Purview permission pages; the role is not named on the page | Retention policies are an Audit (Premium) capability ([overview](https://learn.microsoft.com/purview/audit-solutions-overview)) | State | See "Retention" below | [Available](https://learn.microsoft.com/powershell/module/exchangepowershell/get-unifiedauditlogretentionpolicy) | UNVERIFIED (the [GCC planning page](https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-deployments) lists Audit (Premium) log retention as Available; it does not name the cmdlet) | UNVERIFIED (the [GCC High planning page](https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-high-deployments) lists Audit (Premium) log retention as Available; it does not name the cmdlet) |

### Retention (starting question 5)

From the [retention policies](https://learn.microsoft.com/purview/audit-log-retention-policies) page and the
[auditing solutions overview](https://learn.microsoft.com/purview/audit-solutions-overview):

* Audit (Standard): 180 days for records generated on or after 17 October 2023. Records generated before that
  date were kept 90 days.
* One year only for Exchange, SharePoint, OneDrive, and Microsoft Entra records of users with Office 365 or
  Microsoft 365 E5, or the Microsoft Purview Suite (formerly Microsoft 365 E5 Compliance), or the E5 eDiscovery
  and Audit add-on. Non-E5 users and guest users stay at 180 days. Other workloads stay at 180 days.
  A custom retention policy overrides the default and can be shorter.
* Ten-year retention needs the 10-Year Audit Log Retention add-on in addition to an E5 licence, and the policy
  is not retroactive. Records from non-user entities (service principals, system events, application activity)
  are kept a fixed one year, and custom policies do not apply to them.
* The audit item lifetime is set when the record is added to the pipeline. Later licence or policy changes
  "change the expiration time of the audit data after updating" and "don't affect any previously committed items".
* `Get-UnifiedAuditLogRetentionPolicy` does not return the default policy. An empty result is not "no one-year
  retention". Per-user licence comes from the licence collectors in
  [license-utilization.md](license-utilization.md); the join is the library's logic, not a Microsoft report.

## Event identifiers the report would filter on

Operation names below are copied from [audit log activities](https://learn.microsoft.com/purview/audit-log-activities)
and the cmdlet page. The page lists many more; the build issue should choose the final list.

| Question | Operations seen on the pages read |
| --- | --- |
| Admin role changes | `Add member to role.`, `Remove member from role.` (Microsoft Entra role administration activities) |
| Policy changes | Exchange transport rule operations `New-TransportRule`, `Set-TransportRule`, `Disable-TransportRule`, `Enable-TransportRule`, `Remove-TransportRule` ([transport rule activities](https://learn.microsoft.com/purview/audit-log-search-mailbox-rules)). Other policy families are on the activities page; not enumerated here |
| Files and sharing | `FileAccessed` with record type `SharePointFileOperation` (cmdlet example 4); `SharingSet`, `SharingRevoked`, `SharingInvitationCreated`, `AnonymousLinkCreated`, `CompanyLinkCreated`, `SecureLinkCreated`, `AccessRequestCreated` |
| Mailbox access by someone other than the owner | `FolderBind` (admin and delegate; delegate binds are consolidated to one record per folder per 24 hours), `MessageBind` (admin only, and only users without E5, A5 or G5 licences), `MailItemsAccessed` (Audit (Standard), on by default for Office 365 E3/E5 or Microsoft 365 E3/E5; the `SensitivityLabel` property is Audit (Premium)), `SendAs`, `SendOnBehalf` ([investigate accounts](https://learn.microsoft.com/purview/audit-log-investigate-accounts), [overview](https://learn.microsoft.com/purview/audit-solutions-overview#audit-premium-activity-properties), [mailbox auditing](https://learn.microsoft.com/purview/audit-mailboxes)) |
| Mailbox permission changes | `Add-MailboxPermission` and `Remove-MailboxPermission` (FullAccess; `Add-MailboxPermission` is also written by a system account doing maintenance on the DiscoverySearchMailbox, so filter those out). Folder permission changes that are audited are `UpdateFolderPermissions`. `AddFolderPermissions` and `ModifyFolderPermissions` are listed on the activities page, but the mailbox auditing page says they are not audited separately and not to use those values. Also `UpdateCalendarDelegation` and `Set-Mailbox` ([activities](https://learn.microsoft.com/purview/audit-log-activities), [mailbox auditing](https://learn.microsoft.com/purview/audit-mailboxes)) |

How an event is judged "not the owner" uses fields inside `AuditData`. `Logon_type` is Owner (0), Admin (1), or
Delegate (2). `MailboxUPN` is the mailbox that holds the message and `User` is the UPN of the reader
([investigate accounts](https://learn.microsoft.com/purview/audit-log-investigate-accounts)). `MailItemsAccessed`
is logged for the owner as well as for admins and delegates, so the operation name alone is not "not the owner".
Absence of `MailItemsAccessed` on a licence below E3 is not proof of no access. The `SensitivityLabel` property
on that record is Audit (Premium), and a missing Premium licence is not the same as zero access events.

## Proposed report pages

| # | Page | Reads | What it shows |
| --- | --- | --- | --- |
| 1 | Overview | 1 or 3 (2 outside GCC High), 4 | Records per day by workload and by user; coverage banner: audit ingestion on or off, earliest record held, sources skipped in this cloud |
| 2 | Operations by workload and user | 1 or 3 (2 outside GCC High) | Operation counts by `Workload` (record type), user and day; top operations; drill to the record |
| 3 | Admin role and policy changes | 1 or 3 (2 outside GCC High) | `Add member to role.` and `Remove member from role.`, transport-rule operations: who, what, when |
| 4 | File and sharing activity | 1 or 3 (2 outside GCC High) | `FileAccessed` and the sharing and link operations in SharePoint and OneDrive, by user, site and day |
| 5 | Mailbox access and permission changes | 1 or 3 (2 outside GCC High) | Non-owner `FolderBind`, `MessageBind`, `MailItemsAccessed`; `Add-MailboxPermission` and folder or calendar permission changes; system-account noise excluded |
| 6 | Retention and lookback | 4, 5, licence data | Retention policies in force, the lookback each licence level gets, and the oldest record actually returned |
| 7 | Collection health | `run.log`, 1 to 3 | Which read path ran, which were skipped in this cloud and why, queries throttled, any 50,000 or record-count limit hit, any content blob past `contentExpiration` |

Source 3 cannot fill pages 1 to 5 beyond seven days. Those pages use source 1, or source 2 outside GCC High, for retained history. Source 3 is the feed for the last seven days in every cloud.

## Starting questions

| # | Question | Status |
| --- | --- | --- |
| 1 | Which admin and user operations happen, by workload, user and day | Covered by sources 1, 2 and 3 (workload is the record type or `serviceFilter`; user and time are record fields). Source 3 is the only read path with a Learn page for all three clouds, and it lists only the last 7 days of content blobs. |
| 2 | Which admin role assignments and policy changes happened, by whom and when | Covered by the operations in the table above through sources 1 to 3. Role changes: `Add member to role.` and `Remove member from role.`. Policy changes: only the transport rule family was verified; the build issue picks the rest from the activities page. |
| 3 | Which file and sharing operations happened in SharePoint and OneDrive | Covered by sources 1 to 3 with the file and sharing operations above. |
| 4 | Which mailboxes were accessed by someone other than the owner, and which mailbox permissions changed | Covered by sources 1 to 3 with the mailbox operations above. "Not the owner" is `Logon_type` other than Owner, with `MailboxUPN` and `User`. `MailItemsAccessed` is Audit (Standard) for Office 365 E3/E5 or Microsoft 365 E3/E5, so absence of it below E3 is not proof of no access. |
| 5 | How long records are kept per licence level and how far back a run can read | Covered by the Retention section, source 4 (is auditing on) and source 5 (custom policies). One year is E5, the Microsoft Purview Suite, or the E5 eDiscovery and Audit add-on; guest users stay at 180 days. Source 5 does not return the default policy. GCC and GCC High for source 5 UNVERIFIED. |
| 6 | Which read-only method retrieves the records, with limits, paging and availability per cloud | Covered by sources 1 to 3, limits in each row. Source 2 is `NotAvailable` in GCC High. Source 1 is UNVERIFIED in GCC and GCC High. Source 3 is Available in all three and only the last 7 days of content. |

No starting question is dropped. Question 5 is only partly answerable from Microsoft sources, for the reason
given.

## Overlap with other candidates

* Mailbox exfiltration risk (BRO-317, BRO-323, BRO-328): that report already reads mailbox audit and forwarding
  signals; this report would only show raw non-owner access and permission-change events and does not rebuild
  its risk scoring.
* Oversharing (BRO-359, BRO-370): that report reads current sharing state; this report reads the history of
  sharing events from the audit log and does not rebuild the state snapshot.

## Open items for the build issue

* Read path: source 3 is the only one with a Learn page for all three clouds, so a single code path for every
  cloud would use it. Decide whether source 1 or 2 is also wanted where available, and whether the two `POST`
  calls (query create, subscription start) are acceptable under "read-only".
* Source 1 in GCC and GCC High: UNVERIFIED. A collector may attempt it and record a refusal in `run.log`.
* Source 3 is a 7-day feed. Retained history still needs source 1, or source 2 outside GCC High.
* Source 2 least privileged permission differs between the create and records pages and the get page; confirm.
* Roles for source 4 and source 5 (read only) are not named on the pages read.
