# Microsoft 365 Copilot usage

Who holds a Copilot licence and uses it, in which apps, how the number of active users
changes over time, and which Copilot interactions the audit log and the interaction export
record. Built from the source contract in
[`docs/candidates/copilot-usage.md`](../../docs/candidates/copilot-usage.md) (BRO-384); the
Power BI project is a separate, later issue and is not here.

The collectors write one CSV per source into an output folder of your choosing, plus a
`run.log`. Event files are appended from the last exported timestamp; the usage-report and
reference files get a block of rows stamped with the run date. Nothing is rewritten or
deleted. `samples/` holds fake data (`example.com`) with the exact columns each collector
writes; `samples/gcchigh/` holds the header-only files the usage reports leave in GCC High.

Nothing here has been run against a tenant. Every tenant call is mocked in `tests/`.

## Contents

| Path | What it is |
| --- | --- |
| `collectors/Get-CopilotUsageUserDetail.ps1` | `copilot-usage-user-detail.csv` — source 2, last activity per Copilot app, prompts and active days per licensed user |
| `collectors/Get-CopilotUserCountSummary.ps1` | `copilot-user-count-summary.csv` — source 3, enabled and active users per app, one row per period |
| `collectors/Get-CopilotUserCountTrend.ps1` | `copilot-user-count-trend.csv` — source 4, the daily trend of those counts |
| `collectors/Get-CopilotAuditEvents.ps1` | `copilot-audit-events.csv` — source 6, one row per `CopilotInteraction` audit record |
| `collectors/Get-CopilotInteractions.ps1` | `copilot-interactions.csv` — source 8, one metadata row per prompt or response (no text) |
| `collectors/Get-CopilotFeatureAvailability.ps1` | `copilot-feature-availability.csv` — source 10, which Copilot features each cloud offers (no tenant call) |
| `collectors/Run-All.ps1` | Runs the shared users collector and all six of the above |
| `collectors/CopilotUsageSchema.psd1` | The column order of every CSV, the Learn header each report column is read from, and the availability of every source per cloud |
| `collectors/CopilotUsageHelpers.ps1` | The report-specific helper: availability, throttle retry, the one GET, audit-log paging |
| `New-SampleData.ps1` | Regenerates `samples/` |
| `tests/` | Pester tests; every tenant call is mocked |

### Sources this report does not collect

| Source | Why |
| --- | --- |
| 1 Who holds a Copilot licence | Not re-specified in the contract: it is [license utilization](../../docs/candidates/license-utilization.md) sources 1 and 2 (BRO-364; collectors BRO-374, **not yet built** when this was written). This report reuses them and does not rebuild SKU inventory or user licence assignment. |
| 9 Concealed names (`displayConcealedNames`) | Same: license utilization source 6 (`GET /admin/reportSettings`). See "Concealed names" below. |
| 5 Admin center fallback for GCC High | The same numbers exist in the Microsoft 365 admin center (**Reports > Usage > Microsoft Copilot** and **Copilot Chat**, CSV export), but the contract finds no page saying a script can read them in GCC High. A manual export is an open item, not a collector. |
| 7 Copilot interactions through the Graph Audit Search API | Needs a `POST` to create the query object. This library is GET-only (`tests/ReadOnly.Tests.ps1` and the library-wide read-only test enforce it) and the read-only question is already open in the unified audit log candidate. Source 6 returns the same `CopilotInteraction` records through a read-only cmdlet; source 7 is also `NotAvailable` in GCC High. |

Overlaps: Copilot Studio agent inventory and agent authoring audit events belong to the
[Copilot Studio agents report](../copilot-studio-agents/README.md); this report only reads
the `AgentId` and `AgentName` that appear on a `CopilotInteraction` record. Audit read paths,
paging and retention are owned by the unified audit log candidate.

## Before you run it

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

`Invoke-MgGraphRequest` is the one Graph call (a GET); the Copilot report and interaction
endpoints have no cmdlet of their own.

### Role and licence each collector needs

| Collector | Permission | Role | Licence |
| --- | --- | --- | --- |
| `Get-CopilotUsageUserDetail.ps1` | `Reports.Read.All` (delegated or application) | Company Administrator, AI Administrator, Exchange Administrator, SharePoint Administrator, Lync Administrator, Teams Service Administrator, Teams Communications Administrator or Reports Reader. Global Reader and Usage Summary Reports Reader see no detailed metrics ([authorization](https://learn.microsoft.com/graph/reportroot-authorization)) | A Microsoft 365 Copilot licence is the population. Unlicensed Copilot Chat use is not in this API |
| `Get-CopilotUserCountSummary.ps1`, `Get-CopilotUserCountTrend.ps1` | `Reports.Read.All` | The roles above, plus Global Reader and Usage Summary Reports Reader | As above |
| `Get-CopilotAuditEvents.ps1` | None (Exchange Online) | View-Only Audit Logs or Audit Logs, with auditing turned on | [Audit (Standard)](https://learn.microsoft.com/purview/audit-copilot) covers Microsoft Copilot and Copilot Studio interactions |
| `Get-CopilotInteractions.ps1` | `AiEnterpriseInteraction.Read.All`, **application only** (delegated is not supported). An interactive sign-in does not request it | None; app-only (`-AppId` and `-CertificateThumbprint`) | A Microsoft 365 Copilot licence with the `Microsoft Copilot with Graph-grounded chat` service plan, per user |
| `Get-CopilotFeatureAvailability.ps1` | None | None | None; a public page |

`Get-CopilotInteractions.ps1` is the most sensitive source here: its permission lets the
caller read every user's prompts and responses. The collector keeps metadata only (see the
CSV list), and `Run-All.ps1 -SkipInteractions` leaves it out.

## Cloud availability

Every collector takes `-Environment Commercial|GCC|GCCHigh`, defaulting to `Commercial`.
Graph is `https://graph.microsoft.com` for Commercial and GCC and `Connect-MgGraph
-Environment USGov` for GCC High ([national cloud deployments](https://learn.microsoft.com/graph/deployments)).
Copied from the contract; GCC calls the global service.

| # | Source | Commercial | GCC | GCC High |
| --- | --- | --- | --- | --- |
| 1 | Copilot licence holders | [Available](https://learn.microsoft.com/graph/api/subscribedsku-list) | [Available](https://learn.microsoft.com/graph/deployments) | [Available](https://learn.microsoft.com/graph/api/subscribedsku-list) (license utilization) |
| 2 | Usage per user | [Available](https://learn.microsoft.com/microsoft-365-copilot/extensibility/api/admin-settings/reports/copilotreportroot-getmicrosoft365copilotusageuserdetail) | [Available](https://learn.microsoft.com/graph/deployments) | **NotAvailable** ([US Government L4 ❌](https://learn.microsoft.com/microsoft-365-copilot/extensibility/api/admin-settings/reports/copilotreportroot-getmicrosoft365copilotusageuserdetail)) |
| 3 | User count summary | [Available](https://learn.microsoft.com/microsoft-365-copilot/extensibility/api/admin-settings/reports/copilotreportroot-getmicrosoft365copilotusercountsummary) | [Available](https://learn.microsoft.com/graph/deployments) | **NotAvailable** ([US Government L4 ❌](https://learn.microsoft.com/microsoft-365-copilot/extensibility/api/admin-settings/reports/copilotreportroot-getmicrosoft365copilotusercountsummary)) |
| 4 | User count trend | [Available](https://learn.microsoft.com/microsoft-365-copilot/extensibility/api/admin-settings/reports/copilotreportroot-getmicrosoft365copilotusercounttrend) | [Available](https://learn.microsoft.com/graph/deployments) | **NotAvailable** ([US Government L4 ❌](https://learn.microsoft.com/microsoft-365-copilot/extensibility/api/admin-settings/reports/copilotreportroot-getmicrosoft365copilotusercounttrend)) |
| 5 | Admin center fallback | [Available](https://learn.microsoft.com/microsoft-365/admin/activity-reports/activity-reports#available-usage-reports-in-the-microsoft-365-admin-center) | [Available](https://learn.microsoft.com/microsoft-365/admin/activity-reports/activity-reports#available-usage-reports-in-the-microsoft-365-admin-center) | Partly: available in the admin center; scripted read **UNVERIFIED** |
| 6 | Copilot interactions in the unified audit log | [Available](https://learn.microsoft.com/purview/audit-copilot) | **UNVERIFIED** (the cmdlet is not named for GCC on any page read; the [service description](https://learn.microsoft.com/office365/servicedescriptions/office-365-platform-service-description/microsoft-365-copilot#feature-availability) lists Purview controls for Copilot as `Yes`, which is the feature, not the cmdlet) | **UNVERIFIED** (same) |
| 7 | Graph Audit Search API | [Available](https://learn.microsoft.com/graph/api/security-auditcoreroot-post-auditlogqueries) | [Available](https://learn.microsoft.com/graph/deployments) | **NotAvailable** (US Government L4 ❌) — not collected, see above |
| 8 | Interaction export | [Available](https://learn.microsoft.com/microsoft-365-copilot/extensibility/api/ai-services/interaction-export/aiinteractionhistory-getallenterpriseinteractions) | [Available](https://learn.microsoft.com/graph/deployments) | [Available](https://learn.microsoft.com/microsoft-365-copilot/extensibility/api/ai-services/interaction-export/aiinteractionhistory-getallenterpriseinteractions) (US Government L4 ✅) |
| 9 | Concealed-names setting | [Available](https://learn.microsoft.com/graph/api/adminreportsettings-get) | [Available](https://learn.microsoft.com/graph/deployments) | **UNVERIFIED** (the Graph page marks US Government L4 ❌) — license utilization |
| 10 | Copilot feature availability | [Available](https://learn.microsoft.com/office365/servicedescriptions/office-365-platform-service-description/microsoft-365-copilot#feature-availability) | [Available](https://learn.microsoft.com/office365/servicedescriptions/office-365-platform-service-description/microsoft-365-copilot#feature-availability) | [Available](https://learn.microsoft.com/office365/servicedescriptions/office-365-platform-service-description/microsoft-365-copilot#feature-availability) (with per-feature exceptions) |

A `NotAvailable` source writes its header only and logs why; an `UNVERIFIED` source is
attempted, warns, and logs any refusal. In GCC High this report is therefore a licence
inventory (license utilization) plus audit and interaction records, with no usage-report
numbers unless the admin center is used by hand.

Copilot itself is offered in all three clouds, but not feature for feature: Copilot in Teams
and Copilot in SharePoint are `Not currently available` in GCC High, so a GCC High tenant with
no Teams or SharePoint Copilot activity is expected, not a collection gap. Source 10 holds the
dated copy.

## Limits to know about

* **Counts come from the report APIs, not from the audit log.** Microsoft says audit-log
  counts "might not be consistent" with the usage reports
  ([page](https://learn.microsoft.com/microsoft-365/admin/activity-reports/microsoft-365-copilot-usage#whats-the-difference-between-the-user-activity-table-and-audit-log)).
  Do not set a count from `copilot-audit-events.csv` or `copilot-interactions.csv` beside one
  from the three usage files as the same measure.
* **A record is not a prompt.** A `CopilotInteraction` record typically holds a prompt and a
  response and can hold one prompt with several responses; `MessageCount` and
  `PromptMessageCount` say how many.
* **Periods are windows, not events.** The three usage files take `-Period` `D7`, `D28`, `D90`,
  `D180` or `ALL` (all four periods in one response, not lifetime). `D30` is a `v1` value and is
  refused on the default `v2`. A daily run appends a snapshot stamped with the run date.
* **Active means an intentional action.** Opening the Copilot pane does not count; submitting
  a prompt does. The all-up `Last Activity Date` is a Copilot Chat message date and stays fixed
  when the period changes; each per-app date is any intentional activity in that app regardless
  of the period, so it can be older than the period. The summary's `TotalPromptsSubmitted` and
  `AveragePromptsSubmitted` are Copilot Chat prompts; the user table's
  `PromptsSubmittedAllApps` is prompts across in-scope host applications
  ([FAQ](https://learn.microsoft.com/microsoft-365/admin/activity-reports/microsoft-365-copilot-usage#how-is-a-user-considered-active-in-microsoft-copilot-usage)).
* **Blank dates are not always inactivity.** A user can have no last activity date after using
  Copilot within 24 hours of the licence being assigned and never again, and client-side Office
  events can upload late or leave the date blank. The report validates the past three days daily
  and fills gaps, so history can change between runs.
* **Latency.** The usage page says the report typically becomes available within 48 hours of the
  end of the UTC day; the Copilot reports overview says 72. Both figures are kept.
* **The user table is everyone licensed at any point in the past 180 days** on the admin center;
  the API page says it returns users who have a Microsoft 365 Copilot licence.
* **A 429 is a throttle, not an empty history.** Every Graph read and the audit search
  wait `Retry-After` and retry; if throttling persists the report is logged as throttled,
  not as empty, and a search that stays throttled stops the run instead of writing a
  header and returning. The interaction export is limited to 1,500 requests per second
  per app and 30 per app per tenant
  ([limits](https://learn.microsoft.com/graph/throttling-limits#microsoft-teams-service-limits)).
* **The interaction resume bound includes the last exported second.** The file stores
  whole seconds, so the next read uses `ge` that second and skips rows already stored.
  A caller-supplied `-StartDate` stays `gt`, as the export page's range example does.
  `-StartDate` and `-EndDate` with no time zone are that UTC instant.
* **Audit records can arrive late.** A record ingested after a run, with a time before the
  newest record already exported, is not collected by the next one. A window that reaches the
  50,000-record session cap is not written; re-run it with a smaller `-WindowHours`.
* **Audit (Standard) keeps Copilot records for 180 days by default.** Copilot is not one of
  the workloads the one-year default policy covers. A custom retention policy can keep them
  for up to 10 years, and `-LookbackDays` accepts that span (3653 days). The default
  lookback stays 30 days.
* **The interaction export supports six `appClass` values** (Word, Excel, Teams, BizChat, WebChat,
  CoworkChat). Outlook, PowerPoint, OneNote and Loop are not in it, so it is not the per-app count
  of the usage reports. It does not retrieve Copilot Studio agent interactions.

### Concealed names

When the tenant setting `displayConcealedNames` is `true` (the default), user names, display
names, groups and sites in usage reports are hidden: the user detail report returns 32-character
hashed values for `UserPrincipalName` and `DisplayName`, which `copilot-usage-user-detail.csv`
keeps as returned. A hashed name cannot be joined to `users.csv` or to licence holders on
`UserPrincipalName`. The setting is read by `GET /admin/reportSettings`
([page](https://learn.microsoft.com/graph/api/adminreportsettings-get), permission
`ReportSettings.Read.All`), which belongs to the license utilization report, so read it there
before joining; a report page should say "names concealed, usage not joined" rather than show
zero usage. `samples/copilot-usage-user-detail.csv` shows readable names, as when the setting
is `false`.

## The CSVs

Timestamps are UTC. List columns are joined with `;`. Files are keyed so a re-run never
repeats a row.

### `copilot-usage-user-detail.csv`

Key `RunDate`, `UserPrincipalName`, `ReportPeriod`. Each column is the Learn header of the same
name (spaces removed); the version 2 columns are empty if the service returns a version 1 report.

| Column | Meaning |
| --- | --- |
| `RunDate` | UTC date of the run |
| `ReportRefreshDate`, `ReportPeriod` | When the report was refreshed, and the period in days |
| `UserPrincipalName`, `DisplayName` | The user (hashes when names are concealed) |
| `LastActivityDate` | All-up last activity; fixed across periods |
| `CopilotChat…`, `MicrosoftTeamsCopilot…`, `WordCopilot…`, `ExcelCopilot…`, `PowerPointCopilot…`, `OutlookCopilot…`, `OneNoteCopilot…`, `LoopCopilot…`, `CopilotChatWork…`, `CopilotChatWeb…`, `Microsoft365Copilot…`, `Edge…`, `CopilotAgent…` `LastActivityDate` | Last intentional activity in that app |
| `PromptsSubmittedAllApps`, `PromptsSubmittedCopilotChatWork`, `PromptsSubmittedCopilotChatWeb` | Prompts in the period |
| `ActiveUsageDaysAllApps` | Days with activity in the period |

### `copilot-user-count-summary.csv`

Key `RunDate`, `ReportPeriod`. One `…EnabledUsers` and one `…ActiveUsers` column per app: Microsoft
Teams, Word, PowerPoint, Outlook, Excel, OneNote, Loop, Any App, Copilot Chat, and (version 2) Edge,
Microsoft 365 Copilot, Copilot Chat (work) and Copilot Chat (web). Enabled users are the unique users
who held a Copilot licence over the period, not the licences assigned on the run date.
`TotalPromptsSubmitted` and `AveragePromptsSubmitted` are Copilot Chat prompts.

### `copilot-user-count-trend.csv`

Key `RunDate`, `ReportDate`, `ReportPeriod`. The same enabled and active columns as the summary,
one row per `ReportDate`, and `PromptsSubmitted` (the day's prompts, not the summary's total or
average). A day is re-read each run and can change.

### `copilot-audit-events.csv`

Key `Id`.

| Column | Meaning |
| --- | --- |
| `CreationTime`, `Id`, `UserId` | When, which record, and who |
| `Operation`, `RecordType`, `Workload` | `CopilotInteraction`; the record type (collected with `-Formatted`, so its name; without it the cmdlet returns the integer 261); `Copilot` |
| `AppHost`, `AppIdentity` | Where the interaction happened (`Teams`, `Word`, `BizChat`, …) and which Copilot app, from the record |
| `AgentId`, `AgentName` | Kept wherever the record places them |
| `MessageCount`, `PromptMessageCount` | Messages in the record, and how many are prompts (`isPrompt`) |
| `AccessedResourceCount` | Resources Copilot accessed |
| `PluginIds` | `AISystemPlugin.ID` values; `BingWebSearch` means Copilot used the public web |

### `copilot-interactions.csv`

Key `UserId`, `Id`. Metadata only: `body`, `attachments`, `links` and `mentions` are never read.

| Column | Meaning |
| --- | --- |
| `CreatedDateTime`, `Id`, `UserId` | When, which interaction, and whose history it came from |
| `SessionId`, `RequestId` | Session, and the id that pairs a `userPrompt` with its `aiResponse` |
| `AppClass`, `ConversationType`, `Locale` | Which Copilot app, kind of conversation, locale |
| `InteractionType` | `userPrompt` or `aiResponse` |
| `ContextCount`, `ContextTypes` | The contexts (for example `TeamsMeeting`) on the interaction |

### `copilot-feature-availability.csv`

Key `RunDate`, `Feature`. A dated copy of the service description's feature table, limited to what
the contract records. `Commercial`, `GCC` and `GCCHigh` are `Yes`, `NotCurrentlyAvailable`,
`Limited` or `NotStated` (the contract does not say; not "absent"). `PageReadDate` is the day the
page was read. To refresh the copy, re-read the page and edit `FeatureRows` in the schema.
