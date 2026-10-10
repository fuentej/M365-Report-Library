# Teams usage and activity

How Microsoft Teams is used: chat, call and meeting counts per user, the platforms people use,
tenant totals per day, calls and meetings with the platform of each endpoint, and a count of
Teams events from the unified audit log that the library makes itself.

This folder holds the collectors, fake sample CSVs and tests. **There is no Power BI project
yet**; it is a separate, later issue. The sources and their per-cloud availability were verified
in [`docs/candidates/teams-activity.md`](../../docs/candidates/teams-activity.md) (BRO-382); the
tables below are copied from it.

The collectors write one CSV per source into an output folder, plus a `run.log`. State sources
append a snapshot stamped with `RunDate`. `teams-audit-events.csv` appends from the day after the
latest date already in the file. `call-records.csv` re-reads the 30-day retention window and
appends only rows that are not already there, because a later version of a call can arrive after
its start time was stored. Nothing is rewritten or deleted. Everything is read-only against the tenant.

## Contents

| Path | What it is |
| --- | --- |
| `collectors/Get-TeamsUserActivityUserDetail.ps1` | `teams-user-activity-user-detail.csv`: chat, call and meeting counts per user (source 1) |
| `collectors/Get-TeamsUserActivityCounts.ps1` | `teams-user-activity-counts.csv`: tenant totals per day (source 2) |
| `collectors/Get-TeamsDeviceUsageUserDetail.ps1` | `teams-device-usage-user-detail.csv`: which platforms each user used Teams on (source 3) |
| `collectors/Get-ReportSettings.ps1` | `report-settings.csv`: whether usage reports conceal names (source 5) |
| `collectors/Get-CallRecords.ps1` | `call-records.csv`: calls and meetings of the last 30 days, one row per session, with the caller and callee platform (source 6) |
| `collectors/Get-TeamsAuditEvents.ps1` | `teams-audit-events.csv`: Teams meeting, call, message and chat-created events counted per day, user and operation (source 8) |
| `collectors/Run-All.ps1` | Signs in once to Graph and once to Exchange and runs all of the above, plus the lifecycle report's team activity collector (source 4) |
| `collectors/TeamsActivitySchema.psd1` | Column order of every CSV, the per-cloud availability of every source, and the audited operations |
| `collectors/TeamsActivityHelpers.ps1` | Report-specific helpers (below) |
| `New-SampleData.ps1` | Regenerates `samples/` |
| `samples/` | Fake CSVs (everything is on `example.com`); `samples/gcchigh/` holds the header-only files a GCC High run leaves for the three Graph usage reports |
| `tests/` | Pester 5 tests; every tenant call is mocked |

### Sources that are not collectors here

* **Source 4, Teams activity per team** is the same API as the Teams and Groups lifecycle report's
  team activity. It is not rebuilt: [`reports/teams-groups-lifecycle`](../teams-groups-lifecycle/)
  `Get-TeamActivity.ps1` already writes every column the contract lists (`team-activity.csv`).
  `Run-All.ps1` runs it into the same output folder so the file sits beside the others.
* **Source 7, chat and channel message export** is not built. It reads message bodies, and the contract
  leaves "should the library read message content at all" as a policy decision for the build issue;
  nothing in the library reads content, and `tests/ReadOnly.Tests.ps1` fails if a collector calls
  `getAllMessages`. In GCC High it is the only Graph route to message counts, so the question stays open.
* **Source 9, the admin center reports** (Microsoft 365 admin center **Reports > Usage > Microsoft Teams**,
  Teams admin center **Analytics & reports**) are a manual export. No page read says a script can read
  them, so there is no collector. They are `Yes` for GCC and GCC High, which is the fallback for the
  usage reports in GCC High.
* The **license utilization** report ([`docs/candidates/license-utilization.md`](../../docs/candidates/license-utilization.md))
  shares the `getTeamsUserActivityUserDetail` endpoint and the `displayConcealedNames` check, and is not
  rebuilt here beyond these two collectors; it is its own issue (BRO-374).

`TeamsActivityHelpers.ps1` is report-specific because the shared layer has no usage-report,
Graph paging or audit-search logic: the availability rule, the usage-report download, a Graph
read that follows `@odata.nextLink` (refusing any host that is not Microsoft Graph) and retries
HTTP 429 and 503, the unified audit log search with its `ReturnLargeSet` paging and 50,000-record
guard, the per-day audit count, and the call-record row builder. The functions follow the ones in
`reports/exchange-activity` and `reports/sharepoint-onedrive-activity`. The shared connection
(`Connect-M365Service`) and CSV append functions are used as they are.

## Before you run it

```powershell
Install-Module Microsoft.Graph, ExchangeOnlineManagement -Scope CurrentUser
```

```powershell
./collectors/Run-All.ps1 -OutputPath ./out
./collectors/Run-All.ps1 -OutputPath ./out -Environment GCCHigh -AppId $appId `
    -CertificateThumbprint $thumbprint -TenantId $tenantId -Organization contoso.onmicrosoft.com
```

`-Environment` is `Commercial` (default), `GCC` or `GCCHigh`. Exchange Online connects with
`-ExchangeEnvironmentName O365USGovGCCHigh` in GCC High and with the default environment otherwise
([app-only authentication](https://learn.microsoft.com/powershell/exchange/app-only-auth-powershell-v2)).
Graph uses `https://graph.microsoft.com` for Commercial and GCC and `Connect-MgGraph -Environment USGov`
for GCC High ([national cloud deployments](https://learn.microsoft.com/graph/deployments)).

## Role and licence each collector needs

| Collector | Permission or role | Licence |
| --- | --- | --- |
| Usage reports (1 to 3) | Graph `Reports.Read.All` (application and delegated). A delegated sign-in also needs an Entra limited admin role: Company Administrator, Exchange Administrator, SharePoint Administrator, Lync Administrator, Teams Service Administrator, Teams Communications Administrator or Reports Reader. Global Reader and Usage Summary Reports Reader see tenant-level data only, so they get source 2 but not the per-user detail ([authorization](https://learn.microsoft.com/graph/reportroot-authorization)) | None named on the pages |
| Report settings (5) | Graph `ReportSettings.Read.All`. A delegated sign-in also needs an Entra limited admin role | None named |
| Call records (6) | Graph `CallRecords.Read.All`, **application only**; delegated is not supported on either API page, and an administrator must grant it ([FAQ](https://learn.microsoft.com/graph/callrecords-api-faq)) | None named |
| Audit events (8) | Audit Logs or View-Only Audit Logs in the Microsoft Purview portal, and the same roles in the Exchange admin center to run `Search-UnifiedAuditLog` ([audit search](https://learn.microsoft.com/purview/audit-search)) | Audit (Standard) is enough for the operations used; `MessagesExported` and `MessageDeleted` need Audit (Premium) and are not searched |

`Run-All.ps1` requests `Reports.Read.All` and `ReportSettings.Read.All` on an interactive Graph
sign-in in addition to the library's default scopes. `CallRecords.Read.All` is an application
permission and is not a delegated scope, so it is not requested on that sign-in. Source 6 returns
data only on an app-only sign-in (`-AppId` and `-CertificateThumbprint`) after an administrator
has granted that permission to the app.

## Availability per cloud

`Available` and `NotAvailable` appear only where a Microsoft page says so. `UNVERIFIED` means no page found
says either way, or the pages disagree: the collector attempts it, logs a warning to `run.log` and records a
refusal rather than guessing. A `NotAvailable` source writes its CSV header only and logs why.

| # | Source | Commercial | GCC | GCC High |
| --- | --- | --- | --- | --- |
| 1 | `getTeamsUserActivityUserDetail` | [Available](https://learn.microsoft.com/graph/api/reportroot-getteamsuseractivityuserdetail) | UNVERIFIED: the API page marks the global service `✅`, but the [usage reports cloud table](https://learn.microsoft.com/graph/api/resources/report#cloud-deployments) marks Microsoft Cloud for US Government `➖` | [NotAvailable](https://learn.microsoft.com/graph/api/reportroot-getteamsuseractivityuserdetail) (US Government L4 `❌`) |
| 2 | `getTeamsUserActivityCounts` | [Available](https://learn.microsoft.com/graph/api/reportroot-getteamsuseractivitycounts) | UNVERIFIED (same conflict) | [NotAvailable](https://learn.microsoft.com/graph/api/reportroot-getteamsuseractivitycounts) |
| 3 | `getTeamsDeviceUsageUserDetail` | [Available](https://learn.microsoft.com/graph/api/reportroot-getteamsdeviceusageuserdetail) | UNVERIFIED (same conflict) | [NotAvailable](https://learn.microsoft.com/graph/api/reportroot-getteamsdeviceusageuserdetail) |
| 4 | `getTeamsTeamActivityDetail` | Collected by the lifecycle report: [Available](https://learn.microsoft.com/graph/api/reportroot-getteamsteamactivitydetail) | UNVERIFIED | NotAvailable |
| 5 | `GET /admin/reportSettings` | [Available](https://learn.microsoft.com/graph/api/adminreportsettings-get) | UNVERIFIED (global service `✅`, cloud table `➖`) | UNVERIFIED: the [Graph page](https://learn.microsoft.com/graph/api/adminreportsettings-get) marks US Government L4 `❌`, while the [usage reports overview](https://learn.microsoft.com/microsoft-365/admin/activity-reports/activity-reports#use-adminreportsettings-to-show-user-group-or-site-details) says an API in all environments changes the setting. Sources 1 to 4 are `NotAvailable` there anyway, so there is nothing to conceal |
| 6 | `GET /communications/callRecords` | [Available](https://learn.microsoft.com/graph/api/callrecords-cloudcommunications-list-callrecords) | [Available](https://learn.microsoft.com/graph/deployments) (global service; the API pages have no GCC column) | [Available](https://learn.microsoft.com/graph/api/callrecords-cloudcommunications-list-callrecords) (US Government L4 `✅`) |
| 7 | Teams export APIs (message content) | Not built | Not built | Not built (see above) |
| 8 | `Search-UnifiedAuditLog`, Teams operations | [Available](https://learn.microsoft.com/purview/audit-log-activities#teams-activities) | Audit (Standard) [Available](https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-deployments). `MeetingDetail` and `MeetingParticipantDetail` Available ([footnote 9](https://learn.microsoft.com/purview/audit-log-activities#teams-activities)); `CallParticipantDetail`, `MessageSent`, `ChatCreated` UNVERIFIED | Audit (Standard) [Available](https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-high-deployments). Same split of operations |
| 9 | Admin center reports | Manual export, [Available](https://learn.microsoft.com/microsoftteams/teams-analytics-and-reports/teams-reporting-reference) | Manual export, Available | Manual export, Available (`Yes` in the GCCH column) |

In GCC High a run therefore writes headers only for sources 1 to 3 (and the lifecycle report's `team-activity.csv`),
and still collects sources 5, 6 and 8. Source 6 there sees calls and meetings only, not chat or channel messages,
and only 30 days back. **No read-only Teams PowerShell cmdlet or Graph call was found that gives chat counts,
device use or per-team activity in GCC High**; the admin center export is the only documented route.

## The CSVs

Dates and times are UTC. `RunDate` is the UTC date of the run.

| File | Kind | One row is | Columns |
| --- | --- | --- | --- |
| `teams-user-activity-user-detail.csv` | State | A user, a run, a period or a day | `RunDate`, `QueryDate`, then the report's columns: `ReportRefreshDate`, `UserId`, `UserPrincipalName`, `LastActivityDate`, `IsDeleted`, `DeletedDate`, `AssignedProducts`, `TeamChatMessageCount` (includes posts and replies), `PrivateChatMessageCount`, `CallCount` (1:1 calls), `MeetingCount`, `PostMessages`, `ReplyMessages`, `UrgentMessages`, organized and attended counts (all, ad hoc, scheduled one-time, scheduled recurring), `AudioDuration`, `VideoDuration`, `ScreenShareDuration` (ISO 8601) and the same three `...InSeconds`, `HasOtherAction`, `IsLicensed`, `ReportPeriod`. `Tenant Display Name` and `Shared Channel Tenant Display Names` are not kept |
| `teams-user-activity-counts.csv` | State | A report date in a run | `RunDate`, `ReportRefreshDate`, `ReportDate`, `TeamChatMessages`, `PostMessages`, `ReplyMessages`, `PrivateChatMessages`, `Calls`, `Meetings`, `AudioDuration`, `VideoDuration`, `ScreenShareDuration`, `MeetingsOrganized`, `MeetingsAttended`, `ReportPeriod` |
| `teams-device-usage-user-detail.csv` | State | A user, a run, a period or a day | `RunDate`, `QueryDate`, `ReportRefreshDate`, `UserId`, `UserPrincipalName`, `LastActivityDate`, `IsDeleted`, `DeletedDate`, then `UsedWeb`, `UsedWindowsPhone`, `UsediOS`, `UsedMac`, `UsedAndroidPhone`, `UsedWindows`, `UsedChromeOS`, `UsedLinux` (Yes or No in the period, not counts), `IsLicensed`, `ReportPeriod` |
| `report-settings.csv` | State | A run | `RunDate`, `DisplayConcealedNames` |
| `call-records.csv` | Event | A session of a call record, per record version (one row with empty session columns if a record has none; a repeated session id stays when the endpoints or times differ) | `CallRecordId`, `Version`, `Type` (`groupCall`, `peerToPeer`, `unknown`), `Modalities` (joined with `;`), `StartDateTime`, `EndDateTime`, `LastModifiedDateTime`, `SessionId`, `SessionStartDateTime`, `SessionEndDateTime`, `CallerUserId`, `CallerPlatform`, `CalleeUserId`, `CalleePlatform` (`windows`, `macOS`, `iOS`, `android`, `web`, ...) |
| `teams-audit-events.csv` | Event | A UTC day, workload, user and operation | `Date`, `Workload`, `UserId`, `Operation`, `EventCount` |
| `team-activity.csv` | State | Written by the lifecycle report | See [`reports/teams-groups-lifecycle`](../teams-groups-lifecycle/README.md) |

### Things the data does not say

* **Usage reports are periods, not events.** The count columns aggregate 7, 30, 90 or 180 days (`-Period`, default `D30`).
  A daily run appends a snapshot stamped with the run date. `LastActivityDate` is the most recent activity whatever the
  period. Reports typically become available 24 to 72 hours late, and a recent day can change between runs because the
  system re-checks the past three days ([Team usage report](https://learn.microsoft.com/microsoftteams/teams-analytics-and-reports/teams-usage-report)).
* **Per day, per user** is not a Microsoft-documented result. Source 2 is per day but for the whole tenant. `-Date` on
  sources 1 and 3 is described as "users who performed any activity" on that date, and the page does not say the count
  columns are single-day values (28 days back for source 3, 30 for source 1). Until a test tenant settles that, per-user
  daily counts come from the difference between snapshots, which is the library's derivation.
* **Do not add counts that overlap.** `TeamChatMessageCount` already includes `PostMessages` and `ReplyMessages`;
  `MeetingCount` equals the attended count and is being phased out; `MeetingsOrganizedCount` need not equal the sum of its
  three parts because unclassified meetings are not in the CSV.
* **Missing rows are not zero.** The Graph pages do not say whether an inactive user is returned, and a deleted user's data
  is removed within 30 days. A no-activity question needs a full user list to join against, and a test tenant should
  confirm whether source 1 returns zero-activity rows first.
* **Activity from apps is not counted.** Metric counts include Teams client built-in features but not Teams app posts or
  replies, or emails in the channel.
* **Audio and video duration** count the whole call or meeting if audio or video was enabled, not speaking or camera-on time.
* **Call records** appear up to 150 minutes after a call ends, and a later version can arrive after that, so every run
  reads `startDateTime` from now minus `-LookbackDays` (30) through now minus `-DelayMinutes` (default 180). The list
  filter is the call's start, so resuming from the latest start already stored would skip a call that started earlier
  and was not readable yet, and would skip a later version of it. Rows already exported are skipped. `Version` is part of
  the row key, so each version appends its own rows and a reader keeps the highest `Version` per `CallRecordId`. A session
  id can repeat when a transfer involves more than one service identity; those rows stay because the session times and
  endpoint ids are part of the key. An older record, or one not readable yet, answers 404; the collector logs and skips it.
  Records are kept 30 days, so run it daily to keep history. `organizer` and `participants` stopped returning data on
  2026-06-30, so the user id comes from the endpoint's `associatedIdentity`. Each read sends
  `Prefer: include-unknown-enum-members`. Participants who stream a live event are not returned.
* **Audit event counts** are made by the library and will not match the usage report counts. `MessageSent` is in public preview
  and is generated for chat only when guests, federated or anonymous users are present, so it is not a complete chat count;
  `ChatCreated` is logged only when the chat is created through a Graph API call; `MeetingParticipantDetail` already includes
  recorded or transcribed calls. Audit (Standard) keeps these records 180 days (`-LookbackDays` accepts up to 180). A window that
  reaches the 50,000-record session cap is not written and the run stops.
* **Shared channels.** The usage reports can undercount active shared channels because of telemetry limits (lifecycle report's
  `team-activity.csv`).

## Concealed names

Usage reports hide user and team names by default. The setting is **Settings > Org settings > Services > Reports > "Conceal
user, group, and site names in all reports"** in the Microsoft 365 admin center, and it also applies to the reports in Graph
and the Teams admin center. `Get-ReportSettings.ps1` records its state in `report-settings.csv` (`DisplayConcealedNames`). When
it is `True`, the user columns in sources 1 and 3 (and the team name in `team-activity.csv`) hold concealed identifiers and
**cannot be joined** to the users or groups files: show "names concealed, usage not joined" instead of zero-activity users or
teams. Showing identifiable names is a Global Administrator action and a logged Purview audit event; the library never changes
the setting. Run the report-settings collector first (`Run-All.ps1` does).

## Open items (from the contract)

* GCC: the usage report APIs stay UNVERIFIED because the API pages and the cloud table disagree; a read-only probe in a GCC tenant
  settles it, once for the library.
* GCC High: skip sources 1 to 3 (the library rule, which this report follows) or document a manual export from source 9.
* Whether source 1 with `-Date` returns single-day counts, and whether sources 1 and 4 return zero-activity rows. Both need a test tenant.
* Whether to build source 7 at all. It is the only GCC High route to message counts and it reads message content.
