# Exchange Online usage and activity

How mailboxes are used: storage against quota, who sends and reads mail, which email apps
people connect with, which mailboxes are idle, and a send and receive count the library
makes itself where the Graph usage reports are not available.

This folder holds the collectors, fake sample CSVs and tests. **There is no Power BI
project yet**; it is a separate, later issue. The sources and their per-cloud availability
were verified in [`docs/candidates/exchange-activity.md`](../../docs/candidates/exchange-activity.md)
(BRO-380); the tables below are copied from it.

The collectors write one CSV per source into an output folder, plus a `run.log`. State
sources append a snapshot stamped with `RunDate`; event sources (`message-trace.csv`,
`graph-message-trace.csv`) append from the latest timestamp already in the file. Nothing is
rewritten or deleted. Everything is read-only against the tenant.

## Contents

| Path | What it is |
| --- | --- |
| `collectors/Get-MailboxUsageDetail.ps1` | `mailbox-usage-detail.csv`: storage, item counts, quotas, last activity, archive flag, derived quota status (source 1) |
| `collectors/Get-MailboxUsageStorage.ps1` | `mailbox-usage-storage.csv`: tenant mailbox storage over time (source 2) |
| `collectors/Get-EmailActivityUserDetail.ps1` | `email-activity-user-detail.csv`: send, receive and read counts per user (source 3) |
| `collectors/Get-EmailAppUsageUserDetail.ps1` | `email-app-usage-user-detail.csv`: the email apps each user connected with (source 4) |
| `collectors/Get-ReportSettings.ps1` | `report-settings.csv`: whether usage reports conceal names (source 5) |
| `collectors/Get-Mailboxes.ps1` | `mailboxes.csv`: every mailbox, recipient type, quotas as configured (source 6) |
| `collectors/Get-MailboxStatistics.ps1` | `mailbox-statistics.csv`: size, items, storage limit status, last logon, from Exchange (source 7) |
| `collectors/Get-MessageTrace.ps1` | `message-trace.csv`: sent and received per recipient, `Get-MessageTraceV2` (source 8) |
| `collectors/Get-MobileDevices.ps1` | `mobile-devices.csv`: mobile devices syncing to each mailbox (source 9) |
| `collectors/Get-GraphMessageTrace.ps1` | `graph-message-trace.csv`: sent and received per recipient, Graph message trace (source 11) |
| `collectors/Run-All.ps1` | Signs in once to Exchange and once to Graph and runs all ten |
| `collectors/ExchangeActivitySchema.psd1` | Column order of every CSV and the per-cloud availability of every source |
| `collectors/ExchangeActivityHelpers.ps1` | Report-specific helpers (below) |
| `New-SampleData.ps1` | Regenerates `samples/` |
| `samples/` | Fake CSVs (everything is on `example.com`); `samples/gcchigh/` holds the header-only files a GCC High run leaves for the four Graph usage reports |
| `tests/` | Pester 5 tests; every tenant call is mocked |

Source 10 in the contract (the Microsoft 365 admin center usage reports, a manual export
fallback for GCC High) is **not** a collector. Nothing page-read says a script can read those
reports in GCC High, so it stays an open item in the contract.

`ExchangeActivityHelpers.ps1` is report-specific because the shared layer has no usage-report,
quota or message-trace logic: the availability rule, the usage-report download and the quota
status, the 10-day message-trace windows and the `Get-MessageTraceV2` continuation, a 95-in-5-minutes
request limiter, and the Graph paging read. The shared connection (`Connect-M365Service`) and CSV
append functions are used as they are.

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

Trial run: add `-MailboxLimit 5`. The per-mailbox sources (7, 8, 9) make one or more calls per mailbox;
size a full run for the largest tenant.

## Role and licence each collector needs

| Collector | Permission or role | Licence |
| --- | --- | --- |
| Usage reports (1 to 4) | Graph `Reports.Read.All`. A delegated sign-in also needs an Entra limited admin role: Company Administrator, Exchange Administrator, SharePoint Administrator, Lync Administrator, Teams Service Administrator, Teams Communications Administrator or Reports Reader. Global Reader and Usage Summary Reports Reader see tenant-level data only ([authorization](https://learn.microsoft.com/graph/reportroot-authorization)) | None named on the pages |
| Report settings (5) | Graph `ReportSettings.Read.All` (`ReportSettings.ReadWrite.All` is not used) | None named |
| Exchange cmdlets (6 to 9) | App-only needs `Exchange.ManageAsApp` plus an Exchange role. The pages do not state which role runs each cmdlet. Organization Management and Recipient Management are listed for recipients and mobile devices; Help Desk is the least of the three groups listed for message trace ([feature permissions](https://learn.microsoft.com/exchange/permissions-exo/feature-permissions)). The least-privileged role is **UNVERIFIED** until `Get-ManagementRole` is run in the tenant | None named |
| Graph message trace (11) | Application permission `ExchangeMessageTrace.Read.All` (no delegated permission or Entra role is listed), and a service principal for app `8bd644d1-64a1-4d4b-ae52-2e0cbf64e373` in the tenant. Until that provisioning finishes the API returns 401, and a 401 is not an empty trace | None named |

`Run-All.ps1` requests `Reports.Read.All` and `ReportSettings.Read.All` on an interactive Graph sign-in in
addition to the library's default scopes. `ExchangeMessageTrace.Read.All` is an application permission, so
source 11 returns data only on an app-only sign-in.

## Availability per cloud

`Available` and `NotAvailable` appear only where a Microsoft page says so. `UNVERIFIED` means no page found says
either way: the collector attempts it, logs a warning to `run.log` and records a refusal rather than guessing.
A `NotAvailable` source writes its CSV header only and logs why.

| # | Source | Commercial | GCC | GCC High |
| --- | --- | --- | --- | --- |
| 1 | `getMailboxUsageDetail` | [Available](https://learn.microsoft.com/graph/api/reportroot-getmailboxusagedetail) | [Available](https://learn.microsoft.com/graph/deployments) (global service) | [NotAvailable](https://learn.microsoft.com/graph/api/reportroot-getmailboxusagedetail) |
| 2 | `getMailboxUsageStorage` | [Available](https://learn.microsoft.com/graph/api/reportroot-getmailboxusagestorage) | [Available](https://learn.microsoft.com/graph/deployments) | [NotAvailable](https://learn.microsoft.com/graph/api/reportroot-getmailboxusagestorage) |
| 3 | `getEmailActivityUserDetail` | [Available](https://learn.microsoft.com/graph/api/reportroot-getemailactivityuserdetail) | [Available](https://learn.microsoft.com/graph/deployments) | [NotAvailable](https://learn.microsoft.com/graph/api/reportroot-getemailactivityuserdetail) |
| 4 | `getEmailAppUsageUserDetail` | [Available](https://learn.microsoft.com/graph/api/reportroot-getemailappusageuserdetail) | [Available](https://learn.microsoft.com/graph/deployments) | [NotAvailable](https://learn.microsoft.com/graph/api/reportroot-getemailappusageuserdetail) |
| 5 | `GET /admin/reportSettings` | [Available](https://learn.microsoft.com/graph/api/adminreportsettings-get) | [Available](https://learn.microsoft.com/graph/deployments) | UNVERIFIED: the [Graph page](https://learn.microsoft.com/graph/api/adminreportsettings-get) marks US Government L4 unsupported, while the [usage reports overview](https://learn.microsoft.com/microsoft-365/admin/activity-reports/activity-reports) says an API in all environments changes this setting |
| 6 | `Get-EXOMailbox` | [Available](https://learn.microsoft.com/powershell/module/exchangepowershell/get-exomailbox) | UNVERIFIED | UNVERIFIED |
| 7 | `Get-EXOMailboxStatistics` | [Available](https://learn.microsoft.com/powershell/module/exchangepowershell/get-exomailboxstatistics) | UNVERIFIED | UNVERIFIED |
| 8 | `Get-MessageTraceV2` | [Available](https://learn.microsoft.com/powershell/module/exchangepowershell/get-messagetracev2) | UNVERIFIED for the V2 cmdlet (the service-level "Message trace" row is `Yes`) | UNVERIFIED (same) |
| 9 | `Get-EXOMobileDeviceStatistics` | [Available](https://learn.microsoft.com/powershell/module/exchangepowershell/get-exomobiledevicestatistics) | UNVERIFIED | UNVERIFIED |
| 11 | Graph message trace (`/beta`) | UNVERIFIED (no national-cloud table) | UNVERIFIED | UNVERIFIED |

For sources 6 to 9 no cmdlet page states cloud availability; the only evidence is the
[service description](https://learn.microsoft.com/office365/servicedescriptions/office-365-platform-service-description/office-365-us-government/exchange-online-for-us-government-environments), which says `Yes` for remote PowerShell and message trace in GCC and GCC High and
names no cmdlet. **Learn disagrees with itself for sources 1 to 4 in GCC High:** the per-API pages mark
US Government L4 unsupported, while the service description lists Microsoft Graph Reports as `Yes`. The collectors
follow the per-API pages. The contract asks for one read-only operator probe of sources 1, 7, 8 and 11 in a GCC
High tenant before the skip is relied on.

## The CSVs

Dates and times are UTC. `RunDate` is the UTC date of the run.

| File | Kind | One row is | Columns |
| --- | --- | --- | --- |
| `mailbox-usage-detail.csv` | State | A mailbox, a run, a period | `RunDate`, the report's own columns (`ReportRefreshDate`, `UserPrincipalName`, `DisplayName`, `IsDeleted`, `DeletedDate`, `CreatedDate`, `LastActivityDate`, `ItemCount`, `StorageUsedByte`, `IssueWarningQuotaByte`, `ProhibitSendQuotaByte`, `ProhibitSendReceiveQuotaByte`, `DeletedItemCount`, `DeletedItemSizeByte`, `DeletedItemQuotaByte`, `HasArchive`, `ReportPeriod`), and `QuotaStatus` |
| `mailbox-usage-storage.csv` | State | A report date in a run | `RunDate`, `ReportRefreshDate`, `StorageUsedByte`, `ReportDate`, `ReportPeriod` |
| `email-activity-user-detail.csv` | State | A user, a run, a period or a day | `RunDate`, `QueryDate`, `ReportRefreshDate`, `UserPrincipalName`, `DisplayName`, `IsDeleted`, `DeletedDate`, `LastActivityDate`, `SendCount`, `ReceiveCount`, `ReadCount`, `MeetingCreatedCount`, `MeetingInteractedCount`, `AssignedProducts`, `ReportPeriod` |
| `email-app-usage-user-detail.csv` | State | A user, a run, a period or a day | `RunDate`, `QueryDate`, `ReportRefreshDate`, `UserPrincipalName`, `DisplayName`, `IsDeleted`, `DeletedDate`, `LastActivityDate`, then `MailForMac`, `OutlookForMac`, `OutlookForWindows`, `OutlookForMobile`, `OtherForMobile`, `OutlookForWeb`, `POP3App`, `IMAP4App`, `SMTPApp` (Yes or No in the period), `ReportPeriod` |
| `report-settings.csv` | State | A run | `RunDate`, `DisplayConcealedNames` |
| `mailboxes.csv` | State | A mailbox in a run | `RunDate`, `ExternalDirectoryObjectId`, `UserPrincipalName`, `PrimarySmtpAddress`, `RecipientType`, `RecipientTypeDetails` (UserMailbox, SharedMailbox, ...), the five configured quotas as the cmdlet prints them, `UseDatabaseQuotaDefaults` |
| `mailbox-statistics.csv` | State | A mailbox in a run | `RunDate`, `ExternalDirectoryObjectId`, `UserPrincipalName`, `ItemCount`, `TotalItemSize` / `TotalItemSizeBytes`, `DeletedItemCount`, `TotalDeletedItemSize` / `TotalDeletedItemSizeBytes`, `StorageLimitStatus`, `LastLogonTime`, `LastLogoffTime`, `LastLoggedOnUserAccount`, `IsArchiveMailbox`, the three database quotas |
| `message-trace.csv` | Event | A recipient of a message (not a distinct message) | `Received`, `MessageTraceId`, `SenderAddress`, `RecipientAddress`, `Status`. Sent for a mailbox is `SenderAddress` = the mailbox; received is `RecipientAddress` = the mailbox. No read count |
| `mobile-devices.csv` | State | A mobile device syncing to a mailbox | `RunDate`, `MailboxUserPrincipalName`, `DeviceId`, `DeviceType`, `DeviceOS`, `DeviceAccessState`, `DeviceUserAgent` |
| `graph-message-trace.csv` | Event | A recipient of a message | `ReceivedDateTime`, `Id`, `SenderAddress`, `RecipientAddress`, `Status`, `Size` |

`QuotaStatus` uses the admin center's four categories, and the boundary is at or above a quota
([mailbox usage report](https://learn.microsoft.com/microsoft-365/admin/activity-reports/mailbox-usage)):
`Good` is below the issue-warning quota; `Warning` is at or above it and below prohibit send; `CantSend` is at or
above prohibit send and below prohibit send/receive; `CantSendReceive` is at or above prohibit send/receive. It is
empty when a quota or the storage is missing. The API header list includes `Deleted Item Quota (Byte)` and
`Has Archive`, which the example schema on the same page omits; a CSV without them leaves those two columns empty.

### Things the data does not say

* **Usage reports are periods, not events.** The count columns aggregate 7, 30, 90 or 180 days (`-Period`, default `D30` for mailbox usage and email activity, `D180` for tenant storage). A daily run
  appends a snapshot stamped with the run date. `LastActivityDate` is the most recent activity whatever the period, so a
  snapshot can show inactivity older than the period. Reports typically become available 24 to 72 hours late.
* **Last activity** in the mailbox usage report is the last email send or read. `mailbox-statistics.csv` `LastLogonTime` is a
  separate column and is not assumed equal to it. `LastUserActionTime` is not read: Learn says it is being deprecated and is not the last
  active time.
* **Daily series.** `-Date` (one day) on sources 3 and 4 is sent instead of `-Period`. It reaches back 28 days for source 3 and 30 for source 4;
  running at least every 28 days keeps daily rows contiguous (an inference, not a Learn statement). A per-day series is one `-Date` call per day.
* **Deleted users.** A deleted user's usage data is removed within 30 days; a missing row is not a zero-activity row.
* **Shared mailboxes.** The Graph usage CSV has no recipient-type column. Join `UserPrincipalName` to `mailboxes.csv` (`RecipientTypeDetails`).
* **Message trace** covers the last 90 days and 10 days per query. The older window is queried first, so a failed later window does not move the watermark past the gap. A message that already has 1000 recipient rows is queried again with `-MessageTraceId` (required when a message was sent to more than 1000 recipients). `Received` is UTC; an Unspecified time is not shifted into the local zone. An unfinished run leaves `message-trace.pending` and the next run repeats that start instead of the newest row. A date-only start or end uses the session's regional short date, so the collectors pass full timestamps.
* **Mobile devices** is mobile sync only, not Outlook for Windows, Mac or web. The cmdlet page lists no output properties; the columns are
  the ones Learn's troubleshooting page selects.
* **No Outlook version breakdown** (the Graph columns carry no version) and **no read count in GCC High** (no read-only call found returns it).

## Concealed names

Usage reports hide user names by default. The setting is **Settings > Org settings > Services > Reports >
"Conceal user, group, and site names in all reports"** in the Microsoft 365 admin center, and `Get-ReportSettings.ps1` records its
state in `report-settings.csv` (`DisplayConcealedNames`). When it is `True`, the user columns in sources 1, 3 and 4 hold concealed
identifiers and **cannot be joined** to `mailboxes.csv` or `mailbox-statistics.csv`: show "names concealed, not joined" instead of
zero-activity mailboxes. Showing identifiable names is a logged Purview audit event, and the library never changes the setting.
Run the report-settings collector first (`Run-All.ps1` does).

## Overlap with other reports

* **Mailbox exfiltration risk** ([`reports/mailbox-exfiltration-risk`](../mailbox-exfiltration-risk/README.md)) shares `Get-EXOMailbox` and the Exchange
  connection. This report reads quotas and statistics only; forwarding, delegation, rules and audit settings stay there and are not rebuilt.
* **License utilization** ([`docs/candidates/license-utilization.md`](../../docs/candidates/license-utilization.md)) shares the
  `getEmailActivityUserDetail` endpoint and the `displayConcealedNames` check. This report is the Exchange-only view with mailbox storage
  and clients, and does not rebuild licences or the other workloads. The unified audit log (`MailItemsAccessed`) stays with the exfiltration report.

## Open items

* One read-only operator probe in a GCC High tenant of sources 1, 7, 8 and 11 (see Availability), and a decision on source 10.
* Confirm the Exchange role for sources 6 to 9 with `Get-ManagementRole` before naming a least-privileged role.
* Source 7 is one call per mailbox; decide whether source 1 is enough where it is available.
