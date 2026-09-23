# Purview information protection

Sensitivity labels, DLP, retention, Content Explorer and Copilot activity, ported from
[`fuentej/purview-ip-report`](https://github.com/fuentej/purview-ip-report) onto this
library's `shared/` layer. That repository's Power BI project (semantic model, report,
PBIP tests) is a separate port, tracked in a follow-up issue.

The collectors write one CSV per source into an output folder of your choosing, plus a
`run.log`. Snapshot files get a new block of rows on every run; event files are appended
from where the last run stopped. Nothing is ever rewritten or deleted, so the folder is a
history you can chart over time.

## Contents

| Path | What it is |
| --- | --- |
| `collectors/Get-Policies.ps1` | `policies.csv` — a snapshot of the label, DLP and retention configuration |
| `collectors/Get-ActivityExplorerEvents.ps1` | `activity-explorer-events.csv` — labeling, protection, DLP and retention activity |
| `collectors/Get-ContentExplorerSnapshot.ps1` | `content-explorer-snapshot.csv` — item counts per label/SIT per workload |
| `collectors/Get-CopilotAccessedResources.ps1` | `copilot-accessed-resources.csv` — resources Microsoft 365 Copilot accessed |
| `collectors/Run-All.ps1` | Runs the users collector and all four of the above |
| `collectors/PurviewIpSchema.psd1` | The column order of every CSV, the activity categories, and the per-cloud source availability table |
| `collectors/PurviewIpHelpers.ps1` | Helpers specific to this report (activity/DLP/Copilot field mapping), dot-sourced by every collector |
| `New-SampleData.ps1` | Regenerates `samples/` |
| `tests/` | Pester tests; every tenant call is mocked |

`users.csv` comes from `Invoke-EntraUserCollector` in the shared module, not from this
folder, because every report needs it.

## Before you run it

### Modules

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Install-Module Microsoft.Graph.Users -Scope CurrentUser
```

`ExchangeOnlineManagement` provides `Connect-IPPSSession`, `Connect-ExchangeOnline`,
`Export-ActivityExplorerData`, `Export-ContentExplorerData`, `Search-UnifiedAuditLog` and
the policy `Get-*` cmdlets. `Microsoft.Graph.Users` provides the shared module's
`Connect-MgGraph` / `Get-MgUser` for `users.csv`.

### Licensing

| Source | Licence needed |
| --- | --- |
| Activity Explorer, Content Explorer | Microsoft 365 E5, E5 Compliance, or the E5 Information Protection & Governance add-on |
| Copilot audit records | Microsoft 365 Copilot, plus Purview Audit (Standard or Premium) |
| Sensitivity labels, DLP, retention configuration | Microsoft 365 E3 for the basics; E5 for auto-labeling and advanced DLP |

See the
[Microsoft Purview service description](https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/microsoft-purview-service-description)
for the authoritative breakdown.

### Roles

Each collector needs read access to its own source and nothing more.

| Collector | Access needed |
| --- | --- |
| `Get-ActivityExplorerEvents.ps1` | A role group that can read Activity Explorer: Information Protection Analyst, Information Protection Investigator, Compliance Administrator, Compliance Data Administrator, Security Administrator, Security Reader, or Global Reader. See [Get started with activity explorer](https://learn.microsoft.com/purview/data-classification-activity-explorer). |
| `Get-ContentExplorerSnapshot.ps1` | **Content Explorer List Viewer** (counts and locations) or **Content Explorer Content Viewer** (also opens items — more than this collector needs). For app-only runs, add the role group to the service principal. See [Export-ContentExplorerData](https://learn.microsoft.com/powershell/module/exchangepowershell/export-contentexplorerdata). |
| `Get-CopilotAccessedResources.ps1` | **Audit Reader** or **Audit Manager** in Purview, or the View-Only Audit Logs / Audit Logs role in Exchange Online. See [Audit log search permissions](https://learn.microsoft.com/purview/audit-search). |
| `Get-Policies.ps1` | Read access to the label, DLP and retention configuration: Global Reader, or the read-only role groups (Sensitivity Label Reader, View-Only DLP Compliance Management, View-Only Retention Management). See [Permissions in the Microsoft Purview portal](https://learn.microsoft.com/purview/microsoft-365-compliance-center-permissions). |

## Cloud availability

Every collector takes `-Environment Commercial|GCC|GCCHigh`, defaulting to `Commercial`.
The shared module holds the connection endpoints; the table below is what each source
itself is documented to support. Recorded in `Get-PurviewSourceAvailability` /
`SourceAvailability` in `collectors/PurviewIpSchema.psd1`.

| Source | Commercial | GCC | GCC High |
| --- | --- | --- | --- |
| Activity Explorer | [Available](https://learn.microsoft.com/purview/data-classification-activity-explorer) | [Available](https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-deployments) | [Available](https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-high-deployments) |
| Content Explorer | [Available](https://learn.microsoft.com/purview/data-classification-content-explorer) | [Available](https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-deployments) | [Available](https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-high-deployments) |
| Content Explorer — Teams data | [Available](https://learn.microsoft.com/purview/data-classification-content-explorer) | [Not available](https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-deployments) — listed as in development | [Not available](https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-high-deployments) — listed as in development |
| Copilot audit records | [Available](https://learn.microsoft.com/purview/audit-copilot) | `UNVERIFIED` | `UNVERIFIED` |
| Policy configuration | [Available](https://learn.microsoft.com/powershell/exchange/connect-to-scc-powershell) | [Available](https://learn.microsoft.com/powershell/exchange/connect-to-scc-powershell) | [Available](https://learn.microsoft.com/powershell/exchange/connect-to-scc-powershell) |
| Users (Microsoft Graph) | [Available](https://learn.microsoft.com/graph/deployments) | [Available](https://learn.microsoft.com/graph/deployments) | [Available](https://learn.microsoft.com/graph/deployments) |

Where a row says `UNVERIFIED`, the collector does not assume either way: it asks for the
data, and if the service refuses, it writes the header only and records the reason in
`run.log`. Content Explorer's Teams workload is the one row that is documented as
unavailable rather than unverified, so `Get-ContentExplorerSnapshot.ps1` drops it
automatically in GCC and GCC High rather than recording a run of zero counts.

## Running it

```powershell
# Everything, interactively, against the commercial cloud.
./collectors/Run-All.ps1 -OutputPath ./out

# GCC High, app-only.
./collectors/Run-All.ps1 -OutputPath ./out -Environment GCCHigh `
    -AppId $appId -CertificateThumbprint $thumbprint `
    -TenantId $tenantId -Organization contoso.onmicrosoft.com

# One collector.
./collectors/Get-Policies.ps1 -OutputPath ./out -Environment GCC
```

`-Organization` is the tenant's `*.onmicrosoft.com` domain, and app-only sign-in to
Exchange Online (`Get-CopilotAccessedResources.ps1`) needs it.

### Resuming

`Get-ActivityExplorerEvents.ps1` and `Get-CopilotAccessedResources.ps1` start at the
latest timestamp already in their CSV, so running them daily costs one short query.
Activity Explorer keeps only 30 days, so a run at least that often builds a history
longer than the source itself retains.

### Re-running on the same day

`policies.csv` is keyed on `RunDate` plus `ObjectType` and `ObjectId`, so running twice in
one day does not duplicate a snapshot. `content-explorer-snapshot.csv` is keyed on
`RunDate` plus `TagType`, `TagName` and `Workload`. `activity-explorer-events.csv` is keyed
on `RecordIdentity`; `copilot-accessed-resources.csv` is keyed on `RecordId` plus
`ResourceId`.

## The files

### `policies.csv` — snapshot

One table with an `ObjectType` discriminator, filtered to get back any one of seven
collections:

| `ObjectType` | Source cmdlet |
| --- | --- |
| `SensitivityLabel` | `Get-Label` |
| `LabelPolicy` | `Get-LabelPolicy` |
| `AutoLabelingPolicy` | `Get-AutoSensitivityLabelPolicy` |
| `DlpPolicy` | `Get-DlpCompliancePolicy` |
| `DlpRule` | `Get-DlpComplianceRule` |
| `RetentionPolicy` | `Get-RetentionCompliancePolicy` |
| `RetentionLabel` | `Get-ComplianceTag` |

Columns that do not apply to a row are empty. `ParentId` / `ParentName` link a sublabel to
its parent label, and a DLP rule to its policy. `AppliesToCopilot` marks a DLP policy
scoped to Microsoft 365 Copilot or Copilot Chat, detected from the `CopilotExperiences`
enforcement plane or the Copilot location GUID `470f2276-e011-4e9d-a6ec-20768be3a4b0` in
the policy's `Locations` JSON
([reference](https://learn.microsoft.com/powershell/module/exchangepowershell/new-dlpcompliancepolicy)).
List-valued columns (`Locations`, `LabelIds`, `LabelNames`, `SensitiveInformationTypes`)
are joined with `;`.

### `activity-explorer-events.csv` — events, keyed on `RecordIdentity`

Source: [`Export-ActivityExplorerData`](https://learn.microsoft.com/powershell/module/exchangepowershell/export-activityexplorerdata).
One row per event: sensitivity label applied, changed or removed; protection changes; DLP
matches, enforcements and overrides with their justification; retention label changes; and
Copilot / AI app interactions.

Activity Explorer keeps 30 days, so **each run appends** and resumes from the watermark
already in the file — the latest `Happened` timestamp, read the same way every event
collector in this library reads one (`Get-CsvWatermark`). Running at least every 30 days
builds a history longer than the source keeps.

Most columns carry the name `Export-ActivityExplorerData` uses. The derived ones:

| Column | How it is derived |
| --- | --- |
| `EventDate` | The UTC date part of `Happened` |
| `Activity` | Normalised to the filter-enum name (`LabelApplied`). Taken from `ActivityId` when the record has one, otherwise mapped from `Activity` — a record can carry the portal's display name (`Label applied`) there. A value that maps to nothing is kept as-is and logged as a warning at the end of the run |
| `ActivityRaw` | Exactly what the cmdlet returned |
| `ActivityCategory` | `Activity` grouped into Labeling, Protection, Dlp, Retention, Ai, Discovery, Endpoint or Other (`ActivityCategories` in `PurviewIpSchema.psd1`) |
| `IsLabelDowngrade` | `LabelEventType -eq 'LabelDowngraded'` |
| `SensitiveInfoTypeName` / `Count` / `Confidence` | `SensitiveInfoTypeData` flattened: names joined with `;`, counts totalled, highest confidence kept |

`SensitivityLabel` and `OldSensitivityLabel` hold label GUIDs — join them to
`policies.csv` on `ObjectId` where `ObjectType` is `SensitivityLabel`.

### `content-explorer-snapshot.csv` — snapshot

Item counts per tag per workload, stamped with the run date: one row per (`TagType`,
`TagName`, `Workload`) with the aggregate `TotalCount`, from
[`Export-ContentExplorerData`](https://learn.microsoft.com/powershell/module/exchangepowershell/export-contentexplorerdata).

`TagType` is `Sensitivity`, `Retention` or `SensitiveInformationType`. `Workload` is `EXO`,
`ODB`, `SPO` or `Teams`. **Each run appends**, so repeated runs give the report a trend
rather than a single point.

### `copilot-accessed-resources.csv` — events, keyed on `RecordId` plus `ResourceId`

One row per resource Copilot accessed, flattened out of
`CopilotEventData.AccessedResources` on each `CopilotInteraction` record from
`Search-UnifiedAuditLog`. An interaction's own fields (`RecordId`, `UserId`, `AppHost`,
`ThreadId`) repeat on each of its resource rows. An interaction that touched no resource
keeps one row with the resource columns empty.

`Search-UnifiedAuditLog` is used rather than the Graph audit log query API because
`CopilotInteraction` and its `CopilotEventData.AccessedResources` payload are documented
against the unified audit log
([Audit logs for Copilot and AI applications](https://learn.microsoft.com/purview/audit-copilot),
[Copilot interaction events overview](https://learn.microsoft.com/office/office-365-management-api/copilot-schema)).

Carries `SensitivityLabelId`, `Status` (`success` / `failure`), and `PolicyId` /
`PolicyName` / `PolicyRules` from the resource's `PolicyDetails`, which the record only
includes when a policy blocked or restricted access. `AccessBlocked` is derived: true when
a policy is actually named, or `Status` is not `success` — an empty `PolicyDetails` is not
a block. **Each run appends**, resuming from the watermark on `CreationTime`.

A `ReturnLargeSet` session returns at most 50,000 records, so the search window is split
into `-WindowHours` windows (one day by default), each with its own session. A window that
still reaches the limit is reported as an error naming the window, rather than silently
dropping the rest — re-run it with a smaller `-WindowHours`.

### `users.csv` — snapshot

Written by the shared module so that every report can join to it. See the root README.

### `run.log`

One line per event, in the shared module's format:
`2026-09-01T10:00:00Z [Info] [policies] policies.csv: 26 rows written, 0 skipped.`

## What was not ported from `fuentej/purview-ip-report`

- **`Export-Users.ps1`** — the source repo's own Graph users export. This library already
  has `Invoke-EntraUserCollector` in `shared/`, which every report uses instead of a
  second copy (per BRO-259's own instructions). `Users.csv`'s `GraphUsers` source-
  availability row from the source repo is likewise dropped: the shared module's own
  cloud-endpoint table already documents Graph availability, and every cloud is
  `Available` there.
- **The bespoke Activity Explorer watermark file** (`.purview-watermark.*.json`, with its
  own `ExportedRecordIdentity` list for records sharing the watermark's exact second).
  This library's `Get-CsvWatermark` plus `Export-AppendCsv -KeyColumn RecordIdentity`
  gives the same resume-without-duplicating behaviour — the same mechanism every other
  event collector in this library already uses (`guest-invitations.csv`,
  `sharing-events.csv`) — without a second per-dataset watermark format.
- **`Environment` / `TenantId` / `CollectedAt` columns.** The source repo stamped every
  row with these because its CSVs are single-snapshot, replaced-per-run files with no
  `RunDate`. This library's convention is the opposite: a `RunDate` (or the source's own
  event timestamp) on every row, one output folder per collection run against one
  environment, so the same information is already implicit in *which* file a row came
  from. Matches the guest-access report, which carries neither.
- **The Python package** (`src/purview_ip_report/`) and the Power BI project
  (`report/*.pbip`) — out of scope for this issue; the Power BI port is a separate,
  blocked issue.
- **`New-SampleData.ps1`'s exact entity counts.** This report's sample set is smaller
  (60 members, ~130 events) than the source's (200 members, 1,000+ events) — enough to
  exercise every code path and every activity category without the extra generation and
  review time a byte-for-byte-matched set would cost. The column shapes and per-source
  availability are ported exactly.

## Tests

```powershell
pwsh -NoProfile -Command "Invoke-Pester -Path ./shared/tests, ./reports/purview-ip/tests -CI"
```

Run from the root of the repository. No tenant is contacted: the Security & Compliance,
Exchange Online and Microsoft Graph cmdlets are stubbed in
`shared/tests/TenantCmdletStubs.ps1` and mocked per test.
