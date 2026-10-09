# Copilot Studio agents

Which Copilot Studio agents exist and in which environments, who owns each, how widely each is
shared and how it authenticates, which channels it is published to, which connectors, knowledge
sources and actions it uses, when it was created, last modified and last published, and the
authoring events behind those changes.

The collectors write one CSV per source into an output folder of your choosing, plus a
`run.log`. Snapshot files get a new block of rows tagged `RunDate` on every run; the audit file
is appended from where the last run stopped. Nothing is rewritten or deleted, so the folder is a
history you can chart over time.

The sources were verified against Microsoft Learn in
[`docs/candidates/copilot-studio-agents.md`](../../docs/candidates/copilot-studio-agents.md),
which is the contract for this folder. The Power BI project is in `report/`; saving it as a `.pbit` needs
Power BI Desktop and is a separate, manual step.

## Sign-in: four of the six collectors are delegated and interactive

The Power Platform inventory API does not support service principals or managed identities and
returns HTTP 403 to them ([Inventory API](https://learn.microsoft.com/power-platform/admin/inventory-api#authentication)).
So `Get-PowerPlatformEnvironments.ps1`, `Get-CopilotStudioAgents.ps1` and `Get-AgentConnectors.ps1`
open an interactive sign-in as a user (`Connect-AzAccount`, then `Get-AzAccessToken`). They take
no `-AppId` or `-CertificateThumbprint`. The two Dataverse collectors sign in the same way,
because app-only access through an application user is not covered by any Learn page found.
Run these at a keyboard before each release; they cannot be scheduled unattended.

Only `Get-AgentAuditEvents.ps1` follows the library's usual pattern (interactive, or app-only
with `-AppId`, `-CertificateThumbprint`, `-TenantId` and `-Organization`).

## Contents

| Path | What it is |
| --- | --- |
| `collectors/Get-PowerPlatformEnvironments.ps1` | `environments.csv` — every Power Platform environment |
| `collectors/Get-CopilotStudioAgents.ps1` | `agents.csv` — every agent: owner, sharing counts, authentication, channels, publish state, capability counts |
| `collectors/Get-AgentConnectors.ps1` | `agent-connectors.csv` — the connector operations each agent uses |
| `collectors/Get-AgentComponents.ps1` | `agent-components.csv` — knowledge sources, tools, HTTP request actions, prompts and MCP actions, from Dataverse |
| `collectors/Get-AgentModifications.ps1` | `agent-modifications.csv` — last modified date and modifier, from Dataverse |
| `collectors/Get-AgentAuditEvents.ps1` | `agent-audit-events.csv` — authoring events from the unified audit log |
| `collectors/Run-All.ps1` | Runs all six |
| `collectors/CopilotStudioSchema.psd1` | The column order of every CSV and the per-cloud availability of every source |
| `collectors/CopilotStudioHelpers.ps1` | Helpers the collectors dot-source: the delegated token, inventory paging and Dataverse paging |
| `report/` | The Power BI project (PBIP): `CopilotStudioAgents.SemanticModel` (TMDL) and `CopilotStudioAgents.Report` (PBIR, seven pages). One parameter, `CsvFolder`, holds the folder the collectors wrote; it is built against `samples/` |
| `New-SampleData.ps1` | Regenerates `samples/` |
| `samples/` | Generated fake data; `samples/gcc/` and `samples/gcchigh/` hold the header-only `agent-connectors.csv` |
| `tests/` | Pester tests; every tenant call, including the sign-in, is mocked |

## Before you run it

### Modules

```powershell
Install-Module Az.Accounts -Scope CurrentUser
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

`Az.Accounts` provides `Connect-AzAccount` and `Get-AzAccessToken -AsSecureString`.

### Roles

| Collector | Least privileged role | Notes |
| --- | --- | --- |
| Environments, agents, connectors | AI Reader ([access requirements](https://learn.microsoft.com/power-platform/admin/power-platform-inventory#access-requirements)) | Global Reader sees every inventory resource and is broader than needed. Delegated user only |
| Components, modifications | A Dataverse security role with Organization-level read on `bot` and `botcomponent` ([security roles](https://learn.microsoft.com/microsoft-copilot-studio/guidance/sec-gov-phase3#assign-copilot-studio-authoring-permissions-by-using-security-roles)) | System Administrator and System Customizer are documented; the least privileged read-only role is **UNVERIFIED** |
| Audit events | Audit Reader role group, plus the Exchange View-Only Audit Logs or Audit Logs role ([audit permissions](https://learn.microsoft.com/purview/audit-get-started#step-2-assign-permissions-to-search-the-audit-log)) | Auditing must be on |

### Licences

Users need a Microsoft 365 licence for Copilot Studio to record audit events
([prerequisites](https://learn.microsoft.com/microsoft-copilot-studio/admin-logging-copilot-studio#prerequisites)).
Whether pay-as-you-go billing is also required for these events is **UNVERIFIED**.

## Availability by cloud

`Available` and `NotAvailable` appear only where a Microsoft page says so. **UNVERIFIED** means no
page found says either way: the collector asks for the data anyway and records a refusal in
`run.log`. Only `NotAvailable` skips a source, writing the CSV header only.

| Source | Commercial | GCC | GCC High |
| --- | --- | --- | --- |
| Environments | [Available](https://learn.microsoft.com/power-platform/admin/power-platform-inventory) | [Available](https://learn.microsoft.com/power-platform/admin/power-platform-inventory#sovereign-clouds) | [Available](https://learn.microsoft.com/power-platform/admin/power-platform-inventory#sovereign-clouds) |
| Agents | [Available](https://learn.microsoft.com/power-platform/admin/power-platform-inventory) | [Available](https://learn.microsoft.com/power-platform/admin/power-platform-inventory#sovereign-clouds) (Standard harness only) | [Available](https://learn.microsoft.com/power-platform/admin/power-platform-inventory#sovereign-clouds) (Standard harness only) |
| Connectors | [Available](https://learn.microsoft.com/power-platform/admin/power-platform-inventory#connector-inventory-preview) | [NotAvailable](https://learn.microsoft.com/power-platform/admin/power-platform-inventory#known-limitations) | [NotAvailable](https://learn.microsoft.com/power-platform/admin/power-platform-inventory#known-limitations) |
| Components, modifications (Dataverse) | [Available](https://learn.microsoft.com/microsoft-copilot-studio/guidance/custom-analytics-strategy#copilot-studio,-dataverse,-and-analytics) | UNVERIFIED | UNVERIFIED |
| Audit events | [Available](https://learn.microsoft.com/microsoft-copilot-studio/admin-logging-copilot-studio) | UNVERIFIED | UNVERIFIED |

**The Power Platform API host.** Commercial is `https://api.powerplatform.com`
([Query Resources](https://learn.microsoft.com/rest/api/power-platform/resourcequery/resource-query/query-resources)).
No page found names the host for GCC or GCC High, so it is **UNVERIFIED** and never guessed: in those
clouds pass `-ApiHost`, or the inventory collectors write a header only and say why.

**Dataverse URLs.** The inventory does not expose each environment's Dataverse URL, so pass every
environment to read with `-DataverseUrl`. Without it the two Dataverse collectors write a header only.
The Web API hosts for GCC and GCC High are listed on
[Dynamics 365 US Government](https://learn.microsoft.com/power-platform/admin/microsoft-dynamics-365-government#dynamics-365-us-government-urls).

## What the inventory does and does not show

* **Freshness.** Changes typically appear in the inventory after **about 15 to 20 minutes**
  (the [inventory page](https://learn.microsoft.com/power-platform/admin/power-platform-inventory) says 15,
  the [agent schema](https://learn.microsoft.com/microsoft-copilot-studio/admin-agent-inventory) says 20).
* **Drafts are included**, and `LastPublishedAt` is empty for an agent never published. For a
  published agent the inventory shows the published version; newer unpublished edits are withheld.
* **V1 (Power Virtual Agents classic) agents are not in the inventory.**
* **The 200-item limit.** When an agent has more than 200 resources of one type the inventory returns
  a **random** 200 of that type, not the first 200. `capabilitiesCounts` holds the complete count, so
  `agents.csv` writes both: the `Listed*` columns count what came back and the `Capabilities*` columns
  carry `capabilitiesCounts`. A `Listed` value below its `Capabilities` value means the rows in
  `agent-connectors.csv` for that agent are partial, and the collector logs a warning.
* **Paging.** Queries follow `skipToken` until a page returns none. The inventory page shows
  `resultTruncated` of `1` beside a `skipToken`, while the REST reference describes `0` as truncated;
  the collector does not rely on the flag.
* **Most agent properties are Preview** (channels, authentication, sharing counts, orchestration, model,
  `capabilitiesCounts`, connector detail). Microsoft says Preview features are not meant for production use.
* **Sharing identities are not collected.** Only viewer and editor user and group counts and the
  entire-tenant flag are documented.
* **Audit scope.** Authoring events are read with `Search-UnifiedAuditLog -Operations` and the labels in
  `CopilotStudioSchema.psd1`; `-RecordType CopilotInteraction` does not return them and agent usage events
  are not collected. A search returns 100 records unless paged with `-SessionCommand ReturnLargeSet`,
  capped at 50,000 per session; a window that reaches the cap is not written and the log names it so it
  can be re-run with a smaller `-WindowHours`. Retention is 180 days: Audit (Premium)'s one-year default
  does not cover Copilot Studio.

## Running it

```powershell
./collectors/Run-All.ps1 -OutputPath ./out -DataverseUrl https://org12345.crm.dynamics.com
```

Or one collector at a time, for example `./collectors/Get-CopilotStudioAgents.ps1 -OutputPath ./out`.
Every collector takes `-Environment Commercial|GCC|GCCHigh` (default `Commercial`). Nothing here writes to
the tenant. The inventory query is a POST that carries a query specification and changes nothing; the
read-only test allows it in one function and nowhere else.

## The CSVs

Timestamps are UTC (`yyyy-MM-ddTHH:mm:ssZ`). Lists are joined with `;`.

### `environments.csv` — snapshot, key `RunDate` + `EnvironmentId`

| Column | Meaning |
| --- | --- |
| `RunDate` | Date of the run |
| `EnvironmentId` | Environment id (`name` in the inventory) |
| `DisplayName` | Environment name |
| `Location` | Geography, such as `unitedstates` |
| `EnvironmentType` | Production, Default, Sandbox, Trial, Developer or Dataverse for Teams |
| `IsManaged` | Whether it is a managed environment |
| `EnvironmentGroup` | Environment group name, if any |
| `LastModifiedAt` | Last modified |

### `agents.csv` — snapshot, key `RunDate` + `EnvironmentId` + `AgentId`

| Column | Meaning |
| --- | --- |
| `RunDate`, `EnvironmentId` | As above |
| `AgentId` | The CDS bot id (`properties.name`). `properties.botId` is the same id when the Entra identity block is present; Agent Builder agents leave that block empty |
| `DisplayName` | Agent name |
| `Harness` | Harness type, such as `Standard` |
| `CreatedIn` | Copilot Studio or Microsoft Copilot Agent Builder |
| `CreatedAt`, `CreatedBy` | Creation time and the creator's Entra object id |
| `OwnerId` | The current owner's Entra object id. Join to `users.csv` for the name |
| `LastPublishedAt` | Empty for a draft that was never published |
| `IsPublished` | `True` when `LastPublishedAt` is set |
| `IsQuarantined`, `IsManaged` | Preview. Empty when the inventory returns null |
| `Orchestration`, `Model` | Preview. Classic or Generative; the model name |
| `Authentication` | Preview. None, Microsoft Entra or Generic OAuth 2.0 |
| `Channels` | Preview. Channels the agent is published to |
| `ViewerUserCount`, `ViewerGroupCount` | Preview. Users and groups it is shared with as viewers |
| `ViewerEntireTenant` | Preview. `True` when the maker shared it with the whole tenant, even if an admin did not approve it |
| `EditorUserCount`, `EditorGroupCount` | Preview. Users and groups it is shared with as editors |
| `ListedConnectorCount`, `ListedConnectorOperationCount` | Connectors and operations the inventory listed |
| `CapabilitiesDistinctConnectors`, `CapabilitiesDistinctConnectorOperations` | Preview. The complete counts from `capabilitiesCounts` |
| `IsWebSearchEnabledForKnowledge` | Preview. Whether web search is a knowledge source |

### `agent-connectors.csv` — snapshot, one row per connector operation

| Column | Meaning |
| --- | --- |
| `RunDate`, `EnvironmentId`, `AgentId` | As above |
| `ConnectorId`, `OperationId` | For example `shared_excelonlinebusiness` and `AddRowV2` |
| `UsedAs` | `Tool`, `Topic Tool` or `Knowledge` |
| `IsEnabled`, `RequiresEndUserConsent` | Operation settings |
| `WhenCanBeUsed` | `Anytime`, `ViaDirectReferenceOnly` or `Conditional` |
| `ConnectionProvider` | `User` or `Maker`. A connector returned with an empty `operations` array (tabular connectors such as SharePoint) is still one row, with the operation columns empty |

### `agent-components.csv` — snapshot

| Column | Meaning |
| --- | --- |
| `RunDate` | Date of the run |
| `DataverseUrl` | The environment it was read from |
| `AgentId` | The parent bot id |
| `ComponentId`, `ComponentName`, `ComponentType` | The `botcomponent` row; `ComponentType` is the raw `componenttype` value |
| `Category` | `KnowledgeSource`, `Tool`, `HttpRequest`, `Prompt` or `Mcp`, from markers in the component `data` |

### `agent-modifications.csv` — snapshot

| Column | Meaning |
| --- | --- |
| `RunDate`, `DataverseUrl` | As above |
| `AgentId`, `Name`, `CreatedOn` | The `bot` row |
| `ModifiedOn`, `ModifiedBy` | Last modified and the modifier's system user id. The inventory has no modified field |
| `PublishedOn` | `bot.published` |

### `agent-audit-events.csv` — event, appended from the latest `CreationTime`, key `Id`

| Column | Meaning |
| --- | --- |
| `CreationTime` | When the event happened (UTC) |
| `Id` | Unique id of the audit record |
| `Operation` | The event label, such as `BotUpdateOperation-BotPublish` |
| `UserId` | The actor's Entra id (`UserKey`) |
| `ResultStatus` | Status of the logged row |
| `BotId`, `BotSchemaName` | The agent |
| `BotComponentId`, `BotComponentType` | The component, for component events |

## The Power BI report

Open `report/CopilotStudioAgents.pbip` in Power BI Desktop and set the `CsvFolder` parameter to the folder
the collectors wrote (or to `samples/`). Seven pages follow the contract: Overview, Ownership, Sharing and
authentication, Channels, Connectors and actions, Knowledge and Lifecycle. Every page has a date-range
slicer, an agent slicer and the Anonymize toggle, which swaps agent names, environment names and the
owner, creator and modifier ids for stable pseudonyms.

* An empty value is **Unknown**, never zero. A header-only file (the connector file in GCC and GCC High)
  is **not collected**: its measures are blank, and the Overview shows how many of the six files hold data.
* The inventory returns a random 200 of one resource type when an agent has more. The report shows the
  listed count next to the `capabilitiesCounts` total and counts the agents whose rows are partial.
* Owners, creators and modifiers are ids; no directory file exists in this folder, so owners are not
  named and an owner who is gone from the directory cannot be detected here.
* Sharing is counts only; users and groups are not named.

The report tests (`ReportSchema`, `ReportModel`, `TmdlDataType`, `TmdlExpressionIndent`) reuse the
schema validator, vendored schemas and TMDL reader under `reports/guest-access/tests/`.
