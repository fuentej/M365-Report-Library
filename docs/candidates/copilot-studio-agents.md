# Candidate: Copilot Studio agents

Status: sources verified against Microsoft Learn; nothing built. Written for BRO-316 so the build issue can
be written without guessing. Nothing here was run against a tenant.

Cloud names follow the library: `Commercial`, `GCC`, `GCCHigh` (`-Environment`).

## Finding

A supported read-only surface exists: the **Power Platform inventory**
([Power Platform inventory](https://learn.microsoft.com/power-platform/admin/power-platform-inventory),
[agent fields](https://learn.microsoft.com/microsoft-copilot-studio/admin-agent-inventory)). It covers four
of the five starting questions in one tenant-wide query. It has two constraints the library has not met
before:

1. **Delegated user sign-in only.** The inventory API and the Power Platform for Admins V2 connector do not
   support service principals or managed identities and return HTTP 403 for them
   ([Inventory API](https://learn.microsoft.com/power-platform/admin/inventory-api),
   [Power Platform inventory](https://learn.microsoft.com/power-platform/admin/power-platform-inventory#programmatic-access)).
   The existing collectors support app-only certificate sign-in; this one cannot. A scheduled unattended run
   would need a different source (Dataverse, below) or an operator signing in.
2. **Most agent properties are Preview.** `channels`, `authentication`, `sharedWithViewers`,
   `sharedWithEditors`, `orchestration`, `model`, `capabilitiesCounts` and the connector detail are marked
   Preview in the schema
   ([agent fields](https://learn.microsoft.com/microsoft-copilot-studio/admin-agent-inventory#configuration-properties)).
   The page itself says preview features are not meant for production use.

## How to read the table

* `Available` / `NotAvailable` appear only where a Microsoft page says so; the link is in the cell.
* `UNVERIFIED` means no Microsoft page found says either way. The collector should ask for the data and
  record a refusal in `run.log`, as the existing reports do.
* The inventory includes unpublished draft agents and published agents. `lastPublishedAt` is empty when
  the agent has never been published. For an agent that has been published, the inventory shows the
  published version: newer unpublished changes are omitted until that agent is published again
  ([agent fields](https://learn.microsoft.com/microsoft-copilot-studio/admin-agent-inventory#agent-properties)).
* It excludes V1 agents (Power Virtual Agents classic bots). When an agent has more than 200 configured
  resources of one type, the inventory returns a random 200 of that type, not the first 200 and not the
  full set. `capabilitiesCounts` (Preview) is the complete count for each type
  ([known limitations](https://learn.microsoft.com/microsoft-copilot-studio/admin-agent-inventory#known-limitations)).

## Sources

| # | Source | Endpoint or cmdlet | Least privileged role | License | Event / state | Retention | Commercial | GCC | GCC High |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 1 | Environments (the list of Power Platform environments) | `POST {PowerPlatformAPI url}/resourcequery/resources/query?api-version=2024-10-01` with `TableName` `PowerPlatformResources` and a where clause on `type` for `microsoft.powerplatform/environments` ([Inventory API](https://learn.microsoft.com/power-platform/admin/inventory-api)). Repeat with `Options.SkipToken` while `resultTruncated` is set and `skipToken` is returned. One response is not the full set. The government host for `{PowerPlatformAPI url}` is UNVERIFIED (see notes) | AI Reader, the least privileged role for this report. It can see agents, agentic apps, agent flows, environments and environment groups ([access requirements](https://learn.microsoft.com/power-platform/admin/power-platform-inventory#access-requirements)). Global Reader can see every inventory resource and is broader than this report needs. Delegated sign-in only | None named | State | n/a | [Available](https://learn.microsoft.com/power-platform/admin/power-platform-inventory) | [Available](https://learn.microsoft.com/power-platform/admin/power-platform-inventory#sovereign-clouds) (Environments: GCC Yes) | [Available](https://learn.microsoft.com/power-platform/admin/power-platform-inventory#sovereign-clouds) (Environments: GCC High Yes) |
| 2 | Agents with ownership, creation, last publish, channels, authentication, sharing counts, orchestration, model, capability counts | Same query as source 1, including `SkipToken` paging, with `type` `microsoft.copilotstudio/agents`; fields in [agent fields](https://learn.microsoft.com/microsoft-copilot-studio/admin-agent-inventory) | As source 1 | None named | State | n/a; changes typically appear within 20 minutes ([agent fields](https://learn.microsoft.com/microsoft-copilot-studio/admin-agent-inventory)) | [Available](https://learn.microsoft.com/power-platform/admin/power-platform-inventory) | [Available](https://learn.microsoft.com/power-platform/admin/power-platform-inventory#sovereign-clouds) (Agents with the Standard harness: GCC Yes; GitHub Copilot harness, Copilot Chat harness and Agent Builder agents: No) | [Available](https://learn.microsoft.com/power-platform/admin/power-platform-inventory#sovereign-clouds) (same rows: GCC High Yes for Standard harness) |
| 3 | Connectors and operations each agent uses (connector id, operation id, connection provider) | `properties.powerPlatformConnectors` on source 2 (Preview). Each operation `usedAs` is `Tool`, `Topic Tool` or `Knowledge` ([connector properties](https://learn.microsoft.com/microsoft-copilot-studio/admin-agent-inventory#connector-properties)) | As source 1 | None named | State | n/a | [Available](https://learn.microsoft.com/power-platform/admin/power-platform-inventory#connector-inventory-preview) | [NotAvailable](https://learn.microsoft.com/power-platform/admin/power-platform-inventory#known-limitations) ("Connector inventory isn't available in GCC, GCC High, or DoD") | [NotAvailable](https://learn.microsoft.com/power-platform/admin/power-platform-inventory#known-limitations) (same) |
| 4 | Knowledge sources and actions per agent (file attachments, knowledge source configuration, HTTP request actions, prompts, MCP) | Dataverse tables `bot` and `botcomponent` in each environment; the Copilot Agent Kit documents which columns and `data` markers hold each item ([Agent Inventory data source](https://learn.microsoft.com/microsoft-copilot-studio/guidance/kit-agent-inventory-data-source#agent-details-table)). Read with the Dataverse Web API | Dataverse security role with Organization-level read on `bot` and `botcomponent`; Learn lists System Administrator and System Customizer as Organization (CRUD). Least privileged read-only role: UNVERIFIED ([security roles](https://learn.microsoft.com/microsoft-copilot-studio/guidance/sec-gov-phase3#assign-copilot-studio-authoring-permissions-by-using-security-roles)). App-only access through an application user is not covered by any page found | Dataverse environment (a Copilot Studio license is not stated as required to read the tables) | State | n/a | [Available](https://learn.microsoft.com/microsoft-copilot-studio/guidance/custom-analytics-strategy#copilot-studio,-dataverse,-and-analytics) | UNVERIFIED. The Dataverse Web API host is listed (`https://*.api.crm9.dynamics.com/api/data/v9.1/`) ([Dynamics 365 US Government URLs](https://learn.microsoft.com/power-platform/admin/microsoft-dynamics-365-government#dynamics-365-us-government-urls)). No page found says `bot` and `botcomponent` hold the same knowledge and action columns in GCC | UNVERIFIED. Host listed: `https://*.api.crm.microsoftdynamics.us/api/data/v9.1/` ([same URLs](https://learn.microsoft.com/power-platform/admin/microsoft-dynamics-365-government#dynamics-365-us-government-urls)). No page found says the `bot` tables hold the same columns in GCC High |
| 5 | Last modified date and last modifier | `bot` columns `modifiedon` and `modifiedby`; last published: `bot.published` ([Agent Inventory data source](https://learn.microsoft.com/microsoft-copilot-studio/guidance/kit-agent-inventory-data-source#agent-details-table)). The inventory exposes `createdAt` and `lastPublishedAt` but no modified field | As source 4 | As source 4 | State | n/a | As source 4 | As source 4 | As source 4 |
| 6 | Audit events for agent changes and use | Microsoft Purview audit. Authoring operations are the event labels on the Copilot Studio page, including `BotCreate`, `BotDelete`, `BotUpdateOperation-BotPublish` and `BotUpdateOperation-BotShare` ([Audit Copilot Studio activities](https://learn.microsoft.com/microsoft-copilot-studio/admin-logging-copilot-studio)). Usage is `Search-UnifiedAuditLog -RecordType CopilotInteraction`, then filter `AppIdentity` offline for the prefix `Copilot.Studio.` because search cannot filter `AppIdentity` ([Copilot audit logs](https://learn.microsoft.com/purview/audit-copilot)). A `CopilotInteraction` filter does not return the authoring operations. Read those labels with `Search-UnifiedAuditLog -Operations` and the same `-SessionId` paging. No row on the supported-services page names a record type for them ([services that support auditing](https://learn.microsoft.com/purview/audit-supported-services)). The cmdlet returns at most 100 records unless the same `-SessionId` is repeated with `-SessionCommand ReturnLargeSet` (session cap 50,000; `-ResultSize` maximum 5,000). `StartDate` and `EndDate` are UTC. For a full download the cmdlet page recommends the Microsoft 365 Management Activity API; the Copilot Studio page calls that developer path the Office 365 Management API ([Search-UnifiedAuditLog](https://learn.microsoft.com/powershell/module/exchangepowershell/search-unifiedauditlog)) | Audit Reader role group, which grants View-Only Audit Logs ([audit permissions](https://learn.microsoft.com/purview/audit-get-started#step-2-assign-permissions-to-search-the-audit-log)). Audit Logs and the Audit Manager role group can also search and can change auditing. `Search-UnifiedAuditLog` also needs the Exchange admin center View-Only Audit Logs or Audit Logs role | Users need a Microsoft 365 license for Copilot Studio to record events ([prerequisites](https://learn.microsoft.com/microsoft-copilot-studio/admin-logging-copilot-studio#prerequisites)). Custom applications built with Copilot Studio are included in Audit (Standard), and that page says pay-as-you-go does not apply to them ([Copilot audit logs](https://learn.microsoft.com/purview/audit-copilot)). Get started with auditing also lists Microsoft Copilot Studio under pay-as-you-go billing for non-Microsoft 365 AI interactions ([audit setup](https://learn.microsoft.com/purview/audit-get-started)). Whether pay-as-you-go is required: UNVERIFIED | Event | 180 days for Audit (Standard) records generated on or after 17 October 2023; records from before that date are kept 90 days. Audit (Premium)'s default one-year policy covers only Exchange, SharePoint, OneDrive and Microsoft Entra, so Copilot Studio records stay at 180 days unless a custom retention policy applies. Retaining a record longer than 180 days requires the user who generated it to have an Office 365 or Microsoft 365 E5 license, a Microsoft Purview Suite, or an E5 eDiscovery and Audit add-on ([retention](https://learn.microsoft.com/purview/audit-log-retention-policies)). `CreationTime` on the Copilot Studio schema is UTC ([audit fields](https://learn.microsoft.com/microsoft-copilot-studio/admin-logging-copilot-studio#schema-audit-fields)) | [Available](https://learn.microsoft.com/microsoft-copilot-studio/admin-logging-copilot-studio) | UNVERIFIED. The Copilot Studio audit page lists as a prerequisite that the tenant "isn't a Federal Risk and Authorization Management Program (FedRAMP) tenant" ([prerequisites](https://learn.microsoft.com/microsoft-copilot-studio/admin-logging-copilot-studio#prerequisites)); no page found says whether that excludes GCC | UNVERIFIED (same prerequisite; GCC High is a FedRAMP-aligned cloud, but no page states the effect on this feature) |

Consolidated notes:

* Copilot Studio itself is available in GCC and GCC High. A report should not treat these documented
  gaps as findings
  ([Copilot Studio for US Government](https://learn.microsoft.com/microsoft-copilot-studio/requirements-licensing-gcc#copilot-studio-us-government-feature-limitations)):
  Triggers / Autonomous Agents, Preview Models and Azure AI Search as a knowledge source are not
  available in GCC or GCC High. Generative Answers Enhanced Search is not available in GCC. In GCC High
  only, these are not available: the Copilot Studio Microsoft Teams app, the Teams channel in the web
  app, Transfer to agents, the Teams and Microsoft Copilot channel, Copilot Agents extend M365, and
  Prompt Action. Generative Orchestration is available in both clouds. An absent capability that this
  list does not name is still a finding.
* Admin center and Power Platform API hosts differ in government clouds. The admin center hosts are
  `gcc.admin.powerplatform.microsoft.us` (GCC) and `high.admin.powerplatform.microsoft.us` (GCC High)
  ([Power Apps US Government](https://learn.microsoft.com/power-platform/admin/powerapps-us-government#power-apps-us-government-service-urls)).
  A Power Platform API base URL for government clouds was not found on any page; treat the API endpoint for
  GCC and GCC High as UNVERIFIED until a page names it.
* Source 1 and 2 are served through Azure Resource Graph. The inventory page notes Power Platform inventory
  needs access to Azure Resource Manager and can fail to load if conditional access requires MFA for it
  ([known limitations](https://learn.microsoft.com/power-platform/admin/power-platform-inventory#known-limitations)).

## Proposed report pages

| # | Page | Reads | What it shows |
| --- | --- | --- | --- |
| 1 | Overview | 1, 2 | Agent count by environment and by cloud harness; published vs draft; trend by snapshot date |
| 2 | Ownership | 2 | Owner and creator per agent; agents whose owner is gone from the directory (join to the library's `users.csv`) |
| 3 | Sharing and authentication | 2 | Viewer and editor counts, entire-tenant sharing, authentication mode (None, Microsoft Entra, Generic OAuth 2.0) |
| 4 | Channels | 2 | Channels per agent; agents published with no authentication |
| 5 | Connectors and actions | 3, 4 | Connectors and operations used (Commercial only for source 3); HTTP request actions, prompts, MCP |
| 6 | Knowledge | 4 | Knowledge sources per agent; web search as a knowledge source (`IsWebSearchEnabledForKnowledge`, Preview) |
| 7 | Lifecycle | 2, 5, 6 | Created, last published, last modified; inactive agents; audit events where available |

## Starting questions

| # | Question | Status |
| --- | --- | --- |
| 1 | Which agents exist, in which environments | Covered by sources 1 and 2 |
| 2 | Who owns each agent | Covered by source 2 (`ownerId`, `createdBy`) |
| 2 | ...who has access, and how widely each is shared | **Partly dropped.** Source 2 gives counts of viewer and editor users and groups and an entire-tenant flag, not the identities. Naming the users or groups was not found in any page; dropped until a Learn page documents it. |
| 3 | Which knowledge sources, connectors and actions each agent uses | Connectors: source 3 (not in GCC or GCC High), including operations whose `usedAs` is `Knowledge`. Other knowledge and actions: source 4 through Dataverse. The inventory's own knowledge property is only `IsWebSearchEnabledForKnowledge` (Preview) ([knowledge properties](https://learn.microsoft.com/microsoft-copilot-studio/admin-agent-inventory#knowledge-properties)). In GCC and GCC High, connectors are dropped from the inventory source. |
| 4 | Which channels, and the authentication setting | Covered by source 2 (`channels`, `authentication`, both Preview) |
| 5 | When each agent was last modified or published | Published: source 2 `lastPublishedAt`. Modified: source 5 only. |

## Open items for the build issue

* Decide how to handle the delegated-only sign-in: operator-run collector for source 1 and 2, or Dataverse
  (sources 4 and 5) for unattended runs.
* Find or confirm the Power Platform API base URL for GCC and GCC High, or drop sources 1 to 3 in those
  clouds in favour of Dataverse.
* Confirm in a test tenant the least privileged Dataverse role that can read `bot` and `botcomponent`.
* Confirm whether Copilot Studio audit events (source 6) are recorded in GCC and GCC High.
* Confirm whether pay-as-you-go billing is required for Copilot Studio audit events. The Copilot audit
  page includes those applications in Audit (Standard); the auditing get-started page lists them under
  pay-as-you-go.
