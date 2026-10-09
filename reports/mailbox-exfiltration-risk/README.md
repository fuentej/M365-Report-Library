# Mailbox exfiltration risk

Which mailboxes forward mail and where, which inbox rules forward, redirect or delete mail, which
mail flow rules redirect or blind-copy it, who holds Full Access, Send As and Send on Behalf, which
applications hold consent to read mail, and the audit events and audit settings behind those changes.

The collectors write one CSV per source into an output folder of your choosing, plus a `run.log`.
Snapshot files get a new block of rows tagged `RunDate` on every run; the two audit-event files are
appended from where the last run stopped. Nothing is rewritten or deleted, so the folder is a
history you can chart over time.

The sources were verified against Microsoft Learn in
[`docs/candidates/mailbox-exfiltration-risk.md`](../../docs/candidates/mailbox-exfiltration-risk.md),
which is the contract for this folder. The Power BI project is a separate, later piece of work.
Nothing here writes to the tenant.

## Contents

| Path | What it is |
| --- | --- |
| `collectors/Get-AcceptedDomains.ps1` | `accepted-domains.csv` — the accepted domains used to decide what is external |
| `collectors/Get-MailboxForwarding.ps1` | `mailbox-forwarding.csv` — mailboxes that forward, with an external flag |
| `collectors/Get-SendOnBehalf.ps1` | `send-on-behalf.csv` — Send on Behalf grants |
| `collectors/Get-InboxRules.ps1` | `inbox-rules.csv` — inbox rules that forward, redirect or delete |
| `collectors/Get-TransportRules.ps1` | `transport-rules.csv` — mail flow rules that redirect or blind-copy |
| `collectors/Get-MailboxFullAccess.ps1` | `mailbox-full-access.csv` — Full Access grants |
| `collectors/Get-SendAsPermissions.ps1` | `send-as-permissions.csv` — Send As grants |
| `collectors/Get-DelegatedConsents.ps1` | `delegated-consents.csv` — delegated permission grants |
| `collectors/Get-AppRoleAssignments.ps1` | `app-role-assignments.csv` — app roles held on the Microsoft Graph service principal |
| `collectors/Get-MailboxChangeEvents.ps1` | `mailbox-change-events.csv` — audit events for rule, forwarding and permission changes |
| `collectors/Get-MailAccessEvents.ps1` | `mail-access-events.csv` — audit events for mail accessed and sent |
| `collectors/Get-AuditConfiguration.ps1` | `audit-configuration.csv` — whether auditing is on and what each mailbox logs |
| `collectors/Run-All.ps1` | Runs the shared users collector and all twelve of the above; signs in once to Exchange Online and once to Graph |
| `collectors/MailboxExfiltrationSchema.psd1` | The column order of every CSV, the per-cloud availability of every source |
| `collectors/MailboxExfiltrationHelpers.ps1` | Helpers the collectors dot-source: availability, external-domain test, audit search |
| `New-SampleData.ps1` | Regenerates `samples/` |
| `samples/` | Generated fake data |
| `tests/` | Pester tests; every tenant call is mocked |

`users.csv` comes from `Invoke-EntraUserCollector` in the shared module, not from this folder, because
every report needs it. `Run-All.ps1` writes it.

## Before you run it

### Modules

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Install-Module Microsoft.Graph.Authentication, Microsoft.Graph.Identity.SignIns, Microsoft.Graph.Applications, Microsoft.Graph.Users -Scope CurrentUser
```

### Roles and permissions

The Exchange cmdlet pages do not state the least privileged role; they defer to the
[permissions finder](https://learn.microsoft.com/powershell/exchange/find-exchange-cmdlet-permissions).
Where this table names a role for an Exchange cmdlet it is the role whose description matches
([roles in Exchange Online](https://learn.microsoft.com/exchange/permissions-exo/permissions-exo#roles-in-exchange-online)).
That is an inference, not a Learn statement, so it stays marked **UNVERIFIED** until confirmed in a
test tenant.

| Collector | Role or permission | Notes |
| --- | --- | --- |
| Accepted domains | UNVERIFIED | |
| Mailbox forwarding, send on behalf, full access, send as | View-Only Recipients ("View recipient properties") — **UNVERIFIED** | |
| Inbox rules | UNVERIFIED | The cmdlet page says it does **not** work for View-Only Organization Management or the Entra Global Reader role |
| Transport rules | View-Only Configuration ("Views all of the organization and mail flow (non-recipient) settings") — **UNVERIFIED** | |
| Delegated consents | `Directory.Read.All` (application and delegated); signed in, Global Reader or Directory Readers ([list oauth2PermissionGrants](https://learn.microsoft.com/graph/api/oauth2permissiongrant-list)) | |
| App role assignments | `Application.Read.All` (application and delegated); signed in, Directory Readers ([list appRoleAssignedTo](https://learn.microsoft.com/graph/api/serviceprincipal-list-approleassignedto)) | Global Reader is not in that page's role list |
| Audit events | Audit Reader role group (View-Only Audit Logs); Audit Manager also works ([audit permissions](https://learn.microsoft.com/purview/audit-get-started#step-2-assign-permissions-to-search-the-audit-log)) | `Set-Mailbox` records are visible only to unrestricted admins ([admin units](https://learn.microsoft.com/purview/audit-search#scoping-access-to-audit-logs-using-administrative-units)) |
| Audit configuration | Organization settings role — **UNVERIFIED**; Audit Reader for `Get-AdminAuditLogConfig` | |

A source the sign-in cannot read leaves a header-only CSV and a line in `run.log` naming the role, and
does not stop the others.

### Audit licence and retention

* `MailItemsAccessed` is **Audit (Standard)** and is on by default for users with Office 365 E3/E5 or
  Microsoft 365 E3/E5 ([investigate accounts](https://learn.microsoft.com/purview/audit-log-investigate-accounts)).
  Its `SensitivityLabel` property is an Audit (Premium) property and is not collected. A missing Premium
  licence is not the same as zero access events.
* Retention ([retention policies](https://learn.microsoft.com/purview/audit-log-retention-policies)):
  **180 days** for Audit (Standard) records generated on or after 17 October 2023 (90 days before);
  **one year** for Exchange records of users with E5 or an Audit (Premium) add-on; **ten years** needs
  the 10-year audit log retention add-on in addition to that licence, plus a retention policy — Audit
  (Premium) alone does not retain for ten years. Non-user records (service principals, system) stay at
  one year.
* Exchange records are typically searchable 60 to 90 minutes after the event; Microsoft does not
  guarantee a time. Exchange admin cmdlet records can take up to 30 minutes.
* A search returns 100 records unless paged. Each window is paged with the same `-SessionId` and
  `-SessionCommand ReturnLargeSet`, which is unsorted and stops at 50,000. A window that reaches the cap
  is **not written**: `run.log` names its exact `-StartDate` and `-EndDate` and the collector stops, so
  re-run that window with a smaller `-WindowHours`. The mailbox-change search and the Exchange admin
  search share that cutoff: a cap in either one holds back every record at or after the capped window,
  including records the other search already has, so the watermark cannot pass events that were never
  returned. A capped window is a failed window, not a complete export.
* `AuditEnabled` on a mailbox is not a per-mailbox switch: `Get-Mailbox` always shows `True` when
  mailbox auditing on by default is on. `audit-configuration.csv` reads the organization setting instead.

## Availability by cloud

`Available` and `NotAvailable` appear only where a Microsoft page says so. **UNVERIFIED** means no page
found says either way: the collector asks for the data anyway, logs a warning, and records a refusal in
`run.log`. Only `NotAvailable` skips a source, writing the CSV header only. The contract doc marks no
source `NotAvailable` in any cloud.

Exchange availability in GCC and GCC High comes from the
[Exchange Online for US government](https://learn.microsoft.com/office365/servicedescriptions/office-365-platform-service-description/office-365-us-government/exchange-online-for-us-government-environments#exchange-online-features)
service description, which lists features, not cmdlets.

| Source | Commercial | GCC | GCC High |
| --- | --- | --- | --- |
| Accepted domains | [Available](https://learn.microsoft.com/powershell/module/exchangepowershell/get-accepteddomain) | UNVERIFIED | UNVERIFIED |
| Mailbox forwarding, send on behalf | [Available](https://learn.microsoft.com/powershell/module/exchangepowershell/get-mailbox) | [Available](https://learn.microsoft.com/office365/servicedescriptions/office-365-platform-service-description/office-365-us-government/exchange-online-for-us-government-environments#exchange-online-features) ("Remote Windows PowerShell access" row) | Available (same row) |
| Inbox rules | [Available](https://learn.microsoft.com/powershell/module/exchangepowershell/get-inboxrule) | Available ("Inbox rules" row) | Available (same row) |
| Transport rules | [Available](https://learn.microsoft.com/exchange/security-and-compliance/mail-flow-rules/mail-flow-rules) | Available ("Mail flow rules" row) | Available (same row) |
| Full Access, Send As | [Available](https://learn.microsoft.com/exchange/recipients-in-exchange-online/manage-permissions-for-recipients) | Available ("Delegate access" row) | Available (same row) |
| Delegated consents, app role assignments | [Available](https://learn.microsoft.com/graph/api/oauth2permissiongrant-list) | [Available](https://learn.microsoft.com/graph/deployments) (global endpoint) | Available (US Government L4 endpoint) |
| Mailbox change and mail access audit events | [Available](https://learn.microsoft.com/purview/audit-solutions-overview#comparison-of-key-capabilities) | [Available](https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-deployments) (Audit (Standard)) | [Available](https://learn.microsoft.com/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/plan-for-microsoft-purview-gcc-high-deployments) (Audit (Standard)) |
| Audit configuration (`Get-OrganizationConfig`, `Get-AdminAuditLogConfig`) | [Available](https://learn.microsoft.com/purview/audit-mailboxes#verify-mailbox-auditing-on-by-default-is-turned-on) | UNVERIFIED | UNVERIFIED |

Connections follow the library: `-Environment` is `Commercial` (default), `GCC` or `GCCHigh`. Exchange
Online in GCC High uses `-ExchangeEnvironmentName O365USGovGCCHigh`
([app-only auth](https://learn.microsoft.com/powershell/exchange/app-only-auth-powershell-v2)). Graph uses
`https://graph.microsoft.com` for Commercial and GCC and `Connect-MgGraph -Environment USGov` for GCC High
([deployments](https://learn.microsoft.com/graph/deployments)). Interactive sign-in is the default;
`-AppId`, `-CertificateThumbprint`, `-TenantId` and `-Organization` sign in app-only.

## What is not collected

* **Who consented to a tenant-wide grant.** For `consentType` `AllPrincipals` the grant does not name the
  approver ([resource page](https://learn.microsoft.com/graph/api/resources/oauth2permissiongrant)); that
  needs a Microsoft Entra audit source that was not verified. Per-user consent is covered by `PrincipalId`.
* **Exchange-side app access** (for example EWS). Not researched; no source verified.
* **The tenant-level automatic forwarding policy** (outbound spam filter). Not in the starting questions.
* **Hidden inbox rules are read but not flagged.** `Get-InboxRule -IncludeHidden` returns them, and the
  cmdlet documents no marker for a hidden rule.

## Running it

```powershell
./collectors/Run-All.ps1 -OutputPath ./out
```

Or one collector at a time, for example `./collectors/Get-MailboxForwarding.ps1 -OutputPath ./out`.
The per-mailbox collectors (`Get-InboxRules.ps1`, `Get-MailboxFullAccess.ps1`) make one call per mailbox;
`-MailboxLimit` reads only the first N for a trial run. The audit-event collectors take `-StartDate`,
`-EndDate` (UTC; a value with no time zone is midnight UTC), `-LookbackDays` (first run, default 90)
and `-WindowHours` (default 24). The search uses `-Formatted`, so `RecordType` is a display name such
as `ExchangeAdmin`.

## The CSVs

Timestamps are UTC (`yyyy-MM-ddTHH:mm:ssZ`). Lists are joined with `;`. "Snapshot" files add one block of
rows per run, tagged `RunDate`.

### `accepted-domains.csv` — snapshot

| Column | Meaning |
| --- | --- |
| `RunDate` | Date of the run |
| `Name`, `DomainName` | The accepted domain; a wildcard such as `*.example.com` covers its subdomains |
| `DomainType` | Authoritative, InternalRelay or ExternalRelay |
| `IsDefault` | Whether it is the default domain |

### `mailbox-forwarding.csv` — snapshot; only mailboxes that forward

| Column | Meaning |
| --- | --- |
| `RunDate` | Date of the run |
| `ExternalDirectoryObjectId`, `UserPrincipalName`, `PrimarySmtpAddress` | The mailbox. Join to `users.csv` |
| `ForwardingAddress` | An internal recipient to forward to. Not an SMTP forward |
| `ForwardingSmtpAddress` | An SMTP forward, as `smtp:user@domain` |
| `ForwardingSmtpDomain` | The domain of `ForwardingSmtpAddress` |
| `DeliverToMailboxAndForward` | `True` when a copy stays in the mailbox |
| `IsExternal` | `True` when `ForwardingSmtpDomain` is not an accepted domain; `False` when it is, when it is a subdomain of an accepted domain with `MatchSubdomains`, or when only `ForwardingAddress` is set; empty when the accepted domains could not be read |

### `send-on-behalf.csv` — snapshot, one row per delegate

| Column | Meaning |
| --- | --- |
| `RunDate`, `ExternalDirectoryObjectId`, `UserPrincipalName`, `PrimarySmtpAddress` | The mailbox |
| `Delegate` | A recipient in `GrantSendOnBehalfTo` |

### `inbox-rules.csv` — snapshot; only rules that forward, redirect or delete

| Column | Meaning |
| --- | --- |
| `RunDate` | Date of the run |
| `MailboxUserPrincipalName`, `MailboxExternalDirectoryObjectId` | The mailbox |
| `RuleIdentity`, `RuleName`, `Enabled`, `Priority` | The rule |
| `ForwardTo`, `ForwardAsAttachmentTo`, `RedirectTo` | Recipients, as returned |
| `DeleteMessage` | `True` when the rule deletes the message |
| `TargetDomains` | SMTP domains found in the three recipient lists |
| `HasExternalTarget` | `True` when any target domain is not an accepted domain; empty when unknown or no address was found |

### `transport-rules.csv` — snapshot; only rules with a redirect, blind-copy or visible-copy action

| Column | Meaning |
| --- | --- |
| `RunDate` | Date of the run |
| `Name`, `Guid`, `State`, `Priority` | The rule; `State` is Enabled or Disabled |
| `RedirectMessageTo`, `BlindCopyTo` | The redirect and blind-copy actions |
| `CopyTo`, `AddToRecipients` | Actions that add visible Cc and To recipients, for context |
| `TargetDomains`, `HasExternalTarget` | As in `inbox-rules.csv` |

### `mailbox-full-access.csv` — snapshot

| Column | Meaning |
| --- | --- |
| `RunDate` | Date of the run |
| `MailboxExternalDirectoryObjectId`, `MailboxUserPrincipalName` | The mailbox that is accessed |
| `User` | Who holds Full Access. `NT AUTHORITY\SELF`, Deny rows and inherited rows are dropped |
| `AccessRights` | The rights, as returned |

### `send-as-permissions.csv` — snapshot

| Column | Meaning |
| --- | --- |
| `RunDate` | Date of the run |
| `Identity` | The mailbox or group that can be sent as |
| `Trustee` | Who can send as it |
| `AccessRights`, `AccessControlType`, `IsInherited` | As returned, so a report can filter default entries |

### `delegated-consents.csv` — snapshot

| Column | Meaning |
| --- | --- |
| `RunDate`, `Id` | Run date and grant id |
| `ClientId` | The application's service principal |
| `ConsentType` | `AllPrincipals` for tenant-wide admin consent, `Principal` for one user |
| `PrincipalId` | The consenting user for `Principal`; empty for `AllPrincipals` |
| `ResourceId` | The API the grant is for |
| `Scope` | Space-separated permissions |
| `HasMailScope` | `True` when a scope starts with `Mail.` or `MailboxSettings.` (this report's rule) |

### `app-role-assignments.csv` — snapshot

| Column | Meaning |
| --- | --- |
| `RunDate`, `AssignmentId` | Run date and assignment id |
| `AppRoleId`, `AppRoleValue` | The role and its name, such as `Mail.Read`, resolved from the service principal |
| `PrincipalId`, `PrincipalDisplayName`, `PrincipalType` | The application holding the role |
| `ResourceId`, `ResourceDisplayName` | The service principal the role is on (Microsoft Graph) |
| `CreatedDateTime` | When it was assigned |
| `IsMailRole` | `True` when `AppRoleValue` starts with `Mail.` or `MailboxSettings.` (this report's rule) |

### `mailbox-change-events.csv` and `mail-access-events.csv` — events, appended, key `Id`

| Column | Meaning |
| --- | --- |
| `CreationTime` | When the event happened (UTC). The next run starts here |
| `Id` | Unique id of the audit record |
| `RecordType` | For example `ExchangeAdmin` |
| `Operation` | `New-InboxRule`, `Set-InboxRule`, `UpdateInboxRules`, `Set-Mailbox`, `Add-MailboxPermission`, `Remove-MailboxPermission` and mail flow rule cmdlets in the change file; `MailItemsAccessed`, `Send`, `SendAs`, `SendOnBehalf` in the access file |
| `UserId` | Who did it |
| `Workload`, `ObjectId` | As recorded |
| `MailboxOwnerUPN`, `ClientIP`, `ResultStatus` | As recorded |
| `Parameters` | `Name=Value` pairs of an Exchange admin record, such as `ForwardingSmtpAddress=smtp:user@example.net` |

Mail flow rule changes are `-RecordType ExchangeAdmin` records. The activities page publishes no fixed
operation list for them, so the collector keeps the records whose operation contains `TransportRule`
(this report's filter) and does not name individual cmdlets. `-SkipExchangeAdmin` leaves that search out.

### `audit-configuration.csv` — snapshot

| Column | Meaning |
| --- | --- |
| `RunDate` | Date of the run |
| `Scope`, `Identity` | `Organization`, or `Mailbox` with its user principal name |
| `AuditDisabled` | Organization row. `False` means mailbox auditing on by default is on, and overrides a mailbox-level off |
| `UnifiedAuditLogIngestionEnabled` | Organization row. Read through Exchange Online; always `False` in Security & Compliance PowerShell |
| `DefaultAuditSet`, `AuditAdmin`, `AuditDelegate`, `AuditOwner` | Mailbox rows. `Admin;Delegate;Owner` in `DefaultAuditSet` means the default actions |
