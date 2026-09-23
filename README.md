# M365 Report Library

PowerShell collectors that pull Microsoft 365 data into flat CSV files, and the Power BI
reports built on top of them.

Each report is a folder: a handful of read-only collectors, a sample data set so the
report can be built without a tenant, tests, and a README saying what the data means and
what it takes to collect it. A shared module underneath handles the things every report
needs — signing in to the right cloud, appending to CSVs without losing history, and
knowing where the last run stopped.

## Why CSV

A report is only useful if you can see it change. The collectors never overwrite: a
snapshot file gets a new block of rows tagged with `RunDate` each run, and an event file is
appended from the timestamp the last run reached. Point Power BI at the folder and you
have history, without a database to run.

## Layout

```
shared/
  M365ReportLibrary.psm1        the shared layer
  tests/                        its tests, plus the tenant cmdlet stubs every test uses
reports/
  guest-access/
    collectors/                 one script per CSV, plus Run-All.ps1
    samples/                    generated fake data
    tests/
    New-SampleData.ps1
    README.md
  purview-ip/
    collectors/                 one script per CSV, plus Run-All.ps1
    samples/                    generated fake data
    tests/
    New-SampleData.ps1
    README.md
```

## Reports

| Report | What it answers |
| --- | --- |
| [Guest and external access](reports/guest-access/README.md) | Who the guests are, who invited them, whether they ever signed in, what groups they are in, and what has been shared outside the organisation |
| [Purview information protection](reports/purview-ip/README.md) | What sensitivity labels, DLP and retention policies are configured, how they're being used, what Content Explorer counts, and what Microsoft 365 Copilot has accessed |

## The shared layer

`shared/M365ReportLibrary.psm1`:

| Function | What it does |
| --- | --- |
| `Connect-M365Service` | Signs in to Graph, Exchange Online or Security & Compliance PowerShell in the right cloud, interactively or app-only |
| `Get-M365ServiceEndpoint` | The endpoints a given service uses in a given cloud, separately from connecting, so they can be asserted |
| `Export-AppendCsv` | Writes the header once, appends rows, skips keys the file already holds, and throws if the columns have drifted |
| `Get-CsvWatermark` | The latest timestamp in an existing CSV — where the next run starts |
| `Get-CsvLatestSnapshot` | The rows of the most recent snapshot in a snapshot CSV |
| `Invoke-EntraUserCollector` | Writes `users.csv`, which every report joins to |
| `Write-CollectorLog` | Appends to `run.log` in the output folder |
| `Split-DateRange` | Splits a range into windows, so no one query asks a service for too much |

## Clouds

Every collector takes `-Environment Commercial|GCC|GCCHigh`, defaulting to `Commercial`.

| Service | Commercial and GCC | GCC High |
| --- | --- | --- |
| Microsoft Graph | `https://graph.microsoft.com` (`-Environment Global`) | `Connect-MgGraph -Environment USGov`, `https://graph.microsoft.us` |
| Exchange Online | `-ExchangeEnvironmentName O365Default` | `-ExchangeEnvironmentName O365USGovGCCHigh` |
| Security & Compliance | the cmdlet's defaults | `-ConnectionUri https://ps.compliance.protection.office365.us/powershell-liveid/ -AzureADAuthorizationEndpointUri https://login.microsoftonline.us/organizations` |

Sources: [Graph national cloud deployments](https://learn.microsoft.com/graph/deployments),
[Exchange Online app-only auth](https://learn.microsoft.com/powershell/exchange/app-only-auth-powershell-v2),
[Connect to Security & Compliance PowerShell](https://learn.microsoft.com/powershell/exchange/connect-to-scc-powershell).

## Read-only

Against a tenant this library calls nothing but `Get-*`, `Search-*`, `Connect-*` and
`Disconnect-*`, and issues no HTTP request other than GET. That is not a convention — a
test parses every script in the repository and fails on a call that breaks it, so a
collector that could change something cannot be merged.

## Authentication

Interactive sign-in by default, which asks only for read scopes. For unattended runs, pass
`-AppId`, `-CertificateThumbprint`, `-TenantId`, and `-Organization` (the tenant's
`*.onmicrosoft.com` domain, needed by the Exchange-based services). App-only Exchange
Online also needs the `Exchange.ManageAsApp` application permission and the audit role
on the service principal — see each report's README.

## Requirements

- PowerShell 7
- The Microsoft Graph and Exchange Online modules each report's README lists
- [Pester 5](https://pester.dev) to run the tests

## Tests

```powershell
pwsh -NoProfile -Command "Invoke-Pester -Path ./shared/tests, ./reports/guest-access/tests, ./reports/purview-ip/tests, ./.github/scripts/tests -CI"
```

Every tenant call is mocked; the tests never reach a tenant. They also check that each
sample CSV has exactly the columns its collector produces, in the same order, so the
committed sample data cannot drift away from the collectors.

The same command runs on every pull request — see `.github/workflows/tests.yml`.

### Running tests in a cloud agent session

CI installs Pester from the PowerShell Gallery as usual — that is unchanged. A scheduled
or on-demand Claude Code cloud session cannot: its egress proxy denies
`www.powershellgallery.com` and `codeload.github.com`, so `Install-Module Pester` fails
there. `.claude/hooks/session-start.sh` is a `SessionStart` hook (registered in
`.claude/settings.json`, [documented
here](https://code.claude.com/docs/en/cloud-environments#install-dependencies-with-a-sessionstart-hook))
that installs PowerShell 7 and Pester 5 from sources the proxy does allow — a GitHub
release for `pwsh`, and the `api.nuget.org` flat-container feed for the `pester` NuGet
package, whose `tools/` directory is the module — so `Invoke-Pester` works in that
session too. It only runs when `CLAUDE_CODE_REMOTE` is `true`, skips anything already
installed, and is safe to run more than once.

To pin a different Pester version, edit `PESTER_VERSION` in the hook script. Check
`https://api.nuget.org/v3-flatcontainer/pester/index.json` first — nuget.org does not
mirror every version PSGallery has.

## Merging

Cursor's review automation ends every run by pushing an empty commit to the pull request
branch, with Joshua's GitHub account. `.github/workflows/auto-merge.yml` watches that
commit: when its first line is exactly `Verdict: Ready to merge` and every other check
run or commit status on it has passed, the workflow marks the pull request ready for
review (if it was a draft), squash-merges it with the pull request title as the commit
title, and deletes the branch. `Verdict: Not ready`, a failed check, a skipped or
neutral check, or no other check at all leaves the pull request alone; a check that is
still running makes it wait. GitHub does not deliver `check_suite` for a suite created
by GitHub Actions, so the workflow runs again when the `tests` workflow finishes, and
when any other check suite finishes. The decision itself lives in
`.github/scripts/Get-MergeDecision.ps1`, which the workflow calls and which
`.github/scripts/tests/` tests on its own.

## Adding a report

1. `reports/<report-name>/` with `collectors/`, `samples/`, `tests/` and a `README.md`.
2. One collector per CSV, each writing a single file and taking `-OutputPath`,
   `-Environment` and the four authentication parameters. Add a `Run-All.ps1` that runs
   them in order.
3. Import the shared module by its relative path, never with `-Force`. From `collectors/`
   that is
   `Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')` —
   `-Force` removes the loaded module and imports it again, which drops any Pester mock a caller
   installed against it before running the collector. Without `-Force`, a session that already
   imported the module keeps that copy, so an edit to `shared/M365ReportLibrary.psm1` is not
   loaded until a new session. Scripts under `tests/` are outside this rule: they import with
   `-Force` so the file on disk is the one under test, and they install mocks only after that.
4. Put the column order of every CSV in one data file the collectors, the sample
   generator and the tests all read, so the three cannot disagree.
5. Snapshot data gets a `RunDate` column and a composite key of `RunDate` plus the
   object's id. Event data is keyed on the source's own id, and resumes from
   `Get-CsvWatermark`.
6. A source that is missing from a cloud, or unlicensed in the tenant, writes its header
   and a line in `run.log`. It never stops the run and never writes a half-populated row.
7. A `New-SampleData.ps1` that is deterministic and uses only `example.com`, writing
   through the same column lists as the collectors.
8. Tests covering the column order against the samples, the endpoints per cloud, the
   watermark, and the header-only path. New scripts are picked up by the read-only scan
   automatically.
9. The report's README lists prerequisites, per-cloud availability with a Learn link or
   `UNVERIFIED`, how to run it, and what every column means.
10. Add the report to the table above and to the `Invoke-Pester -Path` list in
    `.github/workflows/tests.yml` (and in this README's own copy of that command, above).

Why relative-path import rather than a module manifest on `PSModulePath`: putting
`shared/M365ReportLibrary.psm1` on `PSModulePath` would let a collector write
`Import-Module M365ReportLibrary` by name, but only after every machine (and CI job) that
runs a collector adds this repo's `shared/` folder to `PSModulePath` first — an extra setup
step the current path-based import needs from nobody, since a fresh `git clone` already has
everything `$PSScriptRoot`-relative imports need. It also keeps each collector's import
statement an unambiguous pointer to exactly one module file, with no possibility of a
same-named module elsewhere on `PSModulePath` shadowing it. Not worth the trade for a
single-repo shared module; revisit if the shared module ever ships to more than one repo.

## Licence

MIT. See [LICENSE](LICENSE).
