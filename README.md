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
```

## Reports

| Report | What it answers |
| --- | --- |
| [Guest and external access](reports/guest-access/README.md) | Who the guests are, who invited them, whether they ever signed in, what groups they are in, and what has been shared outside the organisation |

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
`*.onmicrosoft.com` domain, needed by the Exchange-based services).

## Requirements

- PowerShell 7
- The Microsoft Graph and Exchange Online modules each report's README lists
- [Pester 5](https://pester.dev) to run the tests

## Tests

```powershell
pwsh -NoProfile -Command "Invoke-Pester -Path ./shared/tests, ./reports/guest-access/tests -CI"
```

Every tenant call is mocked; the tests never reach a tenant. They also check that each
sample CSV has exactly the columns its collector produces, in the same order, so the
committed sample data cannot drift away from the collectors.

The same command runs on every pull request — see `.github/workflows/tests.yml`.

## Adding a report

1. `reports/<report-name>/` with `collectors/`, `samples/`, `tests/` and a `README.md`.
2. One collector per CSV, each writing a single file and taking `-OutputPath`,
   `-Environment` and the four authentication parameters. Add a `Run-All.ps1` that runs
   them in order.
3. Put the column order of every CSV in one data file the collectors, the sample
   generator and the tests all read, so the three cannot disagree.
4. Snapshot data gets a `RunDate` column and a composite key of `RunDate` plus the
   object's id. Event data is keyed on the source's own id, and resumes from
   `Get-CsvWatermark`.
5. A source that is missing from a cloud, or unlicensed in the tenant, writes its header
   and a line in `run.log`. It never stops the run and never writes a half-populated row.
6. A `New-SampleData.ps1` that is deterministic and uses only `example.com`, writing
   through the same column lists as the collectors.
7. Tests covering the column order against the samples, the endpoints per cloud, the
   watermark, and the header-only path. New scripts are picked up by the read-only scan
   automatically.
8. The report's README lists prerequisites, per-cloud availability with a Learn link or
   `UNVERIFIED`, how to run it, and what every column means.
9. Add the report to the table above and to the `Invoke-Pester -Path` list in
   `.github/workflows/tests.yml`.

## Licence

MIT. See [LICENSE](LICENSE).
