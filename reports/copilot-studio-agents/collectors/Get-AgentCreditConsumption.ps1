#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes agent-credit-consumption.csv: Copilot Credits consumed by each Copilot Studio agent, read
        from the consumption report an admin downloads from the Power Platform admin center.

    .DESCRIPTION
        Nothing is called in the tenant. An admin downloads the detailed consumption report for each
        environment (Power Platform admin center, Licensing, Copilot Studio, Environments, select the
        environment, Copilot Credits capacity, "Copilot credit consumption details", Download) and saves
        the CSV into -ConsumptionFolder. Name each file after its environment: the report names an agent,
        not the environment, so the file name is the environment unless the report has an Environment column.
        https://learn.microsoft.com/power-platform/admin/manage-copilot-studio-copilot-credits-capacity#download-detailed-copilot-consumption-for-an-environment

        The Learn page lists these report columns: Agent name; Product / AI feature; Channel; LLM model;
        Knowledge sources / Tool used; Billed credits / Non-billed credits. It does not give the file
        format, the exact header text, a date column or an environment column. So this reads *.csv only,
        matches headers ignoring case, spaces and punctuation, and treats the date and environment columns
        as UNVERIFIED: when the report has no date column, Date is empty and run.log says the rows are
        totals for the downloaded timeframe, not per day. A column that cannot be found is left empty and
        named in run.log.

        Rows with the same date, environment, agent, product, feature, channel, model and knowledge source or
        tool are summed within a file. A row already in agent-credit-consumption.csv with the same key is not
        written again, so overlapping downloads do not double count, and a re-download does not update a
        number already written.

        If the folder is missing or holds no CSV, the file is header only and run.log says the download is
        missing. No supported API for per-agent consumption was found: the Power Platform API entitlement
        insight operations return a resource id, a consumed value and a metadata object, with no documented
        entitlement id for Copilot Credits, agent name, channel or day.

    .PARAMETER ConsumptionFolder
        The folder holding the downloaded consumption reports (*.csv).

    .EXAMPLE
        ./Get-AgentCreditConsumption.ps1 -OutputPath ./out -ConsumptionFolder ./downloads
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$OutputPath,

    [ValidateSet('Commercial', 'GCC', 'GCCHigh')]
    [string]$Environment = 'Commercial',

    [string]$ConsumptionFolder
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Imported by path without -Force: a forced re-import builds a fresh module session
# state, which would discard anything the caller has arranged around the module - and
# that is exactly how the tests put mocks in front of the tenant cmdlets.
Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

. (Join-Path $PSScriptRoot 'CopilotStudioHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'CopilotStudioSchema.psd1')
$columns = $schema.AgentCreditConsumption
$source = 'agent-credit-consumption'
$csvPath = Join-Path $OutputPath 'agent-credit-consumption.csv'

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$availability = Get-CopilotStudioSourceAvailability -Source 'AgentCreditConsumption' -Environment $Environment -Schema $schema
if ($availability.ShouldSkip) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        "Skipping agent-credit-consumption.csv. $($availability.Reason) $($availability.Reference)")
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}
if ($availability.Status -eq 'Unverified') {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message $availability.Reason
}

$howToDownload = 'Download it from the Power Platform admin center (Licensing, Copilot Studio, Environments, select the environment, Copilot credit consumption details, Download), save it as a CSV named after the environment, and pass the folder as -ConsumptionFolder.'

$files = @()
if (-not [string]::IsNullOrWhiteSpace($ConsumptionFolder) -and (Test-Path -LiteralPath $ConsumptionFolder -PathType Container)) {
    $files = @(Get-ChildItem -LiteralPath $ConsumptionFolder -File | Sort-Object Name)
}
$csvFiles = @($files | Where-Object { $_.Extension -eq '.csv' })

if ($csvFiles.Count -eq 0) {
    $where = if ([string]::IsNullOrWhiteSpace($ConsumptionFolder)) { 'No -ConsumptionFolder was given.' }
    elseif (-not (Test-Path -LiteralPath $ConsumptionFolder -PathType Container)) { "The folder '$ConsumptionFolder' does not exist." }
    else { "The folder '$ConsumptionFolder' holds no .csv file." }
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        'The Copilot Studio credit consumption download is missing. {0} {1} Writing the header only.' -f $where, $howToDownload)
    Export-AppendCsv -Path $csvPath -Column $columns
    return
}

$notRead = @($files | Where-Object { $_.Extension -ne '.csv' })
if ($notRead.Count -gt 0) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        'Not read, because only .csv files are read: {0}. Save the report as CSV if the admin center offered another format.' -f (($notRead.Name) -join ', '))
}

function ConvertTo-HeaderKey {
    param([string]$Name)
    return ($Name.ToLowerInvariant() -replace '[^a-z0-9]', '')
}

# Output column -> accepted header keys, in order of preference. The first group are the report columns
# the Learn page lists; the rest are guesses and UNVERIFIED.
$aliases = [ordered]@{
    Date                  = @('date', 'day', 'usagedate', 'reportdate', 'consumptiondate')
    EnvironmentName       = @('environment', 'environmentname')
    AgentName             = @('agentname', 'agent')
    Product               = @('product', 'productname')
    Feature               = @('aifeature', 'productaifeature', 'feature', 'featurename', 'productfeature')
    Channel               = @('channel')
    Model                 = @('llmmodel', 'model')
    KnowledgeSourceOrTool = @('knowledgesourcestoolused', 'knowledgesourcetoolused', 'knowledgesourceortool', 'toolused', 'knowledgesource')
    BilledCredits         = @('billedcredits', 'billed')
    NonBilledCredits      = @('nonbilledcredits', 'nonbilled', 'notbilledcredits', 'nonbillablecredits')
}
$required = 'AgentName', 'BilledCredits', 'NonBilledCredits'
$keyColumns = 'Date', 'EnvironmentName', 'AgentName', 'Product', 'Feature', 'Channel', 'Model', 'KnowledgeSourceOrTool'

function ConvertTo-CreditNumber {
    param([string]$Text)
    $number = 0.0
    $clean = $Text.Trim() -replace ',', ''
    if ([double]::TryParse($clean, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
        return $number
    }
    return $null
}

$allRows = [System.Collections.Generic.List[object]]::new()

foreach ($file in $csvFiles) {
    try {
        $records = @(Import-Csv -LiteralPath $file.FullName)
    }
    catch {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            '{0} could not be read as CSV ({1}). It was skipped.' -f $file.Name, $_.Exception.Message)
        continue
    }
    if ($records.Count -eq 0) {
        Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message ('{0} has no data rows. It was skipped.' -f $file.Name)
        continue
    }

    $headerByKey = @{}
    foreach ($name in $records[0].PSObject.Properties.Name) {
        $key = ConvertTo-HeaderKey $name
        if (-not $headerByKey.ContainsKey($key)) { $headerByKey[$key] = $name }
    }
    $map = @{}
    foreach ($column in $aliases.Keys) {
        $found = $aliases[$column] | Where-Object { $headerByKey.ContainsKey($_) } | Select-Object -First 1
        if ($found) { $map[$column] = $headerByKey[$found] }
    }

    $missingRequired = @($required | Where-Object { -not $map.ContainsKey($_) })
    if ($missingRequired.Count -gt 0) {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $source -Message (
            '{0} does not look like the Copilot credit consumption report: no column found for {1}. Columns found: {2}. It was skipped.' -f `
                $file.Name, ($missingRequired -join ', '), (($records[0].PSObject.Properties.Name) -join ', '))
        continue
    }
    $missingOptional = @($aliases.Keys | Where-Object { -not $map.ContainsKey($_) })
    if ($missingOptional.Count -gt 0) {
        $note = 'in {0}: no column found for {1}; those columns are left empty.' -f $file.Name, ($missingOptional -join ', ')
        if ($missingOptional -contains 'Date') {
            $note += ' With no date column the rows are totals for the downloaded timeframe, not per day.'
        }
        Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message $note
    }

    $fileEnvironment = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
    $groups = [ordered]@{}
    foreach ($record in $records) {
        $row = [ordered]@{}
        foreach ($column in $aliases.Keys) {
            $row[$column] = if ($map.ContainsKey($column)) { [string]$record.($map[$column]) } else { '' }
            $row[$column] = $row[$column].Trim()
        }
        if ($row['Date']) {
            $parsed = [datetime]::MinValue
            if ([datetime]::TryParse($row['Date'], [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal, [ref]$parsed)) {
                $row['Date'] = $parsed.ToString('yyyy-MM-dd')
            }
        }
        if (-not $row['EnvironmentName']) { $row['EnvironmentName'] = $fileEnvironment }
        if (-not $row['AgentName']) { continue }

        $groupKey = ($keyColumns | ForEach-Object { $row[$_] }) -join "`u{1F}"
        if (-not $groups.Contains($groupKey)) {
            $groups[$groupKey] = [pscustomobject]@{ Row = $row; Billed = 0.0; NonBilled = 0.0; BilledText = $null; NonBilledText = $null }
        }
        $group = $groups[$groupKey]
        foreach ($pair in @(@('BilledCredits', 'Billed', 'BilledText'), @('NonBilledCredits', 'NonBilled', 'NonBilledText'))) {
            $text = $row[$pair[0]]
            $number = ConvertTo-CreditNumber $text
            if ($null -ne $number) { $group.($pair[1]) += $number }
            elseif ($text -and $null -eq $group.($pair[2])) { $group.($pair[2]) = $text }
        }
    }

    foreach ($group in $groups.Values) {
        $row = $group.Row
        $billed = if ($null -ne $group.BilledText) { $group.BilledText } elseif ($row['BilledCredits'] -eq '') { '' } else { [string]$group.Billed.ToString('0.####', [System.Globalization.CultureInfo]::InvariantCulture) }
        $nonBilled = if ($null -ne $group.NonBilledText) { $group.NonBilledText } elseif ($row['NonBilledCredits'] -eq '') { '' } else { [string]$group.NonBilled.ToString('0.####', [System.Globalization.CultureInfo]::InvariantCulture) }
        $allRows.Add([pscustomobject]@{
                Date                  = $row['Date']
                EnvironmentName       = $row['EnvironmentName']
                AgentName             = $row['AgentName']
                Product               = $row['Product']
                Feature               = $row['Feature']
                Channel               = $row['Channel']
                Model                 = $row['Model']
                KnowledgeSourceOrTool = $row['KnowledgeSourceOrTool']
                BilledCredits         = $billed
                NonBilledCredits      = $nonBilled
                SourceFile            = $file.Name
            })
    }
}

$result = Export-AppendCsv -Path $csvPath -Rows @($allRows) -Column $columns -KeyColumn $keyColumns -PassThru
if ($allRows.Count -eq 0) {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source $source -Message (
        'No file in {0} held consumption rows. Writing the header only.' -f $ConsumptionFolder)
}
Write-CollectorLog -OutputPath $OutputPath -Source $source -Message (
    'agent-credit-consumption.csv: {0} rows written, {1} skipped, from {2} file(s).' -f $result.Written, $result.Skipped, $csvFiles.Count)
