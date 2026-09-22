#Requires -Version 7.0

<#
    Behavioral checks the structural field-reference tests do not cover:
    how the model reads collector timestamps, how the date slicer is bound,
    how snapshot metrics follow that slicer, and how the anonymize toggle
    treats display names.
#>

# Populated at discovery time: Pester evaluates -ForEach while dot-sourcing
# this file, before any BeforeAll runs.
$script:Root = Join-Path $PSScriptRoot '../../..' | Resolve-Path | Select-Object -ExpandProperty Path
$script:ReportRoot = Join-Path $script:Root 'reports/guest-access/report'
$script:TablesFolder = Join-Path $script:ReportRoot 'GuestAccess.SemanticModel/definition/tables'
$script:PagesFolder = Join-Path $script:ReportRoot 'GuestAccess.Report/definition/pages'
$script:UtcColumns = @{
    'Guests.tmdl'           = @('CreatedDateTime', 'ExternalUserStateChangeDateTime', 'LastSignInDateTime', 'LastNonInteractiveSignInDateTime', 'LastSuccessfulSignInDateTime')
    'Users.tmdl'            = @('CreatedDateTime')
    'GuestInvitations.tmdl' = @('ActivityDateTime')
    'GuestSignIns.tmdl'     = @('CreatedDateTime')
    'SharingEvents.tmdl'    = @('CreationTime')
}

Describe 'Collector UTC timestamps are parsed with an explicit Z format' {
    It 'parses <File> column <Column> with DateTime.FromText' -ForEach @(
        foreach ($file in $script:UtcColumns.Keys) {
            foreach ($column in $script:UtcColumns[$file]) {
                @{ File = $file; Column = $column }
            }
        }
    ) {
        $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder $File) -Raw
        $tmdl | Should -Match 'DateTime\.FromText'
        $tmdl | Should -Match "yyyy-MM-dd'T'HH:mm:ss'Z'"
        # A type-datetime cast of the raw text drops or rejects the Z designator.
        $tmdl | Should -Not -Match ('"' + [regex]::Escape($Column) + '",\s*type datetime')
        $tmdl | Should -Match ([regex]::Escape('{"' + $Column + '", ParseUtc, type datetime}'))
    }
}
