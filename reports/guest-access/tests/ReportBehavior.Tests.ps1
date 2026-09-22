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
BeforeAll {
    $script:Root = Join-Path $PSScriptRoot '../../..' | Resolve-Path | Select-Object -ExpandProperty Path
    $script:ReportRoot = Join-Path $script:Root 'reports/guest-access/report'
    $script:TablesFolder = Join-Path $script:ReportRoot 'GuestAccess.SemanticModel/definition/tables'
    $script:PagesFolder = Join-Path $script:ReportRoot 'GuestAccess.Report/definition/pages'
}

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

Describe 'The date slicer is a date range on DateDim[Date]' {
    It 'formats YearMonth with the documented four-digit year token' {
        $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'DateDim.tmdl') -Raw
        $tmdl | Should -MatchExactly 'FORMAT\(DateDim\[Date\], "yyyy-MM"\)'
        ($tmdl -cmatch 'FORMAT\(DateDim\[Date\], "YYYY-MM"\)') | Should -BeFalse
    }

    It 'binds <Page> to DateDim[Date] in Between mode' -ForEach @(
        Get-ChildItem -LiteralPath $script:PagesFolder -Directory | ForEach-Object {
            @{ Page = $_.Name }
        }
    ) {
        $path = Join-Path $script:PagesFolder "$Page/visuals/slicer-date-range/visual.json"
        $visual = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
        $projection = $visual.visual.query.queryState.Values.projections[0]
        $projection.field.Column.Expression.SourceRef.Entity | Should -Be 'DateDim'
        $projection.field.Column.Property | Should -Be 'Date'
        $visual.visual.objects.data[0].properties.mode.expr.Literal.Value | Should -Be "'Between'"
    }
}

Describe 'The anonymize toggle covers every displayed name' {
    It 'does not bind a visual to a raw DisplayName column' {
        $leaks = foreach ($file in (Get-ChildItem -LiteralPath $script:PagesFolder -Filter 'visual.json' -Recurse -File)) {
            $json = Get-Content -LiteralPath $file.FullName -Raw | ConvertFrom-Json
            $raw = $json | ConvertTo-Json -Depth 30
            if ($raw -match '"Property":\s*"DisplayName"') {
                $file.FullName.Substring($script:PagesFolder.Length + 1)
            }
        }
        $leaks -join '; ' | Should -BeNullOrEmpty
    }

    It 'filters the guest slicer on Pseudonym' {
        $slicers = @(Get-ChildItem -LiteralPath $script:PagesFolder -Recurse -Filter 'visual.json' |
                Where-Object { $_.Directory.Name -eq 'slicer-guest' })
        $slicers.Count | Should -Be 7
        foreach ($file in $slicers) {
            $json = Get-Content -LiteralPath $file.FullName -Raw | ConvertFrom-Json
            $json.visual.query.queryState.Values.projections[0].field.Column.Property | Should -Be 'Pseudonym'
        }
    }

    It 'shows member names through Member Display Name' {
        foreach ($relative in @(
                'sharing/visuals/table-by-sharer/visual.json',
                'invitations/visuals/table-top-inviters/visual.json'
            )) {
            $json = Get-Content -LiteralPath (Join-Path $script:PagesFolder $relative) -Raw
            $json | Should -Match 'Member Display Name'
            $json | Should -Match 'Pseudonym'
        }
    }
}

Describe 'Snapshot metrics follow the date slicer and stay on UTC dates' {
    It 'does not measure age with TODAY()' {
        $hits = Get-ChildItem -LiteralPath $script:TablesFolder -Filter '*.tmdl' |
            Select-String -Pattern 'TODAY\(' -SimpleMatch:$false
        @($hits).Count | Should -Be 0
    }

    It 'counts guests on the snapshot inside the selected date range' {
        $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'Guests.tmdl') -Raw
        $tmdl | Should -Match "measure 'Selected Snapshot Date'"
        $tmdl | Should -Match 'Guests\[RunDate\] = SnapshotDate'
        $tmdl | Should -Match 'DATEDIFF\(Guests\[CreatedDateTime\], Guests\[RunDate\], DAY\)'
        $tmdl | Should -Match 'DATEDIFF\(d, Guests\[RunDate\], DAY\)'
        $tmdl | Should -Match 'GuestSignIns\[CreatedDateTime\] < AsOfEnd'
    }

    It 'counts memberships on the snapshot inside the selected date range' {
        $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'GuestMemberships.tmdl') -Raw
        $tmdl | Should -Match 'GuestMemberships\[RunDate\] = SnapshotDate'
    }

    It 'does not treat a blank sign-in error code as a failure' {
        $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'GuestSignIns.tmdl') -Raw
        $tmdl | Should -Match 'NOT ISBLANK\(GuestSignIns\[ErrorCode\]\)'
        $tmdl | Should -Match 'GuestSignIns\[ErrorCode\] <> 0'
    }
}
