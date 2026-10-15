#Requires -Version 7.0

<#
    The Power BI report (reports/identity-posture/report/) is built from the same CSVs
    the collectors write. These tests hold the report to that:

    - every table, column and measure a visual's query references must exist in the
      TMDL semantic model (a mistyped field reference is silently blank in Power BI,
      not an error, so nothing else would catch it);
    - every semantic-model table sourced directly from a CSV must have exactly the
      columns that CSV has, in the same order, for the populated sample and for the
      header-only file a tenant without the licence gets;
    - every relationship names columns that exist, and the report builds no page on
      the two beta-only sources the README lists as not collected;
    - the report has the one approved page (BRO-501: Accounts to fix), the report
      filter pane carries the anonymize toggle, and names reach a visual only through
      the toggle.

    UsersCurrent (latest-snapshot import of users.csv), the calculated RoleDefinitions, the
    'User Display' import table built from the UsersCurrent query (BRO-501, chart axis
    names) and the helper tables that are not sourced from any CSV (DateDim and the
    AnonymizeMode toggle) are exempt from the CSV-mapping check. $CsvBackedTables below is
    the complete list of what is checked.
#>

BeforeAll {
    $script:Root = Join-Path $PSScriptRoot '../../..' | Resolve-Path | Select-Object -ExpandProperty Path
    $script:Samples = Join-Path $script:Root 'reports/identity-posture/samples'
    $script:ReportRoot = Join-Path $script:Root 'reports/identity-posture/report'
    $script:DefinitionFolder = Join-Path $script:ReportRoot 'IdentityPosture.SemanticModel/definition'
    $script:TablesFolder = Join-Path $script:DefinitionFolder 'tables'
    $script:PagesFolder = Join-Path $script:ReportRoot 'IdentityPosture.Report/definition/pages'

    Import-Module (Join-Path $script:Root 'shared/M365ReportLibrary.psm1') -Force
    . (Join-Path $PSScriptRoot '../../guest-access/tests/TmdlModel.ps1')

    $script:Model = Get-TmdlModel -TablesFolder $script:TablesFolder

    # Table name -> the sample CSV its `m` partition imports.
    $script:CsvBackedTables = [ordered]@{
        Users                     = 'users.csv'
        AuthenticationMethods     = 'authentication-methods.csv'
        ConditionalAccessPolicies = 'conditional-access-policies.csv'
        ActiveRoleAssignments     = 'role-assignments-active.csv'
        EligibleRoleAssignments   = 'role-assignments-eligible.csv'
        RoleAssignments           = 'role-assignments.csv'
        UserSignInActivity        = 'user-signin-activity.csv'
        RiskyUsers                = 'risky-users.csv'
        SignIns                   = 'signins.csv'
    }

    function Get-VisualFieldReference {
        <#
            .SYNOPSIS
                Every {Entity, Property, Kind} a visual.json's query pulls from the
                semantic model, found by walking the parsed JSON for `Column` /
                `Measure` field containers (semanticQuery schema: a field is
                `{ "Column": { "Expression": { "SourceRef": { "Entity": ... } },
                "Property": ... } }`, or the same shape under `"Measure"`).
        #>
        param([Parameter(Mandatory)]$Node)

        $results = [System.Collections.Generic.List[hashtable]]::new()

        function Walk($n) {
            if ($n -is [System.Collections.IDictionary]) {
                foreach ($kind in 'Column', 'Measure') {
                    if ($n.ContainsKey($kind)) {
                        $inner = $n[$kind]
                        $entity = $inner.Expression.SourceRef.Entity
                        $property = $inner.Property
                        if ($entity -and $property) {
                            $results.Add(@{ Kind = $kind; Entity = $entity; Property = $property })
                        }
                    }
                }
                foreach ($key in $n.Keys) { Walk $n[$key] }
            }
            elseif (($n -is [System.Collections.IEnumerable]) -and ($n -isnot [string])) {
                foreach ($item in $n) { Walk $item }
            }
        }

        Walk $Node
        return , $results.ToArray()
    }

    $script:VisualFiles = @(Get-ChildItem -LiteralPath $script:PagesFolder -Filter 'visual.json' -Recurse -File)
}

AfterAll {
    Remove-Module M365ReportLibrary -Force -ErrorAction SilentlyContinue
}

Describe 'The semantic model has tables to check' {
    It 'parses the nine CSV-backed tables plus the calculated and helper tables' {
        $script:Model.Keys.Count | Should -Be 14
        foreach ($name in $script:CsvBackedTables.Keys + @('UsersCurrent', 'RoleDefinitions', 'DateDim', 'AnonymizeMode', 'User Display')) {
            $script:Model.ContainsKey($name) | Should -BeTrue -Because "the model should have a $name table"
        }
    }

    It 'finds visual.json files to check' {
        $script:VisualFiles.Count | Should -BeGreaterThan 50
    }

    It 'lists every table in model.tmdl' {
        $refs = @(Get-Content -LiteralPath (Join-Path $script:DefinitionFolder 'model.tmdl') |
                Select-String -Pattern '^ref table (.+)$' | ForEach-Object { ConvertFrom-TmdlName -Name $_.Matches[0].Groups[1].Value })
        ($refs | Sort-Object) -join ',' | Should -Be (($script:Model.Keys | Sort-Object) -join ',')
    }
}

Describe 'Every visual field reference resolves against the TMDL model' {
    BeforeAll {
        $script:AllReferences = foreach ($file in $script:VisualFiles) {
            $json = Get-Content -LiteralPath $file.FullName -Raw | ConvertFrom-Json -AsHashtable
            foreach ($ref in (Get-VisualFieldReference -Node $json)) {
                [pscustomobject]@{
                    File     = $file.FullName.Substring($script:ReportRoot.Length + 1)
                    Kind     = $ref.Kind
                    Entity   = $ref.Entity
                    Property = $ref.Property
                }
            }
        }
    }

    It 'has field references to check' {
        # Guards the tests below against silently passing because nothing was found.
        # BRO-501 replaced the seven old pages with one approved page (about 45 references).
        $script:AllReferences.Count | Should -BeGreaterThan 30
    }

    It 'references only tables that exist in the model' {
        $missing = @($script:AllReferences | Where-Object { -not $script:Model.ContainsKey($_.Entity) } |
                ForEach-Object { '{0}: {1}' -f $_.File, $_.Entity } | Sort-Object -Unique)
        $missing -join '; ' | Should -BeNullOrEmpty
    }

    It 'references only columns that exist on their table' {
        $missing = @($script:AllReferences | Where-Object { $_.Kind -eq 'Column' } | ForEach-Object {
                $table = $script:Model[$_.Entity]
                if ($table -and ($table.Columns.Name -notcontains $_.Property)) {
                    '{0}: {1}.{2}' -f $_.File, $_.Entity, $_.Property
                }
            } | Sort-Object -Unique)
        $missing -join '; ' | Should -BeNullOrEmpty
    }

    It 'references only measures that exist on their table' {
        $missing = @($script:AllReferences | Where-Object { $_.Kind -eq 'Measure' } | ForEach-Object {
                $table = $script:Model[$_.Entity]
                if ($table -and ($table.Measures -notcontains $_.Property)) {
                    '{0}: {1}.{2}' -f $_.File, $_.Entity, $_.Property
                }
            } | Sort-Object -Unique)
        $missing -join '; ' | Should -BeNullOrEmpty
    }

    It 'matches queryRef to the Entity.Property it projects' {
        $bad = [System.Collections.Generic.List[string]]::new()
        foreach ($file in $script:VisualFiles) {
            $json = Get-Content -LiteralPath $file.FullName -Raw | ConvertFrom-Json -AsHashtable
            foreach ($role in $json.visual.query.queryState.Values) {
                foreach ($projection in $role.projections) {
                    $inner = if ($projection.field.ContainsKey('Column')) { $projection.field.Column } else { $projection.field.Measure }
                    $expected = '{0}.{1}' -f $inner.Expression.SourceRef.Entity, $inner.Property
                    if ($projection.queryRef -ne $expected) { $bad.Add("$($file.Name): $($projection.queryRef) != $expected") }
                }
            }
        }
        $bad -join '; ' | Should -BeNullOrEmpty
    }

    It 'gives every page a visual that uses the anonymized display name or only pseudonyms' {
        # No page may project a raw name or sign-in name column. Display names reach a
        # visual only through the 'User Display Name' measure, which honours the toggle.
        $raw = @($script:AllReferences | Where-Object {
                $_.Kind -eq 'Column' -and $_.Property -in 'DisplayName', 'UserDisplayName', 'UserPrincipalName', 'Mail', 'ManagerUserPrincipalName' -and $_.Entity -ne 'ConditionalAccessPolicies'
            } | ForEach-Object { '{0}: {1}.{2}' -f $_.File, $_.Entity, $_.Property } | Sort-Object -Unique)
        $raw -join '; ' | Should -BeNullOrEmpty
    }
}

Describe 'Every CSV-backed table matches its sample CSV' {
    It 'checks <Table>' -ForEach @(
        @{ Table = 'Users' }
        @{ Table = 'AuthenticationMethods' }
        @{ Table = 'ConditionalAccessPolicies' }
        @{ Table = 'ActiveRoleAssignments' }
        @{ Table = 'EligibleRoleAssignments' }
        @{ Table = 'RoleAssignments' }
        @{ Table = 'UserSignInActivity' }
        @{ Table = 'RiskyUsers' }
        @{ Table = 'SignIns' }
    ) {
        $script:Model.ContainsKey($Table) | Should -BeTrue -Because "the model should have a $Table table"

        $csvPath = Join-Path $script:Samples $script:CsvBackedTables[$Table]
        $csvHeader = Get-CsvHeaderColumn -Path $csvPath
        $csvHeader | Should -Not -BeNullOrEmpty -Because "$csvPath should be readable"

        # Only the imported (sourceColumn-backed) columns, in file order: a calculated
        # column layered on a CSV-backed table (DaysSinceLastSuccessfulSignIn, Outcome)
        # is not part of the CSV and is excluded, as it would be from an Import-Csv header.
        $modelColumns = @($script:Model[$Table].Columns | Where-Object { -not $_.IsCalculated } |
                ForEach-Object { $_.SourceColumn })

        ($modelColumns -join ',') | Should -Be ($csvHeader -join ',')
    }

    It 'names every CSV-backed table exactly once' {
        $script:CsvBackedTables.Keys.Count | Should -Be 9
        (@($script:CsvBackedTables.Values) | Sort-Object -Unique).Count | Should -Be 9
    }

    It 'reads <Table> from the CsvFolder parameter and the CSV named in the map' -ForEach @(
        @{ Table = 'Users' }, @{ Table = 'AuthenticationMethods' }, @{ Table = 'ConditionalAccessPolicies' },
        @{ Table = 'ActiveRoleAssignments' }, @{ Table = 'EligibleRoleAssignments' }, @{ Table = 'RoleAssignments' },
        @{ Table = 'UserSignInActivity' }, @{ Table = 'RiskyUsers' }, @{ Table = 'SignIns' }
    ) {
        $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder "$Table.tmdl") -Raw
        $tmdl | Should -Match ([regex]::Escape('File.Contents(CsvFolder & "\' + $script:CsvBackedTables[$Table] + '")'))
    }

    It 'keeps the CsvFolder parameter as the only parameter query' {
        $expressions = Get-Content -LiteralPath (Join-Path $script:DefinitionFolder 'expressions.tmdl') -Raw
        ([regex]::Matches($expressions, '(?m)^expression ')).Count | Should -Be 1
        $expressions | Should -Match '(?m)^expression CsvFolder = '
    }
}

Describe 'The header-only files of an unlicensed tenant still fit the model' {
    # samples/unlicensed/ holds what a tenant without Entra ID P2 or Governance gets.
    # Power BI reads a header-only file as a table with no rows, so the columns must match.
    It 'matches the model columns of <Table> to samples/unlicensed/<Csv>' -ForEach @(
        @{ Table = 'ActiveRoleAssignments'; Csv = 'role-assignments-active.csv' }
        @{ Table = 'EligibleRoleAssignments'; Csv = 'role-assignments-eligible.csv' }
        @{ Table = 'RiskyUsers'; Csv = 'risky-users.csv' }
    ) {
        $path = Join-Path $script:Samples "unlicensed/$Csv"
        $header = Get-CsvHeaderColumn -Path $path
        $header | Should -Not -BeNullOrEmpty
        @(Get-Content -LiteralPath $path).Count | Should -Be 1 -Because 'an unlicensed file is the header only'

        $modelColumns = @($script:Model[$Table].Columns | Where-Object { -not $_.IsCalculated } | ForEach-Object { $_.SourceColumn })
        ($modelColumns -join ',') | Should -Be ($header -join ',')
    }

    It 'leaves a source with no rows blank, never zero, in the snapshot measures' {
        # Every count reads its snapshot date first; a source with no rows has none.
        foreach ($name in 'AuthenticationMethods', 'ConditionalAccessPolicies', 'ActiveRoleAssignments', 'EligibleRoleAssignments', 'RoleAssignments', 'UserSignInActivity', 'RiskyUsers', 'Users') {
            $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder "$name.tmdl") -Raw
            $tmdl | Should -Match 'IF\(ISBLANK\(Snap\), BLANK\(\)' -Because "$name measures must return blank when the file is header-only"
        }
    }

    It 'leaves legacy sign-in counts blank when signins.csv has no rows' {
        # COUNTROWS of an empty table is BLANK, and BLANK() + 0 is 0
        # (https://learn.microsoft.com/dax/dax-operator-reference). A header-only
        # signins.csv is not collected, so the empty check has to win.
        $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'SignIns.tmdl') -Raw
        foreach ($measure in 'Legacy Sign-Ins', 'Legacy Sign-Ins Succeeded', 'Legacy Sign-Ins Failed or Blocked', 'Legacy Sign-Ins Outcome Unknown', 'Legacy Sign-In Users') {
            $tmdl | Should -Match ([regex]::Escape("measure '$measure' = IF(COUNTROWS(ALL(SignIns)) = 0, BLANK(),")) -Because "$measure must be blank when the file has no rows"
        }
    }

    It 'falls back to role-assignments.csv for active holders when PIM data is absent' {
        # Source 4c replaces 4a only when the PIM schedule API was not collected.
        # A slicer that merely excludes the PIM snapshot must not switch files.
        # https://learn.microsoft.com/graph/api/rbacapplication-list-roleassignments
        $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'ActiveRoleAssignments.tmdl') -Raw
        $tmdl | Should -Match ([regex]::Escape('IF(NOT ISBLANK([Active Assignments]), [Active Assignments], IF(COUNTROWS(ALL(ActiveRoleAssignments)) = 0, [Active Assignments (No PIM)], BLANK()))'))
        $tmdl | Should -Match ([regex]::Escape('IF(NOT ISBLANK([Active Role Holders]), [Active Role Holders], IF(COUNTROWS(ALL(ActiveRoleAssignments)) = 0, [Active Role Holders (No PIM)], BLANK()))'))
        $tmdl | Should -Not -Match ([regex]::Escape('IF(NOT ISBLANK([Active Assignments]), [Active Assignments], [Active Assignments (No PIM)])'))
    }
}

Describe 'Empty values are Unknown, never zero' {
    It 'derives days since last sign-in as blank when the timestamp is empty' {
        $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'UserSignInActivity.tmdl') -Raw
        $tmdl | Should -Match ([regex]::Escape('IF(ISBLANK(UserSignInActivity[LastSuccessfulSignInDateTime]), BLANK(), DATEDIFF('))
        $tmdl | Should -Match 'IF\(ISBLANK\(d\), "Unknown"'
    }

    It 'reads 0001-01-01 sign-in timestamps as empty, not as a date' {
        # List users returns 0001-01-01T00:00:00Z when lastNonInteractiveSignInDateTime has no value.
        # https://learn.microsoft.com/graph/api/user-list#example-11-get-users-including-their-last-sign-in-time
        $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'UserSignInActivity.tmdl') -Raw
        $tmdl | Should -Match ([regex]::Escape('Text.StartsWith(value, "0001-01-01")'))
    }

    It 'counts unknown last sign-ins separately from stale ones' {
        $model = $script:Model['UserSignInActivity']
        $model.Measures | Should -Contain 'Unknown Last Sign-In'
        $model.Measures | Should -Contain 'Stale Accounts (90+ days)'
    }

    It 'never compares a nullable boolean to FALSE without excluding blanks' {
        # In DAX, BLANK() = FALSE() is true, so an empty flag would be counted as No.
        $offenders = [System.Collections.Generic.List[string]]::new()
        foreach ($file in Get-ChildItem -LiteralPath $script:TablesFolder -Filter '*.tmdl') {
            $lines = Get-Content -LiteralPath $file.FullName
            for ($i = 0; $i -lt $lines.Count; $i++) {
                if ($lines[$i] -match '(\w+\[\w+\]) = FALSE\(\)') {
                    $column = $Matches[1]
                    if ($lines[$i] -notmatch ('ISBLANK\(' + [regex]::Escape($column) + '\)')) {
                        $offenders.Add("$($file.Name):$($i + 1) $column")
                    }
                }
            }
        }
        $offenders -join '; ' | Should -BeNullOrEmpty
    }

    It 'labels risk Unknown only for a user the riskyUsers list actually returned' {
        # List riskyUsers returns users Entra has evaluated, not the directory.
        # unknownFutureValue is a real riskLevel; a missing row is not that value.
        # https://learn.microsoft.com/graph/api/resources/riskyuser
        $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'RiskyUsers.tmdl') -Raw
        $tmdl | Should -Match ([regex]::Escape("measure 'Listed On Snapshot' ="))
        ([regex]::Matches($tmdl, [regex]::Escape('IF(NOT (Listed > 0), BLANK(), IF(ISBLANK(V), "Unknown", V))'))).Count | Should -Be 3 -Because 'Risk Level, Risk State and Role Overlap must stay blank for a user the list did not return'
        # BRO-501: the Accounts to fix table draws risk from [Risk Level], so an account the
        # list did not return shows no pill, and only high, medium and low get one.
        $users = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'UsersCurrent.tmdl') -Raw
        $users | Should -Match "measure 'Fix Signals Image'[\s\S]*?VAR RiskLevelText = \[Risk Level\]"
        $users | Should -Match ([regex]::Escape('SWITCH(RiskLevelText, "high", "High", "medium", "Medium", "low", "Low")'))
    }

    It 'keeps the registration userType measure apart from the directory userType' {
        # userRegistrationDetails.userType is member, guest, or unknownFutureValue.
        # The directory user resource uses Member and Guest.
        # https://learn.microsoft.com/graph/api/resources/userregistrationdetails
        $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'AuthenticationMethods.tmdl') -Raw
        $tmdl | Should -Match ([regex]::Escape("measure 'User Type (Display)' ="))
        $tmdl | Should -Match ([regex]::Escape('IF(NOT (Listed > 0), BLANK(), IF(ISBLANK(V), "Unknown", V))'))
        # BRO-501: "Guest" under an account name is the directory userType (users.csv, Guest).
        $users = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'UsersCurrent.tmdl') -Raw
        $users | Should -Match ([regex]::Escape('VAR Place = IF(UserKind = "Guest", "Guest", SELECTEDVALUE(UsersCurrent[Department]))'))
    }

    It 'treats a blank MethodsRegistered as unknown, not zero methods' {
        $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'AuthenticationMethods.tmdl') -Raw
        $tmdl | Should -Match ([regex]::Escape('IF(ISBLANK(AuthenticationMethods[MethodsRegistered]), BLANK()'))
        $tmdl | Should -Match ([regex]::Escape('"Unknown"'))
    }
}

Describe 'Relationships and the anonymize toggle' {
    BeforeAll {
        $text = Get-Content -LiteralPath (Join-Path $script:DefinitionFolder 'relationships.tmdl') -Raw
        $script:Relationships = @([regex]::Matches($text, '(?ms)^relationship (?<id>\S+)\s+fromColumn: (?<from>\S+)\s+toColumn: (?<to>\S+)') |
                ForEach-Object { @{ Id = $_.Groups['id'].Value; From = $_.Groups['from'].Value; To = $_.Groups['to'].Value } })
    }

    It 'declares relationships' {
        $script:Relationships.Count | Should -Be 20
    }

    It 'uses unique relationship ids' {
        @($script:Relationships.Id | Sort-Object -Unique).Count | Should -Be $script:Relationships.Count
    }

    It 'names a column that exists at both ends of <From> -> <To>' -ForEach @(
        $text = Get-Content -LiteralPath (Join-Path $PSScriptRoot '../report/IdentityPosture.SemanticModel/definition/relationships.tmdl') -Raw
        [regex]::Matches($text, '(?ms)^relationship \S+\s+fromColumn: (?<from>\S+)\s+toColumn: (?<to>\S+)') |
            ForEach-Object { @{ From = $_.Groups['from'].Value; To = $_.Groups['to'].Value } }
    ) {
        foreach ($end in $From, $To) {
            $tableName, $columnName = $end -split '\.', 2
            $script:Model.ContainsKey($tableName) | Should -BeTrue -Because "$end names table $tableName"
            $script:Model[$tableName].Columns.Name | Should -Contain $columnName -Because "$end should be a column"
        }
    }

    It 'points the many side at the one side, as the TMDL overview does' {
        # fromColumn is the many side and toColumn the one side
        # (https://learn.microsoft.com/analysis-services/tmdl/tmdl-overview#relationship).
        $oneSide = 'DateDim.Date', 'UsersCurrent.Id', 'RoleDefinitions.RoleDefinitionId'
        $wrong = @($script:Relationships | Where-Object { $_.To -notin $oneSide -or $_.From -in $oneSide })
        ($wrong | ForEach-Object { "$($_.From) -> $($_.To)" }) -join '; ' | Should -BeNullOrEmpty
    }

    It 'offers Show names and Anonymize on the toggle table' {
        $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'AnonymizeMode.tmdl') -Raw
        $tmdl | Should -Match ([regex]::Escape('{{"Show names"}, {"Anonymize"}}'))
    }

    It 'switches the user display name on the toggle and uses an Id-derived pseudonym' {
        $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'UsersCurrent.tmdl') -Raw
        $tmdl | Should -Match ([regex]::Escape('SELECTEDVALUE(AnonymizeMode[Mode], "Show names")'))
        $tmdl | Should -Match ([regex]::Escape('IF(Mode = "Anonymize", Pseudo, RealName)'))
        $tmdl | Should -Match ([regex]::Escape('VAR IdText = UsersCurrent[Id]'))
    }

    It 'keeps each user at that user''s own latest snapshot, not only the newest file date' {
        # A global MAX(RunDate) drops anyone who left before the newest snapshot,
        # and the one-side relationship then removes their historical rows.
        # UsersCurrent is an import table (BRO-373), so the per-user maximum is a
        # Power Query group by Id joined back to the rows.
        $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'UsersCurrent.tmdl') -Raw
        $tmdl | Should -Match ([regex]::Escape('Table.Group(Selected, {"Id"}, {{"LatestRunDate", each List.Max(List.RemoveNulls([RunDate]))'))
        $tmdl | Should -Match ([regex]::Escape('each [RunDate] = [LatestRunDate]'))
        $tmdl | Should -Not -Match ([regex]::Escape('List.Max(List.RemoveNulls(Selected[RunDate]))'))
    }
}

Describe 'No page is built on a beta-only source' {
    It 'has no privileged-role flag or non-interactive sign-in column in the model' {
        # README: sources 4d (isPrivileged) and 6b (signInEventTypes) are documented, not collected.
        $names = foreach ($table in $script:Model.Values) { $table.Columns.Name; $table.Measures }
        $beta = @($names | Where-Object { $_ -match 'IsPrivileged|Privileged Flag|SignInEventType|NonInteractiveUser' })
        $beta -join '; ' | Should -BeNullOrEmpty
    }

    It 'describes the legacy sign-in count as a lower bound' {
        # The approved page (BRO-501) titles the chart "Legacy sign-ins by account" with no
        # "at least"; the model description and the README keep the caveat.
        $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'SignIns.tmdl') -Raw
        $tmdl | Should -Match "/// At least: signins.csv holds interactive and successful federated sign-ins only[^\n]*\n(?:\s*///[^\n]*\n)*\s*measure 'Legacy Sign-Ins' ="
    }

    It 'names roles by the collected RoleDisplayName, with no hard-coded lookup, and never reads a privileged flag' {
        # BRO-554: the collectors write RoleDisplayName (BRO-553, from roleDefinitions), so the
        # page shows built-in and custom roles by their real name. The 16-ID lookup from BRO-501
        # and its "Role " + ID fallback are gone.
        # https://learn.microsoft.com/graph/api/resources/unifiedroledefinition?view=graph-rest-1.0
        $roles = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'RoleDefinitions.tmdl') -Raw
        $roles | Should -Not -Match "column 'Role Name'"
        $roles | Should -Not -Match '62e90394-69f5-4237-9190-012177145e10'
        $users = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'UsersCurrent.tmdl') -Raw
        $users | Should -Not -Match '"Role " & LEFT\('
        $users | Should -Match "measure 'Fix Admin Role' =[\s\S]*?ActiveRoleAssignments\[RoleDisplayName\][\s\S]*?RoleAssignments\[RoleDisplayName\][\s\S]*?EligibleRoleAssignments\[RoleDisplayName\]"
        $roles | Should -Not -Match 'isPrivileged'
    }
}

Describe 'The report has the approved Accounts to fix page and its filter' {
    # BRO-501 replaced the seven pages (Overview ... Risky users) with the approved page 1
    # from docs/design/identity-posture. The design has no slicers: the anonymize toggle is a
    # report-level filter, as on the Teams and Groups lifecycle pages (BRO-499).
    BeforeAll {
        $script:PageVisuals = Join-Path $script:PagesFolder 'accounts-to-fix/visuals'
        $script:UsersTmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'UsersCurrent.tmdl') -Raw
    }

    It 'lists exactly the approved page in pages.json' {
        $pages = Get-Content -LiteralPath (Join-Path $script:PagesFolder 'pages.json') -Raw | ConvertFrom-Json -AsHashtable
        @($pages.pageOrder) -join ',' | Should -Be 'accounts-to-fix'
        $page = Get-Content -LiteralPath (Join-Path $script:PagesFolder 'accounts-to-fix/page.json') -Raw | ConvertFrom-Json -AsHashtable
        $page.displayName | Should -Be 'Accounts to fix'
    }

    It 'carries the house-style headline, red ONE FIX box and dashed Premium box, with figures from measures' {
        # BRO-554: the approved image has none of the three; Joshua decided on 2026-10-14 the page follows the house style.
        $read = { param($name) Get-Content -LiteralPath (Join-Path $script:PageVisuals "$name/visual.json") -Raw | ConvertFrom-Json -AsHashtable }
        (& $read 'fix-headline').visual.query.queryState.Data.projections[0].nativeQueryRef | Should -Be 'Accounts To Fix Headline'
        $panel = & $read 'fix-onefix-panel'
        $panel.visual.objects.outline[0].properties.lineColor.solid.color.expr.Literal.Value | Should -Be "'#C62828'"
        (& $read 'fix-onefix-title-1').visual.query.queryState.Data.projections[0].nativeQueryRef | Should -Be 'Accounts To Fix One Fix Title Line 1'
        (& $read 'fix-onefix-detail').visual.query.queryState.Data.projections[0].nativeQueryRef | Should -Be 'Accounts To Fix One Fix Detail'
        (& $read 'fix-premium-box').visual.objects.general[0].properties.imageUrl.expr.ResourcePackageItem.ItemName | Should -Be 'premium-dashed-box.svg'
        (& $read 'fix-premium-body').visual.objects.general[0].properties.paragraphs[0].textRuns[0].value | Should -Match 'Premium edition at backlogkrew\.com$'
        $script:UsersTmdl | Should -Match "measure 'Accounts To Fix Headline' =[\s\S]*?\[Fix Now Accounts\]"
        $script:UsersTmdl | Should -Match "measure 'Accounts To Fix One Fix Detail' =[\s\S]*?\[Legacy Sign-Ins Succeeded\]"
    }

    It 'carries the anonymize toggle in the report filter pane' {
        $report = Get-Content -LiteralPath (Join-Path $script:ReportRoot 'IdentityPosture.Report/definition/report.json') -Raw | ConvertFrom-Json -AsHashtable
        $fields = @($report.filterConfig.filters | ForEach-Object { '{0}.{1}' -f $_.field.Column.Expression.SourceRef.Entity, $_.field.Column.Property })
        $fields | Should -Contain 'AnonymizeMode.Mode'
    }

    It 'names accounts through the toggle in the table, the chart and the caption' {
        $script:UsersTmdl | Should -Match "measure 'Fix Account Image'[\s\S]*?VAR ShownName = \[User Display Name\]"
        $script:UsersTmdl | Should -Match "measure 'Legacy Caption'[\s\S]*?\[User Display Name\]"
        foreach ($measure in 'Legacy Succeeded (Fix Now)', 'Legacy Succeeded (Other)', 'Legacy Blocked', 'Legacy Chart Order') {
            $script:UsersTmdl | Should -Match ("measure '" + [regex]::Escape($measure) + "'[\s\S]*?SELECTEDVALUE\(AnonymizeMode\[Mode\], `"Show names`"\)[\s\S]*?'User Display'\[Mode\] = ModeShown")
        }
        $chart = Get-Content -LiteralPath (Join-Path $script:PageVisuals 'fix-legacy-chart/visual.json') -Raw | ConvertFrom-Json -AsHashtable
        $chart.visual.query.queryState.Category.projections[0].queryRef | Should -Be 'User Display.Name'
    }

    It 'builds the User Display pseudonym the way UsersCurrent[Pseudonym] does' {
        # DAX: SUMX(GENERATESERIES(1, LEN(Id)), UNICODE(MID(Id, i, 1)) * i), MOD 9000, + 1000.
        $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'User Display.tmdl') -Raw
        $tmdl | Should -Match ([regex]::Escape('Number.Mod(List.Sum(List.Transform({1..Text.Length(id)}, (i) => Character.ToNumber(Text.At(id, i - 1)) * i)), 9000) + 1000'))
        $tmdl | Should -Match ([regex]::Escape('"User " & Text.PadStart('))
        $tmdl | Should -Match "partition 'User Display' = m"
        $script:UsersTmdl | Should -Match ([regex]::Escape('MOD(HashSeed, 9000) + 1000'))
    }

    It 'sorts the accounts table on the hidden Fix Order column, worst first' {
        $table = Get-Content -LiteralPath (Join-Path $script:PageVisuals 'fix-table/visual.json') -Raw | ConvertFrom-Json -AsHashtable
        $key = @($table.visual.query.queryState.Values.projections | Where-Object { $_.queryRef -eq 'UsersCurrent.Fix Order' })
        $key.Count | Should -Be 1
        $key[0].hidden | Should -BeTrue
        $table.visual.query.sortDefinition.sort[0].field.Column.Property | Should -Be 'Fix Order'
        $table.visual.query.sortDefinition.sort[0].direction | Should -Be 'Ascending'
    }

    It 'gives every account its own Fix Order, so no two accounts merge into one table row' {
        # The table has Fix Order as its only column; two accounts with one key (same name, same
        # band, same legacy count) would show as one row with blank cells and drop out of the list.
        $measure = [regex]::Match($script:UsersTmdl, "measure 'Fix Sort Key' =[\s\S]*?lineageTag").Value
        $measure | Should -Not -BeNullOrEmpty
        $measure | Should -Match ([regex]::Escape('UsersCurrent[Id] < UserKey'))
        # Bands are scaled so the name position (1 to the number of accounts) cannot reach the legacy count digits.
        $scales = [regex]::Match($measure, 'Band \* (\d+) \+ \(99 - MIN\(LegacyOk, 99\)\) \* (\d+) \+ NamePosition')
        $scales.Success | Should -BeTrue
        [int64]$scales.Groups[2].Value | Should -BeGreaterOrEqual 1000000
        [int64]$scales.Groups[1].Value | Should -BeGreaterOrEqual (100 * [int64]$scales.Groups[2].Value)
    }

    It 'flags fix now only with a successful legacy sign-in, as the footnote says' {
        $script:UsersTmdl | Should -Match ([regex]::Escape('LegacyOk > 0 && [Fix At Risk] = 1 && [Fix No MFA] = 1, 1,'))
        $script:UsersTmdl | Should -Match ([regex]::Escape('LegacyOk > 0 && [Fix Has Active Role] = 1, 2,'))
        $script:UsersTmdl | Should -Match ([regex]::Escape('LegacyOk > 0 && [Fix Disabled] = 1, 3,'))
    }

    It 'avoids the DAX VAR names Desktop 2.158 rejects' {
        # Desktop 2.158 reports "The syntax for '<name>' is incorrect" for these (BRO-501, BRO-502).
        $bad = foreach ($file in Get-ChildItem -LiteralPath $script:TablesFolder -Filter '*.tmdl') {
            Select-String -LiteralPath $file.FullName -Pattern '\bVAR (Value|Status|Scope|Disabled|Hidden|FirstDate|LastDate|Lead|At|Rows|Level)\b' |
                ForEach-Object { '{0}:{1}' -f $file.Name, $_.LineNumber }
        }
        $bad -join '; ' | Should -BeNullOrEmpty
    }

    It 'never filters CALCULATE on a column compared with a measure' {
        # Desktop 2.158 put the nine Conditional Access detail measures in error ("A function
        # 'PLACEHOLDER' has been used in a True/False expression that is used as a table filter
        # expression"); BRO-501 reads the snapshot date into a VAR first.
        $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'ConditionalAccessPolicies.tmdl') -Raw
        $tmdl | Should -Not -Match ([regex]::Escape('ConditionalAccessPolicies[RunDate] = [Conditional Access Snapshot Date]'))
    }
}
