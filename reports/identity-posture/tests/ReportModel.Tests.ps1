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
      the two beta-only sources the README lists as not collected.

    Calculated tables (UsersCurrent, RoleDefinitions) and the helper tables that are
    not sourced from any CSV (DateDim and the AnonymizeMode toggle) are exempt from the
    CSV-mapping check. $CsvBackedTables below is the complete list of what is checked.
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
        $script:Model.Keys.Count | Should -Be 13
        foreach ($name in $script:CsvBackedTables.Keys + @('UsersCurrent', 'RoleDefinitions', 'DateDim', 'AnonymizeMode')) {
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
        $script:AllReferences.Count | Should -BeGreaterThan 100
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
        $visual = Get-Content -LiteralPath (Join-Path $script:PagesFolder 'risky-users/visuals/table-users/visual.json') -Raw
        $visual | Should -Match ([regex]::Escape('"Property": "Listed On Snapshot"'))
        $visual | Should -Match '"ComparisonKind": 0'
    }

    It 'shows the registration userType on the authentication methods table' {
        # userRegistrationDetails.userType is member, guest, or unknownFutureValue.
        # The directory user resource uses Member and Guest.
        # https://learn.microsoft.com/graph/api/resources/userregistrationdetails
        $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'AuthenticationMethods.tmdl') -Raw
        $tmdl | Should -Match ([regex]::Escape("measure 'User Type (Display)' ="))
        $tmdl | Should -Match ([regex]::Escape('IF(NOT (Listed > 0), BLANK(), IF(ISBLANK(V), "Unknown", V))'))
        $visual = Get-Content -LiteralPath (Join-Path $script:PagesFolder 'authentication-methods/visuals/table-users/visual.json') -Raw
        $visual | Should -Match ([regex]::Escape('"Property": "User Type (Display)"'))
        $visual | Should -Not -Match 'UsersCurrent.UserType'
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
        $tmdl = Get-Content -LiteralPath (Join-Path $script:TablesFolder 'UsersCurrent.tmdl') -Raw
        $tmdl | Should -Match ([regex]::Escape('ALLEXCEPT(Users, Users[Id])'))
        $tmdl | Should -Not -Match ([regex]::Escape('Users[RunDate] = MAX(Users[RunDate])'))
    }
}

Describe 'No page is built on a beta-only source' {
    It 'has no privileged-role flag or non-interactive sign-in column in the model' {
        # README: sources 4d (isPrivileged) and 6b (signInEventTypes) are documented, not collected.
        $names = foreach ($table in $script:Model.Values) { $table.Columns.Name; $table.Measures }
        $beta = @($names | Where-Object { $_ -match 'IsPrivileged|Privileged Flag|SignInEventType|NonInteractiveUser' })
        $beta -join '; ' | Should -BeNullOrEmpty
    }

    It 'says in the page titles that the legacy sign-in counts are a lower bound' {
        $cards = Get-ChildItem -LiteralPath (Join-Path $script:PagesFolder 'legacy-authentication/visuals') -Filter visual.json -Recurse |
            ForEach-Object { (Get-Content -LiteralPath $_.FullName -Raw | ConvertFrom-Json -AsHashtable).visual.visualContainerObjects.title[0].properties.text }
        @($cards | Where-Object { $_ -match 'at least' }).Count | Should -BeGreaterOrEqual 4
    }

    It 'says the role page groups by role definition ID because names are not collected' {
        $path = Join-Path $script:PagesFolder 'privileged-roles/visuals/bar-by-role/visual.json'
        (Get-Content -LiteralPath $path -Raw) | Should -Match 'role names and the privileged flag are not collected'
    }
}
