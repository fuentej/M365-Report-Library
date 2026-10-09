#Requires -Version 7.0

<#
    The Power BI report (reports/mailbox-exfiltration-risk/report/) is built from the same CSVs
    the collectors write. These tests hold the report to that:

    - every table, column and measure a visual's query references must exist in the
      TMDL semantic model (a mistyped field reference is silently blank in Power BI,
      not an error, so nothing else would catch it);
    - every semantic-model table sourced directly from a CSV must have exactly the
      columns that CSV has, in the same order;
    - every relationship names columns that exist, and the many side points at the one side;
    - an empty value is Unknown, never zero, and a header-only file is blank, never zero;
    - the Anonymize toggle reaches every name and address a visual shows.

    Calculated tables (UsersCurrent, Apps) and the helper tables that are not sourced from
    any CSV (DateDim and the AnonymizeMode toggle) are exempt from the CSV-mapping check.
    $CsvBackedTables below is the complete list of what is checked.
#>

BeforeAll {
    $script:Root = Join-Path $PSScriptRoot '../../..' | Resolve-Path | Select-Object -ExpandProperty Path
    $script:Samples = Join-Path $script:Root 'reports/mailbox-exfiltration-risk/samples'
    $script:ReportRoot = Join-Path $script:Root 'reports/mailbox-exfiltration-risk/report'
    $script:DefinitionFolder = Join-Path $script:ReportRoot 'MailboxExfiltrationRisk.SemanticModel/definition'
    $script:TablesFolder = Join-Path $script:DefinitionFolder 'tables'
    $script:PagesFolder = Join-Path $script:ReportRoot 'MailboxExfiltrationRisk.Report/definition/pages'

    Import-Module (Join-Path $script:Root 'shared/M365ReportLibrary.psm1') -Force
    . (Join-Path $PSScriptRoot '../../guest-access/tests/TmdlModel.ps1')

    $script:Model = Get-TmdlModel -TablesFolder $script:TablesFolder

    # Table name -> the sample CSV its `m` partition imports.
    $script:CsvBackedTables = [ordered]@{
        Users              = 'users.csv'
        AcceptedDomains    = 'accepted-domains.csv'
        MailboxForwarding  = 'mailbox-forwarding.csv'
        SendOnBehalf       = 'send-on-behalf.csv'
        InboxRules         = 'inbox-rules.csv'
        TransportRules     = 'transport-rules.csv'
        MailboxFullAccess  = 'mailbox-full-access.csv'
        SendAsPermissions  = 'send-as-permissions.csv'
        DelegatedConsents  = 'delegated-consents.csv'
        AppRoleAssignments = 'app-role-assignments.csv'
        MailboxChangeEvents = 'mailbox-change-events.csv'
        MailAccessEvents   = 'mail-access-events.csv'
        AuditConfiguration = 'audit-configuration.csv'
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

    function Get-TableText {
        param([Parameter(Mandatory)][string]$Name)
        Get-Content -LiteralPath (Join-Path $script:TablesFolder "$Name.tmdl") -Raw
    }

    $script:VisualFiles = @(Get-ChildItem -LiteralPath $script:PagesFolder -Filter 'visual.json' -Recurse -File)
}

AfterAll {
    Remove-Module M365ReportLibrary -Force -ErrorAction SilentlyContinue
}

Describe 'The semantic model has tables to check' {
    It 'parses the thirteen CSV-backed tables plus the calculated and helper tables' {
        $script:Model.Keys.Count | Should -Be 17
        foreach ($name in $script:CsvBackedTables.Keys + @('UsersCurrent', 'Apps', 'DateDim', 'AnonymizeMode')) {
            $script:Model.ContainsKey($name) | Should -BeTrue -Because "the model should have a $name table"
        }
    }

    It 'finds visual.json files to check' {
        $script:VisualFiles.Count | Should -BeGreaterThan 80
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

    It 'projects no raw name or address column' {
        # Display names, user principal names and mailbox addresses, and the recipient
        # lists that hold addresses, reach a visual only through the pseudonym columns and
        # the *(Display) measures, which honour the Anonymize toggle. Destinations are shown
        # as domains (DestinationDomain, TargetDomainsLabel), never as addresses.
        $rawColumns = @(
            'DisplayName', 'Mail', 'UserPrincipalName', 'PrimarySmtpAddress', 'MailboxUserPrincipalName',
            'ForwardingAddress', 'ForwardingSmtpAddress', 'ForwardTo', 'ForwardAsAttachmentTo', 'RedirectTo',
            'RedirectMessageTo', 'BlindCopyTo', 'CopyTo', 'AddToRecipients', 'User', 'Trustee', 'Delegate',
            'Identity', 'DelegateKey', 'TrusteeKey', 'MailboxKey', 'PrincipalDisplayName', 'PrincipalId',
            'ClientId', 'AppId', 'AppName', 'UserId', 'MailboxOwnerUPN', 'ObjectId', 'ClientIP', 'Parameters',
            'ManagerUserPrincipalName', 'ExternalDirectoryObjectId', 'MailboxExternalDirectoryObjectId'
        )
        $raw = @($script:AllReferences | Where-Object { $_.Kind -eq 'Column' -and $_.Property -in $rawColumns } |
                ForEach-Object { '{0}: {1}.{2}' -f $_.File, $_.Entity, $_.Property } | Sort-Object -Unique)
        $raw -join '; ' | Should -BeNullOrEmpty
    }
}

Describe 'Every CSV-backed table matches its sample CSV' {
    It 'checks <Table>' -ForEach @(
        @{ Table = 'Users' }, @{ Table = 'AcceptedDomains' }, @{ Table = 'MailboxForwarding' }, @{ Table = 'SendOnBehalf' },
        @{ Table = 'InboxRules' }, @{ Table = 'TransportRules' }, @{ Table = 'MailboxFullAccess' }, @{ Table = 'SendAsPermissions' },
        @{ Table = 'DelegatedConsents' }, @{ Table = 'AppRoleAssignments' }, @{ Table = 'MailboxChangeEvents' },
        @{ Table = 'MailAccessEvents' }, @{ Table = 'AuditConfiguration' }
    ) {
        $script:Model.ContainsKey($Table) | Should -BeTrue -Because "the model should have a $Table table"

        $csvPath = Join-Path $script:Samples $script:CsvBackedTables[$Table]
        $csvHeader = Get-CsvHeaderColumn -Path $csvPath
        $csvHeader | Should -Not -BeNullOrEmpty -Because "$csvPath should be readable"

        # Only the imported (sourceColumn-backed) columns, in file order: a calculated
        # column layered on a CSV-backed table (DestinationType, ActivityDate) is not part
        # of the CSV and is excluded, as it would be from an Import-Csv header.
        $modelColumns = @($script:Model[$Table].Columns | Where-Object { -not $_.IsCalculated } |
                ForEach-Object { $_.SourceColumn })

        ($modelColumns -join ',') | Should -Be ($csvHeader -join ',')
    }

    It 'names every CSV-backed table exactly once' {
        $script:CsvBackedTables.Keys.Count | Should -Be 13
        (@($script:CsvBackedTables.Values) | Sort-Object -Unique).Count | Should -Be 13
    }

    It 'maps every sample CSV the collectors write to a table' {
        $samples = @(Get-ChildItem -LiteralPath $script:Samples -Filter '*.csv' -File | ForEach-Object Name | Sort-Object)
        $samples -join ',' | Should -Be ((@($script:CsvBackedTables.Values) | Sort-Object) -join ',')
    }

    It 'reads <Table> from the CsvFolder parameter and the CSV named in the map' -ForEach @(
        @{ Table = 'Users' }, @{ Table = 'AcceptedDomains' }, @{ Table = 'MailboxForwarding' }, @{ Table = 'SendOnBehalf' },
        @{ Table = 'InboxRules' }, @{ Table = 'TransportRules' }, @{ Table = 'MailboxFullAccess' }, @{ Table = 'SendAsPermissions' },
        @{ Table = 'DelegatedConsents' }, @{ Table = 'AppRoleAssignments' }, @{ Table = 'MailboxChangeEvents' },
        @{ Table = 'MailAccessEvents' }, @{ Table = 'AuditConfiguration' }
    ) {
        Get-TableText -Name $Table | Should -Match ([regex]::Escape('File.Contents(CsvFolder & "\' + $script:CsvBackedTables[$Table] + '")'))
    }

    It 'keeps the CsvFolder parameter as the only parameter query' {
        $expressions = Get-Content -LiteralPath (Join-Path $script:DefinitionFolder 'expressions.tmdl') -Raw
        ([regex]::Matches($expressions, '(?m)^expression ')).Count | Should -Be 1
        $expressions | Should -Match '(?m)^expression CsvFolder = '
    }

    It 'types the audit timestamps as UTC datetimes and the flags as logical' {
        foreach ($name in 'MailboxChangeEvents', 'MailAccessEvents') {
            Get-TableText -Name $name | Should -Match ([regex]::Escape('{"CreationTime", ParseUtc, type datetime}'))
        }
        Get-TableText -Name 'MailboxForwarding' | Should -Match ([regex]::Escape('{"IsExternal", type logical}'))
    }
}

Describe 'Internal and external forwarding are separated' {
    It 'derives the destination type from IsExternal and keeps an empty IsExternal Unknown' {
        $text = Get-TableText -Name 'MailboxForwarding'
        $text | Should -Match ([regex]::Escape('IF(ISBLANK(MailboxForwarding[IsExternal]), "Unknown", IF(MailboxForwarding[IsExternal], "External", "Internal"))'))
    }

    It 'counts external, internal and unknown forwarding mailboxes separately' {
        $measures = $script:Model['MailboxForwarding'].Measures
        $measures | Should -Contain 'External Forwarding Mailboxes'
        $measures | Should -Contain 'Internal Forwarding Mailboxes'
        $measures | Should -Contain 'Unknown Forwarding Destination'
        (Get-TableText -Name 'MailboxForwarding') | Should -Match ([regex]::Escape('MailboxForwarding[IsExternal] = TRUE()'))
    }

    It 'separates inbox rules and mail flow rules by target scope' {
        foreach ($name in 'InboxRules', 'TransportRules') {
            $text = Get-TableText -Name $name
            $text | Should -Match ([regex]::Escape('IF(ISBLANK(' + $name + '[HasExternalTarget]), "Unknown", IF(' + $name + '[HasExternalTarget], "External", "Internal"))'))
        }
    }

    It 'shows forwarding on the forwarding page by destination type' {
        $path = Join-Path $script:PagesFolder 'mailbox-forwarding/visuals/column-by-destination/visual.json'
        (Get-Content -LiteralPath $path -Raw) | Should -Match 'MailboxForwarding.DestinationType'
    }
}

Describe 'Empty values are Unknown, never zero' {
    It 'returns blank, not zero, for a snapshot measure over a header-only file' {
        foreach ($name in 'AcceptedDomains', 'MailboxForwarding', 'SendOnBehalf', 'InboxRules', 'TransportRules', 'MailboxFullAccess',
            'SendAsPermissions', 'DelegatedConsents', 'AppRoleAssignments', 'AuditConfiguration') {
            Get-TableText -Name $name | Should -Match 'IF\(ISBLANK\(Snap\), BLANK\(\)' -Because "$name measures must return blank when the file is header-only"
        }
    }

    It 'returns blank, not zero, for an event measure over a header-only file' {
        foreach ($name in 'MailboxChangeEvents', 'MailAccessEvents') {
            Get-TableText -Name $name | Should -Match ([regex]::Escape("IF(COUNTROWS(ALL($name)) = 0, BLANK()"))
        }
    }

    It 'leaves the delegates-with-all-three measure blank when any of its three files is header-only' {
        $text = Get-TableText -Name 'MailboxFullAccess'
        $text | Should -Match ([regex]::Escape('IF(ISBLANK(FaSnap) || ISBLANK(SaSnap) || ISBLANK(SobSnap), BLANK()'))
    }

    It 'labels an empty destination, target, copy setting, consent type, mail scope or result Unknown' {
        (Get-TableText -Name 'MailboxForwarding') | Should -Match ([regex]::Escape('IF(ISBLANK(MailboxForwarding[DeliverToMailboxAndForward]), "Unknown"'))
        (Get-TableText -Name 'InboxRules') | Should -Match ([regex]::Escape('IF(InboxRules[ForwardsMail], "Unknown", "No recipient")'))
        (Get-TableText -Name 'TransportRules') | Should -Match ([regex]::Escape('"Unknown")') )
        (Get-TableText -Name 'DelegatedConsents') | Should -Match ([regex]::Escape('IF(ISBLANK(DelegatedConsents[ConsentType]), "Unknown"'))
        (Get-TableText -Name 'DelegatedConsents') | Should -Match ([regex]::Escape('IF(ISBLANK(DelegatedConsents[HasMailScope]), "Unknown"'))
        (Get-TableText -Name 'MailboxChangeEvents') | Should -Match ([regex]::Escape('IF(ISBLANK(MailboxChangeEvents[ResultStatus]), "Unknown"'))
    }

    It 'counts the rows whose flag is empty in their own Unknown measure' {
        $script:Model['InboxRules'].Measures | Should -Contain 'Inbox Rules With Unknown Target'
        $script:Model['TransportRules'].Measures | Should -Contain 'Mail Flow Rules With Unknown Target'
        $script:Model['DelegatedConsents'].Measures | Should -Contain 'Consents With Unknown Mail Scope'
        $script:Model['SendAsPermissions'].Measures | Should -Contain 'Send As With Unknown Inheritance'
        $script:Model['AuditConfiguration'].Measures | Should -Contain 'Mailboxes With Unknown Audit Set'
        $script:Model['MailboxChangeEvents'].Measures | Should -Contain 'Change Events With Unknown Result'
    }

    It 'reports the organization audit setting as Unknown when the organization row is empty' {
        $text = Get-TableText -Name 'AuditConfiguration'
        $text | Should -Match ([regex]::Escape('IF(OnRows + OffRows = 0, "Unknown"'))
    }

    It 'leaves the unknown app name Unknown rather than showing an id' {
        (Get-TableText -Name 'Apps') | Should -Match ([regex]::Escape('"Unknown app name"'))
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
}

Describe 'Relationships' {
    BeforeAll {
        $text = Get-Content -LiteralPath (Join-Path $script:DefinitionFolder 'relationships.tmdl') -Raw
        $script:Relationships = @([regex]::Matches($text, '(?ms)^relationship (?<id>\S+)\s+fromColumn: (?<from>\S+)\s+toColumn: (?<to>\S+)') |
                ForEach-Object { @{ Id = $_.Groups['id'].Value; From = $_.Groups['from'].Value; To = $_.Groups['to'].Value } })
    }

    It 'declares relationships' {
        $script:Relationships.Count | Should -Be 24
    }

    It 'uses unique relationship ids' {
        @($script:Relationships.Id | Sort-Object -Unique).Count | Should -Be $script:Relationships.Count
    }

    It 'names a column that exists at both ends of <From> -> <To>' -ForEach @(
        $text = Get-Content -LiteralPath (Join-Path $PSScriptRoot '../report/MailboxExfiltrationRisk.SemanticModel/definition/relationships.tmdl') -Raw
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
        $oneSide = 'DateDim.Date', 'UsersCurrent.Id', 'UsersCurrent.UserPrincipalName', 'Apps.AppId'
        $wrong = @($script:Relationships | Where-Object { $_.To -notin $oneSide -or $_.From -in $oneSide })
        ($wrong | ForEach-Object { "$($_.From) -> $($_.To)" }) -join '; ' | Should -BeNullOrEmpty
    }

    It 'relates every snapshot and event table to the date dimension' {
        $dated = @($script:Relationships | Where-Object { $_.To -eq 'DateDim.Date' } | ForEach-Object { $_.From })
        foreach ($name in $script:CsvBackedTables.Keys) {
            $expected = if ($name -in 'MailboxChangeEvents', 'MailAccessEvents') { "$name.ActivityDate" } else { "$name.RunDate" }
            $dated | Should -Contain $expected
        }
    }

    It 'does not relate a calculated column to the table it is looked up from' {
        # A calculated column that looks a value up in table B and is also the many end of a
        # relationship to B is a circular dependency. The display-name lookups read Users,
        # which has no relationship to the tables that look it up; UsersCurrent is the related table.
        foreach ($relationship in $script:Relationships) {
            $relationship.To | Should -Not -Match '^Users\.'
        }
        (Get-TableText -Name 'SendAsPermissions') | Should -Match ([regex]::Escape('LOOKUPVALUE(Users[UserPrincipalName]'))
        (Get-TableText -Name 'SendAsPermissions') | Should -Not -Match ([regex]::Escape('LOOKUPVALUE(UsersCurrent['))
    }
}

Describe 'The anonymize toggle' {
    It 'offers Show names and Anonymize on the toggle table' {
        Get-TableText -Name 'AnonymizeMode' | Should -Match ([regex]::Escape('{{"Show names"}, {"Anonymize"}}'))
    }

    It 'switches the mailbox and app display names on the toggle and derives the pseudonym from the identity' {
        $users = Get-TableText -Name 'UsersCurrent'
        $users | Should -Match ([regex]::Escape('SELECTEDVALUE(AnonymizeMode[Mode], "Show names")'))
        $users | Should -Match ([regex]::Escape('IF(Mode = "Anonymize", Pseudo, RealName)'))
        $users | Should -Match ([regex]::Escape('LOWER(UsersCurrent[UserPrincipalName])'))
        $apps = Get-TableText -Name 'Apps'
        $apps | Should -Match ([regex]::Escape('IF(Mode = "Anonymize", Pseudo,'))
        $apps | Should -Match ([regex]::Escape('LOWER(Apps[AppId])'))
    }

    It 'switches <Measure> on the toggle' -ForEach @(
        @{ Table = 'MailboxFullAccess'; Measure = 'Full Access Delegate (Display)' }
        @{ Table = 'SendOnBehalf'; Measure = 'Send On Behalf Delegate (Display)' }
        @{ Table = 'SendAsPermissions'; Measure = 'Send As Trustee (Display)' }
        @{ Table = 'SendAsPermissions'; Measure = 'Send As Identity (Display)' }
        @{ Table = 'MailboxChangeEvents'; Measure = 'Actor (Display)' }
        @{ Table = 'MailAccessEvents'; Measure = 'Accessor (Display)' }
    ) {
        $script:Model[$Table].Measures | Should -Contain $Measure
        Get-TableText -Name $Table | Should -Match ([regex]::Escape('IF(Mode = "Anonymize", Pseudo, RealName)'))
    }

    It 'gives the same person the same pseudonym wherever they appear' {
        # One formula, one prefix, one lower-cased key: a user as a mailbox, a Full Access
        # delegate, a Send As trustee and an actor all hash their lower-cased principal name.
        $formula = 'MOD(SUMX(GENERATESERIES(1, LEN(K)), UNICODE(MID(K, [Value], 1)) * [Value]), 9000) + 1000'
        foreach ($pair in @(@('UsersCurrent', 'UserPrincipalName'), @('MailboxFullAccess', 'DelegateKey'), @('SendAsPermissions', 'TrusteeKey'),
                @('MailboxChangeEvents', 'UserId'), @('MailAccessEvents', 'UserId'), @('SendOnBehalf', 'DelegateKey'))) {
            $text = Get-TableText -Name $pair[0]
            $text | Should -Match ([regex]::Escape($formula))
            $text | Should -Match ([regex]::Escape("LOWER($($pair[0])[$($pair[1])])"))
            $text | Should -Match ([regex]::Escape('"Person "'))
        }
    }

    It 'shows a pseudonym column and a display measure for every person, mailbox or app a table lists' {
        foreach ($path in Get-ChildItem -LiteralPath $script:PagesFolder -Filter 'visual.json' -Recurse -File) {
            $json = Get-Content -LiteralPath $path.FullName -Raw | ConvertFrom-Json -AsHashtable
            if ($json.visual.visualType -ne 'tableEx') { continue }
            $fields = @($json.visual.query.queryState.Values.projections | ForEach-Object { $_.queryRef })
            $pseudonyms = @($fields | Where-Object { $_ -like '*Pseudonym' })
            $displays = @($fields | Where-Object { $_ -like '*(Display)*' -or $_ -like '*Display Name' })
            if ($pseudonyms.Count -gt 0) {
                $displays.Count | Should -BeGreaterThan 0 -Because "$($path.FullName) lists pseudonyms but offers no display measure for the toggle"
            }
        }
    }
}

Describe 'The pages say what the data cannot show' {
    It 'says hidden inbox rules are read but not flagged' {
        $path = Join-Path $script:PagesFolder 'inbox-rules/visuals/bar-by-domain/visual.json'
        (Get-Content -LiteralPath $path -Raw) | Should -Match 'hidden rules are read but not flagged'
    }

    It 'says who approved a tenant-wide consent is not collected' {
        $path = Join-Path $script:PagesFolder 'application-access/visuals/column-by-consent/visual.json'
        (Get-Content -LiteralPath $path -Raw) | Should -Match 'who approved a tenant-wide grant is not collected'
    }

    It 'says a group does not join to the mailbox slicer on the Send As table' {
        $path = Join-Path $script:PagesFolder 'delegation/visuals/table-send-as/visual.json'
        (Get-Content -LiteralPath $path -Raw) | Should -Match 'a group does not join to the mailbox slicer'
    }

    It 'plots change events on the days they occurred' {
        # A DateDim axis evaluates the measure for every calendar day from 2020 through 2035.
        $path = Join-Path $script:PagesFolder 'change-history/visuals/line-changes/visual.json'
        $text = Get-Content -LiteralPath $path -Raw
        $text | Should -Match 'MailboxChangeEvents.ActivityDate'
        $text | Should -Not -Match 'DateDim.Date'
    }

    It 'shows destinations as domains, never as addresses' {
        $text = (Get-ChildItem -LiteralPath $script:PagesFolder -Filter 'visual.json' -Recurse -File | ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw }) -join "`n"
        $text | Should -Match 'MailboxForwarding.DestinationDomain'
        $text | Should -Match 'InboxRules.TargetDomainsLabel'
        $text | Should -Match 'TransportRules.TargetDomainsLabel'
    }

    It 'counts only explicit Allow Send As rows, not inherited defaults' {
        # Get-RecipientPermission returns IsInherited. BLANK() = FALSE() in DAX, so an empty
        # flag has to be excluded or it is counted as an explicit grant.
        # https://learn.microsoft.com/powershell/module/exchangepowershell/get-recipientpermission
        $sendAs = Get-TableText -Name 'SendAsPermissions'
        $trustee = @($sendAs -split "`n" | Where-Object { $_ -match 'DISTINCTCOUNT\(SendAsPermissions\[TrusteeKey\]\)' })
        $trustee.Count | Should -Be 1
        $trustee[0] | Should -Match 'IsInherited\] = FALSE\(\)'
        $trustee[0] | Should -Match 'NOT ISBLANK\(SendAsPermissions\[IsInherited\]\)'
        $trend = @($sendAs -split "`n" | Where-Object { $_ -match "Send As Grants \(All Snapshots\)" })
        $trend.Count | Should -Be 1
        $trend[0] | Should -Match 'IsInherited\] = FALSE\(\)'
        $trend[0] | Should -Match 'NOT ISBLANK\(SendAsPermissions\[IsInherited\]\)'
        $allThree = Get-TableText -Name 'MailboxFullAccess'
        $keys = @($allThree -split "`n" | Where-Object { $_ -match 'SendAsPermissions\[TrusteeKey\]' })
        $keys.Count | Should -Be 1
        $keys[0] | Should -Match 'IsInherited\] = FALSE\(\)'
        $keys[0] | Should -Match 'NOT ISBLANK\(SendAsPermissions\[IsInherited\]\)'
        $table = Get-Content -LiteralPath (Join-Path $script:PagesFolder 'delegation/visuals/table-send-as/visual.json') -Raw
        $table | Should -Match 'explicit-send-as-grants'
        $table | Should -Match '"ComparisonKind": 1'
    }

    It 'keeps organization-wide mail flow rules off the mailbox trend' {
        # Get-TransportRule is organization-wide. The other series follow the mailbox slicer.
        # https://learn.microsoft.com/exchange/security-and-compliance/mail-flow-rules/mail-flow-rules
        $text = Get-Content -LiteralPath (Join-Path $script:PagesFolder 'overview/visuals/line-trend/visual.json') -Raw
        $text | Should -Not -Match 'TransportRules'
        $text | Should -Match 'MailboxForwarding.Forwarding Mailboxes \(All Snapshots\)'
        $text | Should -Match 'InboxRules.Forward Or Redirect Rules \(All Snapshots\)'
    }

    It 'plots mail access events on the days they occurred' {
        # MailItemsAccessed, Send, SendAs and SendOnBehalf are dated audit records.
        # https://learn.microsoft.com/purview/audit-log-activities#exchange-mailbox-activities
        $path = Join-Path $script:PagesFolder 'change-history/visuals/line-access/visual.json'
        $path | Should -Exist
        $text = Get-Content -LiteralPath $path -Raw
        $text | Should -Match 'MailAccessEvents.ActivityDate'
        $text | Should -Match 'MailAccessEvents.Operation'
        $text | Should -Match 'Mail Access Events'
        $text | Should -Not -Match 'DateDim.Date'
    }

    It 'reads the earliest audit record across both event files' {
        $text = Get-TableText -Name 'MailboxChangeEvents'
        $text | Should -Match ([regex]::Escape('MIN(MailAccessEvents[CreationTime])'))
        $text | Should -Match ([regex]::Escape('MIN(MailboxChangeEvents[CreationTime])'))
    }
}
