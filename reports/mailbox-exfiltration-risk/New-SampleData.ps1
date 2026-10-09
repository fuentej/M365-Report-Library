#Requires -Version 7.0

<#
    .SYNOPSIS
        Generates the fake mailbox exfiltration risk data in ./samples.

    .DESCRIPTION
        The sample set lets the later Power BI report, and anyone reading this repository,
        work against realistic shapes without a tenant. Nothing here comes from a real
        directory: every address is under example.com, example.net or example.org, which
        RFC 2606 reserves for documentation.

        The generator is deterministic. The same -Seed and -EndDate always produce the
        same files.

        Every file is written through Export-AppendCsv with the column list the collectors
        use, so a sample file cannot drift from its collector's output. The mailbox
        forwarding sample holds internal forwards (ForwardingAddress), SMTP forwards to an
        accepted domain, and external SMTP forwards, so each branch of the external
        decision is present.

    .PARAMETER EndDate
        The "now" the sample set is generated around. Fixed by default to keep the
        committed files stable.

    .EXAMPLE
        ./New-SampleData.ps1
#>
[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path $PSScriptRoot 'samples'),

    [datetime]$EndDate = [datetime]::new(2026, 9, 1, 0, 0, 0, [System.DateTimeKind]::Utc),

    [int]$Seed = 20260901
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot '../../shared/M365ReportLibrary.psm1')

. (Join-Path $PSScriptRoot 'collectors/MailboxExfiltrationHelpers.ps1')

$schema = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'collectors/MailboxExfiltrationSchema.psd1')

$EndDate = $EndDate.ToUniversalTime()
$random = [System.Random]::new($Seed)

# Three snapshots a month apart, the last one on EndDate.
$snapshotDates = 2..0 | ForEach-Object { $EndDate.AddMonths(-$_) }

$memberCount = 40
$givenNames = @('Avery', 'Blair', 'Casey', 'Devon', 'Emery', 'Finley', 'Gray', 'Harper', 'Indigo', 'Jordan', 'Kai', 'Logan', 'Marlowe', 'Nico', 'Oakley', 'Parker', 'Quinn', 'Reese', 'Sage', 'Tatum')
$familyNames = @('Abara', 'Bergstrom', 'Chaudhry', 'Dlamini', 'Ibarra', 'Jovanovic', 'Kowalski', 'Lindqvist', 'Moreau', 'Nakamura', 'Quintero', 'Silva')
$externalDomains = @('fabrikam.example.net', 'northwind.example.org', 'mail.example.net')
$acceptedDomains = @(
    @{ Name = 'example.com'; DomainName = 'example.com'; DomainType = 'Authoritative'; Default = 'True' }
    @{ Name = 'example.onmicrosoft.com'; DomainName = 'example.onmicrosoft.com'; DomainType = 'Authoritative'; Default = 'False' }
    @{ Name = 'mail.example.com'; DomainName = 'mail.example.com'; DomainType = 'InternalRelay'; Default = 'False' }
)

function New-DeterministicGuid {
    $bytes = [byte[]]::new(16)
    $random.NextBytes($bytes)
    return [guid]::new($bytes).ToString()
}

function Format-Stamp {
    param([AllowNull()][object]$Value)
    return ConvertTo-CsvTimestamp $Value
}

if (Test-Path -LiteralPath $OutputPath) {
    Get-ChildItem -LiteralPath $OutputPath -Recurse -File -Filter '*.csv' | Remove-Item -Force
}
New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null

$users = foreach ($i in 1..$memberCount) {
    $given = $givenNames[($i - 1) % $givenNames.Count]
    $family = $familyNames[($i * 5) % $familyNames.Count]
    $upn = ('{0}.{1}{2}@example.com' -f $given, $family, $(if ($i -gt $givenNames.Count) { $i } else { '' })).ToLowerInvariant()
    [pscustomobject]@{
        Id          = New-DeterministicGuid
        DisplayName = "$given $family"
        Upn         = $upn
        Enabled     = (($i % 9) -ne 0)
        Created     = $EndDate.AddDays(-$random.Next(200, 1200))
    }
}
$users = @($users)

function Get-TargetSummary {
    param([string]$Recipient)
    $summary = Get-RecipientTargetSummary -Recipient $Recipient -AcceptedDomain @($acceptedDomains | ForEach-Object { $_.DomainName })
    return $summary
}

foreach ($snapshot in $snapshotDates) {
    $runDate = $snapshot.ToString('yyyy-MM-dd')

    $userRows = foreach ($user in $users) {
        [pscustomobject]@{
            RunDate                  = $runDate
            Id                       = $user.Id
            DisplayName              = $user.DisplayName
            UserPrincipalName        = $user.Upn
            Mail                     = $user.Upn
            UserType                 = 'Member'
            AccountEnabled           = $user.Enabled
            CreatedDateTime          = Format-Stamp $user.Created
            Department               = 'Operations'
            JobTitle                 = 'Specialist'
            City                     = 'Seattle'
            Country                  = 'US'
            ManagerId                = ''
            ManagerUserPrincipalName = ''
        }
    }
    Export-AppendCsv -Path (Join-Path $OutputPath 'users.csv') -Rows @($userRows) -Column (Get-EntraUserCsvColumn) -KeyColumn @('RunDate', 'Id')

    $domainRows = foreach ($domain in $acceptedDomains) {
        [pscustomobject]@{
            RunDate = $runDate; Name = $domain.Name; DomainName = $domain.DomainName
            DomainType = $domain.DomainType; IsDefault = $domain.Default
        }
    }
    Export-AppendCsv -Path (Join-Path $OutputPath 'accepted-domains.csv') -Rows @($domainRows) -Column $schema.AcceptedDomains -KeyColumn @('RunDate', 'DomainName')

    # Every fifth mailbox forwards; the kind rotates internal recipient, SMTP forward to an
    # accepted domain, SMTP forward to an external domain.
    $forwardRows = foreach ($i in 0..($users.Count - 1)) {
        if (($i % 5) -ne 2) { continue }
        $user = $users[$i]
        $row = [ordered]@{
            RunDate = $runDate; ExternalDirectoryObjectId = $user.Id; UserPrincipalName = $user.Upn
            PrimarySmtpAddress = $user.Upn; ForwardingAddress = ''; ForwardingSmtpAddress = ''
            ForwardingSmtpDomain = ''; DeliverToMailboxAndForward = (($i % 2) -eq 0).ToString(); IsExternal = ''
        }
        switch ([int][math]::Floor($i / 5) % 3) {
            0 { $row.ForwardingAddress = $users[($i + 1) % $users.Count].DisplayName; $row.IsExternal = 'False' }
            1 { $row.ForwardingSmtpAddress = "smtp:$($users[($i + 3) % $users.Count].Upn)"; $row.ForwardingSmtpDomain = 'example.com'; $row.IsExternal = 'False' }
            2 {
                $domain = $externalDomains[$i % $externalDomains.Count]
                $row.ForwardingSmtpAddress = "smtp:inbox$i@$domain"; $row.ForwardingSmtpDomain = $domain; $row.IsExternal = 'True'
            }
        }
        [pscustomobject]$row
    }
    Export-AppendCsv -Path (Join-Path $OutputPath 'mailbox-forwarding.csv') -Rows @($forwardRows) -Column $schema.MailboxForwarding -KeyColumn @('RunDate', 'ExternalDirectoryObjectId')

    $sobRows = foreach ($i in 0..($users.Count - 1)) {
        if (($i % 7) -ne 0) { continue }
        [pscustomobject]@{
            RunDate = $runDate; ExternalDirectoryObjectId = $users[$i].Id; UserPrincipalName = $users[$i].Upn
            PrimarySmtpAddress = $users[$i].Upn; Delegate = $users[($i + 2) % $users.Count].DisplayName
        }
    }
    Export-AppendCsv -Path (Join-Path $OutputPath 'send-on-behalf.csv') -Rows @($sobRows) -Column $schema.SendOnBehalf -KeyColumn @('RunDate', 'ExternalDirectoryObjectId', 'Delegate')

    # Inbox rules: forward, redirect and delete, some to an external domain.
    $ruleRows = foreach ($i in 0..($users.Count - 1)) {
        if (($i % 6) -ne 1) { continue }
        $user = $users[$i]
        $kind = [int][math]::Floor($i / 6) % 3
        $external = (($i % 12) -eq 1)
        $target = if ($external) { "inbox$i@$($externalDomains[$i % $externalDomains.Count])" } else { $users[($i + 4) % $users.Count].Upn }
        $forwardTo = if ($kind -eq 0) { "`"Contact $i`" [SMTP:$target]" } else { '' }
        $redirectTo = if ($kind -eq 1) { "`"Contact $i`" [SMTP:$target]" } else { '' }
        $summary = Get-TargetSummary "$forwardTo $redirectTo"
        [pscustomobject]@{
            RunDate = $runDate; MailboxUserPrincipalName = $user.Upn; MailboxExternalDirectoryObjectId = $user.Id
            RuleIdentity = [string](16752869479666417665 - $i * 7919); RuleName = @('Forward invoices', 'Move to Archive', 'Cleanup')[$kind]
            Enabled = 'True'; Priority = [string]($i % 4 + 1)
            ForwardTo = $forwardTo; ForwardAsAttachmentTo = ''; RedirectTo = $redirectTo
            DeleteMessage = ($kind -eq 2).ToString()
            TargetDomains = $summary.Domains; HasExternalTarget = $summary.HasExternal
        }
    }
    Export-AppendCsv -Path (Join-Path $OutputPath 'inbox-rules.csv') -Rows @($ruleRows) -Column $schema.InboxRules -KeyColumn @('RunDate', 'MailboxExternalDirectoryObjectId', 'RuleIdentity')

    $transportRows = foreach ($i in 0..3) {
        $external = ($i -eq 1)
        $target = if ($external) { "archive@$($externalDomains[0])" } else { "compliance@example.com" }
        $redirect = if ($i -eq 0 -or $i -eq 1) { $target } else { '' }
        $blind = if ($i -eq 2) { $target } else { '' }
        $copy = if ($i -eq 3) { $target } else { '' }
        $summary = Get-TargetSummary "$redirect $blind $copy"
        [pscustomobject]@{
            RunDate = $runDate; Name = @('Redirect legal hold', 'Archive outbound', 'Blind copy finance', 'Copy to compliance')[$i]
            Guid = [guid]::new($i + 1, 2, 3, [byte[]](1, 2, 3, 4, 5, 6, 7, 8)).ToString(); State = $(if ($i -eq 3) { 'Disabled' } else { 'Enabled' })
            Priority = [string]$i; RedirectMessageTo = $redirect; BlindCopyTo = $blind; CopyTo = $copy; AddToRecipients = ''
            TargetDomains = $summary.Domains; HasExternalTarget = $summary.HasExternal
        }
    }
    Export-AppendCsv -Path (Join-Path $OutputPath 'transport-rules.csv') -Rows @($transportRows) -Column $schema.TransportRules -KeyColumn @('RunDate', 'Guid')

    $faRows = foreach ($i in 0..($users.Count - 1)) {
        if (($i % 8) -ne 3) { continue }
        [pscustomobject]@{
            RunDate = $runDate; MailboxExternalDirectoryObjectId = $users[$i].Id; MailboxUserPrincipalName = $users[$i].Upn
            User = $users[($i + 5) % $users.Count].Upn; AccessRights = 'FullAccess'
        }
    }
    Export-AppendCsv -Path (Join-Path $OutputPath 'mailbox-full-access.csv') -Rows @($faRows) -Column $schema.MailboxFullAccess -KeyColumn @('RunDate', 'MailboxExternalDirectoryObjectId', 'User')

    $saRows = foreach ($i in 0..($users.Count - 1)) {
        if (($i % 9) -ne 4) { continue }
        [pscustomobject]@{
            RunDate = $runDate; Identity = $users[$i].DisplayName; Trustee = $users[($i + 6) % $users.Count].Upn
            AccessRights = 'SendAs'; AccessControlType = 'Allow'; IsInherited = 'False'
        }
    }
    Export-AppendCsv -Path (Join-Path $OutputPath 'send-as-permissions.csv') -Rows @($saRows) -Column $schema.SendAsPermissions -KeyColumn @('RunDate', 'Identity', 'Trustee', 'AccessControlType')

    $scopes = @('User.Read', 'Mail.Read offline_access', 'Mail.ReadWrite Mail.Send', 'Files.Read.All', 'MailboxSettings.Read')
    $consentRows = foreach ($i in 0..9) {
        $scope = $scopes[$i % $scopes.Count]
        $tokens = @($scope -split '\s+')
        [pscustomobject]@{
            RunDate = $runDate; Id = [guid]::new($i + 1, 4, 5, [byte[]](1, 2, 3, 4, 5, 6, 7, 8)).ToString()
            ClientId = [guid]::new($i + 100, 4, 5, [byte[]](8, 7, 6, 5, 4, 3, 2, 1)).ToString()
            ConsentType = $(if ($i % 3 -eq 0) { 'AllPrincipals' } else { 'Principal' })
            PrincipalId = $(if ($i % 3 -eq 0) { '' } else { $users[$i].Id })
            ResourceId = '2f7d1b6a-4c1e-4d0e-9b1e-1a2b3c4d5e6f'; Scope = $scope
            HasMailScope = [bool](@($tokens | Where-Object { $_ -like 'Mail.*' -or $_ -like 'MailboxSettings.*' }).Count -gt 0) -as [string]
        }
    }
    Export-AppendCsv -Path (Join-Path $OutputPath 'delegated-consents.csv') -Rows @($consentRows) -Column $schema.DelegatedConsents -KeyColumn @('RunDate', 'Id')

    $roles = @(
        @{ Id = 'b633e1c5-b582-4048-a93e-9f11b44c7e96'; Value = 'Mail.Send' }
        @{ Id = '810c84a8-4a9e-49e6-bf7d-12d183f40d01'; Value = 'Mail.Read' }
        @{ Id = '7ab1d382-f21e-4acd-a863-ba3e13f7da61'; Value = 'Directory.Read.All' }
        @{ Id = 'e2a3a72e-5f79-4c64-b1b1-878b674786c9'; Value = 'Mail.ReadWrite' }
    )
    $assignmentRows = foreach ($i in 0..7) {
        $role = $roles[$i % $roles.Count]
        [pscustomobject]@{
            RunDate = $runDate; AssignmentId = [guid]::new($i + 1, 6, 7, [byte[]](1, 1, 2, 3, 5, 8, 13, 21)).ToString()
            AppRoleId = $role.Id; AppRoleValue = $role.Value
            PrincipalId = [guid]::new($i + 200, 6, 7, [byte[]](2, 3, 5, 7, 11, 13, 17, 19)).ToString()
            PrincipalDisplayName = "Contoso Integration $i"; PrincipalType = 'ServicePrincipal'
            ResourceId = '2f7d1b6a-4c1e-4d0e-9b1e-1a2b3c4d5e6f'; ResourceDisplayName = 'Microsoft Graph'
            CreatedDateTime = Format-Stamp $snapshot.AddDays(-30 - $i)
            IsMailRole = ($role.Value -like 'Mail.*').ToString()
        }
    }
    Export-AppendCsv -Path (Join-Path $OutputPath 'app-role-assignments.csv') -Rows @($assignmentRows) -Column $schema.AppRoleAssignments -KeyColumn @('RunDate', 'AssignmentId')

    $auditRows = @([pscustomobject]@{
            RunDate = $runDate; Scope = 'Organization'; Identity = 'Organization'; AuditDisabled = 'False'
            UnifiedAuditLogIngestionEnabled = 'True'; DefaultAuditSet = ''; AuditAdmin = ''; AuditDelegate = ''; AuditOwner = ''
        })
    $auditRows += foreach ($i in 0..9) {
        [pscustomobject]@{
            RunDate = $runDate; Scope = 'Mailbox'; Identity = $users[$i].Upn; AuditDisabled = ''
            UnifiedAuditLogIngestionEnabled = ''; DefaultAuditSet = 'Admin;Delegate;Owner'
            AuditAdmin = 'Update;MoveToDeletedItems;SoftDelete;HardDelete;SendAs;SendOnBehalf;Create;UpdateFolderPermissions'
            AuditDelegate = 'Update;MoveToDeletedItems;SoftDelete;HardDelete;SendAs;SendOnBehalf;Create;UpdateFolderPermissions'
            AuditOwner = 'Update;MoveToDeletedItems;SoftDelete;HardDelete;UpdateFolderPermissions;UpdateInboxRules'
        }
    }
    Export-AppendCsv -Path (Join-Path $OutputPath 'audit-configuration.csv') -Rows @($auditRows) -Column $schema.AuditConfiguration -KeyColumn @('RunDate', 'Scope', 'Identity')
}

# Audit events, inside the last 90 days.
function New-SampleEvent {
    param([string]$Operation, [string]$RecordType, [int]$Index, [string]$Parameters = '')
    $user = $users[$Index % $users.Count]
    [pscustomobject]@{
        CreationTime    = Format-Stamp $EndDate.AddMinutes(-$random.Next(60, 60 * 24 * 90))
        Id              = New-DeterministicGuid
        RecordType      = $RecordType
        Operation       = $Operation
        UserId          = $user.Upn
        Workload        = 'Exchange'
        ObjectId        = $user.Upn
        MailboxOwnerUPN = $user.Upn
        ClientIP        = "203.0.113.$(10 + $Index)"
        ResultStatus    = 'Succeeded'
        Parameters      = $Parameters
    }
}

$changeEvents = foreach ($i in 0..11) {
    switch ($i % 6) {
        0 { New-SampleEvent 'New-InboxRule' 'ExchangeAdmin' $i "Name=Forward invoices;ForwardTo=inbox$i@$($externalDomains[0])" }
        1 { New-SampleEvent 'Set-InboxRule' 'ExchangeAdmin' $i 'Name=Forward invoices;Enabled=True' }
        2 { New-SampleEvent 'UpdateInboxRules' 'ExchangeItem' $i }
        3 { New-SampleEvent 'Set-Mailbox' 'ExchangeAdmin' $i "Identity=$($users[$i].Upn);ForwardingSmtpAddress=smtp:inbox$i@$($externalDomains[1])" }
        4 { New-SampleEvent 'Add-MailboxPermission' 'ExchangeAdmin' $i "Identity=$($users[$i].Upn);User=$($users[$i + 1].Upn);AccessRights=FullAccess" }
        5 { New-SampleEvent 'New-TransportRule' 'ExchangeAdmin' $i 'Name=Redirect legal hold;RedirectMessageTo=compliance@example.com' }
    }
}
$changeEvents = @($changeEvents | Sort-Object CreationTime)
Export-AppendCsv -Path (Join-Path $OutputPath 'mailbox-change-events.csv') -Rows $changeEvents -Column $schema.AuditEvents -KeyColumn 'Id'

$accessEvents = foreach ($i in 0..15) {
    New-SampleEvent @('MailItemsAccessed', 'Send', 'SendAs', 'SendOnBehalf')[$i % 4] 'ExchangeItemAggregated' $i
}
$accessEvents = @($accessEvents | Sort-Object CreationTime)
Export-AppendCsv -Path (Join-Path $OutputPath 'mail-access-events.csv') -Rows $accessEvents -Column $schema.AuditEvents -KeyColumn 'Id'
