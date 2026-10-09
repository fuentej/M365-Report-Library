#Requires -Version 7.0

BeforeAll {
    $script:Root = Join-Path $PSScriptRoot '../../..' | Resolve-Path | Select-Object -ExpandProperty Path
    $script:Collectors = Join-Path $script:Root 'reports/mailbox-exfiltration-risk/collectors'
    $script:Samples = Join-Path $script:Root 'reports/mailbox-exfiltration-risk/samples'

    . (Join-Path $script:Root 'shared/tests/TenantCmdletStubs.ps1')
    . (Join-Path $PSScriptRoot 'MailboxStubs.ps1')
    Import-Module (Join-Path $script:Root 'shared/M365ReportLibrary.psm1') -Force

    $script:Schema = Import-PowerShellDataFile -LiteralPath (Join-Path $script:Collectors 'MailboxExfiltrationSchema.psd1')

    function New-TestFolder {
        $path = Join-Path ([System.IO.Path]::GetTempPath()) ('mailbox-' + [guid]::NewGuid().ToString('N'))
        New-Item -Path $path -ItemType Directory -Force | Out-Null
        return $path
    }

    function Get-HeaderText {
        param([Parameter(Mandatory)][string]$Path)
        return ((Get-CsvHeaderColumn -Path $Path) -join ',')
    }

    function Get-LogText {
        param([Parameter(Mandatory)][string]$Folder)
        return (Get-Content -LiteralPath (Join-Path $Folder 'run.log') -Raw)
    }

    function Invoke-CollectorScript {
        param([Parameter(Mandatory)][string]$Name, [hashtable]$Arguments = @{})
        & (Join-Path $script:Collectors $Name) @Arguments
    }

    # Get-EXOMailbox: the property sets are at
    # https://learn.microsoft.com/powershell/exchange/cmdlet-property-sets. Minimum holds
    # ExternalDirectoryObjectId, PrimarySmtpAddress and UserPrincipalName; Delivery holds
    # DeliverToMailboxAndForward, ForwardingAddress, ForwardingSmtpAddress and
    # GrantSendOnBehalfTo; Audit holds AuditAdmin, AuditDelegate, AuditOwner and
    # DefaultAuditSet. The cmdlet pages do not document the output object, so the property
    # types follow the parameter descriptions: ForwardingSmtpAddress is rendered
    # smtp:user@domain, GrantSendOnBehalfTo is a list of recipients.
    function New-MockMailbox {
        param(
            [string]$Id = 'mbx-1',
            [string]$Upn = 'avery.abara@example.com',
            [string]$ForwardingAddress = $null,
            [string]$ForwardingSmtpAddress = $null,
            [bool]$DeliverToMailboxAndForward = $false,
            [string[]]$GrantSendOnBehalfTo = @()
        )

        [pscustomobject]@{
            ExternalDirectoryObjectId  = $Id
            UserPrincipalName          = $Upn
            PrimarySmtpAddress         = $Upn
            RecipientTypeDetails       = 'UserMailbox'
            ForwardingAddress          = $ForwardingAddress
            ForwardingSmtpAddress      = $ForwardingSmtpAddress
            DeliverToMailboxAndForward = $DeliverToMailboxAndForward
            GrantSendOnBehalfTo        = $GrantSendOnBehalfTo
            DefaultAuditSet            = @('Admin', 'Delegate', 'Owner')
            AuditAdmin                 = @('Update', 'SendAs')
            AuditDelegate              = @('Update', 'SendAs')
            AuditOwner                 = @('Update', 'UpdateInboxRules')
        }
    }

    # Get-AcceptedDomain: Name, DomainName, DomainType, Default are the properties the
    # cmdlet and its Format-Table examples use. No output object is documented.
    # https://learn.microsoft.com/powershell/module/exchangepowershell/get-accepteddomain
    function New-MockAcceptedDomain {
        param([string]$Name = 'example.com', [string]$DomainType = 'Authoritative', [bool]$Default = $true)
        [pscustomobject]@{ Name = $Name; DomainName = $Name; DomainType = $DomainType; Default = $Default }
    }

    # Get-InboxRule: ForwardTo, ForwardAsAttachmentTo, RedirectTo and DeleteMessage are the
    # properties named on New-InboxRule; RuleIdentity (for example 16752869479666417665)
    # and Name come from the Identity parameter description of Get-InboxRule.
    # https://learn.microsoft.com/powershell/module/exchangepowershell/get-inboxrule
    function New-MockInboxRule {
        param(
            [string]$Identity = '16752869479666417665',
            [string]$Name = 'Forward invoices',
            [string[]]$ForwardTo = @(),
            [string[]]$RedirectTo = @(),
            [bool]$DeleteMessage = $false
        )

        [pscustomobject]@{
            RuleIdentity          = $Identity
            Name                  = $Name
            Enabled               = $true
            Priority              = 1
            ForwardTo             = $ForwardTo
            ForwardAsAttachmentTo = @()
            RedirectTo            = $RedirectTo
            DeleteMessage         = $DeleteMessage
        }
    }

    # Get-TransportRule: RedirectMessageTo, BlindCopyTo, CopyTo and AddToRecipients are the
    # action names on the mail flow rule actions page; Name, Guid, State and Priority are
    # rule properties. No output object is documented.
    # https://learn.microsoft.com/exchange/security-and-compliance/mail-flow-rules/mail-flow-rule-actions
    function New-MockTransportRule {
        param(
            [string]$Name = 'Redirect legal hold',
            [string]$Guid = '11111111-2222-3333-4444-555555555555',
            [string[]]$RedirectMessageTo = @(),
            [string[]]$BlindCopyTo = @(),
            [string[]]$CopyTo = @()
        )

        [pscustomobject]@{
            Name              = $Name
            Guid              = $Guid
            State             = 'Enabled'
            Priority          = 0
            RedirectMessageTo = $RedirectMessageTo
            BlindCopyTo       = $BlindCopyTo
            CopyTo            = $CopyTo
            AddToRecipients   = @()
        }
    }

    # Get-MailboxPermission: the verification command on the manage-permissions page
    # prints User, Deny, IsInherited and AccessRights.
    # https://learn.microsoft.com/exchange/recipients-in-exchange-online/manage-permissions-for-recipients
    function New-MockMailboxPermission {
        param(
            [string]$User = 'raymond@example.com',
            [string[]]$AccessRights = @('FullAccess'),
            [bool]$Deny = $false,
            [bool]$IsInherited = $false
        )
        [pscustomobject]@{ Identity = 'avery.abara'; User = $User; AccessRights = $AccessRights; Deny = $Deny; IsInherited = $IsInherited }
    }

    # Get-RecipientPermission: Identity, Trustee and AccessRights are the parameters of
    # Add-RecipientPermission on the same page; AccessControlType and IsInherited are the
    # remaining properties of the permission object. No output object is documented.
    function New-MockRecipientPermission {
        param([string]$Identity = 'Contoso Printer Support', [string]$Trustee = 'printers@example.com')
        [pscustomobject]@{ Identity = $Identity; Trustee = $Trustee; AccessRights = @('SendAs'); AccessControlType = 'Allow'; IsInherited = $false }
    }

    # oAuth2PermissionGrant: id, clientId, consentType, principalId, resourceId, scope.
    # https://learn.microsoft.com/graph/api/resources/oauth2permissiongrant
    function New-MockGrant {
        param([string]$Id = 'grant-1', [string]$ConsentType = 'Principal', [string]$Scope = ' Mail.Read offline_access')
        [pscustomobject]@{
            Id = $Id; ClientId = 'client-1'; ConsentType = $ConsentType
            PrincipalId = $(if ($ConsentType -eq 'AllPrincipals') { $null } else { 'user-1' })
            ResourceId = 'graph-sp'; Scope = $Scope
        }
    }

    # servicePrincipal.appRoles (id, value) and appRoleAssignment (id, appRoleId,
    # principalId, principalDisplayName, principalType, resourceId, resourceDisplayName,
    # createdDateTime).
    # https://learn.microsoft.com/graph/api/resources/approleassignment
    function New-MockGraphServicePrincipal {
        [pscustomobject]@{
            Id       = 'graph-sp'
            AppId    = '00000003-0000-0000-c000-000000000000'
            AppRoles = @(
                [pscustomobject]@{ Id = 'role-mail-read'; Value = 'Mail.Read' }
                [pscustomobject]@{ Id = 'role-dir-read'; Value = 'Directory.Read.All' }
            )
        }
    }

    function New-MockAppRoleAssignment {
        param([string]$Id = 'assignment-1', [string]$RoleId = 'role-mail-read')
        [pscustomobject]@{
            Id = $Id; AppRoleId = $RoleId; PrincipalId = 'app-sp-1'; PrincipalDisplayName = 'Contoso Mailer'
            PrincipalType = 'ServicePrincipal'; ResourceId = 'graph-sp'; ResourceDisplayName = 'Microsoft Graph'
            CreatedDateTime = [datetime]'2026-07-01T10:00:00Z'
        }
    }

    # Unified audit log: https://learn.microsoft.com/purview/audit-log-activities (operation
    # names) and https://learn.microsoft.com/powershell/module/exchangepowershell/search-unifiedauditlog
    # (ResultCount, AuditSearchRequestMetadata). The detail is a JSON string in AuditData.
    function New-MockAuditRecord {
        param(
            [string]$Id = 'event-1',
            [string]$Operation = 'New-InboxRule',
            [object]$ResultCount = $null,
            [string]$RecordType = 'ExchangeAdmin',
            [string]$CreationTime = '2026-08-10T12:00:00'
        )

        $auditData = [ordered]@{
            CreationTime    = $CreationTime
            Id              = $Id
            Operation       = $Operation
            UserId          = 'avery.abara@example.com'
            Workload        = 'Exchange'
            ObjectId        = 'avery.abara@example.com'
            MailboxOwnerUPN = 'avery.abara@example.com'
            ClientIP        = '203.0.113.5'
            ResultStatus    = 'True'
            Parameters      = @(
                [ordered]@{ Name = 'Name'; Value = 'Forward invoices' }
                [ordered]@{ Name = 'ForwardTo'; Value = 'inbox@fabrikam.example.net' }
            )
        } | ConvertTo-Json -Compress -Depth 5

        $record = [ordered]@{ RecordType = $RecordType; AuditData = $auditData }
        if ($null -ne $ResultCount) { $record['ResultCount'] = $ResultCount }
        return [pscustomobject]$record
    }
}

AfterAll {
    Remove-Variable -Name MxCalls, MxSessions, MxRecordTypes -Scope Global -ErrorAction SilentlyContinue
    Remove-Module M365ReportLibrary -Force -ErrorAction SilentlyContinue
}

Describe 'Collector output matches the committed sample files' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
        Mock Disconnect-ExchangeOnline -MockWith { }
        Mock Get-AcceptedDomain -MockWith { New-MockAcceptedDomain }
        Mock Get-EXOMailbox -MockWith { New-MockMailbox -ForwardingSmtpAddress 'smtp:inbox@fabrikam.example.net' -GrantSendOnBehalfTo @('Devon Moreau') }
        Mock Get-InboxRule -MockWith { New-MockInboxRule -ForwardTo @('"Contact" [SMTP:inbox@fabrikam.example.net]') }
        Mock Get-TransportRule -MockWith { New-MockTransportRule -RedirectMessageTo @('archive@fabrikam.example.net') }
        Mock Get-MailboxPermission -MockWith { New-MockMailboxPermission }
        Mock Get-RecipientPermission -MockWith { New-MockRecipientPermission }
        Mock Get-OrganizationConfig -MockWith { [pscustomobject]@{ AuditDisabled = $false } }
        Mock Get-AdminAuditLogConfig -MockWith { [pscustomobject]@{ UnifiedAuditLogIngestionEnabled = $true } }
        Mock Get-MgOauth2PermissionGrant -MockWith { New-MockGrant }
        Mock Get-MgServicePrincipal -MockWith { New-MockGraphServicePrincipal }
        Mock Get-MgServicePrincipalAppRoleAssignedTo -MockWith { New-MockAppRoleAssignment }
        Mock Search-UnifiedAuditLog -MockWith { New-MockAuditRecord -Operation 'Set-Mailbox' }
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It '<Csv> columns match the sample and the schema' -ForEach @(
        @{ Csv = 'accepted-domains.csv'; Script = 'Get-AcceptedDomains.ps1'; Key = 'AcceptedDomains' }
        @{ Csv = 'mailbox-forwarding.csv'; Script = 'Get-MailboxForwarding.ps1'; Key = 'MailboxForwarding' }
        @{ Csv = 'send-on-behalf.csv'; Script = 'Get-SendOnBehalf.ps1'; Key = 'SendOnBehalf' }
        @{ Csv = 'inbox-rules.csv'; Script = 'Get-InboxRules.ps1'; Key = 'InboxRules' }
        @{ Csv = 'transport-rules.csv'; Script = 'Get-TransportRules.ps1'; Key = 'TransportRules' }
        @{ Csv = 'mailbox-full-access.csv'; Script = 'Get-MailboxFullAccess.ps1'; Key = 'MailboxFullAccess' }
        @{ Csv = 'send-as-permissions.csv'; Script = 'Get-SendAsPermissions.ps1'; Key = 'SendAsPermissions' }
        @{ Csv = 'delegated-consents.csv'; Script = 'Get-DelegatedConsents.ps1'; Key = 'DelegatedConsents' }
        @{ Csv = 'app-role-assignments.csv'; Script = 'Get-AppRoleAssignments.ps1'; Key = 'AppRoleAssignments' }
        @{ Csv = 'audit-configuration.csv'; Script = 'Get-AuditConfiguration.ps1'; Key = 'AuditConfiguration' }
    ) {
        Invoke-CollectorScript $Script @{ OutputPath = $script:folder; SkipConnect = $true }

        $produced = Join-Path $script:folder $Csv
        @(Import-Csv -LiteralPath $produced).Count | Should -BeGreaterThan 0
        Get-HeaderText -Path $produced | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples $Csv))
        Get-HeaderText -Path $produced | Should -Be ($script:Schema[$Key] -join ',')
    }

    It '<Csv> columns match the sample and the schema' -ForEach @(
        @{ Csv = 'mailbox-change-events.csv'; Script = 'Get-MailboxChangeEvents.ps1' }
        @{ Csv = 'mail-access-events.csv'; Script = 'Get-MailAccessEvents.ps1' }
    ) {
        Invoke-CollectorScript $Script @{
            OutputPath = $script:folder; SkipConnect = $true
            StartDate = [datetime]'2026-08-10T00:00:00Z'; EndDate = [datetime]'2026-08-11T00:00:00Z'
        }

        $produced = Join-Path $script:folder $Csv
        @(Import-Csv -LiteralPath $produced).Count | Should -BeGreaterThan 0
        Get-HeaderText -Path $produced | Should -Be (Get-HeaderText -Path (Join-Path $script:Samples $Csv))
        Get-HeaderText -Path $produced | Should -Be ($script:Schema.AuditEvents -join ',')
    }

    It 'every committed sample has the schema columns, and the sets cover external and internal forwarding' {
        foreach ($pair in @(
                @{ Csv = 'accepted-domains.csv'; Key = 'AcceptedDomains' }, @{ Csv = 'mailbox-forwarding.csv'; Key = 'MailboxForwarding' }
                @{ Csv = 'send-on-behalf.csv'; Key = 'SendOnBehalf' }, @{ Csv = 'inbox-rules.csv'; Key = 'InboxRules' }
                @{ Csv = 'transport-rules.csv'; Key = 'TransportRules' }, @{ Csv = 'mailbox-full-access.csv'; Key = 'MailboxFullAccess' }
                @{ Csv = 'send-as-permissions.csv'; Key = 'SendAsPermissions' }, @{ Csv = 'delegated-consents.csv'; Key = 'DelegatedConsents' }
                @{ Csv = 'app-role-assignments.csv'; Key = 'AppRoleAssignments' }, @{ Csv = 'audit-configuration.csv'; Key = 'AuditConfiguration' }
                @{ Csv = 'mailbox-change-events.csv'; Key = 'AuditEvents' }, @{ Csv = 'mail-access-events.csv'; Key = 'AuditEvents' }
            )) {
            Get-HeaderText -Path (Join-Path $script:Samples $pair.Csv) | Should -Be ($script:Schema[$pair.Key] -join ',')
        }
        $forwarding = Import-Csv -LiteralPath (Join-Path $script:Samples 'mailbox-forwarding.csv')
        ($forwarding | Where-Object IsExternal -eq 'True') | Should -Not -BeNullOrEmpty
        ($forwarding | Where-Object IsExternal -eq 'False') | Should -Not -BeNullOrEmpty
        ($forwarding | Where-Object { $_.ForwardingAddress }) | Should -Not -BeNullOrEmpty
        ($forwarding | Where-Object { $_.ForwardingSmtpAddress }) | Should -Not -BeNullOrEmpty
    }
}

Describe 'Connections target the right endpoints for each -Environment' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-ExchangeOnline -ModuleName M365ReportLibrary -MockWith { }
        Mock Disconnect-ExchangeOnline -MockWith { }
        Mock Connect-MgGraph -ModuleName M365ReportLibrary -MockWith { }
        Mock Get-AcceptedDomain -MockWith { New-MockAcceptedDomain }
        Mock Get-MgOauth2PermissionGrant -MockWith { New-MockGrant }
        Mock Get-MgServicePrincipal -MockWith { New-MockGraphServicePrincipal }
        Mock Get-MgServicePrincipalAppRoleAssignedTo -MockWith { New-MockAppRoleAssignment }
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'connects Exchange Online to <ExchangeName> in <Cloud>' -ForEach @(
        @{ Cloud = 'Commercial'; ExchangeName = 'O365Default' }
        @{ Cloud = 'GCC'; ExchangeName = 'O365Default' }
        @{ Cloud = 'GCCHigh'; ExchangeName = 'O365USGovGCCHigh' }
    ) {
        Invoke-CollectorScript 'Get-AcceptedDomains.ps1' @{ OutputPath = $script:folder; Environment = $Cloud; WarningAction = 'SilentlyContinue' }

        Should -Invoke Connect-ExchangeOnline -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter { $ExchangeEnvironmentName -eq $ExchangeName }
        Should -Invoke Disconnect-ExchangeOnline -Times 1 -Exactly
    }

    It 'connects Graph to <GraphEnvironment> in <Cloud>' -ForEach @(
        @{ Cloud = 'Commercial'; GraphEnvironment = 'Global' }
        @{ Cloud = 'GCC'; GraphEnvironment = 'Global' }
        @{ Cloud = 'GCCHigh'; GraphEnvironment = 'USGov' }
    ) {
        Invoke-CollectorScript 'Get-DelegatedConsents.ps1' @{ OutputPath = $script:folder; Environment = $Cloud }

        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter {
            $Environment -eq $GraphEnvironment -and $Scopes -contains 'Directory.Read.All'
        }
    }

    It 'asks for Application.Read.All to read app role assignments' {
        Invoke-CollectorScript 'Get-AppRoleAssignments.ps1' @{ OutputPath = $script:folder }

        Should -Invoke Connect-MgGraph -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter { $Scopes -contains 'Application.Read.All' }
    }

    It 'signs in app-only to Exchange Online when given a certificate' {
        Invoke-CollectorScript 'Get-AcceptedDomains.ps1' @{
            OutputPath = $script:folder; AppId = 'app'; CertificateThumbprint = 'thumb'; Organization = 'example.onmicrosoft.com'
            WarningAction = 'SilentlyContinue'
        }

        Should -Invoke Connect-ExchangeOnline -ModuleName M365ReportLibrary -Times 1 -Exactly -ParameterFilter {
            $AppId -eq 'app' -and $CertificateThumbprint -eq 'thumb' -and $Organization -eq 'example.onmicrosoft.com'
        }
    }

    It 'uses the existing session and does not disconnect it with -SkipConnect' {
        Invoke-CollectorScript 'Get-AcceptedDomains.ps1' @{ OutputPath = $script:folder; SkipConnect = $true }

        Should -Not -Invoke Connect-ExchangeOnline -ModuleName M365ReportLibrary
        Should -Not -Invoke Disconnect-ExchangeOnline
    }
}

Describe 'Mailbox forwarding' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Disconnect-ExchangeOnline -MockWith { }
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'reads the Delivery property set for every mailbox' {
        Mock Get-EXOMailbox -MockWith { New-MockMailbox -ForwardingSmtpAddress 'smtp:x@example.com' }
        Mock Get-AcceptedDomain -MockWith { New-MockAcceptedDomain }

        Invoke-CollectorScript 'Get-MailboxForwarding.ps1' @{ OutputPath = $script:folder; SkipConnect = $true }

        Should -Invoke Get-EXOMailbox -Times 1 -Exactly -ParameterFilter { $ResultSize -eq 'Unlimited' -and $PropertySets -contains 'Delivery' }
        Should -Invoke Get-AcceptedDomain -Times 1 -Exactly -ParameterFilter { $ResultSize -eq 'Unlimited' }
    }

    It 'decides external forwarding by comparing the SMTP domain with the accepted domains' {
        Mock Get-AcceptedDomain -MockWith { @((New-MockAcceptedDomain -Name 'Example.com'), (New-MockAcceptedDomain -Name 'mail.example.com' -Default $false)) }
        Mock Get-EXOMailbox -MockWith {
            @(
                (New-MockMailbox -Id 'a' -Upn 'a@example.com' -ForwardingSmtpAddress 'smtp:boss@EXAMPLE.com')
                (New-MockMailbox -Id 'b' -Upn 'b@example.com' -ForwardingSmtpAddress 'smtp:someone@fabrikam.example.net')
                (New-MockMailbox -Id 'c' -Upn 'c@example.com' -ForwardingSmtpAddress 'smtp:someone@sub.example.com')
                (New-MockMailbox -Id 'd' -Upn 'd@example.com' -ForwardingAddress 'Devon Moreau')
                (New-MockMailbox -Id 'e' -Upn 'e@example.com')
            )
        }

        Invoke-CollectorScript 'Get-MailboxForwarding.ps1' @{ OutputPath = $script:folder; SkipConnect = $true }

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:folder 'mailbox-forwarding.csv'))
        $rows.Count | Should -Be 4
        ($rows | Where-Object ExternalDirectoryObjectId -eq 'a').IsExternal | Should -Be 'False'
        ($rows | Where-Object ExternalDirectoryObjectId -eq 'a').ForwardingSmtpDomain | Should -Be 'example.com'
        ($rows | Where-Object ExternalDirectoryObjectId -eq 'b').IsExternal | Should -Be 'True'
        ($rows | Where-Object ExternalDirectoryObjectId -eq 'c').IsExternal | Should -Be 'True'
        ($rows | Where-Object ExternalDirectoryObjectId -eq 'd').IsExternal | Should -Be 'False'
        $rows.ExternalDirectoryObjectId | Should -Not -Contain 'e'
    }

    It 'treats a wildcard accepted domain as covering its subdomains' {
        Mock Get-AcceptedDomain -MockWith { New-MockAcceptedDomain -Name '*.example.com' }
        Mock Get-EXOMailbox -MockWith { New-MockMailbox -ForwardingSmtpAddress 'smtp:x@sub.example.com' }

        Invoke-CollectorScript 'Get-MailboxForwarding.ps1' @{ OutputPath = $script:folder; SkipConnect = $true }

        (Import-Csv -LiteralPath (Join-Path $script:folder 'mailbox-forwarding.csv')).IsExternal | Should -Be 'False'
    }

    It 'leaves IsExternal empty and warns when the accepted domains cannot be read' {
        Mock Get-AcceptedDomain -MockWith { throw 'The term is not recognized' }
        Mock Get-EXOMailbox -MockWith { New-MockMailbox -ForwardingSmtpAddress 'smtp:x@fabrikam.example.net' }

        Invoke-CollectorScript 'Get-MailboxForwarding.ps1' @{ OutputPath = $script:folder; SkipConnect = $true; WarningAction = 'SilentlyContinue' }

        $row = Import-Csv -LiteralPath (Join-Path $script:folder 'mailbox-forwarding.csv')
        $row.IsExternal | Should -Be ''
        $row.ForwardingSmtpDomain | Should -Be 'fabrikam.example.net'
        Get-LogText -Folder $script:folder | Should -Match 'Get-AcceptedDomain is unavailable'
    }

    It 'writes the header only and names the role when the mailbox read is refused' {
        Mock Get-AcceptedDomain -MockWith { New-MockAcceptedDomain }
        Mock Get-EXOMailbox -MockWith { throw 'Insufficient permissions' }

        Invoke-CollectorScript 'Get-MailboxForwarding.ps1' @{ OutputPath = $script:folder; SkipConnect = $true; WarningAction = 'SilentlyContinue' }

        @(Get-Content -LiteralPath (Join-Path $script:folder 'mailbox-forwarding.csv')).Count | Should -Be 1
        Get-LogText -Folder $script:folder | Should -Match 'UNVERIFIED'
    }

    It 'writes one send-on-behalf row per delegate' {
        Mock Get-EXOMailbox -MockWith {
            @((New-MockMailbox -Id 'a' -GrantSendOnBehalfTo @('Devon Moreau', 'Casey Dlamini')), (New-MockMailbox -Id 'b'))
        }

        Invoke-CollectorScript 'Get-SendOnBehalf.ps1' @{ OutputPath = $script:folder; SkipConnect = $true }

        @(Import-Csv -LiteralPath (Join-Path $script:folder 'send-on-behalf.csv')).Delegate | Should -Be @('Devon Moreau', 'Casey Dlamini')
        Should -Invoke Get-EXOMailbox -ParameterFilter { $PropertySets -contains 'Delivery' }
    }

    It 'lists every accepted domain with its type' {
        Mock Get-AcceptedDomain -MockWith { @((New-MockAcceptedDomain), (New-MockAcceptedDomain -Name 'mail.example.com' -DomainType 'InternalRelay' -Default $false)) }

        Invoke-CollectorScript 'Get-AcceptedDomains.ps1' @{ OutputPath = $script:folder; SkipConnect = $true }

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:folder 'accepted-domains.csv'))
        $rows.Count | Should -Be 2
        $rows[0].IsDefault | Should -Be 'True'
        $rows[1].DomainType | Should -Be 'InternalRelay'
    }
}

Describe 'Rules' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Disconnect-ExchangeOnline -MockWith { }
        Mock Get-AcceptedDomain -MockWith { New-MockAcceptedDomain }
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'reads inbox rules per mailbox, with hidden rules and no result cap' {
        Mock Get-EXOMailbox -MockWith { @((New-MockMailbox -Id 'a' -Upn 'a@example.com'), (New-MockMailbox -Id 'b' -Upn 'b@example.com')) }
        Mock Get-InboxRule -MockWith { New-MockInboxRule }

        Invoke-CollectorScript 'Get-InboxRules.ps1' @{ OutputPath = $script:folder; SkipConnect = $true }

        Should -Invoke Get-InboxRule -Times 2 -Exactly -ParameterFilter { $IncludeHidden -and $ResultSize -eq 'Unlimited' }
        Should -Invoke Get-InboxRule -Times 1 -Exactly -ParameterFilter { $Mailbox -eq 'a@example.com' }
    }

    It 'writes only rules that forward, redirect or delete, and flags external targets' {
        Mock Get-EXOMailbox -MockWith { New-MockMailbox }
        Mock Get-InboxRule -MockWith {
            @(
                (New-MockInboxRule -Identity '1' -Name 'Move to folder')
                (New-MockInboxRule -Identity '2' -Name 'Internal fwd' -ForwardTo @('"Devon Moreau" [SMTP:devon.moreau@example.com]'))
                (New-MockInboxRule -Identity '3' -Name 'External redirect' -RedirectTo @('"Out" [SMTP:inbox@fabrikam.example.net]'))
                (New-MockInboxRule -Identity '4' -Name 'Delete' -DeleteMessage $true)
            )
        }

        Invoke-CollectorScript 'Get-InboxRules.ps1' @{ OutputPath = $script:folder; SkipConnect = $true }

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:folder 'inbox-rules.csv'))
        $rows.RuleIdentity | Should -Be @('2', '3', '4')
        $rows[0].HasExternalTarget | Should -Be 'False'
        $rows[0].TargetDomains | Should -Be 'example.com'
        $rows[1].HasExternalTarget | Should -Be 'True'
        $rows[1].TargetDomains | Should -Be 'fabrikam.example.net'
        $rows[2].DeleteMessage | Should -Be 'True'
    }

    It 'skips a mailbox it cannot read and logs it' {
        Mock Get-EXOMailbox -MockWith { @((New-MockMailbox -Id 'a' -Upn 'a@example.com'), (New-MockMailbox -Id 'b' -Upn 'b@example.com')) }
        $global:MxCalls = 0
        Mock Get-InboxRule -MockWith {
            $global:MxCalls++
            if ($global:MxCalls -eq 1) { throw 'Not allowed for Global Reader' }
            New-MockInboxRule -ForwardTo @('x@fabrikam.example.net')
        }

        Invoke-CollectorScript 'Get-InboxRules.ps1' @{ OutputPath = $script:folder; SkipConnect = $true; WarningAction = 'SilentlyContinue' }

        @(Import-Csv -LiteralPath (Join-Path $script:folder 'inbox-rules.csv')).Count | Should -Be 1
        Get-LogText -Folder $script:folder | Should -Match 'a@example.com'
    }

    It 'writes the header only when no mailbox can be read' {
        Mock Get-EXOMailbox -MockWith { New-MockMailbox }
        Mock Get-InboxRule -MockWith { throw 'Not allowed' }

        Invoke-CollectorScript 'Get-InboxRules.ps1' @{ OutputPath = $script:folder; SkipConnect = $true; WarningAction = 'SilentlyContinue' }

        @(Get-Content -LiteralPath (Join-Path $script:folder 'inbox-rules.csv')).Count | Should -Be 1
        Get-LogText -Folder $script:folder | Should -Match 'Global Reader'
    }

    It 'reads transport rules in full and keeps only redirect and blind-copy style rules' {
        Mock Get-TransportRule -MockWith {
            @(
                (New-MockTransportRule -Name 'Plain' -Guid '1')
                (New-MockTransportRule -Name 'Redirect' -Guid '2' -RedirectMessageTo @('legal@fabrikam.example.net'))
                (New-MockTransportRule -Name 'Bcc' -Guid '3' -BlindCopyTo @('audit@example.com'))
            )
        }

        Invoke-CollectorScript 'Get-TransportRules.ps1' @{ OutputPath = $script:folder; SkipConnect = $true }

        Should -Invoke Get-TransportRule -Times 1 -Exactly -ParameterFilter { $ResultSize -eq 'Unlimited' -and $ExcludeConditionActionDetails -eq $false }
        $rows = @(Import-Csv -LiteralPath (Join-Path $script:folder 'transport-rules.csv'))
        $rows.Name | Should -Be @('Redirect', 'Bcc')
        $rows[0].HasExternalTarget | Should -Be 'True'
        $rows[1].HasExternalTarget | Should -Be 'False'
        $rows[1].BlindCopyTo | Should -Be 'audit@example.com'
    }
}

Describe 'Delegation' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Disconnect-ExchangeOnline -MockWith { }
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'drops SELF, Deny rows and inherited rows from Full Access, and keeps the explicit grant' {
        Mock Get-EXOMailbox -MockWith { New-MockMailbox }
        Mock Get-MailboxPermission -MockWith {
            @(
                (New-MockMailboxPermission -User 'NT AUTHORITY\SELF')
                (New-MockMailboxPermission -User 'denied@example.com' -Deny $true)
                (New-MockMailboxPermission -User 'EXAMPLE\Domain Admins' -IsInherited $true)
                (New-MockMailboxPermission -User 'raymond@example.com')
                (New-MockMailboxPermission -User 'reader@example.com' -AccessRights @('ReadPermission'))
            )
        }

        Invoke-CollectorScript 'Get-MailboxFullAccess.ps1' @{ OutputPath = $script:folder; SkipConnect = $true }

        $rows = @(Import-Csv -LiteralPath (Join-Path $script:folder 'mailbox-full-access.csv'))
        $rows.Count | Should -Be 1
        $rows[0].User | Should -Be 'raymond@example.com'
        $rows[0].AccessRights | Should -Be 'FullAccess'
    }

    It 'calls Get-MailboxPermission once per mailbox with an identity and no result cap' {
        Mock Get-EXOMailbox -MockWith { @((New-MockMailbox -Id 'a' -Upn 'a@example.com'), (New-MockMailbox -Id 'b' -Upn 'b@example.com')) }
        Mock Get-MailboxPermission -MockWith { New-MockMailboxPermission }

        Invoke-CollectorScript 'Get-MailboxFullAccess.ps1' @{ OutputPath = $script:folder; SkipConnect = $true }

        Should -Invoke Get-MailboxPermission -Times 2 -Exactly -ParameterFilter { $ResultSize -eq 'Unlimited' -and $Identity -in @('a@example.com', 'b@example.com') }
    }

    It 'honours -MailboxLimit' {
        Mock Get-EXOMailbox -MockWith { @((New-MockMailbox -Id 'a' -Upn 'a@example.com'), (New-MockMailbox -Id 'b' -Upn 'b@example.com')) }
        Mock Get-MailboxPermission -MockWith { New-MockMailboxPermission }

        Invoke-CollectorScript 'Get-MailboxFullAccess.ps1' @{ OutputPath = $script:folder; SkipConnect = $true; MailboxLimit = 1 }

        Should -Invoke Get-MailboxPermission -Times 1 -Exactly
    }

    It 'reads Send As with no result cap' {
        Mock Get-RecipientPermission -MockWith { New-MockRecipientPermission }

        Invoke-CollectorScript 'Get-SendAsPermissions.ps1' @{ OutputPath = $script:folder; SkipConnect = $true }

        Should -Invoke Get-RecipientPermission -Times 1 -Exactly -ParameterFilter { $ResultSize -eq 'Unlimited' }
        $row = Import-Csv -LiteralPath (Join-Path $script:folder 'send-as-permissions.csv')
        $row.Trustee | Should -Be 'printers@example.com'
        $row.AccessRights | Should -Be 'SendAs'
    }
}

Describe 'Application access to mail' {
    BeforeEach {
        $script:folder = New-TestFolder
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'follows paging by passing -All, and flags mail scopes and tenant-wide consent' {
        Mock Get-MgOauth2PermissionGrant -MockWith {
            @(
                (New-MockGrant -Id 'g1' -Scope ' Mail.Read offline_access')
                (New-MockGrant -Id 'g2' -Scope 'User.Read' -ConsentType 'AllPrincipals')
                (New-MockGrant -Id 'g3' -Scope 'MailboxSettings.Read')
            )
        }

        Invoke-CollectorScript 'Get-DelegatedConsents.ps1' @{ OutputPath = $script:folder; SkipConnect = $true }

        Should -Invoke Get-MgOauth2PermissionGrant -Times 1 -Exactly -ParameterFilter { $All }
        $rows = @(Import-Csv -LiteralPath (Join-Path $script:folder 'delegated-consents.csv'))
        $rows.HasMailScope | Should -Be @('True', 'False', 'True')
        $rows[1].ConsentType | Should -Be 'AllPrincipals'
        $rows[1].PrincipalId | Should -Be ''
        $rows[0].Scope | Should -Be 'Mail.Read offline_access'
    }

    It 'resolves app role names from the Graph service principal and flags mail roles' {
        Mock Get-MgServicePrincipal -MockWith { New-MockGraphServicePrincipal }
        Mock Get-MgServicePrincipalAppRoleAssignedTo -MockWith {
            @((New-MockAppRoleAssignment -Id 'a1' -RoleId 'role-mail-read'), (New-MockAppRoleAssignment -Id 'a2' -RoleId 'role-dir-read'))
        }

        Invoke-CollectorScript 'Get-AppRoleAssignments.ps1' @{ OutputPath = $script:folder; SkipConnect = $true }

        Should -Invoke Get-MgServicePrincipal -Times 1 -Exactly -ParameterFilter { $Filter -like "*00000003-0000-0000-c000-000000000000*" }
        Should -Invoke Get-MgServicePrincipalAppRoleAssignedTo -Times 1 -Exactly -ParameterFilter { $All -and $ServicePrincipalId -eq 'graph-sp' }
        $rows = @(Import-Csv -LiteralPath (Join-Path $script:folder 'app-role-assignments.csv'))
        $rows.AppRoleValue | Should -Be @('Mail.Read', 'Directory.Read.All')
        $rows.IsMailRole | Should -Be @('True', 'False')
        $rows[0].CreatedDateTime | Should -Be '2026-07-01T10:00:00Z'
    }

    It 'writes the header only and names the permission when Graph refuses' {
        Mock Get-MgOauth2PermissionGrant -MockWith { throw 'Forbidden' }

        Invoke-CollectorScript 'Get-DelegatedConsents.ps1' @{ OutputPath = $script:folder; SkipConnect = $true; WarningAction = 'SilentlyContinue' }

        @(Get-Content -LiteralPath (Join-Path $script:folder 'delegated-consents.csv')).Count | Should -Be 1
        Get-LogText -Folder $script:folder | Should -Match 'Directory.Read.All'
    }
}

Describe 'Audit events' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Disconnect-ExchangeOnline -MockWith { }
        $script:range = @{
            StartDate = [datetime]'2026-08-10T00:00:00Z'
            EndDate   = [datetime]'2026-08-11T00:00:00Z'
        }
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'searches the six mailbox change operations in a ReturnLargeSet session' {
        Mock Search-UnifiedAuditLog -MockWith { New-MockAuditRecord }

        Invoke-CollectorScript 'Get-MailboxChangeEvents.ps1' (@{ OutputPath = $script:folder; SkipConnect = $true; SkipExchangeAdmin = $true } + $script:range)

        Should -Invoke Search-UnifiedAuditLog -Times 1 -Exactly -ParameterFilter {
            $SessionCommand -eq 'ReturnLargeSet' -and -not [string]::IsNullOrEmpty($SessionId) -and $ResultSize -eq 5000 -and
            $Operations.Count -eq 6 -and $Operations -ccontains 'New-InboxRule' -and $Operations -ccontains 'Set-InboxRule' -and
            $Operations -ccontains 'UpdateInboxRules' -and $Operations -ccontains 'Set-Mailbox' -and
            $Operations -ccontains 'Add-MailboxPermission' -and $Operations -ccontains 'Remove-MailboxPermission'
        }
    }

    It 'searches the mail access operations' {
        Mock Search-UnifiedAuditLog -MockWith { New-MockAuditRecord -Operation 'MailItemsAccessed' -RecordType 'ExchangeItemAggregated' }

        Invoke-CollectorScript 'Get-MailAccessEvents.ps1' (@{ OutputPath = $script:folder; SkipConnect = $true } + $script:range)

        Should -Invoke Search-UnifiedAuditLog -Times 1 -Exactly -ParameterFilter {
            $Operations.Count -eq 4 -and $Operations -ccontains 'MailItemsAccessed' -and $Operations -ccontains 'Send' -and
            $Operations -ccontains 'SendAs' -and $Operations -ccontains 'SendOnBehalf' -and $SessionCommand -eq 'ReturnLargeSet'
        }
        (Import-Csv -LiteralPath (Join-Path $script:folder 'mail-access-events.csv')).Operation | Should -Be 'MailItemsAccessed'
    }

    It 'reads the actor, mailbox and parameters out of AuditData' {
        Mock Search-UnifiedAuditLog -MockWith { New-MockAuditRecord -Id 'event-9' }

        Invoke-CollectorScript 'Get-MailboxChangeEvents.ps1' (@{ OutputPath = $script:folder; SkipConnect = $true; SkipExchangeAdmin = $true } + $script:range)

        $row = Import-Csv -LiteralPath (Join-Path $script:folder 'mailbox-change-events.csv')
        $row.Id | Should -Be 'event-9'
        $row.Operation | Should -Be 'New-InboxRule'
        $row.RecordType | Should -Be 'ExchangeAdmin'
        $row.UserId | Should -Be 'avery.abara@example.com'
        $row.MailboxOwnerUPN | Should -Be 'avery.abara@example.com'
        $row.Parameters | Should -Be 'Name=Forward invoices;ForwardTo=inbox@fabrikam.example.net'
        $row.CreationTime | Should -Be '2026-08-10T12:00:00Z'
    }

    It 'pages past 100 records by repeating the same session until a page comes back empty' {
        $global:MxCalls = 0
        $global:MxSessions = [System.Collections.Generic.List[string]]::new()
        Mock Search-UnifiedAuditLog -MockWith {
            $global:MxSessions.Add($SessionId)
            $global:MxCalls++
            if ($global:MxCalls -le 3) {
                1..100 | ForEach-Object { New-MockAuditRecord -Id "p$($global:MxCalls)-$_" -ResultCount 300 }
            }
        }

        Invoke-CollectorScript 'Get-MailboxChangeEvents.ps1' (@{ OutputPath = $script:folder; SkipConnect = $true; SkipExchangeAdmin = $true } + $script:range)

        @(Import-Csv -LiteralPath (Join-Path $script:folder 'mailbox-change-events.csv')).Count | Should -Be 300
        @($global:MxSessions | Select-Object -Unique).Count | Should -Be 1
        $global:MxCalls | Should -BeGreaterOrEqual 3
    }

    It 'stops at the 50,000-record cap, writes nothing from that window, and names the window to re-run' {
        Mock Search-UnifiedAuditLog -MockWith { New-MockAuditRecord -ResultCount 60000 }

        { Invoke-CollectorScript 'Get-MailboxChangeEvents.ps1' (@{ OutputPath = $script:folder; SkipConnect = $true; SkipExchangeAdmin = $true; WarningAction = 'SilentlyContinue' } + $script:range) } |
            Should -Throw '*50,000*'

        @(Get-Content -LiteralPath (Join-Path $script:folder 'mailbox-change-events.csv')).Count | Should -Be 1
        Get-LogText -Folder $script:folder | Should -Match '-StartDate 2026-08-10T00:00:00Z -EndDate 2026-08-11T00:00:00Z'
    }

    It 'stops paging after 50,000 records have been read in a session' {
        $global:MxCalls = 0
        Mock Search-UnifiedAuditLog -MockWith {
            $global:MxCalls++
            1..5000 | ForEach-Object { New-MockAuditRecord -Id "r$($global:MxCalls)-$_" }
        }

        { Invoke-CollectorScript 'Get-MailAccessEvents.ps1' (@{ OutputPath = $script:folder; SkipConnect = $true; WarningAction = 'SilentlyContinue' } + $script:range) } |
            Should -Throw '*50,000*'

        Should -Invoke Search-UnifiedAuditLog -Times 10 -Exactly
    }

    It 'also searches -RecordType ExchangeAdmin for mail flow rule changes and keeps only TransportRule operations' {
        $global:MxRecordTypes = [System.Collections.Generic.List[string]]::new()
        Mock Search-UnifiedAuditLog -MockWith {
            $global:MxRecordTypes.Add([string]($RecordType -join ','))
            if ($RecordType -contains 'ExchangeAdmin') {
                @((New-MockAuditRecord -Id 'rule' -Operation 'Set-TransportRule'), (New-MockAuditRecord -Id 'other' -Operation 'Set-OrganizationConfig'))
            }
            else {
                New-MockAuditRecord -Id 'inbox' -Operation 'New-InboxRule'
            }
        }

        Invoke-CollectorScript 'Get-MailboxChangeEvents.ps1' (@{ OutputPath = $script:folder; SkipConnect = $true } + $script:range)

        $global:MxRecordTypes | Should -Contain 'ExchangeAdmin'
        @(Import-Csv -LiteralPath (Join-Path $script:folder 'mailbox-change-events.csv')).Id | Should -Be @('inbox', 'rule')
    }

    It 'does not write events at or after a capped Exchange admin window' {
        # ReturnLargeSet stops at 50,000 and is unsorted. Writing the mailbox events
        # from later in the range would move the watermark past the admin events that
        # the capped window never returned.
        # https://learn.microsoft.com/powershell/module/exchangepowershell/search-unifiedauditlog
        Mock Search-UnifiedAuditLog -MockWith {
            if ($RecordType -contains 'ExchangeAdmin') {
                New-MockAuditRecord -ResultCount 60000
            }
            else {
                $stamp = $StartDate.ToUniversalTime().AddHours(12).ToString('yyyy-MM-ddTHH:mm:ss', [cultureinfo]::InvariantCulture)
                New-MockAuditRecord -Id ('mbx-' + $StartDate.ToUniversalTime().ToString('dd')) -Operation 'Set-Mailbox' -CreationTime $stamp
            }
        }

        {
            Invoke-CollectorScript 'Get-MailboxChangeEvents.ps1' @{
                OutputPath = $script:folder; SkipConnect = $true; WarningAction = 'SilentlyContinue'
                StartDate  = [datetime]'2026-08-10T00:00:00Z'; EndDate = [datetime]'2026-08-12T00:00:00Z'; WindowHours = 24
            }
        } | Should -Throw '*50,000*'

        @(Get-Content -LiteralPath (Join-Path $script:folder 'mailbox-change-events.csv')).Count | Should -Be 1
    }

    It 'keeps earlier windows from both searches when a later window hits the cap' {
        Mock Search-UnifiedAuditLog -MockWith {
            $firstWindow = $StartDate.ToUniversalTime() -lt [datetime]'2026-08-11T00:00:00Z'
            if ($RecordType -contains 'ExchangeAdmin') {
                if ($firstWindow) {
                    New-MockAuditRecord -Id 'admin-rule' -Operation 'Set-TransportRule' -CreationTime '2026-08-10T06:00:00' -ResultCount 1
                }
                else {
                    New-MockAuditRecord -Id 'admin-late' -Operation 'Set-TransportRule' -CreationTime '2026-08-11T06:00:00' -ResultCount 1
                }
            }
            elseif ($firstWindow) {
                New-MockAuditRecord -Id 'day1' -Operation 'New-InboxRule' -CreationTime '2026-08-10T12:00:00' -ResultCount 1
            }
            else {
                New-MockAuditRecord -Id 'day2' -Operation 'Set-Mailbox' -ResultCount 60000
            }
        }

        {
            Invoke-CollectorScript 'Get-MailboxChangeEvents.ps1' @{
                OutputPath = $script:folder; SkipConnect = $true; WarningAction = 'SilentlyContinue'
                StartDate  = [datetime]'2026-08-10T00:00:00Z'; EndDate = [datetime]'2026-08-12T00:00:00Z'; WindowHours = 24
            }
        } | Should -Throw '*50,000*'

        @(Import-Csv -LiteralPath (Join-Path $script:folder 'mailbox-change-events.csv')).Id | Should -Be @('day1', 'admin-rule')
    }

    It 'starts at the latest CreationTime already collected' {
        Export-AppendCsv -Path (Join-Path $script:folder 'mail-access-events.csv') -Column $script:Schema.AuditEvents -Rows @(
            [pscustomobject]@{
                CreationTime = '2026-08-09T05:00:00Z'; Id = 'old'; RecordType = 'ExchangeItem'; Operation = 'Send'; UserId = 'u'
                Workload = 'Exchange'; ObjectId = ''; MailboxOwnerUPN = ''; ClientIP = ''; ResultStatus = ''; Parameters = ''
            }
        )
        Mock Search-UnifiedAuditLog -MockWith { $null }

        Invoke-CollectorScript 'Get-MailAccessEvents.ps1' @{ OutputPath = $script:folder; SkipConnect = $true; EndDate = [datetime]'2026-08-09T06:00:00Z' }

        Should -Invoke Search-UnifiedAuditLog -ParameterFilter { $StartDate.ToUniversalTime() -eq [datetime]'2026-08-09T05:00:00Z' }
    }

    It 'rejects an inverted range instead of reporting success' {
        Mock Search-UnifiedAuditLog -MockWith { $null }

        { Invoke-CollectorScript 'Get-MailAccessEvents.ps1' @{ OutputPath = $script:folder; SkipConnect = $true; StartDate = [datetime]'2026-08-11'; EndDate = [datetime]'2026-08-10' } } |
            Should -Throw '*range is empty*'
    }

    It 'writes the header only and names the Audit Reader role when the search is refused' {
        Mock Search-UnifiedAuditLog -MockWith { throw 'The term is not recognized' }

        Invoke-CollectorScript 'Get-MailAccessEvents.ps1' (@{ OutputPath = $script:folder; SkipConnect = $true; WarningAction = 'SilentlyContinue' } + $script:range)

        @(Get-Content -LiteralPath (Join-Path $script:folder 'mail-access-events.csv')).Count | Should -Be 1
        Get-LogText -Folder $script:folder | Should -Match 'Audit Reader'
    }
}

Describe 'Audit configuration' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Disconnect-ExchangeOnline -MockWith { }
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'writes the organization row and a row per mailbox' {
        Mock Get-OrganizationConfig -MockWith { [pscustomobject]@{ AuditDisabled = $false } }
        Mock Get-AdminAuditLogConfig -MockWith { [pscustomobject]@{ UnifiedAuditLogIngestionEnabled = $true } }
        Mock Get-EXOMailbox -MockWith { New-MockMailbox }

        Invoke-CollectorScript 'Get-AuditConfiguration.ps1' @{ OutputPath = $script:folder; SkipConnect = $true }

        Should -Invoke Get-EXOMailbox -Times 1 -Exactly -ParameterFilter { $PropertySets -contains 'Audit' -and $ResultSize -eq 'Unlimited' }
        $rows = @(Import-Csv -LiteralPath (Join-Path $script:folder 'audit-configuration.csv'))
        $rows.Count | Should -Be 2
        $rows[0].Scope | Should -Be 'Organization'
        $rows[0].AuditDisabled | Should -Be 'False'
        $rows[0].UnifiedAuditLogIngestionEnabled | Should -Be 'True'
        $rows[1].Scope | Should -Be 'Mailbox'
        $rows[1].DefaultAuditSet | Should -Be 'Admin;Delegate;Owner'
    }

    It 'logs the refused call, keeps the others, and attempts GCC High with a warning' {
        Mock Get-OrganizationConfig -MockWith { throw 'not recognized' }
        Mock Get-AdminAuditLogConfig -MockWith { [pscustomobject]@{ UnifiedAuditLogIngestionEnabled = $false } }
        Mock Get-EXOMailbox -MockWith { New-MockMailbox }

        Invoke-CollectorScript 'Get-AuditConfiguration.ps1' @{ OutputPath = $script:folder; SkipConnect = $true; Environment = 'GCCHigh'; WarningAction = 'SilentlyContinue' }

        $org = Import-Csv -LiteralPath (Join-Path $script:folder 'audit-configuration.csv') | Where-Object Scope -eq 'Organization'
        $org.AuditDisabled | Should -Be ''
        $org.UnifiedAuditLogIngestionEnabled | Should -Be 'False'
        $log = Get-LogText -Folder $script:folder
        $log | Should -Match 'UNVERIFIED'
        $log | Should -Match 'Get-OrganizationConfig is unavailable'
    }

    It 'writes the header only when every call is refused' {
        Mock Get-OrganizationConfig -MockWith { throw 'no' }
        Mock Get-AdminAuditLogConfig -MockWith { throw 'no' }
        Mock Get-EXOMailbox -MockWith { throw 'no' }

        Invoke-CollectorScript 'Get-AuditConfiguration.ps1' @{ OutputPath = $script:folder; SkipConnect = $true; WarningAction = 'SilentlyContinue' }

        @(Get-Content -LiteralPath (Join-Path $script:folder 'audit-configuration.csv')).Count | Should -Be 1
    }
}

Describe 'Source availability' {
    It 'covers every source in all three clouds with a Learn link and a valid status' {
        foreach ($source in $script:Schema.SourceAvailability.Keys) {
            foreach ($cloud in 'Commercial', 'GCC', 'GCCHigh') {
                $entry = $script:Schema.SourceAvailability[$source][$cloud]
                $entry.Status | Should -BeIn @('Available', 'NotAvailable', 'Unverified')
                $entry.Reference | Should -Match '^https://learn\.microsoft\.com/'
            }
        }
        $script:Schema.SourceAvailability.Keys.Count | Should -Be 12
    }

    It 'skips only a NotAvailable source' {
        . (Join-Path $script:Collectors 'MailboxExfiltrationHelpers.ps1')
        $synthetic = @{ SourceAvailability = @{ Demo = @{
                    Commercial = @{ Status = 'Available'; Reference = 'https://learn.microsoft.com/a' }
                    GCC        = @{ Status = 'Unverified'; Reference = 'https://learn.microsoft.com/b' }
                    GCCHigh    = @{ Status = 'NotAvailable'; Reference = 'https://learn.microsoft.com/c' }
                } } }

        (Get-MailboxSourceAvailability -Source Demo -Environment Commercial -Schema $synthetic).ShouldSkip | Should -BeFalse
        (Get-MailboxSourceAvailability -Source Demo -Environment GCC -Schema $synthetic).ShouldSkip | Should -BeFalse
        (Get-MailboxSourceAvailability -Source Demo -Environment GCCHigh -Schema $synthetic).ShouldSkip | Should -BeTrue
        { Get-MailboxSourceAvailability -Source Nope -Environment GCC -Schema $synthetic } | Should -Throw '*Unknown source*'
    }

    It 'attempts an UNVERIFIED source with a warning in GCC' {
        $folder = New-TestFolder
        try {
            Mock Get-AcceptedDomain -MockWith { New-MockAcceptedDomain }
            Mock Disconnect-ExchangeOnline -MockWith { }

            Invoke-CollectorScript 'Get-AcceptedDomains.ps1' @{ OutputPath = $folder; SkipConnect = $true; Environment = 'GCC'; WarningAction = 'SilentlyContinue' }

            Should -Invoke Get-AcceptedDomain -Times 1 -Exactly
            Get-LogText -Folder $folder | Should -Match 'UNVERIFIED'
        }
        finally {
            Remove-Item -LiteralPath $folder -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'writes the header only, logs why, and calls nothing when the schema marks a source NotAvailable' {
        # The contract doc marks no source NotAvailable, so run a copy of the collector
        # beside a copy of the schema that does, to exercise the real skip path.
        $sandbox = New-TestFolder
        $folder = New-TestFolder
        try {
            $copy = Join-Path $sandbox 'reports/mailbox-exfiltration-risk/collectors'
            New-Item -Path $copy -ItemType Directory -Force | Out-Null
            Copy-Item -Path (Join-Path $script:Collectors '*') -Destination $copy -Recurse
            Copy-Item -Path (Join-Path $script:Root 'shared') -Destination (Join-Path $sandbox 'shared') -Recurse

            $schemaCopy = Join-Path $copy 'MailboxExfiltrationSchema.psd1'
            $text = Get-Content -LiteralPath $schemaCopy -Raw
            $edited = $text -replace "(AcceptedDomains = @\{\s+Commercial = @\{ Status = 'Available'.+\r?\n\s+GCC\s+= @\{ Status = )'Unverified'", "`$1'NotAvailable'"
            $edited | Should -Not -Be $text
            Set-Content -LiteralPath $schemaCopy -Value $edited

            Mock Get-AcceptedDomain -MockWith { throw 'must not be called' }
            Mock Connect-M365Service -MockWith { throw 'must not be called' }

            & (Join-Path $copy 'Get-AcceptedDomains.ps1') -OutputPath $folder -Environment GCC -WarningAction SilentlyContinue

            Should -Not -Invoke Get-AcceptedDomain
            Should -Not -Invoke Connect-M365Service
            $path = Join-Path $folder 'accepted-domains.csv'
            @(Get-Content -LiteralPath $path).Count | Should -Be 1
            Get-HeaderText -Path $path | Should -Be ($script:Schema.AcceptedDomains -join ',')
            Get-LogText -Folder $folder | Should -Match 'unavailable in GCC'
        }
        finally {
            Remove-Item -LiteralPath $sandbox, $folder -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

Describe 'Run-All.ps1' {
    BeforeEach {
        $script:folder = New-TestFolder
        Mock Connect-M365Service -MockWith { }
        Mock Disconnect-ExchangeOnline -MockWith { }
        Mock Get-AcceptedDomain -MockWith { New-MockAcceptedDomain }
        Mock Get-EXOMailbox -MockWith { New-MockMailbox -ForwardingSmtpAddress 'smtp:x@fabrikam.example.net' -GrantSendOnBehalfTo @('Devon Moreau') }
        Mock Get-InboxRule -MockWith { New-MockInboxRule -ForwardTo @('x@fabrikam.example.net') }
        Mock Get-TransportRule -MockWith { New-MockTransportRule -RedirectMessageTo @('archive@fabrikam.example.net') }
        Mock Get-MailboxPermission -MockWith { New-MockMailboxPermission }
        Mock Get-RecipientPermission -MockWith { New-MockRecipientPermission }
        Mock Get-OrganizationConfig -MockWith { [pscustomobject]@{ AuditDisabled = $false } }
        Mock Get-AdminAuditLogConfig -MockWith { [pscustomobject]@{ UnifiedAuditLogIngestionEnabled = $true } }
        Mock Get-MgOauth2PermissionGrant -MockWith { New-MockGrant }
        Mock Get-MgServicePrincipal -MockWith { New-MockGraphServicePrincipal }
        Mock Get-MgServicePrincipalAppRoleAssignedTo -MockWith { New-MockAppRoleAssignment }
        Mock Search-UnifiedAuditLog -MockWith { New-MockAuditRecord }
        Mock Get-MgUser -MockWith { [pscustomobject]@{ Id = 'u1'; DisplayName = 'Avery'; UserPrincipalName = 'avery@example.com'; Mail = 'avery@example.com'; UserType = 'Member'; AccountEnabled = $true } }
    }

    AfterEach {
        Remove-Item -LiteralPath $script:folder -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'signs in to Exchange Online and Graph once each and writes all thirteen CSVs' {
        Invoke-CollectorScript 'Run-All.ps1' @{
            OutputPath = $script:folder; WarningAction = 'SilentlyContinue'
            StartDate = [datetime]'2026-08-10T00:00:00Z'; EndDate = [datetime]'2026-08-11T00:00:00Z'
        }

        Should -Invoke Connect-M365Service -Times 1 -Exactly -ParameterFilter { $Service -eq 'ExchangeOnline' }
        Should -Invoke Connect-M365Service -Times 1 -Exactly -ParameterFilter { $Service -eq 'Graph' }
        foreach ($csv in 'users', 'accepted-domains', 'mailbox-forwarding', 'send-on-behalf', 'inbox-rules', 'transport-rules', 'mailbox-full-access',
            'send-as-permissions', 'delegated-consents', 'app-role-assignments', 'audit-configuration', 'mailbox-change-events', 'mail-access-events') {
            Test-Path -LiteralPath (Join-Path $script:folder "$csv.csv") | Should -BeTrue -Because "$csv.csv should exist"
        }
        Should -Invoke Disconnect-ExchangeOnline -Times 1 -Exactly
    }
}
