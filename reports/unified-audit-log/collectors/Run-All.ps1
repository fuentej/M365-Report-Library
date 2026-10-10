#Requires -Version 7.0

<#
    .SYNOPSIS
        Runs the five unified audit log collectors into one output folder.

    .DESCRIPTION
        Exchange Online (sources 1 and 4), Security & Compliance PowerShell (source 5) and Microsoft
        Graph (source 2) are signed in to one after another and each is signed out before the next, so
        two Exchange-family sessions never overlap. Source 3 reads the Management Activity API with a
        token you supply; without -AccessToken, -TenantId and -PublisherIdentifier it is skipped with a
        line in run.log and no file.

        A collector whose source is refused, or documented as unavailable in this cloud, leaves a
        header-only CSV and a line in run.log without stopping the others. In GCC High the Graph Audit
        Search API (source 2) is skipped that way.

        Two collectors send a POST that starts a read and changes no tenant data (decision D-007):
        source 2 creates an audit log query, and source 3 starts a Management Activity subscription for
        a content type that has none. Pass -NoStartSubscription to withhold the second.

    .PARAMETER Organization
        The tenant's *.onmicrosoft.com domain. Required for app-only sign-in to Exchange Online and
        Security & Compliance PowerShell.

    .EXAMPLE
        ./Run-All.ps1 -OutputPath ./out

    .EXAMPLE
        ./Run-All.ps1 -OutputPath ./out -Environment GCCHigh -AppId $appId `
            -CertificateThumbprint $thumbprint -TenantId $tenantId -Organization contoso.onmicrosoft.us `
            -PublisherIdentifier $publisherId -AccessToken $token
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$OutputPath,

    [ValidateSet('Commercial', 'GCC', 'GCCHigh')]
    [string]$Environment = 'Commercial',

    [string]$AppId,
    [string]$CertificateThumbprint,
    [string]$TenantId,
    [string]$Organization,

    [string]$PublisherIdentifier,
    [securestring]$AccessToken,
    [switch]$NoStartSubscription,

    [ValidateRange(1, 180)]
    [int]$LookbackDays = 7
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Imported by path without -Force: a forced re-import builds a fresh module session
# state, which would discard anything the caller has arranged around the module - and
# that is exactly how the tests put mocks in front of the tenant cmdlets.
Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}
$OutputPath = (Resolve-Path -LiteralPath $OutputPath).Path

$auth = @{
    Environment           = $Environment
    AppId                 = $AppId
    CertificateThumbprint = $CertificateThumbprint
    TenantId              = $TenantId
    Organization          = $Organization
}
$common = @{ OutputPath = $OutputPath; Environment = $Environment; SkipConnect = $true }

Write-CollectorLog -OutputPath $OutputPath -Source 'run-all' -Message (
    'Starting the unified audit log collectors against the {0} cloud.' -f $Environment)

function Invoke-Step {
    param([string]$Name, [string]$Script, [hashtable]$Arguments)
    try {
        & (Join-Path $PSScriptRoot $Script) @Arguments
    }
    catch {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source $Name -Message (
            'The {0} collector stopped ({1}). The remaining collectors still run.' -f $Name, $_.Exception.Message)
    }
}

function Invoke-SignedInGroup {
    param([string]$Service, [scriptblock]$Steps, [string[]]$Scopes)
    $connected = $false
    try {
        $connectArgs = $auth.Clone()
        if ($Scopes) { $connectArgs['Scopes'] = $Scopes }
        Connect-M365Service -Service $Service @connectArgs
        $connected = $true
    }
    catch {
        Write-CollectorLog -OutputPath $OutputPath -Level Error -Source 'run-all' -Message (
            '{0} sign-in failed ({1}). Its collectors will write headers only.' -f $Service, $_.Exception.Message)
    }
    try { & $Steps }
    finally {
        if ($connected -and $Service -ne 'Graph') { Disconnect-ExchangeOnline -Confirm:$false }
    }
}

Invoke-SignedInGroup -Service ExchangeOnline -Steps {
    Invoke-Step 'audit-ingestion' 'Get-AuditIngestion.ps1' $common
    Invoke-Step 'audit-search-cmdlet' 'Get-AuditSearchCmdlet.ps1' ($common + @{ LookbackDays = $LookbackDays })
}

Invoke-SignedInGroup -Service SecurityCompliance -Steps {
    Invoke-Step 'audit-retention-policies' 'Get-AuditRetentionPolicies.ps1' $common
}

# In GCC High the collector skips itself before it would use the session, so the sign-in is not needed.
if ($Environment -ne 'GCCHigh') {
    Invoke-SignedInGroup -Service Graph -Scopes @('AuditLogsQuery.Read.All', 'ThreatIntelligence.Read.All') -Steps {
        Invoke-Step 'audit-graph-records' 'Get-AuditGraphRecords.ps1' ($common + @{ LookbackDays = $LookbackDays })
    }
}
else {
    Invoke-Step 'audit-graph-records' 'Get-AuditGraphRecords.ps1' ($common + @{ LookbackDays = $LookbackDays })
}

if ($AccessToken -and $TenantId -and $PublisherIdentifier) {
    $feed = @{
        OutputPath          = $OutputPath
        Environment         = $Environment
        TenantId            = $TenantId
        PublisherIdentifier = $PublisherIdentifier
        AccessToken         = $AccessToken
        LookbackDays        = [math]::Min($LookbackDays, 7)
    }
    if ($NoStartSubscription) { $feed['NoStartSubscription'] = $true }
    Invoke-Step 'audit-activity-feed' 'Get-AuditActivityFeed.ps1' $feed
}
else {
    Write-CollectorLog -OutputPath $OutputPath -Level Warning -Source 'audit-activity-feed' -Message (
        'Skipped: the Management Activity API needs -AccessToken, -TenantId and -PublisherIdentifier. No file was written.')
}

Write-CollectorLog -OutputPath $OutputPath -Source 'run-all' -Message 'Finished.'
