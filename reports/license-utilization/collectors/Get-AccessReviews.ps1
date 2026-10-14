#Requires -Version 7.0

<#
    .SYNOPSIS
        Writes access-reviews.csv: the access review definitions the tenant has, so the
        License value page can say whether the E5 step-up feature "Access reviews" is in
        use. A snapshot, appended each run.

    .DESCRIPTION
        Source 9: Microsoft Graph list accessReviewScheduleDefinition
        (https://learn.microsoft.com/graph/api/accessreviewset-list-definitions), a GET to
        /v1.0/identityGovernance/accessReviews/definitions?$top=100. The default page is
        100 definitions; @odata.nextLink is followed until it is absent, and a next link
        on another host is refused.

        Needs AccessReview.Read.All. Signed in, the caller needs a role such as Global
        Reader, Security Reader or Identity Governance Administrator for group and app
        reviews, and Privileged Role Administrator or Security Reader for Entra role
        reviews. The feature needs Microsoft Entra ID P2 or Microsoft Entra ID Governance
        (https://learn.microsoft.com/entra/id-governance/access-reviews-overview#license-requirements).
        Available in every cloud the page lists, including US Government L4.

    .EXAMPLE
        ./Get-AccessReviews.ps1 -OutputPath ./out
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

    # An alternate schema file. The tests use it to exercise the NotAvailable path.
    [string]$SchemaPath = (Join-Path $PSScriptRoot 'LicenseUtilizationSchema.psd1'),

    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Imported by path without -Force: a forced re-import builds a fresh module session
# state, which would discard anything the caller has arranged around the module - and
# that is exactly how the tests put mocks in front of the tenant cmdlets.
Import-Module (Join-Path $PSScriptRoot '../../../shared/M365ReportLibrary.psm1')

. (Join-Path $PSScriptRoot 'LicenseUtilizationHelpers.ps1')

$base = (Get-M365ServiceEndpoint -Service Graph -Environment $Environment).ResourceEndpoint
$uri = '{0}/v1.0/identityGovernance/accessReviews/definitions?$top=100' -f $base

Invoke-LicenseUtilizationSnapshot -Source 'AccessReviewDefinitions' -CsvName 'access-reviews.csv' `
    -Description 'Access review definitions' `
    -License 'AccessReview.Read.All and, when signed in, Global Reader, Security Reader or Identity Governance Administrator; Microsoft Entra ID P2 or ID Governance' `
    -KeyColumn @('RunDate', 'DefinitionId') -OutputPath $OutputPath -Environment $Environment -AppId $AppId `
    -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -Organization $Organization `
    -SchemaPath $SchemaPath -SkipConnect:$SkipConnect `
    -Fetch { Get-GraphReportJson -Uri $uri } `
    -Map {
        param($definition, $runDate)

        $id = [string](Get-GraphAdditionalProperty -Object $definition -Name 'id')
        if ([string]::IsNullOrWhiteSpace($id)) { return }

        $scope = Get-GraphAdditionalProperty -Object $definition -Name 'scope'
        $settings = Get-GraphAdditionalProperty -Object $definition -Name 'settings'
        $recurrence = Get-GraphAdditionalProperty -Object $settings -Name 'recurrence'
        $pattern = Get-GraphAdditionalProperty -Object $recurrence -Name 'pattern'

        [pscustomobject]@{
            RunDate                   = $runDate
            DefinitionId              = $id
            DisplayName               = [string](Get-GraphAdditionalProperty -Object $definition -Name 'displayName')
            Status                    = [string](Get-GraphAdditionalProperty -Object $definition -Name 'status')
            ScopeQuery                = [string](Get-GraphAdditionalProperty -Object $scope -Name 'query')
            ReviewerCount             = @(Get-GraphAdditionalProperty -Object $definition -Name 'reviewers').Count
            RecurrencePattern         = [string](Get-GraphAdditionalProperty -Object $pattern -Name 'type')
            AutoApplyDecisionsEnabled = [string](Get-GraphAdditionalProperty -Object $settings -Name 'autoApplyDecisionsEnabled')
            RecommendationsEnabled    = [string](Get-GraphAdditionalProperty -Object $settings -Name 'recommendationsEnabled')
        }
    }
