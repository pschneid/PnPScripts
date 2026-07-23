<#
.SYNOPSIS
    Predicts the crawled properties SharePoint Online Search is likely to emit for a given file.

.DESCRIPTION
    SharePoint Online does not expose any API (CSOM, REST, Graph, or PnP PowerShell) that reports which
    crawled properties were actually produced for a specific item - that information only becomes visible,
    after a crawl, in the tenant Search Schema admin center. This script works around that gap by:

      1. Reading the file's list item and its populated field values.
      2. Reading the field definitions (type, and whether the field is an out-of-the-box/base-list-template
         column vs. a custom site column).
      3. Applying the documented SharePoint crawled-property naming conventions to PREDICT the crawled
         property name(s) each populated field is expected to produce:
           - ows_<InternalName>                     -> raw property, emitted for every populated column
           - ows_taxId_<InternalName>                -> Managed Metadata (Taxonomy) columns
           - ows_q_<TYPECODE>_<InternalName>         -> custom site columns (non rich-text)
           - ows_r_<TYPECODE>_<InternalName>         -> custom site columns of type "Multiple lines of text"
         Each prediction is also tagged with its search schema property set (name + GUID), using the same
         well-known property set GUIDs this module's own Add-PnPTenantSearchCrawledProperty cmdlet uses
         (SharePointDefault/SharePointTaxonomy/SharePointStructured/SharePointRich).
      4. Cross-referencing each predicted name against the Managed Property mappings that already exist at
         all three levels search configuration can be set at - Web, Site (site collection), and Tenant
         (subscription) - to show whether, and at which level(s), it is already mapped to a managed property.
         The tenant/subscription level requires the SharePoint Online Admin Center context; this script
         mirrors what PnP.PowerShell's PnPSharePointOnlineAdminCmdlet base class does internally, by
         elevating the current connection to the tenant "-admin" site (via Connect-PnPOnline -ReturnConnection)
         after the Web/Site level checks have been performed, and skipping it gracefully if elevation fails
         (e.g. no SharePoint Administrator role, or a connection type that cannot be silently re-used, such as
         device code login).

    IMPORTANT: These are PREDICTIONS based on well-documented but unofficial SharePoint search schema
    conventions, not a guarantee of what the search index actually contains. Always confirm the final
    answer for a given field by searching for its internal name in:
    https://<tenant>-admin.sharepoint.com/_layouts/15/searchadmin/crawledproperties.aspx

.PARAMETER FileUrl
    Server-relative, site-relative, or absolute URL of the file to analyze,
    e.g. "/sites/Team/Shared Documents/Report.docx".

.PARAMETER IncludeSystemFields
    Include hidden/system fields (e.g. _dlc_*, owshiddenversion, ContentTypeId) in the prediction.
    Off by default since these rarely matter for search scenarios.

.PARAMETER SkipManagedPropertyLookup
    Skip the Web/Site/Tenant managed property cross-reference entirely and only show predicted names.

.PARAMETER SkipTenantElevation
    Perform the Web and Site level managed property lookups, but do not attempt to elevate the connection to
    the tenant admin center for the Subscription level lookup, and suppress the reminder about it. Equivalent
    to simply omitting -AdminConnectionParameters, provided just to make the intent explicit/silence the note.

.PARAMETER TenantAdminUrl
    Explicit SharePoint Online Admin Center URL to elevate to for the tenant/subscription level lookup, e.g.
    "https://contoso-admin.sharepoint.com". When omitted, it is derived from the current connection's URL
    using the standard "<tenant>-admin.sharepoint.<tld>" convention.

.PARAMETER AdminConnectionParameters
    Opt-in hashtable of additional parameters to splat into the Connect-PnPOnline call used to elevate to the
    tenant admin center for the Tenant-level lookup. When this parameter is not supplied at all, the
    Tenant-level lookup is skipped with a warning, since re-authenticating to the "-admin" host may need a
    secret (certificate, client secret, ...) that PnP.PowerShell cannot silently re-derive from the current
    connection. Pass an empty hashtable, @{}, to opt in for connections that need no extra secret (Interactive,
    Managed Identity, ...), or supply the same secret you originally connected with, e.g.:
      @{ CertificatePath = '.\AppCertNoPass.pfx' }
      @{ CertificatePath = '.\AppCert.pfx'; CertificatePassword = (ConvertTo-SecureString 'pwd' -AsPlainText -Force) }
      @{ Thumbprint = '<thumbprint-in-local-cert-store>' }
      @{ ClientSecret = '...'; Realm = '...' }                   # legacy ACS app-only
    ClientId/Tenant/AzureEnvironment are always auto-reused from the current connection and do not need to be
    repeated here unless you want to override them. Not needed for interactive/managed identity connections.

.EXAMPLE
    Connect-PnPOnline -Url https://contoso.sharepoint.com/sites/Team -Interactive
    .\Get-PredictedCrawledProperties.ps1 -FileUrl "/sites/Team/Shared Documents/Report.docx"

.EXAMPLE
    .\Get-PredictedCrawledProperties.ps1 -FileUrl "/sites/Team/Shared Documents/Report.docx" -IncludeSystemFields | Format-Table -AutoSize

.EXAMPLE
    .\Get-PredictedCrawledProperties.ps1 -FileUrl "/sites/Team/Shared Documents/Report.docx" -SkipTenantElevation

.NOTES
    Requires an existing PnP PowerShell connection (Connect-PnPOnline) with at least read access to the file's
    site. The Web/Site level managed-property cross-reference needs read access on the site; the Tenant/
    Subscription level needs the SharePoint Administrator role (Sites.FullControl.All / AllSites.FullControl).
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$FileUrl,

    [switch]$IncludeSystemFields,

    [switch]$SkipManagedPropertyLookup,

    [switch]$SkipTenantElevation,

    [string]$TenantAdminUrl,

    [hashtable]$AdminConnectionParameters
)

# Documented 4-letter SharePoint search schema type codes used in the ows_q_/ows_r_ crawled property names.
# Source: SharePoint search schema conventions (undocumented by Microsoft, but widely observed/stable).
$typeCodeMap = @{
    'Text'        = 'TEXT'
    'Note'        = 'MTXT'   # Multiple lines of text -> rich text extraction (ows_r_)
    'Choice'      = 'CHCS'
    'MultiChoice' = 'CHCM'
    'DateTime'    = 'DATE'
    'Number'      = 'NMBR'
    'Integer'     = 'INTG'
    'Currency'    = 'CURR'
    'Boolean'     = 'BOOL'
    'User'        = 'USER'
    'UserMulti'   = 'USER'
    'URL'         = 'URLH'
    'Guid'        = 'GUID'
}

# Property set GUIDs used by SharePoint's search schema for each crawled-property naming convention. These
# are the same GUIDs this module's own Add-PnPTenantSearchCrawledProperty cmdlet uses when creating a tenant
# crawled property in one of these well-known property sets (see AddTenantSearchCrawledProperty.cs).
$propertySets = @{
    'SharePointDefault'    = [pscustomobject]@{ Name = 'SharePointDefault'; Guid = '00130329-0000-0130-C000-000000131346' }
    'SharePointTaxonomy'   = [pscustomobject]@{ Name = 'SharePointTaxonomy'; Guid = '158D7563-AEFF-4DBF-BF16-4A1445F0366C' }
    'SharePointStructured' = [pscustomobject]@{ Name = 'SharePointStructured'; Guid = 'ED280121-B677-4E2A-8FBC-0D9E2325B0A2' }
    'SharePointRich'       = [pscustomobject]@{ Name = 'SharePointRich'; Guid = 'FEA84DF6-A0FC-492C-9AA7-D28B8DCB08B3' }
}

function Get-PropertySetForCrawledProperty {
    # Mirrors GetExpectedPropertySet() in AddTenantSearchCrawledProperty.cs, matching on the crawled property
    # name's prefix rather than the field's internal name so it works whether the raw ows_ or the typed
    # ows_q_/ows_r_/ows_taxId_ variant is passed in.
    param([string]$CrawledPropertyName)

    if ($CrawledPropertyName -like 'ows_taxId_*') { return $propertySets.SharePointTaxonomy }
    if ($CrawledPropertyName -like 'ows_q_*') { return $propertySets.SharePointStructured }
    if ($CrawledPropertyName -like 'ows_r_*') { return $propertySets.SharePointRich }
    if ($CrawledPropertyName -like 'ows_*') { return $propertySets.SharePointDefault }
    return $null
}

# Internal field names (or prefixes) that are SharePoint system/plumbing fields and rarely relevant to search.
$systemFieldPrefixes = @(
    '_', 'owshiddenversion', 'ContentTypeId', 'FileRef', 'FileDirRef', 'FSObjType', 'SortBehavior',
    'PermMask', 'UniqueId', 'SyncClientId', 'ProgId', 'ScopeId', 'MetaInfo', 'ItemChildCount',
    'FolderChildCount', 'Restricted', 'OriginatorId', 'NoExecute', 'ContentVersion'
)

function Test-SystemField {
    param([string]$InternalName)
    foreach ($prefix in $systemFieldPrefixes) {
        if ($InternalName -like "$prefix*") { return $true }
    }
    return $false
}

function Get-TenantAdminUrl {
    # Mirrors the derivation logic used internally by PnP.PowerShell's PnPSharePointOnlineAdminCmdlet /
    # Connect-PnPOnline: <tenant>[-my].sharepoint.<tld> -> <tenant>-admin.sharepoint.<tld>
    param([string]$SiteUrl)

    $uri = [Uri]$SiteUrl
    $hostParts = $uri.Host.Split('.')
    $tenantPart = $hostParts[0]

    if ($tenantPart.EndsWith('-admin')) {
        return "$($uri.Scheme)://$($uri.Host)"
    }
    if ($tenantPart.EndsWith('-my')) {
        $tenantPart = $tenantPart.Substring(0, $tenantPart.Length - 3)
    }

    $remainder = ($hostParts | Select-Object -Skip 1) -join '.'
    return "$($uri.Scheme)://$tenantPart-admin.$remainder"
}

function Get-ManagedPropertyMappings {
    # Wraps Get-PnPSearchConfiguration for a given scope/connection, returning $null on failure instead of throwing,
    # so that a missing permission at one level doesn't stop the Web/Site/Tenant checks at the other levels.
    param(
        [string]$Scope,
        [object]$Connection
    )
    try {
        $params = @{ Scope = $Scope; OutputFormat = 'ManagedPropertyMappings'; ErrorAction = 'Stop' }
        if ($Connection) { $params.Connection = $Connection }
        return Get-PnPSearchConfiguration @params
    }
    catch {
        Write-Warning "Could not read managed property mappings at '$Scope' scope: $($_.Exception.Message)"
        return $null
    }
}

Write-Host "Resolving file '$FileUrl' ..." -ForegroundColor Cyan
$fileListItem = Get-PnPFile -Url $FileUrl -AsListItem -ThrowExceptionIfFileNotFound -ErrorAction Stop
$list = Get-PnPProperty -ClientObject $fileListItem -Property ParentList
$itemId = Get-PnPProperty -ClientObject $fileListItem -Property Id

# FieldValues is a client-side aggregate of whatever fields have actually been loaded on the ListItem - it
# cannot itself be requested via Get-PnPProperty/Load(item => item.FieldValues). Re-fetch the item through
# Get-PnPListItem instead, which loads it the supported way (a plain, non-expression Load() call that
# populates FieldValues for every field) and can also bring back the ContentType in the same round-trip.
Write-Host "Reading field values for list item $itemId in list '$($list.Title)' ..." -ForegroundColor Cyan
$item = Get-PnPListItem -List $list.Id -Id $itemId -IncludeContentType -ErrorAction Stop

Write-Host "Reading field definitions for list '$($list.Title)' ..." -ForegroundColor Cyan
$fields = Get-PnPField -List $list.Id

$mappingsByScope = @{ Web = $null; Site = $null; Tenant = $null }

if (-not $SkipManagedPropertyLookup) {

    Write-Host "Reading Web-level managed property mappings ..." -ForegroundColor Cyan
    $mappingsByScope.Web = Get-ManagedPropertyMappings -Scope Web

    Write-Host "Reading Site-level managed property mappings ..." -ForegroundColor Cyan
    $mappingsByScope.Site = Get-ManagedPropertyMappings -Scope Site

    if (-not $SkipTenantElevation) {
        if (-not $PSBoundParameters.ContainsKey('AdminConnectionParameters')) {
            # Re-authenticating to the "-admin" host may need a secret (certificate, client secret, etc.) that
            # PnP.PowerShell does not - and for security reasons should not - expose back out of an existing
            # Connection object, so we don't have enough context to safely attempt this automatically. Requiring
            # an explicit opt-in avoids unexpectedly popping an interactive/browser login in an unattended script.
            Write-Warning "Skipping the Tenant-level lookup: pass -AdminConnectionParameters to enable it (use an empty hashtable, @{}, if your connection is Interactive/Managed Identity and needs no extra secret; otherwise e.g. @{ CertificatePath = '...pfx' } or @{ ClientSecret = '...'; Realm = '...' }), or -SkipTenantElevation to suppress this message."
        }
        else {
            $currentConnection = Get-PnPConnection
            if (-not $TenantAdminUrl) { $TenantAdminUrl = Get-TenantAdminUrl -SiteUrl $currentConnection.Url }

            # Reuse the same app registration/tenant/cloud as the current connection so the elevated
            # connection to the "-admin" site authenticates the same way. Anything explicitly supplied via
            # -AdminConnectionParameters (e.g. CertificatePath/Thumbprint/ClientSecret) takes precedence.
            $elevateParams = @{}
            if ($currentConnection.ClientId) { $elevateParams.ClientId = $currentConnection.ClientId }
            if ($currentConnection.Tenant) { $elevateParams.Tenant = $currentConnection.Tenant }
            if ($currentConnection.AzureEnvironment) { $elevateParams.AzureEnvironment = $currentConnection.AzureEnvironment }
            foreach ($key in $AdminConnectionParameters.Keys) { $elevateParams[$key] = $AdminConnectionParameters[$key] }

            Write-Host "Elevating to the tenant admin context at '$TenantAdminUrl' for the Tenant-level lookup ..." -ForegroundColor Cyan
            try {
                $adminConnection = Connect-PnPOnline -Url $TenantAdminUrl -ReturnConnection @elevateParams -ErrorAction Stop
                $mappingsByScope.Tenant = Get-ManagedPropertyMappings -Scope Subscription -Connection $adminConnection
            }
            catch {
                Write-Warning "Could not elevate to the tenant admin context, skipping the Tenant-level lookup. This requires the SharePoint Administrator role and enough context in -AdminConnectionParameters to re-authenticate. Error: $($_.Exception.Message)"
            }
        }
    }
    else {
        Write-Host "Skipping Tenant-level lookup (-SkipTenantElevation specified) ..." -ForegroundColor Yellow
    }
}

$results = [System.Collections.Generic.List[object]]::new()

foreach ($fieldValue in $item.FieldValues.GetEnumerator()) {

    $internalName = $fieldValue.Key
    $value = $fieldValue.Value

    if ($null -eq $value -or ($value -is [string] -and $value -eq '')) { continue }
    if (-not $IncludeSystemFields -and (Test-SystemField -InternalName $internalName)) { continue }

    $field = $fields | Where-Object { $_.InternalName -eq $internalName } | Select-Object -First 1

    # ows_<InternalName> is emitted for every populated list or site column - this is the "raw" extraction.
    $predictedNames = [System.Collections.Generic.List[string]]::new()
    $predictedNames.Add("ows_$internalName")

    if ($field) {
        $typeAsString = $field.TypeAsString

        if ($typeAsString -like 'TaxonomyFieldType*') {
            $predictedNames.Add("ows_taxId_$internalName")
        }
        elseif (-not $field.FromBaseType) {
            # Custom (non out-of-the-box) site column also gets a typed ows_q_/ows_r_ extraction.
            $typeCode = $typeCodeMap[$typeAsString]
            if (-not $typeCode -and $field.FieldTypeKind) { $typeCode = $typeCodeMap[$field.FieldTypeKind.ToString()] }
            if ($typeCode) {
                if ($typeAsString -eq 'Note') {
                    $predictedNames.Add("ows_r_${typeCode}_$internalName")
                }
                else {
                    $predictedNames.Add("ows_q_${typeCode}_$internalName")
                }
            }
        }
    }

    foreach ($crawledPropertyName in $predictedNames) {

        # A single crawled property can be mapped to more than one managed property at the same scope -
        # collect all matches instead of just the first one.
        $webMappings = @($mappingsByScope.Web | Where-Object { $_.Mappings -contains $crawledPropertyName })
        $siteMappings = @($mappingsByScope.Site | Where-Object { $_.Mappings -contains $crawledPropertyName })
        $tenantMappings = @($mappingsByScope.Tenant | Where-Object { $_.Mappings -contains $crawledPropertyName })

        $mappedAtScopes = @()
        if ($webMappings.Count -gt 0) { $mappedAtScopes += 'Web' }
        if ($siteMappings.Count -gt 0) { $mappedAtScopes += 'Site' }
        if ($tenantMappings.Count -gt 0) { $mappedAtScopes += 'Tenant' }

        $propertySet = Get-PropertySetForCrawledProperty -CrawledPropertyName $crawledPropertyName

        $results.Add([pscustomobject]@{
            FieldInternalName        = $internalName
            FieldType                = if ($field) { $field.TypeAsString } else { 'Unknown (not a list column)' }
            IsCustomSiteColumn       = if ($field) { -not [bool]$field.FromBaseType } else { 'Unknown' }
            Value                    = $value
            PredictedCrawledProperty = $crawledPropertyName
            PropertySet              = if ($propertySet) { $propertySet.Name } else { $null }
            PropertySetGuid          = if ($propertySet) { $propertySet.Guid } else { $null }
            MappedAtScopes           = if ($mappedAtScopes) { $mappedAtScopes -join ', ' } else { $null }
            WebManagedProperties     = if ($webMappings.Count -gt 0) { ($webMappings.Name | Sort-Object -Unique) -join ', ' } else { $null }
            SiteManagedProperties    = if ($siteMappings.Count -gt 0) { ($siteMappings.Name | Sort-Object -Unique) -join ', ' } else { $null }
            TenantManagedProperties  = if ($tenantMappings.Count -gt 0) { ($tenantMappings.Name | Sort-Object -Unique) -join ', ' } else { $null }
        })
    }
}

$results | Sort-Object FieldInternalName, PredictedCrawledProperty
