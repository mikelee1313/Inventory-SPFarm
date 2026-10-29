#Requires -Version 7.4
<#
.SYNOPSIS
    Finds SharePoint sites whose linked Microsoft 365 groups no longer exist.

.DESCRIPTION
    Uses PnP.PowerShell to enumerate sites with their GroupId and RelatedGroupId,
    then Microsoft Graph GET /groups/{id} to check each distinct nonzero GUID.
    No Exchange Online or Microsoft Graph PowerShell module is needed.

    Graph's site resource does not expose GroupId or RelatedGroupId, so a
    Graph-only site inventory cannot reliably discover these orphan candidates.
    Certificate authentication is handled by PnP; Get-PnPAccessToken obtains and
    refreshes the Graph token using the same application and certificate.

    A site is an OrphanedGroupSite only when ALL its distinct linked group IDs
    return HTTP 404. A valid ID in EITHER property makes a site a candidate.
    Both zero/empty properties mean the site is not group connected.
    An existing group plus a missing group is InconsistentGroupLinks, not an
    orphan. Invalid GUIDs, non-Microsoft 365 groups, and failed lookups are
    reported for review and never treated as proof of deletion.

    Outputs timestamped OrphanedGroupSites.csv, GroupSiteAudit.csv, and a log.
    The audit includes every candidate and sites with invalid/missing properties.
    Lookup failures produce reports but also a terminating error so an incomplete
    scan cannot be mistaken for a successful scan with no orphans.

    This is read-only. A 404 means absent from active groups, not proof of
    permanent deletion: soft-deleted groups also qualify. Review candidates
    before remediation, allowing for directory replication delays.

.PARAMETER TenantName
    SharePoint tenant prefix, without .sharepoint.com. Can be set in the configuration below.
.PARAMETER AppId
    Entra application (client) ID. Can be set in the configuration below.
.PARAMETER Thumbprint
    Thumbprint of the certificate with private key installed for PnP authentication.
    Can be set in the configuration below.
.PARAMETER Tenant
    Entra tenant ID or verified domain. Can be set in the configuration below.
.PARAMETER OutputDirectory
    Existing or new report directory. Defaults to the current user's TEMP folder.
.PARAMETER AdminUrl
    Optional SharePoint admin URL. Defaults to https://TenantName-admin.sharepoint.com.
    For Multi-Geo, run against each geography's admin URL with separate reports.

.NOTES
    Requires PnP.PowerShell and a certificate configured on the Entra application.
    Application permissions with admin consent:
      SharePoint: Sites.FullControl.All (tenant administration/site enumeration).
      Microsoft Graph: Group.Read.All (group existence and groupTypes).
    No user, membership, Exchange, or Graph Sites permissions are required.
    OneDrive and redirect sites are excluded. Tenant discovery must succeed
    before reports are created; failures are logged and terminate the scan.

.EXAMPLE
    # Fill in the four connection defaults in the configuration below, then run:
    .\Find-OphanedSiteGroups.ps1

.EXAMPLE
    .\Find-OphanedSiteGroups.ps1 -TenantName contoso -AppId '11111111-1111-1111-1111-111111111111' -Thumbprint 'YOUR_CERTIFICATE_THUMBPRINT' -Tenant 'contoso.onmicrosoft.com'
#>
#region Configuration
[CmdletBinding()]
param(
    # Admin configuration: replace these defaults to run without command-line parameters.
    [ValidateNotNullOrEmpty()][string] $TenantName = 'M365CPI13246019',
    [guid] $AppId = 'abc64618-283f-47ba-a185-50d935d51d57',
    [ValidateNotNullOrEmpty()][string] $Thumbprint = 'B696FDCFE1453F3FBC6031F54DE988DA0ED905A9',
    [ValidateNotNullOrEmpty()][string] $Tenant = '9cfc42cb-51da-4055-87e9-b20a170b6ba3',
    [ValidateNotNullOrEmpty()][string] $OutputDirectory = [System.IO.Path]::GetTempPath(),
    [uri] $AdminUrl
)
#endregion Configuration

#region Configuration validation
$ErrorActionPreference = 'Stop'

$missingSettings = @(
    foreach ($name in 'TenantName', 'Thumbprint', 'Tenant') {
        if ([string]::IsNullOrWhiteSpace((Get-Variable -Name $name -ValueOnly))) { $name }
    }
    if ($AppId -eq [guid]::Empty) { 'AppId' }
)
if ($missingSettings.Count -gt 0) {
    throw "Configure $($missingSettings -join ', ') in the script's parameter defaults, or supply them as command-line parameters."
}
#endregion Configuration validation

#region Logging
function Write-LogEntry {
    param(
        [string] $Message,
        [ValidateSet('INFO', 'WARNING', 'ERROR')][string] $Level = 'INFO'
    )
    Add-Content -LiteralPath $logPath -Value "$(Get-Date -Format o) [$Level] $Message"
}
#endregion Logging

#region Graph requests and retry handling
function Invoke-GraphGet {
    param([Parameter(Mandatory)][string] $Path)

    $maxRetries = 5
    for ($attempt = 0; $attempt -le $maxRetries; $attempt++) {
        # PnP uses its authentication cache and renews expiring Graph tokens.
        $token = Get-PnPAccessToken -ResourceTypeName Graph -Connection $connection -ErrorAction Stop
        if ($token -is [System.Security.SecureString]) {
            $token = ConvertFrom-SecureString -SecureString $token -AsPlainText
        }
        if ([string]::IsNullOrWhiteSpace($token)) {
            throw 'PnP did not return a Graph access token.'
        }
        $statusCode = 0
        $responseHeaders = $null
        $body = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/v1.0/$Path" `
            -Headers @{ Authorization = "Bearer $token" } -SkipHttpErrorCheck `
            -StatusCodeVariable statusCode -ResponseHeadersVariable responseHeaders -ErrorAction Stop

        if ([int]$statusCode -in 408, 429, 500, 502, 503, 504 -and $attempt -lt $maxRetries) {
            $delay = [int][Math]::Pow(2, $attempt + 1)
            $retryAfter = @($responseHeaders['Retry-After'])[0]
            $seconds = 0
            $retryDate = [DateTimeOffset]::MinValue
            if ([int]::TryParse([string]$retryAfter, [ref]$seconds) -and $seconds -ge 0) {
                $delay = $seconds
            }
            elseif ([DateTimeOffset]::TryParse([string]$retryAfter, [ref]$retryDate)) {
                $delay = [int][Math]::Max(0, [Math]::Ceiling(($retryDate - [DateTimeOffset]::UtcNow).TotalSeconds))
            }
            Write-LogEntry "Graph $Path returned HTTP $statusCode. Retry $($attempt + 1)/$maxRetries in $delay seconds." 'WARNING'
            Start-Sleep -Seconds $delay
            continue
        }
        return [pscustomobject]@{ StatusCode = [int]$statusCode; Body = $body }
    }
}
#endregion Graph requests and retry handling

#region Cached group lookups
function Get-GroupLookup {
    param([Parameter(Mandatory)][guid] $GroupId)

    $key = $GroupId.ToString()
    if ($groupCache.ContainsKey($key)) {
        return $groupCache[$key]
    }

    try {
        $response = Invoke-GraphGet "groups/$key`?`$select=id,displayName,groupTypes"
        $state = 'LookupFailed'
        $detail = ''
        if ($response.StatusCode -eq 404) {
            $state = 'Missing'
            $detail = 'HTTP 404: group not found in active directory.'
        }
        elseif ($response.StatusCode -eq 200) {
            if ([string]$response.Body.id -ne $key) {
                $detail = 'Graph returned an unexpected or missing group ID.'
            }
            elseif (@($response.Body.groupTypes) -contains 'Unified') {
                $state = 'Exists'
            }
            else {
                $state = 'NotMicrosoft365Group'
                $detail = 'The linked object exists but is not a Microsoft 365 (Unified) group.'
            }
        }
        else {
            $detail = "HTTP $($response.StatusCode): $($response.Body.error.code) $($response.Body.error.message)"
        }
        $result = [pscustomobject]@{
            State = $state
            DisplayName = if ($state -eq 'Exists') { $response.Body.displayName } else { '' }
            Detail = $detail
        }
    }
    catch {
        $result = [pscustomobject]@{ State = 'LookupFailed'; DisplayName = ''; Detail = $_.Exception.Message }
    }
    if ($result.State -ne 'Exists') {
        Write-LogEntry "Group $key : $($result.State). $($result.Detail)" 'WARNING'
    }
    $groupCache[$key] = $result
    return $result
}
#endregion Cached group lookups

#region Site group assessment
function Get-SiteGroupAssessment {
    param([Parameter(Mandatory)] $Site)

    $links = @{}
    $details = [System.Collections.Generic.List[string]]::new()
    foreach ($property in 'GroupId', 'RelatedGroupId') {
        if ($null -eq $Site.PSObject.Properties[$property]) {
            $links[$property] = [pscustomobject]@{ State = 'PropertyUnavailable'; DisplayName = '' }
            $details.Add("$property was not returned by site discovery.")
            continue
        }
        $value = [string]$Site.$property
        $id = [guid]::Empty
        if ([string]::IsNullOrWhiteSpace($value)) {
            $links[$property] = [pscustomobject]@{ State = 'NotLinked'; DisplayName = '' }
        }
        elseif (-not [guid]::TryParse($value, [ref]$id)) {
            $links[$property] = [pscustomobject]@{ State = 'InvalidGroupId'; DisplayName = '' }
            $details.Add("$property is not a valid GUID: $value")
        }
        elseif ($id -eq [guid]::Empty) {
            $links[$property] = [pscustomobject]@{ State = 'NotLinked'; DisplayName = '' }
        }
        else {
            $links[$property] = Get-GroupLookup -GroupId $id
            if ($links[$property].Detail) {
                $details.Add("$property : $($links[$property].Detail)")
            }
        }
    }

    $states = @($links.GroupId.State, $links.RelatedGroupId.State)
    $status = if ($states -contains 'PropertyUnavailable') { 'PropertyUnavailable' }
        elseif ($states -contains 'InvalidGroupId') { 'InvalidGroupId' }
        elseif ($states -contains 'LookupFailed') { 'LookupFailed' }
        elseif ($states -contains 'NotMicrosoft365Group') { 'NotMicrosoft365Group' }
        elseif ($states -contains 'Missing' -and $states -contains 'Exists') { 'InconsistentGroupLinks' }
        elseif ($states -contains 'Missing') { 'OrphanedGroupSite' }
        elseif ($states -contains 'Exists') { 'GroupFound' }
        else { 'NotGroupConnected' }

    return [pscustomobject][ordered]@{
        URL = $Site.Url
        Title = $Site.Title
        Template = $Site.Template
        'Group ID' = $Site.GroupId
        RelatedGroupId = $Site.RelatedGroupId
        Status = $status
        GroupIdStatus = $links.GroupId.State
        GroupDisplayName = $links.GroupId.DisplayName
        RelatedGroupIdStatus = $links.RelatedGroupId.State
        RelatedGroupDisplayName = $links.RelatedGroupId.DisplayName
        Details = $details -join ' | '
        CheckedAtUtc = [DateTimeOffset]::UtcNow.ToString('o')
    }
}
#endregion Site group assessment

#region Report initialization
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$stamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss_fff'
$prefix = Join-Path $OutputDirectory "GroupSites_${TenantName}_$stamp"
$orphanPath = "${prefix}_OrphanedGroupSites.csv"
$auditPath = "${prefix}_GroupSiteAudit.csv"
$logPath = "${prefix}.log"
$connection = $null
$groupCache = @{}
#endregion Report initialization

#region Scan execution and cleanup
try {
    #region Connection and permission preflight
    Import-Module PnP.PowerShell -ErrorAction Stop
    if (-not $AdminUrl) {
        $AdminUrl = [uri]"https://$TenantName-admin.sharepoint.com"
    }
    Write-LogEntry "Connecting to $AdminUrl."
    $connection = Connect-PnPOnline -Url $AdminUrl.AbsoluteUri -ClientId $AppId.ToString() `
        -Thumbprint $Thumbprint -Tenant $Tenant -ReturnConnection -ErrorAction Stop

    # Verify tenant-wide group read access before accepting any 404 as evidence.
    $probe = Invoke-GraphGet 'groups?$top=1&$select=id'
    if ($probe.StatusCode -ne 200 -or $null -eq $probe.Body.value) {
        throw "Graph group-read preflight failed (HTTP $($probe.StatusCode)). Verify application Group.Read.All and admin consent. $($probe.Body.error.message)"
    }
    #endregion Connection and permission preflight

    #region Site discovery
    Write-Host 'Retrieving tenant site properties...' -ForegroundColor Cyan
    $sites = @(Get-PnPTenantSite -Detailed -Connection $connection -ErrorAction Stop |
        Where-Object { $_.Url -notlike '*-my.sharepoint.com*' -and $_.Template -ne 'REDIRECTSITE#0' })
    Write-LogEntry "Retrieved $($sites.Count) sites."
    #endregion Site discovery

    #region CSV headers
    # Always write headers, even when no candidates are found.
    $schema = Get-SiteGroupAssessment -Site ([pscustomobject]@{
        Url = ''; Title = ''; Template = ''; GroupId = [guid]::Empty; RelatedGroupId = [guid]::Empty
    })
    $header = ($schema | ConvertTo-Csv -NoTypeInformation)[0]
    Set-Content -LiteralPath $orphanPath -Value $header -Encoding utf8BOM
    Set-Content -LiteralPath $auditPath -Value $header -Encoding utf8BOM
    #endregion CSV headers

    #region Site processing and report export
    $checked = 0
    $orphans = 0
    $review = 0
    $failures = 0
    foreach ($site in $sites) {
        $row = Get-SiteGroupAssessment -Site $site
        if ($row.Status -eq 'NotGroupConnected') { continue }
        $checked++
        $row | Export-Csv -LiteralPath $auditPath -Append -NoTypeInformation -Encoding utf8BOM
        if ($row.Status -eq 'OrphanedGroupSite') {
            $orphans++
            $row | Export-Csv -LiteralPath $orphanPath -Append -NoTypeInformation -Encoding utf8BOM
        }
        elseif ($row.Status -ne 'GroupFound') {
            $review++
            Write-LogEntry "$($row.URL) : $($row.Status). $($row.Details)" 'WARNING'
            if ($row.Status -in 'LookupFailed', 'InvalidGroupId', 'PropertyUnavailable') { $failures++ }
        }
    }
    #endregion Site processing and report export

    #region Scan summary
    Write-LogEntry "Scan finished: $checked candidates/review sites; $orphans orphan candidates; $review sites require review; $failures incomplete assessments."
    Write-Host "Checked $checked candidate/review sites. Orphan candidates: $orphans. Review required: $review."
    Write-Host "Orphan candidates: $orphanPath"
    Write-Host "Full audit: $auditPath"
    Write-Host "Log: $logPath"
    if ($failures -gt 0) {
        throw "$failures site assessment(s) could not be completed. Review the audit and log; the orphan report is incomplete."
    }
    #endregion Scan summary
}
catch {
    Write-LogEntry $_.Exception.Message 'ERROR'
    throw
}
finally {
    # -ReturnConnection does not replace the caller's current PnP connection.
    if ($null -ne $connection -and $null -ne $connection.Context) {
        $connection.Context.Dispose()
    }
}
#endregion Scan execution and cleanup
