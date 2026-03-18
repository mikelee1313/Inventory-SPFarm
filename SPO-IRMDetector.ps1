<#
.SYNOPSIS
    Inventories SharePoint Online Information Management (IM) policies and In-Place Records
    Management features across site collections as part of MC1211579 retirement preparedness.

.DESCRIPTION
    SPO-IRMDetector.ps1 helps you inventory your SPO tenant to identify which site collections
    have the Information Management and In-Place Records Management features enabled, and
    surfaces any Information Management Policies configured on document library content types.

    Specifically, it scans for the following retiring features (MC1211579):
      - In Place Records Management  (Site scope)
      - Hold                         (Web scope)
      - Location Based Policy        (Site scope)

    For every site where a matching feature is found, the script additionally inspects all
    non-hidden document libraries and their content types for configured IM policies,
    extracting:
      - Library name and content type name
      - Policy name and description
      - Whether a Retention schedule is defined

    All findings are written to a single CSV file in real time. Sites with active features
    but no IM policies produce one row with blank policy columns so they are still visible
    in the output.

    The script supports two scan modes:
      - Full tenant scan  (default, $SiteListPath left empty)
      - Targeted scan     (set $SiteListPath to a file with one site URL per line)

    Use targeted mode when you already have a list of affected sites from a previous full
    scan and only want to re-inspect or add IM policy detail without re-scanning the whole
    tenant.

.PARAMETER tenant
    The SharePoint tenant name (without .sharepoint.com suffix).
    Example: "m365cpi13246019"

.PARAMETER appID
    The Azure AD App Registration Client ID used for certificate-based authentication.

.PARAMETER thumbprint
    The certificate thumbprint registered in the Azure AD App Registration.

.PARAMETER tenantid
    The Azure AD Tenant ID (GUID).

.PARAMETER SiteListPath
    Optional. Path to a plain-text file containing site URLs to scan (one URL per line).
    Leave empty ("") to scan ALL sites in the tenant via Get-PnPTenantSite.
    Example: "./sites.txt"

.PARAMETER featureIds
    Array of Feature Definition GUIDs to search for.
    Defaults cover the three retiring MC1211579 features. Extend as needed.
    Obtain IDs with:
        Get-PnPFeature -Scope Site | Select-Object DisplayName, DefinitionId
        Get-PnPFeature -Scope Web  | Select-Object DisplayName, DefinitionId

.PARAMETER outputFolder
    Destination folder for the results CSV. Defaults to the current directory ("./").

.PARAMETER outputPrefix
    Prefix for the output filename. A timestamp (yyyyMMdd_HHmmss) is appended automatically.
    Default: "FeatureInventory"

.PARAMETER maxRetries
    Maximum retry attempts for throttled requests (HTTP 429 / 503). Default: 5.

.PARAMETER baseDelaySeconds
    Base delay in seconds for exponential backoff (delay = 2^retry * base). Default: 2.

.PARAMETER delayBetweenSites
    Milliseconds to pause between site scans to reduce request spikes. Default: 500.

.OUTPUTS
    Single CSV file: FeatureInventory_yyyyMMdd_HHmmss.csv
    Columns: SiteUrl, FoundFeatureIds, LibraryTitle, ContentTypeName,
             PolicyName, PolicyDescription, RetentionDefined

.NOTES
    - Author:       Mike Lee
    - Last updated: 3/18/26
    - Related MC:   MC1211579 (SharePoint Online Information Management /
                    In-Place Records Management feature retirement)
    - Requires:     PnP PowerShell module
    - Auth:         Certificate-based (Azure AD App Registration)
    - Throttling:   Exponential backoff, honours Retry-After headers
    - Scopes:       Checks both Site-scope and Web-scope features per site collection
    - IM Policies:  Detected via SharePoint REST API SchemaXml / office.server.policy
                    XmlDocument on list-bound content types
    - Output:       Single flat CSV; one row per CT policy found, or one row per site
                    (with blank policy columns) if features are active but no policies set
    - Scan modes:   Full tenant (default) or targeted list via $SiteListPath

.LINK
    https://admin.microsoft.com/Adminportal/Home#/MessageCenter (search MC1211579)
    https://learn.microsoft.com/en-us/sharepoint/dev/general-development/how-to-avoid-getting-throttled-or-blocked-in-sharepoint-online

.EXAMPLE
    .\SPO-IRMDetector.ps1
    Full tenant scan. Finds all sites with retiring IM/IPRM features active and checks
    each for configured Information Management policies on document libraries.

.EXAMPLE
    # Set $SiteListPath = "./affected-sites.txt" in the CONFIGURATION section, then run:
    .\SPO-IRMDetector.ps1
    Targeted scan against only the URLs listed in affected-sites.txt. Useful when you
    already have a list from a prior full scan and want faster re-inspection.

#>
# SPO-IRMDetector.ps1
# Purpose : Inventory SPO tenant for retiring Information Management / In-Place Records
#           Management features (MC1211579) and surface any configured IM policies.
# Usage   : Set credentials + optional $SiteListPath in CONFIGURATION, then run.
# Output  : Single CSV with feature hits and IM policy detail per content type.
# Ref     : https://learn.microsoft.com/en-us/sharepoint/dev/general-development/how-to-avoid-getting-throttled-or-blocked-in-sharepoint-online

#region ==================== CONFIGURATION ====================
# ============================================================
# MODIFY THE SETTINGS BELOW TO MATCH YOUR ENVIRONMENT
# ============================================================

# ----- Tenant & Authentication -----
$tenant = "m365cpi13246019"                                    # Tenant name (without .sharepoint.com)
$appID = "1e488dc4-1977-48ef-8d4d-9856f4e04536"                # Azure AD App Registration Client ID
$thumbprint = "216f5dd7327719bc8cf15ff3c077adf59ace0c23"       # Certificate thumbprint
$tenantid = "9cfc42cb-51da-4055-87e9-b20a170b6ba3"             # Azure AD Tenant ID

# ----- Site Scope -----
# Leave empty ("") to scan ALL tenant sites
# Set to a file path to scan only specific sites (one URL per line in the file)
$SiteListPath = ""                                              # Example: "./sites.txt"

# ----- Features to Find -----
# Add the Feature Definition IDs (GUIDs) you want to search for
# Get IDs by running:
#   Get-PnPFeature -Scope Site | Select-Object DisplayName, DefinitionId
#   Get-PnPFeature -Scope Web  | Select-Object DisplayName, DefinitionId
$featureIds = @(
    "da2e115b-07e4-49d9-bb2c-35e93bb9fca9",                     # In Place Records Management (Site scope)
    "9e56487c-795a-4077-9425-54a1ecb84282",                     # Hold                          (Web scope)
    "063c26fa-3ccc-4180-8a84-b6f98e991df3"                      # LocationBasedPolicy           (Site scope)
)

# ----- Output -----
$outputFolder = "./"                                            # Folder for the results CSV
$outputPrefix = "FeatureInventory"                              # File name prefix (timestamp will be added)
# Output file will be: FeatureInventory_yyyyMMdd_HHmmss.csv

# ----- Throttling Settings -----
# Adjust these if you experience throttling issues
$maxRetries = 5                                                 # Max retries for throttled requests
$baseDelaySeconds = 2                                           # Base delay for exponential backoff
$delayBetweenSites = 500                                        # Milliseconds between sites

# ============================================================
# END OF CONFIGURATION - DO NOT MODIFY BELOW THIS LINE
# ============================================================
#endregion

#region ==================== FUNCTIONS ====================

# Initialize throttle counter
$script:throttleCount = 0

function Invoke-WithThrottleHandling {
    param(
        [Parameter(Mandatory = $true)]
        [ScriptBlock]$ScriptBlock,
        [Parameter(Mandatory = $false)]
        [string]$OperationName = "Operation",
        [Parameter(Mandatory = $false)]
        [int]$MaxRetries = $script:maxRetries
    )
    
    $retryCount = 0
    $success = $false
    $result = $null
    
    while (-not $success -and $retryCount -le $MaxRetries) {
        try {
            $result = & $ScriptBlock
            $success = $true
        }
        catch {
            $exception = $_.Exception
            $statusCode = $null
            $retryAfter = $null
            
            # Try to get the HTTP status code and Retry-After header
            if ($exception.Response) {
                $statusCode = [int]$exception.Response.StatusCode
                $retryAfter = $exception.Response.Headers["Retry-After"]
            }
            
            # Check for throttling (429) or server busy (503)
            if ($statusCode -eq 429 -or $statusCode -eq 503) {
                $retryCount++
                $script:throttleCount++
                
                if ($retryCount -gt $MaxRetries) {
                    Write-Host "        Max retries ($MaxRetries) exceeded for $OperationName. Skipping." -ForegroundColor Red
                    throw $_
                }
                
                # Determine wait time: use Retry-After header if available, otherwise exponential backoff
                if ($retryAfter) {
                    $waitSeconds = [int]$retryAfter
                    Write-Host "        Throttled (HTTP $statusCode). Retry-After: $waitSeconds seconds. Attempt $retryCount of $MaxRetries" -ForegroundColor DarkYellow
                }
                else {
                    # Exponential backoff: 2^retryCount * baseDelay (2, 4, 8, 16, 32 seconds...)
                    $waitSeconds = [math]::Pow(2, $retryCount) * $baseDelaySeconds
                    Write-Host "        Throttled (HTTP $statusCode). Waiting $waitSeconds seconds (exponential backoff). Attempt $retryCount of $MaxRetries" -ForegroundColor DarkYellow
                }
                
                Start-Sleep -Seconds $waitSeconds
            }
            else {
                # Non-throttling error, rethrow
                throw $_
            }
        }
    }
    
    return $result
}

function Get-SiteIMPolicies {
    <#
    .SYNOPSIS
        Scans all non-hidden document libraries in the currently connected site for
        Information Management Policies configured on their content types.
    .OUTPUTS
        List of PSCustomObjects: SiteUrl, LibraryTitle, ContentTypeName,
        PolicyName, PolicyDescription, RetentionDefined
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl
    )

    $policyResults = [System.Collections.Generic.List[PSObject]]::new()

    try {
        # Get all non-hidden document libraries
        $lists = Invoke-WithThrottleHandling -OperationName "Get document libraries from $SiteUrl" -ScriptBlock {
            Get-PnPList | Where-Object { $_.BaseType -eq "DocumentLibrary" -and $_.Hidden -eq $false }
        }

        Write-Host "      Found $($lists.Count) libraries. Checking for IM policies..." -ForegroundColor Gray

        foreach ($list in $lists) {
            try {
                $listId = $list.Id.ToString()
                $ctResp = Invoke-PnPSPRestMethod -Url "/_api/web/lists(guid'$listId')/ContentTypes?`$select=Name,StringId,SchemaXml" -Method Get -ErrorAction Stop
                $ctItems = if ($ctResp.value) { $ctResp.value } else { @() }

                foreach ($ct in $ctItems) {
                    $schema = $ct.SchemaXml
                    if (-not $schema) { continue }

                    # The IM policy lives in <XmlDocument NamespaceURI="office.server.policy">
                    # inside the CT's SchemaXml <XmlDocuments> section.
                    try { [xml]$schemaDoc = $schema } catch { continue }

                    $policyXmlDocNode = $schemaDoc.SelectSingleNode(
                        "//*[local-name()='XmlDocument'][@NamespaceURI='office.server.policy']")
                    if (-not $policyXmlDocNode) { continue }

                    $policyEl = $policyXmlDocNode.SelectSingleNode("*[local-name()='Policy']")
                    if (-not $policyEl) { continue }

                    $policyName = $policyEl.SelectSingleNode("*[local-name()='Name']").InnerText
                    $policyDesc = $policyEl.SelectSingleNode("*[local-name()='Description']").InnerText
                    if (-not $policyName -or $policyName.Trim() -eq "") { continue }

                    $hasRetention = $null -ne $policyEl.SelectSingleNode(
                        ".//*[contains(@featureId,'Expiration')]")

                    Write-Host "        [IM Policy] $($list.Title) / $($ct.Name): '$policyName' | Retention: $(if ($hasRetention) { 'Yes' } else { 'No' })" -ForegroundColor Cyan

                    $policyResults.Add([PSCustomObject]@{
                            SiteUrl           = $SiteUrl
                            LibraryTitle      = $list.Title
                            ContentTypeName   = $ct.Name
                            PolicyName        = $policyName
                            PolicyDescription = $policyDesc
                            RetentionDefined  = if ($hasRetention) { "Yes" } else { "No" }
                        })
                }
            }
            catch {
                Write-Host "      Warning: Could not check library '$($list.Title)' - $($_.Exception.Message)" -ForegroundColor DarkYellow
            }
        }
    }
    catch {
        Write-Host "    Warning: Could not retrieve libraries for $SiteUrl - $($_.Exception.Message)" -ForegroundColor DarkYellow
    }

    return $policyResults
}
#endregion

#region ==================== MAIN EXECUTION ====================

# Generate output file path with timestamp
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputPath = Join-Path -Path $outputFolder -ChildPath "${outputPrefix}_${timestamp}.csv"

# Initialize CSV with headers (single file: feature + policy detail columns)
"SiteUrl,FoundFeatureIds,LibraryTitle,ContentTypeName,PolicyName,PolicyDescription,RetentionDefined" | Out-File -FilePath $outputPath -Encoding UTF8

# Connect to admin site to get all tenant sites
Write-Host "Connecting to SharePoint Admin..." -ForegroundColor Cyan
Invoke-WithThrottleHandling -OperationName "Connect to Admin" -ScriptBlock {
    Connect-PnPOnline -Url "https://$($script:tenant)-admin.sharepoint.com" -ClientId $script:appID -Thumbprint $script:thumbprint -Tenant $script:tenantid
}

# Get sites: from file if specified, otherwise all tenant sites
if ($SiteListPath -and $SiteListPath.Trim() -ne "") {
    if (Test-Path $SiteListPath) {
        Write-Host "Loading sites from file: $SiteListPath" -ForegroundColor Cyan
        $siteUrls = Get-Content -Path $SiteListPath | Where-Object { $_.Trim() -ne "" }
        $filteredSites = $siteUrls | ForEach-Object {
            [PSCustomObject]@{ Url = $_.Trim() }
        }
        Write-Host "Loaded $($filteredSites.Count) sites from file" -ForegroundColor Green
    }
    else {
        Write-Host "Site list file not found: $SiteListPath" -ForegroundColor Red
        exit 1
    }
}
else {
    Write-Host "Retrieving all tenant sites..." -ForegroundColor Cyan
    $sites = Invoke-WithThrottleHandling -OperationName "Get Tenant Sites" -ScriptBlock {
        Get-PnPTenantSite
    }
    # You may choose to filter sites
    $filteredSites = $sites #| Where-Object { $_.Url -like '*IT*' }
    Write-Host "Found $($filteredSites.Count) sites in tenant" -ForegroundColor Green
}

$totalSites = $filteredSites.Count
Write-Host ""

$currentSite = 0
$sitesWithHits = 0
$totalFeaturesFound = 0
$totalIMPolicies = 0
$startTime = Get-Date

foreach ($site in $filteredSites) {
    $currentSite++
    $percentComplete = [math]::Round(($currentSite / $totalSites) * 100, 1)
    Write-Host "[$currentSite/$totalSites] ($percentComplete%) Scanning: $($site.Url)" -ForegroundColor Yellow
    
    try {
        # Connect to the site with throttle handling
        $siteUrl = $site.Url
        Invoke-WithThrottleHandling -OperationName "Connect to $siteUrl" -ScriptBlock {
            Connect-PnPOnline -Url $siteUrl -ClientId $script:appID -Thumbprint $script:thumbprint -Tenant $script:tenantid -ErrorAction Stop
        }
        
        $siteHasHits = $false
        $matchedFeatureIds = [System.Collections.Generic.List[string]]::new()

        # Check both Site-scope and Web-scope features.
        # All GUIDs in $featureIds are checked against both scopes, so Web-scoped features
        # (e.g. Hold) will be correctly detected when iterating the "Web" scope pass.
        foreach ($scope in @("Site", "Web")) {
            try {
                $currentScope = $scope
                $features = Invoke-WithThrottleHandling -OperationName "Get $currentScope features from $siteUrl" -ScriptBlock {
                    Get-PnPFeature -Scope $currentScope -ErrorAction Stop
                }

                if ($features) {
                    foreach ($feature in $features) {
                        $featureGuid = $feature.DefinitionId.ToString()
                        if ($featureIds -contains $featureGuid) {
                            if (-not $siteHasHits) {
                                $sitesWithHits++
                                $siteHasHits = $true
                            }
                            $totalFeaturesFound++

                            $displayName = if ($feature.DisplayName) { $feature.DisplayName } else { "(no display name)" }
                            Write-Host "    FOUND [$scope]: $displayName  ($featureGuid)" -ForegroundColor Green

                            # Accumulate: "[Scope] DisplayName (GUID)"
                            $matchedFeatureIds.Add("[$scope] $displayName ($featureGuid)")
                        }
                    }
                }
            }
            catch {
                Write-Host "    Warning: Could not retrieve $scope features - $($_.Exception.Message)" -ForegroundColor DarkYellow
            }
        }

        # Write one row per site if any features matched
        if ($matchedFeatureIds.Count -gt 0) {
            # Check for Information Management Policies on libraries in this site
            Write-Host "    Checking for Information Management policies on libraries..." -ForegroundColor Cyan
            $imPolicies = Get-SiteIMPolicies -SiteUrl $site.Url
            $featureStr = $matchedFeatureIds -join "; "

            if ($imPolicies.Count -gt 0) {
                $totalIMPolicies += $imPolicies.Count
                foreach ($policy in $imPolicies) {
                    [PSCustomObject]@{
                        SiteUrl           = $site.Url
                        FoundFeatureIds   = $featureStr
                        LibraryTitle      = $policy.LibraryTitle
                        ContentTypeName   = $policy.ContentTypeName
                        PolicyName        = $policy.PolicyName
                        PolicyDescription = $policy.PolicyDescription
                        RetentionDefined  = $policy.RetentionDefined
                    } | Export-Csv -Path $outputPath -Append -NoTypeInformation
                }
            }
            else {
                # Site has features but no IM policies — write one row with empty policy columns
                [PSCustomObject]@{
                    SiteUrl           = $site.Url
                    FoundFeatureIds   = $featureStr
                    LibraryTitle      = ""
                    ContentTypeName   = ""
                    PolicyName        = ""
                    PolicyDescription = ""
                    RetentionDefined  = ""
                } | Export-Csv -Path $outputPath -Append -NoTypeInformation
            }
        }
    }
    catch {
        Write-Host "    Error accessing site: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    # Add delay between sites to avoid request spikes
    Start-Sleep -Milliseconds $delayBetweenSites
}

$endTime = Get-Date
$duration = $endTime - $startTime

Write-Host ""
Write-Host "========== SCAN COMPLETE ==========" -ForegroundColor Cyan
Write-Host "Total sites scanned:    $totalSites" -ForegroundColor White
Write-Host "Sites with hits:        $sitesWithHits" -ForegroundColor White
Write-Host "Features found:         $totalFeaturesFound" -ForegroundColor White
Write-Host "IM policies found:      $totalIMPolicies" -ForegroundColor White
Write-Host "Throttle events:        $($script:throttleCount)" -ForegroundColor $(if ($script:throttleCount -gt 0) { "Yellow" } else { "White" })
Write-Host "Duration:               $($duration.ToString('hh\:mm\:ss'))" -ForegroundColor White
Write-Host "==================================" -ForegroundColor Cyan
Write-Host "Results exported to: $outputPath" -ForegroundColor Green
#endregion
