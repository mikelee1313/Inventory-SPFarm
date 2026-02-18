<#
.SYNOPSIS
    Inventories SharePoint Online Information Management and In-Place Records Management features
    across site collections as part of MC1211579 retirement preparedness.

.DESCRIPTION
    This script helps you inventory your SPO tenant to identify which site collections have the
    Information Management and In-Place Records Management features enabled, in preparation for
    the SharePoint Online feature retirement announced in Message Center post MC1211579.

    Specifically, it scans for the following retiring features:
      - In Place Records Management  (Site scope)
      - Hold                         (Web scope)
      - Location Based Policy         (Site scope)

    It checks both Site-scoped (site collection) and Web-scoped (root web) feature activations,
    includes comprehensive throttling handling per Microsoft guidance, and uses certificate-based
    authentication with an Azure AD App Registration. Results are exported to a CSV file with
    immediate write capability for large datasets.

    Use the output to prioritize remediation efforts before the retirement date communicated
    in MC1211579.

.PARAMETER tenant
    The SharePoint tenant name (without .sharepoint.com suffix).
    Example: "m365cpi13246019"

.PARAMETER appID
    The Azure AD App Registration Client ID used for authentication.

.PARAMETER thumbprint
    The certificate thumbprint for certificate-based authentication.

.PARAMETER tenantid
    The Azure AD Tenant ID.

.PARAMETER SiteListPath
    Optional file path containing a list of site URLs to scan (one per line).
    If empty, scans all tenant sites.

.PARAMETER featureIds
    Array of Feature Definition IDs (GUIDs) to search for. Obtain IDs using:
    Get-PnPFeature -Scope Site | Select-Object DisplayName, DefinitionId
    Get-PnPFeature -Scope Web  | Select-Object DisplayName, DefinitionId

.PARAMETER outputFolder
    Destination folder for the results CSV file.

.PARAMETER outputPrefix
    Prefix for the output filename. Timestamp is appended automatically.

.PARAMETER maxRetries
    Maximum retry attempts for throttled requests (HTTP 429/503).

.PARAMETER baseDelaySeconds
    Base delay value for exponential backoff calculation.

.PARAMETER delayBetweenSites
    Milliseconds to wait between site scans.

.OUTPUTS
    CSV file containing: SiteUrl, Scope, FeatureDisplayName, FeatureId

.NOTES
    - Author: Mike Lee
    - Date: 2/18/26
    - Related MC post: MC1211579 (SharePoint Online Information Management / In-Place Records
      Management feature retirement)
    - Requires PnP PowerShell module
    - Uses certificate-based authentication
    - Implements exponential backoff for throttling
    - Respects Retry-After headers from responses
    - Checks both Site-scope and Web-scope features per site collection
    - Real-time progress reporting with detailed statistics
    - Only sites with at least one matching feature are written to the output CSV

.LINK
    https://admin.microsoft.com/Adminportal/Home#/MessageCenter (search MC1211579)
    https://learn.microsoft.com/en-us/sharepoint/dev/general-development/how-to-avoid-getting-throttled-or-blocked-in-sharepoint-online

.EXAMPLE
    .\Find-SPOFeatures.ps1
    Scans all tenant sites and reports which site collections have the retiring
    Information Management / In-Place Records Management features activated.

#>
# Find-SPOFeatures.ps1
# Purpose : Inventory SPO tenant for Information Management and In-Place Records Management
#           features that are being retired per MC1211579.
# Usage   : Update $featureIds and $SiteListPath (or leave blank for all sites), then run.
# Output  : CSV with one row per affected site collection listing all matched feature IDs.
# Ref     : https://learn.microsoft.com/en-us/sharepoint/dev/general-development/how-to-avoid-getting-throttled-or-blocked-in-sharepoint-online

#region ==================== CONFIGURATION ====================
# ============================================================
# MODIFY THE SETTINGS BELOW TO MATCH YOUR ENVIRONMENT
# ============================================================

# ----- Tenant & Authentication -----
$tenant = "m365cpi13246019"                                    # Tenant name (without .sharepoint.com)
$appID = "1e488dc4-1977-48ef-8d4d-9856f4e04536"                # Azure AD App Registration Client ID
$thumbprint = "5EAD7303A5C7E27DB4245878AD554642940BA082"       # Certificate thumbprint
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
#endregion

#region ==================== MAIN EXECUTION ====================

# Generate output file path with timestamp
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputPath = Join-Path -Path $outputFolder -ChildPath "${outputPrefix}_${timestamp}.csv"

# Initialize CSV with headers
"SiteUrl,FoundFeatureIds" | Out-File -FilePath $outputPath -Encoding UTF8

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
            [PSCustomObject]@{
                SiteUrl         = $site.Url
                FoundFeatureIds = $matchedFeatureIds -join "; "
            } | Export-Csv -Path $outputPath -Append -NoTypeInformation
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
Write-Host "Throttle events:        $($script:throttleCount)" -ForegroundColor $(if ($script:throttleCount -gt 0) { "Yellow" } else { "White" })
Write-Host "Duration:               $($duration.ToString('hh\:mm\:ss'))" -ForegroundColor White
Write-Host "===================================" -ForegroundColor Cyan
Write-Host "Results exported to: $outputPath" -ForegroundColor Green
#endregion
