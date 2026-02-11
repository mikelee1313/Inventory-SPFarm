<#
.SYNOPSIS
    Inventories specific Web Parts across SharePoint Online sites by directly inspecting pages.

.DESCRIPTION
    This script scans SharePoint Online sites and pages to find and catalog specific Web Parts.
    It includes comprehensive throttling handling per Microsoft guidance and uses certificate-based
    authentication with an Azure AD App Registration. Results are exported to a CSV file with
    immediate write capability for large datasets.

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

.PARAMETER webPartIds
    Array of Web Part IDs to search for. Obtain IDs using:
    Get-PnPPageComponent -Page "page-name" | Select-Object Title, WebPartId

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

.PARAMETER delayBetweenPages
    Milliseconds to wait between page scans.

.OUTPUTS
    CSV file containing: SiteUrl, PageUrl, PageName, WebPartTitle, WebPartId

.NOTES
    - Author: Mike Lee
    - Date: 2/11/26
    - Requires PnP PowerShell module
    - Uses certificate-based authentication
    - Implements exponential backoff for throttling
    - Respects Retry-After headers from responses
    - Skips non-ASPX files and inaccessible pages gracefully
    - Real-time progress reporting with detailed statistics

.EXAMPLE
    .\Find-SPOWebparts.ps1
    Scans all tenant sites for configured Web Parts.

#>
# FindWebParts.ps1 - Inventory specific Web Parts by directly inspecting pages (not search-based)
# Includes throttling handling per Microsoft guidance:
# https://learn.microsoft.com/en-us/sharepoint/dev/general-development/how-to-avoid-getting-throttled-or-blocked-in-sharepoint-online

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

# ----- Web Parts to Find -----
# Add the WebPartIds you want to search for
# Get IDs by running: Get-PnPPageComponent -Page "page-name" | Select-Object Title, WebPartId
$webPartIds = @(
    "e377ea37-9047-43b9-8cdb-a761be2f8e09"                     # Bing Maps
)

# ----- Output -----
$outputFolder = "./"                                            # Folder for the results CSV
$outputPrefix = "WebPartInventory"                              # File name prefix (timestamp will be added)
# Output file will be: WebPartInventory_yyyyMMdd_HHmmss.csv

# ----- Throttling Settings -----
# Adjust these if you experience throttling issues
$maxRetries = 5                                                 # Max retries for throttled requests
$baseDelaySeconds = 2                                           # Base delay for exponential backoff
$delayBetweenSites = 500                                        # Milliseconds between sites
$delayBetweenPages = 100                                        # Milliseconds between pages

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
"SiteUrl,PageUrl,PageName,WebPartTitle,WebPartId" | Out-File -FilePath $outputPath -Encoding UTF8

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
$totalPagesScanned = 0
$totalWebPartsFound = 0
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
        
        # Get all pages from the Site Pages library with throttle handling
        $pages = Invoke-WithThrottleHandling -OperationName "Get Site Pages" -ScriptBlock {
            Get-PnPListItem -List "Site Pages" -Fields "FileLeafRef" -ErrorAction Stop
        }
        
        if ($pages) {
            $siteHasHits = $false
            
            foreach ($pageItem in $pages) {
                $pageName = $pageItem.FieldValues["FileLeafRef"]
                
                # Skip non-aspx files
                if ($pageName -notlike "*.aspx") { continue }
                
                $totalPagesScanned++
                
                # Add small delay between pages to avoid request spikes
                Start-Sleep -Milliseconds $delayBetweenPages
                
                try {
                    # Get all components on the page with throttle handling
                    $currentPageName = $pageName
                    $components = Invoke-WithThrottleHandling -OperationName "Get components from $currentPageName" -ScriptBlock {
                        Get-PnPPageComponent -Page $currentPageName -ErrorAction Stop
                    }
                    
                    if ($components) {
                        foreach ($component in $components) {
                            if ($component.WebPartId -and $webPartIds -contains $component.WebPartId) {
                                if (-not $siteHasHits) {
                                    $sitesWithHits++
                                    $siteHasHits = $true
                                }
                                $totalWebPartsFound++
                                
                                $pageUrl = "$($site.Url)/SitePages/$pageName"
                                Write-Host "    FOUND: $($component.Title) on $pageName" -ForegroundColor Green
                                
                                # Write to CSV immediately
                                [PSCustomObject]@{
                                    SiteUrl      = $site.Url
                                    PageUrl      = $pageUrl
                                    PageName     = $pageName
                                    WebPartTitle = $component.Title
                                    WebPartId    = $component.WebPartId
                                } | Export-Csv -Path $outputPath -Append -NoTypeInformation
                            }
                        }
                    }
                }
                catch {
                    # Skip pages that can't be read (e.g., wiki pages, publishing pages)
                }
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
Write-Host "Total pages scanned:    $totalPagesScanned" -ForegroundColor White
Write-Host "Web parts found:        $totalWebPartsFound" -ForegroundColor White
Write-Host "Throttle events:        $($script:throttleCount)" -ForegroundColor $(if ($script:throttleCount -gt 0) { "Yellow" } else { "White" })
Write-Host "Duration:               $($duration.ToString('hh\:mm\:ss'))" -ForegroundColor White
Write-Host "===================================" -ForegroundColor Cyan
Write-Host "Results exported to: $outputPath" -ForegroundColor Green
#endregion
