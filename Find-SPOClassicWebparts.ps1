<#
.SYNOPSIS
    Inventories selected classic Web Parts across SharePoint Online sites.

.DESCRIPTION
    This script scans classic ASPX pages in document libraries of each selected site
    and all its nested subsites to find and catalog selected Web Parts. Libraries are
    discovered by BaseType (DocumentLibrary), including custom-named, wiki, and
    publishing libraries, rather than by title.
    It includes comprehensive throttling handling per Microsoft guidance and uses certificate-based
    authentication with an Azure AD App Registration. Results are exported to a CSV file with
    immediate write capability for large datasets.
    Page owner is the Created By user. PageLastModified is the page's Modified timestamp
    in ISO 8601 format, not a last-view or last-access timestamp.
    Library discovery is filtered on the server to avoid enumerating unrelated restricted lists.

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
    Each URL and its nested subsites are scanned. If empty, scans all tenant site
    collections that match the tenant-site filter and their subsites. Tenant discovery
    excludes OneDrive, inactive, archived, redirect, and selected system sites.
    Explicit file URLs bypass that filter. Overlapping URLs are scanned only once.

.PARAMETER classicWebPartTypes
    Array of classic Web Part type names to search for.

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
    CSV file containing: SiteUrl, PageUrl, PageName, WebPartTitle, WebPartId, WebPartType,
    PageOwner, PageOwnerEmail, PageLastModified

.NOTES
    - Author: Mike Lee
    - Date: 11/4/2026
    - Requires PnP PowerShell module
    - Uses certificate-based authentication
    - Implements exponential backoff for throttling
    - Respects Retry-After headers from responses
    - Skips non-ASPX files and inaccessible pages gracefully
    - Real-time progress reporting with detailed statistics

.EXAMPLE
    .\Find-SPOClassicWebparts.ps1
    Scans classic pages in the tenant, including subsites, for the configured classic Web Parts.

#>
# FindWebParts.ps1 - Inventory selected classic Web Parts by inspecting classic ASPX pages
# Includes throttling handling per Microsoft guidance:
# https://learn.microsoft.com/en-us/sharepoint/dev/general-development/how-to-avoid-getting-throttled-or-blocked-in-sharepoint-online

#region ==================== CONFIGURATION ====================
# ============================================================
# MODIFY THE SETTINGS BELOW TO MATCH YOUR ENVIRONMENT
# ============================================================

# ----- Tenant & Authentication -----
$tenant = "m365cpi13246019"                                    # Tenant name (without .sharepoint.com)
$appID = "abc64618-283f-47ba-a185-50d935d51d57"              # Azure AD App Registration Client ID
$thumbprint = "B696FDCFE1453F3FBC6031F54DE988DA0ED905A9"       # Certificate thumbprint
$tenantid =  "9cfc42cb-51da-4055-87e9-b20a170b6ba3"             # Azure AD Tenant ID

# ----- Site Scope -----
# Leave empty ("") to scan ALL tenant sites
# Set to a file path to scan only specific sites (one URL per line in the file)
$SiteListPath = ""                                              # Example: "./sites.txt"

# ----- Classic Web Parts to Find -----
# Specify the full type names used in classic Web Part XML
$classicWebPartTypes = @(
    "Microsoft.SharePoint.WebPartPages.ScriptEditorWebPart"
    "Microsoft.SharePoint.WebPartPages.ContentEditorWebPart"
    "Microsoft.Office.Server.Search.WebControls.ContentBySearchWebPart"
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

function Get-ClassicWebPartTypeName {
    param(
        [Parameter(Mandatory = $true)]
        [object]$WebPartXml
    )

    if ($WebPartXml -is [System.Xml.XmlNode]) {
        $xmlDocument = [xml]$WebPartXml.OuterXml
    }
    else {
        $xmlDocument = [xml]([string]$WebPartXml)
    }

    $typeNameNode = $xmlDocument.SelectSingleNode("//*[local-name()='type' and @name]")
    if ($typeNameNode) {
        $typeName = $typeNameNode.Attributes["name"].Value
    }
    else {
        $typeNameNode = $xmlDocument.SelectSingleNode("//*[local-name()='TypeName' or local-name()='PartType']")
        if (-not $typeNameNode) {
            return ""
        }
        $typeName = $typeNameNode.InnerText
    }

    return ($typeName -split ',')[0].Trim()
}
#endregion

#region ==================== MAIN EXECUTION ====================

# Generate output file path with timestamp
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputPath = Join-Path -Path $outputFolder -ChildPath "${outputPrefix}_${timestamp}.csv"

# Initialize CSV with headers
"SiteUrl,PageUrl,PageName,WebPartTitle,WebPartId,WebPartType,PageOwner,PageOwnerEmail,PageLastModified" | Out-File -FilePath $outputPath -Encoding UTF8

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
    Write-Host "Retrieving active, non-archived tenant sites (excluding OneDrive and selected system sites)..." -ForegroundColor Cyan
    $filteredSites = @(Invoke-WithThrottleHandling -OperationName "Get Tenant Sites" -ScriptBlock {
        Get-PnPTenantSite -IncludeOneDriveSites:$false -ErrorAction Stop | Where-Object {
            $_.Template -notmatch "SRCHCEN|MYSITE|APPCATALOG|PWS|POINTPUBLISHINGTOPIC|SPSMSITEHOST|EHS|REVIEWCTR|TENANTADMIN" -and
            $_.Status -eq "Active" -and
            $_.ArchiveStatus -eq "NotArchived" -and
            $_.Template -notlike "*Redirect*" -and
            -not [string]::IsNullOrEmpty($_.Url)
        }
    })
    Write-Host "Found $($filteredSites.Count) tenant sites matching the scan filter" -ForegroundColor Green
}

$siteQueue = [System.Collections.Generic.Queue[object]]::new()
$knownSiteUrls = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach ($site in $filteredSites) {
    if ($knownSiteUrls.Add($site.Url.TrimEnd('/'))) {
        $siteQueue.Enqueue($site)
    }
}
$totalSites = $knownSiteUrls.Count
Write-Host ""

$currentSite = 0
$sitesWithHits = 0
$totalPagesScanned = 0
$totalWebPartsFound = 0
$startTime = Get-Date

while ($siteQueue.Count -gt 0) {
    $site = $siteQueue.Dequeue()
    $currentSite++
    $percentComplete = [math]::Round(($currentSite / $totalSites) * 100, 1)
    Write-Host "[$currentSite/$totalSites] ($percentComplete%) Scanning: $($site.Url)" -ForegroundColor Yellow
    
    try {
        # Connect to the site with throttle handling
        $siteUrl = $site.Url
        Invoke-WithThrottleHandling -OperationName "Connect to $siteUrl" -ScriptBlock {
            Connect-PnPOnline -Url $siteUrl -ClientId $script:appID -Thumbprint $script:thumbprint -Tenant $script:tenantid -ErrorAction Stop
        }

        # Queue immediate children; processing each child discovers every nested level.
        try {
            $subsites = Invoke-WithThrottleHandling -OperationName "Get subsites of $siteUrl" -ScriptBlock {
                Get-PnPSubWeb -ErrorAction Stop
            }
            foreach ($subsite in $subsites) {
                if ($knownSiteUrls.Add($subsite.Url.TrimEnd('/'))) {
                    $siteQueue.Enqueue([PSCustomObject]@{ Url = $subsite.Url })
                }
            }
            $totalSites = $knownSiteUrls.Count
        }
        catch {
            Write-Host "    Unable to discover subsites of $siteUrl`: $($_.Exception.Message). Descendant coverage may be incomplete." -ForegroundColor DarkYellow
        }
        
        # Filter on the server; enumerating unrelated restricted lists can fail with access denied.
        $pageLibraries = Invoke-WithThrottleHandling -OperationName "Get page libraries" -ScriptBlock {
            $libraryUrl = '/_api/web/lists?$select=Id,Title&$filter=BaseType eq 1'
            $libraries = @(
                do {
                    $response = Invoke-PnPSPRestMethod -Url $libraryUrl -Method Get -ErrorAction Stop
                    $response.value
                    $libraryUrl = $response.'odata.nextLink'
                    if (-not $libraryUrl) {
                        $libraryUrl = $response.'@odata.nextLink'
                    }
                } while ($libraryUrl)
            )
            $libraries
        }

        $siteHasHits = $false
        foreach ($pageLibrary in $pageLibraries) {
            try {
                $pages = Invoke-WithThrottleHandling -OperationName "Get pages from $($pageLibrary.Title)" -ScriptBlock {
                    Get-PnPListItem -List $pageLibrary.Id -Fields "FileLeafRef", "FileRef", "Author", "Modified" -PageSize 500 -ErrorAction Stop
                }
            }
            catch {
                Write-Host "    Unable to read library $($pageLibrary.Title) on $siteUrl`: $($_.Exception.Message). Continuing with other libraries." -ForegroundColor DarkYellow
                continue
            }
            
            foreach ($pageItem in $pages) {
                $pageName = $pageItem.FieldValues["FileLeafRef"]
                
                # Skip non-aspx files
                if ($pageName -notlike "*.aspx") { continue }

                $pageAuthor = $pageItem.FieldValues["Author"]
                $pageModified = $pageItem.FieldValues["Modified"]
                $pageLastModified = ""
                if ($null -ne $pageModified) {
                    $pageLastModified = ([datetime]$pageModified).ToString("o")
                }
                if ($null -eq $pageAuthor -or $null -eq $pageModified) {
                    Write-Host "    Missing Created By or Modified metadata on $($pageItem.FieldValues['FileRef']); unavailable values will be blank in the CSV." -ForegroundColor DarkYellow
                }
                
                $totalPagesScanned++
                
                # Add small delay between pages to avoid request spikes
                Start-Sleep -Milliseconds $delayBetweenPages
                
                try {
                    # Retrieve classic web part definitions and their XML type names
                    $currentPageName = $pageName
                    $pageServerRelativeUrl = $pageItem.FieldValues["FileRef"]
                    $webParts = Invoke-WithThrottleHandling -OperationName "Get classic web parts from $currentPageName" -ScriptBlock {
                        Get-PnPWebPart -ServerRelativePageUrl $pageServerRelativeUrl -ErrorAction Stop
                    }
                }
                catch {
                    Write-Host "    Unable to inspect classic web parts on $pageName`: $($_.Exception.Message)" -ForegroundColor DarkYellow
                    continue
                }

                foreach ($webPart in $webParts) {
                    try {
                        $webPartXml = Invoke-WithThrottleHandling -OperationName "Get web part XML from $currentPageName" -ScriptBlock {
                            Get-PnPWebPartXml -ServerRelativePageUrl $pageServerRelativeUrl -Identity $webPart.Id -ErrorAction Stop
                        }
                        $webPartType = Get-ClassicWebPartTypeName -WebPartXml $webPartXml
                    }
                    catch {
                        Write-Host "    Unable to inspect web part $($webPart.Id) on $pageName`: $($_.Exception.Message)" -ForegroundColor DarkYellow
                        continue
                    }

                    if ($webPartType -and $classicWebPartTypes -contains $webPartType) {
                        if (-not $siteHasHits) {
                            $sitesWithHits++
                            $siteHasHits = $true
                        }
                        $totalWebPartsFound++
                        
                        $pageUrl = ([uri]::new([uri]$site.Url, $pageServerRelativeUrl)).AbsoluteUri
                        Write-Host "    FOUND: $webPartType ($($webPart.WebPart.Title)) on $pageName" -ForegroundColor Green
                        
                        # Write to CSV immediately
                        [PSCustomObject]@{
                            SiteUrl          = $site.Url
                            PageUrl          = $pageUrl
                            PageName         = $pageName
                            WebPartTitle     = $webPart.WebPart.Title
                            WebPartId        = $webPart.Id
                            WebPartType      = $webPartType
                            PageOwner        = $pageAuthor.LookupValue
                            PageOwnerEmail   = $pageAuthor.Email
                            PageLastModified = $pageLastModified
                        } | Export-Csv -Path $outputPath -Append -NoTypeInformation
                    }
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
Write-Host "Sites/subsites attempted: $currentSite" -ForegroundColor White
Write-Host "Sites with hits:        $sitesWithHits" -ForegroundColor White
Write-Host "Total pages scanned:    $totalPagesScanned" -ForegroundColor White
Write-Host "Web parts found:        $totalWebPartsFound" -ForegroundColor White
Write-Host "Throttle events:        $($script:throttleCount)" -ForegroundColor $(if ($script:throttleCount -gt 0) { "Yellow" } else { "White" })
Write-Host "Duration:               $($duration.ToString('hh\:mm\:ss'))" -ForegroundColor White
Write-Host "===================================" -ForegroundColor Cyan
Write-Host "Results exported to: $outputPath" -ForegroundColor Green
#endregion
