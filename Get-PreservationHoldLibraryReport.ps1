<#
.SYNOPSIS
    Scans SharePoint Online sites and reports the file count and total size of the PreservationHoldLibrary on each site.

.DESCRIPTION
    This script connects to SharePoint Online using Microsoft Graph API with provided tenant-level credentials and iterates through a list of 
    site URLs specified in an input file. For each site it locates the PreservationHoldLibrary, recursively counts all files (not folders),
    and sums their sizes in bytes. Results are output per-site showing file count and total size in bytes, MB, and GB.
    The script logs its operations and outputs the results to a CSV or Excel file using the ImportExcel module.

.PARAMETER None
    This script does not accept parameters via the command line. Configuration is done within the script.

.INPUTS
    A text file containing SharePoint site URLs to scan (path specified in $inputFilePath variable).

.OUTPUTS
    - Report CSV: One row per site with file count and size totals
      (path: $env:TEMP\PreservationHoldLibrary_Report_[timestamp].csv)
    - Files CSV: One row per file, written when $getFiles is "top" or "all"
      (path: $env:TEMP\PreservationHoldLibrary_Files_[timestamp].csv)
    - Log file documenting the script's execution
      (path: $env:TEMP\PreservationHoldLibrary_Report_[timestamp].txt)

.NOTES
    File Name      : Get-PreservationHoldLibraryReport.ps1
    Author         : Mike Lee
    Date Created   : 4/14/2026

    The script uses app-only authentication with a certificate thumbprint and Microsoft Graph API. Make sure the app has
    proper permissions in your tenant (Sites.Read.All or Sites.ReadWrite.All is recommended).

    PREREQUISITES:
    - Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
    - Install-Module Microsoft.Graph.Sites -Scope CurrentUser
    - Install-Module Microsoft.Graph.Files -Scope CurrentUser

.DISCLAIMER
Disclaimer: The sample scripts are provided AS IS without warranty of any kind. 
Microsoft further disclaims all implied warranties including, without limitation, 
any implied warranties of merchantability or of fitness for a particular purpose. 
The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. 
In no event shall Microsoft, its authors, or anyone else involved in the creation, 
production, or delivery of the scripts be liable for any damages whatsoever 
(including, without limitation, damages for loss of business profits, business interruption, 
loss of business information, or other pecuniary loss) arising out of the use of or inability 
to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages.

.EXAMPLE
    .\Get-PreservationHoldLibraryReport.ps1
    Executes the script with the configured settings. Ensure you've updated the variables at the top
    of the script (appID, thumbprint, tenant, inputFilePath, outputFormat, and debug) before running.

    Example configurations:
    - $debug = $true    # Enable detailed debug output
    - $debug = $false   # Enable informational output only (default)
#>

#region User Configuration
# =================================================================================================
# USER CONFIGURATION - Update the variables in this section
# =================================================================================================

# --- Tenant and App Registration Details ---
$appID = "abc64618-283f-47ba-a185-50d935d51d57"                 # This is your Entra App ID
$thumbprint = "B696FDCFE1453F3FBC6031F54DE988DA0ED905A9"        # This is certificate thumbprint
$tenant = "9cfc42cb-51da-4055-87e9-b20a170b6ba3"                # This is your Tenant ID

# --- Input File Path ---
$inputFilePath = 'C:\temp\SPOSiteList.txt' # Path to the input file containing site URLs

# --- Script Behavior Settings ---
$debug = $false  # Enable debug output: $true for detailed debug info, $false for informational only
$getFiles = "all"  # File enumeration mode: "none" = skip, "top" = top 100 per site (largest first), "all" = every file per site

# =================================================================================================
# END OF USER CONFIGURATION
# =================================================================================================
#endregion User Configuration

#region Module Prerequisites
# Check for required modules
$requiredModules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Sites', 'Microsoft.Graph.Files')

Write-Host "Checking and installing required modules..." -ForegroundColor Cyan

foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Host "Module '$module' is not installed. Installing..." -ForegroundColor Yellow
        try {
            Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber | Out-Null
            Write-Host "Successfully installed module '$module'" -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to install module '$module': $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "Please install the module manually: Install-Module -Name $module -Scope CurrentUser" -ForegroundColor Yellow
            exit 1
        }
    }
    try {
        Import-Module $module -Force | Out-Null
        Write-Host "Successfully imported module '$module'" -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to import module '$module': $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

# Verify that required cmdlets are available
$requiredCmdlets = @('Get-MgSite', 'Get-MgSiteList', 'Get-MgSiteListItem', 'Get-MgSiteDrive', 'Connect-MgGraph')
foreach ($cmdlet in $requiredCmdlets) {
    if (-not (Get-Command $cmdlet -ErrorAction SilentlyContinue)) {
        Write-Host "ERROR: Required cmdlet '$cmdlet' is not available. Please ensure all Microsoft Graph modules are properly installed." -ForegroundColor Red
        exit 1
    }
}
Write-Host "All required cmdlets are available." -ForegroundColor Green
#endregion Module Prerequisites

#region Script Initialization
# Script Parameters
Add-Type -AssemblyName System.Web
$startime = Get-Date -Format "yyyyMMdd_HHmmss"
$logFilePath = "$env:TEMP\PreservationHoldLibrary_Report_$startime.txt"

# Set output file path
$outputFilePath = "$env:TEMP\PreservationHoldLibrary_Report_$startime.csv"

# Initialize results collection (List avoids O(n²) array copy on each += append)
$global:reportData = [System.Collections.Generic.List[object]]::new()
$global:topFilesData = [System.Collections.Generic.List[object]]::new()

# Set top files output path (used when $getFiles = $true)
$topFilesOutputPath = "$env:TEMP\PreservationHoldLibrary_Files_$startime.csv"
#endregion Script Initialization

#region Helper Functions

#region Function: Write-Log
# Setup logging
function Write-Log {
    param (
        [string]$message,
        [string]$level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $level - $message"
    Add-Content -Path $logFilePath -Value $logMessage
    
    # Also display important messages to console with color coding
    switch ($level) {
        "ERROR" { Write-Host $message -ForegroundColor Red }
        "WARNING" { Write-Host $message -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $message -ForegroundColor Green }
        default { 
            if ($level -eq "INFO" -and $message -match "Processing|Completed") {
                Write-Host $message -ForegroundColor Cyan
            }
        }
    }
}
#endregion Function: Write-Log

#region Function: Invoke-WithRetry
# Handle SharePoint Online throttling with exponential backoff
function Invoke-WithRetry {
    param (
        [ScriptBlock]$ScriptBlock,
        [int]$MaxRetries = 5,
        [int]$InitialDelaySeconds = 5
    )
    
    $retryCount = 0
    $delay = $InitialDelaySeconds
    $success = $false
    $result = $null
    
    while (-not $success -and $retryCount -lt $MaxRetries) {
        try {
            $result = & $ScriptBlock
            [void]($success = $true)
        }
        catch {
            $exception = $_.Exception
            
            # Check if this is a throttling error (look for specific status codes or messages)
            [void]($isThrottlingError = $false)
            $retryAfterSeconds = $delay
            
            if ($null -ne $exception.Response) {
                # Check for Retry-After header
                $retryAfterHeader = $exception.Response.Headers['Retry-After']
                if ($retryAfterHeader) {
                    [void]($isThrottlingError = $true)
                    $retryAfterSeconds = [int]$retryAfterHeader
                    Write-Log "Received Retry-After header: $retryAfterSeconds seconds" "WARNING"
                }
                
                # Check for 429 (Too Many Requests) or 503 (Service Unavailable)
                $statusCode = [int]$exception.Response.StatusCode
                if ($statusCode -eq 429 -or $statusCode -eq 503) {
                    [void]($isThrottlingError = $true)
                    Write-Log "Detected throttling response (Status code: $statusCode)" "WARNING"
                }
            }
            
            # Also check for specific throttling error messages
            if ($exception.Message -match "throttl" -or 
                $exception.Message -match "too many requests" -or
                $exception.Message -match "temporarily unavailable") {
                [void]($isThrottlingError = $true)
                Write-Log "Detected throttling error in message: $($exception.Message)" "WARNING"
            }
            
            if ($isThrottlingError) {
                $retryCount++
                if ($retryCount -lt $MaxRetries) {
                    Write-Log "Throttling detected. Retry attempt $retryCount of $MaxRetries. Waiting $retryAfterSeconds seconds..." "WARNING"
                    Write-Host "Throttling detected. Retry attempt $retryCount of $MaxRetries. Waiting $retryAfterSeconds seconds..." -ForegroundColor Yellow
                    Start-Sleep -Seconds $retryAfterSeconds
                    
                    # Implement exponential backoff if no Retry-After header was provided
                    if ($retryAfterSeconds -eq $delay) {
                        $delay = $delay * 2 # Exponential backoff
                    }
                }
                else {
                    Write-Log "Maximum retry attempts reached. Giving up on operation." "ERROR"
                    throw $_
                }
            }
            else {
                # Not a throttling error, rethrow
                $errorMessage = $_.Exception.Message
                $logLevel = "WARNING" # Default to WARNING for unexpected errors

                # Check for common, potentially less critical errors
                if ($errorMessage -match "File Not Found" -or $errorMessage -match "404" -or 
                    $errorMessage -match "Access denied" -or $errorMessage -match "403" -or
                    $errorMessage -match "notSupported" -or $errorMessage -match "422" -or
                    $errorMessage -match "not a folder") {
                    $logLevel = "INFO" # Downgrade to INFO for these specific cases
                }
                Write-Log "General Error occurred During retrieval : $errorMessage" $logLevel
                throw $_
            }
        }
    }
    
    return $result
}
#endregion Function: Invoke-WithRetry

#region Function: Read-SiteURLs
# Read site URLs from input file — skips blank lines and comment lines (starting with #)
function Read-SiteURLs {
    param (
        [string]$filePath
    )
    $urls = Get-Content -Path $filePath | Where-Object { $_.Trim() -ne '' -and -not $_.TrimStart().StartsWith('#') }
    return $urls
}
#endregion Function: Read-SiteURLs

#region Function: Get-SiteIdFromUrl
# Helper function to get site ID from URL using Graph API
function Get-SiteIdFromUrl {
    param (
        [string]$siteUrl
    )
    
    try {
        # Parse the site URL to extract hostname and site path
        $uri = [System.Uri]$siteUrl
        $hostname = $uri.Host
        $sitePath = $uri.AbsolutePath
        
        if ($debug) {
            Write-Log "DEBUG - Parsing URL: $siteUrl" "INFO"
            Write-Log "DEBUG - Hostname: $hostname" "INFO"
            Write-Log "DEBUG - Site Path: $sitePath" "INFO"
        }
        
        # Remove leading slash if present
        if ($sitePath.StartsWith('/')) {
            $sitePath = $sitePath.Substring(1)
        }
        
        # Handle different URL formats
        $siteIdentifier = ""
        if ([string]::IsNullOrEmpty($sitePath) -or $sitePath -eq "/") {
            # Root site collection
            $siteIdentifier = $hostname
        }
        else {
            # Sub-site or site collection with path
            $siteIdentifier = "${hostname}:/${sitePath}"
        }
        
        if ($debug) {
            Write-Log "DEBUG - Site identifier for Graph API: $siteIdentifier" "INFO"
        }
        
        # Get site ID using Graph API
        $siteId = Invoke-WithRetry -ScriptBlock {
            try {
                $site = Get-MgSite -SiteId $siteIdentifier
                if ($debug) {
                    Write-Log "DEBUG - Site found: $($site.DisplayName) | ID: $($site.Id)" "INFO"
                }
                return $site.Id
            }
            catch {
                # If the first format fails, try alternative approaches
                if ($sitePath -and $sitePath -ne "") {
                    # Try with different path format
                    $altIdentifier = "${hostname}:/sites/${sitePath}"
                    if ($debug) {
                        Write-Log "DEBUG - Trying alternative identifier: $altIdentifier" "INFO"
                    }
                    $site = Get-MgSite -SiteId $altIdentifier
                    if ($debug) {
                        Write-Log "DEBUG - Site found with alternative format: $($site.DisplayName) | ID: $($site.Id)" "INFO"
                    }
                    return $site.Id
                }
                else {
                    throw $_
                }
            }
        }
        
        return $siteId
    }
    catch {
        Write-Log "Failed to get site ID for $siteUrl : $($_.Exception.Message)" "ERROR"
        if ($debug) {
            Write-Log "DEBUG - Full error details: $($_.Exception)" "ERROR"
        }
        return $null
    }
}
#endregion Function: Get-SiteIdFromUrl

#region Function: Get-LibraryDriveId
# Helper function to get drive ID for a document library
function Get-LibraryDriveId {
    param (
        [string]$siteId,
        [object]$list
    )
    
    try {
        # First try to get from the list object
        if ($list.Drive -and $list.Drive.Id) {
            if ($debug) {
                Write-Log "DEBUG - Found drive ID from list object: $($list.Drive.Id)" "INFO"
            }
            return $list.Drive.Id
        }
        
        # If that fails, try to get all drives for the site and match by name
        if ($debug) {
            Write-Log "DEBUG - Drive ID not found in list object, trying to get site drives" "INFO"
        }
        
        $siteDrives = Invoke-WithRetry -ScriptBlock {
            return Get-MgSiteDrive -SiteId $siteId -All
        }
        
        if ($debug) {
            Write-Log "DEBUG - Found $($siteDrives.Count) drives for site" "INFO"
            foreach ($drive in $siteDrives) {
                Write-Log "DEBUG - Drive: Name='$($drive.Name)', ID='$($drive.Id)', WebUrl='$($drive.WebUrl)'" "INFO"
            }
        }
        
        # Try to match by library name
        $matchingDrive = $siteDrives | Where-Object { $_.Name -eq $list.DisplayName }
        if ($matchingDrive) {
            if ($debug) {
                Write-Log "DEBUG - Found matching drive by name: $($matchingDrive.Id)" "INFO"
            }
            return $matchingDrive.Id
        }
        
        # If still no match, try alternative approaches
        # For "Documents" library, it's often the default drive
        if ($list.DisplayName -eq "Documents" -or $list.DisplayName -eq "Shared Documents") {
            $defaultDrive = $siteDrives | Where-Object { $_.Name -eq "Documents" -or $_.WebUrl -like "*/Shared%20Documents" }
            if ($defaultDrive) {
                if ($debug) {
                    Write-Log "DEBUG - Found default Documents drive: $($defaultDrive.Id)" "INFO"
                }
                return $defaultDrive.Id
            }
        }
        
        Write-Log "Could not determine drive ID for library '$($list.DisplayName)'" "WARNING"
        return $null
    }
    catch {
        Write-Log "Error getting drive ID for library '$($list.DisplayName)': $($_.Exception.Message)" "ERROR"
        return $null
    }
}
#endregion Function: Get-LibraryDriveId

#region Function: Connect-GraphAPI
# Connect to Microsoft Graph
function Connect-GraphAPI {
    try {
        # Connect using certificate-based authentication
        $clientCertificate = Get-ChildItem -Path "Cert:\CurrentUser\My\$thumbprint" -ErrorAction SilentlyContinue
        if (-not $clientCertificate) {
            $clientCertificate = Get-ChildItem -Path "Cert:\LocalMachine\My\$thumbprint" -ErrorAction SilentlyContinue
        }
        
        if (-not $clientCertificate) {
            Write-Log "Certificate with thumbprint $thumbprint not found in CurrentUser\My or LocalMachine\My" "ERROR"
            return $false
        }
        
        # Connect to Microsoft Graph
        Connect-MgGraph -ClientId $appID -TenantId $tenant -Certificate $clientCertificate -NoWelcome
        
        Write-Log "Connected to Microsoft Graph successfully"
        return $true
    }
    catch {
        Write-Log "Failed to connect to Microsoft Graph: $($_.Exception.Message)" "ERROR"
        return $false
    }
}
#endregion Function: Connect-GraphAPI

#region Function: Write-ReportToFile
# Function to write report data to file (CSV or Excel)
function Write-ReportToFile {
    param (
        [array]$Data,
        [string]$FilePath
    )

    if ($Data.Count -eq 0) {
        Write-Log "No data to write to report file." "WARNING"
        return
    }

    try {
        $Data | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8
        Write-Log "Report CSV created: $FilePath" "SUCCESS"
    }
    catch {
        Write-Log "Failed to write report file: $($_.Exception.Message)" "ERROR"
        throw
    }
}
#endregion Function: Write-ReportToFile

#region Function: Get-TopFilesBySize
# Returns files per site sorted largest first. Pass $Top = [int]::MaxValue to retrieve all files.
function Get-TopFilesBySize {
    param (
        [System.Collections.Generic.List[object]]$Items,
        [string]$SiteURL,
        [int]$Top = 100
    )

    $topFiles = $Items |
    Where-Object { -not $_.deleted -and $null -eq $_.folder -and $null -ne $_.file } |
    Sort-Object { if ($null -ne $_.size) { [long]$_.size } else { 0L } } -Descending |
    Select-Object -First $Top

    $result = [System.Collections.Generic.List[object]]::new()
    foreach ($f in $topFiles) {
        $sizeBytes = if ($null -ne $f.size) { [long]$f.size } else { 0L }
        $result.Add([PSCustomObject]@{
                SiteURL       = $SiteURL
                FileName      = $f.name
                FileSizeBytes = $sizeBytes
                FileSizeMB    = [Math]::Round($sizeBytes / 1MB, 4)
            })
    }
    return $result
}
#endregion Function: Get-TopFilesBySize

#endregion Helper Functions

#region Main Execution

#region Startup and Connect
# Main script execution
$script:startTime = Get-Date
Write-Log "Script started at $($script:startTime)"
Write-Log "Debug mode: $debug"
Write-Log "Get top files by size: $getFiles"
Write-Log "Output format: CSV"
Write-Log "Targeting library: PreservationHoldLibrary"
Write-Log "Output will be saved to: $outputFilePath"

$siteURLs = Read-SiteURLs -filePath $inputFilePath
Write-Log "Found $($siteURLs.Count) sites to process"

# Connect to Microsoft Graph once for all sites
if (-not (Connect-GraphAPI)) {
    Write-Log "Failed to connect to Microsoft Graph. Exiting..." "ERROR"
    exit 1
}
#endregion Startup and Connect

#region Process Sites
foreach ($siteURL in $siteURLs) {
    $siteStartTime = Get-Date
    Write-Log "Starting processing for site: $siteURL" "INFO"

    # Get site ID using Graph API
    $siteId = Get-SiteIdFromUrl -siteUrl $siteURL
    if (-not $siteId) {
        Write-Log "Failed to get site ID for $siteURL, skipping..." "ERROR"
        $global:reportData.Add([PSCustomObject]@{
                SiteURL        = $siteURL
                LibraryFound   = $false
                FileCount      = 0
                TotalSizeBytes = 0
                TotalSizeMB    = 0
                TotalSizeGB    = 0
                ProcessingTime = ((Get-Date) - $siteStartTime).ToString()
                Notes          = "Failed to resolve site ID"
            })
        continue
    }

    if ($debug) {
        Write-Log "DEBUG - Successfully got site ID: $siteId for URL: $siteURL" "INFO"
    }

    try {
        # Look specifically for PreservationHoldLibrary
        $allLists = Invoke-WithRetry -ScriptBlock {
            return Get-MgSiteList -SiteId $siteId -All
        }

        $holdLibrary = $allLists | Where-Object { $_.DisplayName -eq "PreservationHoldLibrary" -or $_.DisplayName -eq "Preservation Hold Library" -or $_.Name -eq "PreservationHoldLibrary" } | Select-Object -First 1

        if (-not $holdLibrary) {
            Write-Log "PreservationHoldLibrary not found on site $siteURL" "WARNING"
            $global:reportData.Add([PSCustomObject]@{
                    SiteURL        = $siteURL
                    LibraryFound   = $false
                    FileCount      = 0
                    TotalSizeBytes = 0
                    TotalSizeMB    = 0
                    TotalSizeGB    = 0
                    ProcessingTime = ((Get-Date) - $siteStartTime).ToString()
                    Notes          = "PreservationHoldLibrary not found"
                })
            continue
        }

        Write-Host "Found PreservationHoldLibrary on site $siteURL. Enumerating items..." -ForegroundColor Cyan
        Write-Log "Found PreservationHoldLibrary on site $siteURL. Getting drive items..." "INFO"

        # Get the drive ID for the library
        $driveId = Get-LibraryDriveId -siteId $siteId -list $holdLibrary

        if ([string]::IsNullOrEmpty($driveId)) {
            Write-Log "Could not get drive ID for PreservationHoldLibrary on $siteURL" "ERROR"
            $global:reportData.Add([PSCustomObject]@{
                    SiteURL        = $siteURL
                    LibraryFound   = $true
                    FileCount      = 0
                    TotalSizeBytes = 0
                    TotalSizeMB    = 0
                    TotalSizeGB    = 0
                    ProcessingTime = ((Get-Date) - $siteStartTime).ToString()
                    Notes          = "Could not resolve drive ID"
                })
            continue
        }

        # Use the drive delta API to enumerate ALL items recursively in a single flat call.
        # This bypasses folder traversal entirely and avoids 422 errors from non-standard items.
        Write-Log "Enumerating all items in PreservationHoldLibrary via delta API..." "INFO"
        $allItems = [System.Collections.Generic.List[object]]::new()
        $deltaUri = "https://graph.microsoft.com/v1.0/drives/$driveId/root/delta?`$select=id,name,size,file,folder,deleted"
        $pageNumber = 0

        try {
            do {
                $pageNumber++
                $response = $null
                # Wrap each page fetch in retry logic to handle mid-pagination throttling
                $currentUri = $deltaUri
                $response = Invoke-WithRetry -ScriptBlock {
                    return Invoke-MgGraphRequest -Uri $currentUri -Method GET -OutputType PSObject
                }
                if ($response.value) {
                    foreach ($item in $response.value) {
                        $allItems.Add($item)
                    }
                }
                $deltaUri = $response.'@odata.nextLink'
                # Report progress every 5 pages (~500 items at default page size)
                if ($pageNumber % 5 -eq 0 -or -not $deltaUri) {
                    Write-Host "  ...retrieved $($allItems.Count) items so far (page $pageNumber)" -ForegroundColor DarkCyan
                }
            } while ($deltaUri)

            Write-Log "Delta API enumeration complete: $($allItems.Count) total items across $pageNumber page(s)" "INFO"
        }
        catch {
            Write-Log "Delta API failed: $($_.Exception.Message). Falling back to list items API..." "WARNING"

            # Fallback: use Get-MgSiteListItem with field expansion
            $allItems = [System.Collections.Generic.List[object]]::new()
            $rawListItems = Invoke-WithRetry -ScriptBlock {
                return Get-MgSiteListItem -SiteId $siteId -ListId $holdLibrary.Id -All -ExpandProperty "driveItem"
            }
            foreach ($li in $rawListItems) {
                if ($li.DriveItem -and -not $li.DriveItem.Folder) {
                    $allItems.Add([PSCustomObject]@{
                            id      = $li.Id
                            name    = $li.DriveItem.Name
                            size    = $li.DriveItem.Size
                            file    = $true
                            folder  = $null
                            deleted = $null
                        })
                }
            }
            if ($debug) {
                Write-Log "DEBUG - Fallback list items API returned $($allItems.Count) file items" "INFO"
            }
        }

        # Count only files: not deleted, no folder property, has file property
        $files = $allItems | Where-Object {
            -not $_.deleted -and
            $null -eq $_.folder -and
            $null -ne $_.file
        }
        $fileCount = $files.Count
        $totalBytes = 0
        foreach ($f in $files) {
            if ($null -ne $f.size) { $totalBytes += [long]$f.size }
        }
        $totalMB = [Math]::Round($totalBytes / 1MB, 2)
        $totalGB = [Math]::Round($totalBytes / 1GB, 4)

        if ($getFiles -ne "none") {
            $topParam = if ($getFiles -eq "all") { [int]::MaxValue } else { 100 }
            $siteFiles = Get-TopFilesBySize -Items $allItems -SiteURL $siteURL -Top $topParam
            foreach ($tf in $siteFiles) { $global:topFilesData.Add($tf) }
            $fileLabel = if ($getFiles -eq "all") { "all $($siteFiles.Count)" } else { "top $($siteFiles.Count)" }
            Write-Log "Collected $fileLabel file(s) for $siteURL" "INFO"
        }

        Write-Log "Site: $siteURL | Files: $fileCount | Size: $totalMB MB ($totalGB GB)" "SUCCESS"

        $global:reportData.Add([PSCustomObject]@{
                SiteURL        = $siteURL
                LibraryFound   = $true
                FileCount      = $fileCount
                TotalSizeBytes = $totalBytes
                TotalSizeMB    = $totalMB
                TotalSizeGB    = $totalGB
                ProcessingTime = ((Get-Date) - $siteStartTime).ToString()
                Notes          = ""
            })
    }
    catch {
        Write-Log "Failed to process PreservationHoldLibrary for site $siteURL. Error: $($_.Exception.Message)" "ERROR"
        $global:reportData.Add([PSCustomObject]@{
                SiteURL        = $siteURL
                LibraryFound   = $false
                FileCount      = 0
                TotalSizeBytes = 0
                TotalSizeMB    = 0
                TotalSizeGB    = 0
                ProcessingTime = ((Get-Date) - $siteStartTime).ToString()
                Notes          = "Error: $($_.Exception.Message)"
            })
    }
}
#endregion Process Sites

#region Write Results
# Write report
Write-ReportToFile -Data $global:reportData -FilePath $outputFilePath

if ($getFiles -ne "none") {
    if ($global:topFilesData.Count -gt 0) {
        try {
            $global:topFilesData | Export-Csv -Path $topFilesOutputPath -NoTypeInformation -Encoding UTF8
            Write-Log "Files CSV created: $topFilesOutputPath" "SUCCESS"
        }
        catch {
            Write-Log "Failed to write files report: $($_.Exception.Message)" "ERROR"
        }
    }
    else {
        Write-Log "No file data collected (getFiles was '$getFiles' but no files were found)." "WARNING"
    }
}
#endregion Write Results

#region Cleanup
# Final log output
$totalTime = (Get-Date) - $script:startTime
Write-Log "Scan completed. Sites processed: $($global:reportData.Count). Total processing time: $totalTime" "SUCCESS"

# Disconnect from Microsoft Graph
try {
    Disconnect-MgGraph | Out-Null
    Write-Log "Disconnected from Microsoft Graph" "INFO"
}
catch {
    Write-Log "Error disconnecting from Microsoft Graph: $($_.Exception.Message)" "WARNING"
}
#endregion Cleanup

#endregion Main Execution
