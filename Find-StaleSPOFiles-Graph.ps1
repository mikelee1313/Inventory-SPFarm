<#
.SYNOPSIS
    Scans SharePoint Online sites to identify all files and folders that have NOT been modified in the specified number of months (stale files).

.DESCRIPTION
    This script connects to SharePoint Online using Microsoft Graph API with provided tenant-level credentials and iterates through a list of 
    site URLs specified in an input file. It recursively scans document libraries and lists (excluding specified folders) 
    to locate all files and folders that have NOT been modified within the specified number of months from today 
    (i.e., files that are considered "stale" or inactive), and outputs their details including
    site URL, item type, library name, item path, item name, creator, created date, and modified date.
    The script logs its operations and outputs the results to an Excel file using the ImportExcel module.

.PARAMETER None
    This script does not accept parameters via the command line. Configuration is done within the script.

.INPUTS
    A text file containing SharePoint site URLs to scan (path specified in $inputFilePath variable).

.OUTPUTS
    - CSV format: A CSV file containing all found stale files (path: $env:TEMP\Stale_Files_Report_[timestamp].csv)
      Plus separate summary files: *_Summary.csv and *_SiteSummary.csv
    - XLSX format: An Excel file containing all found stale files with multiple worksheets (path: $env:TEMP\Stale_Files_Report_[timestamp].xlsx)
    - A log file documenting the script's execution (path: $env:TEMP\Stale_Files_Report_[timestamp].txt)

.NOTES
    File Name      : Find-StaleSPOFiles-Graph.ps1
    Author         : Mike Lee
    Date Created   : 7/28/2025

    The script uses app-only authentication with a certificate thumbprint and Microsoft Graph API. Make sure the app has
    proper permissions in your tenant (Sites.Read.All or Sites.ReadWrite.All is recommended).

    The script ignores several system folders and lists to improve performance and avoid errors.

    PREREQUISITES:
    - Install-Module ImportExcel -Scope CurrentUser
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
    .\Find-StaleSPOFiles.ps1 # Assuming you rename the script
    Executes the script with the configured settings. Ensure you've updated the variables at the top
    of the script (appID, thumbprint, tenant, inputFilePath, monthsBack, outputFormat, and debug) before running.
    
    Example configurations:
    - $monthsBack = 3   # Files NOT modified in the last 3 months (stale for 3+ months)
    - $monthsBack = 6   # Files NOT modified in the last 6 months (stale for 6+ months)
    - $monthsBack = 12  # Files NOT modified in the last 12 months (stale for 1+ year)
    - $monthsBack = 24  # Files NOT modified in the last 24 months (stale for 2+ years)
    - $outputFormat = "CSV"   # For CSV output
    - $outputFormat = "XLSX"  # For Excel output
    - $debug = $true    # Enable detailed debug output
    - $debug = $false   # Enable informational output only (default)
#>

# =================================================================================================
# USER CONFIGURATION - Update the variables in this section
# =================================================================================================

# --- Tenant and App Registration Details ---
$appID = "5baa1427-1e90-4501-831d-a8e67465f0d9"                 # This is your Entra App ID
$thumbprint = "B696FDCFE1453F3FBC6031F54DE988DA0ED905A9"        # This is certificate thumbprint
$tenant = "85612ccb-4c28-4a34-88df-a538cc139a51"                # This is your Tenant ID

# --- Input File Path ---
$inputFilePath = 'C:\temp\SPOSiteList.txt' # Path to the input file containing site URLs

# --- Target Date Filter ---
$monthsBack = 12  # Number of months to look back from today (finds files NOT modified within this timeframe)

# --- Script Behavior Settings ---
$batchSize = 100  # How many items to process before writing to Excel
$maxItemsPerSheet = 5000 # Maximum items per sheet in Excel
$outputFormat = "xlsx"  # Output format: "CSV" or "XLSX"
$debug = $true  # Enable debug output: $true for detailed debug info, $false for informational only

# =================================================================================================
# END OF USER CONFIGURATION
# =================================================================================================

# Calculate target date based on months back from today
try {
    $today = Get-Date
    # Calculate the cutoff date - files older than this are stale
    $cutoffDate = $today.AddMonths(-$monthsBack)
    # Set to the last day of the target month to be more inclusive
    $targetDateParsed = New-Object DateTime($cutoffDate.Year, $cutoffDate.Month, [DateTime]::DaysInMonth($cutoffDate.Year, $cutoffDate.Month))
    $targetDate = $targetDateParsed.ToString("yyyy-MM-dd")
    
    if ($debug) {
        Write-Host "=== DATE CALCULATION DEBUG ===" -ForegroundColor Magenta
        Write-Host "Today's date: $($today.ToString('yyyy-MM-dd dddd'))" -ForegroundColor Cyan
        Write-Host "Months back: $monthsBack" -ForegroundColor Cyan
        Write-Host "Target date calculated: $targetDate (last day of target month)" -ForegroundColor Green
        Write-Host "Files modified on or before $($targetDateParsed.ToString('yyyy-MM-dd dddd')) will be included (stale files)" -ForegroundColor Green
        Write-Host "Example: Files from January 2025 should be INCLUDED (stale)" -ForegroundColor Yellow
        Write-Host "Example: Files from February 2025 and later should be EXCLUDED (too recent)" -ForegroundColor Yellow
        Write-Host "================================" -ForegroundColor Magenta
    }
    else {
        Write-Host "Target date: $targetDate (files modified on or before this date will be included)" -ForegroundColor Green
    }
}
catch {
    Write-Host "ERROR: Invalid monthsBack value. Please use a positive integer (e.g., 1, 6, 12)" -ForegroundColor Red
    exit 1
}

# Check for required modules
$requiredModules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Sites', 'Microsoft.Graph.Files')
if ($outputFormat -eq "XLSX") {
    $requiredModules += 'ImportExcel'
}

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
$requiredCmdlets = @('Get-MgSite', 'Get-MgSiteList', 'Get-MgSiteListItem', 'Get-MgDriveItem', 'Get-MgDriveItemChild', 'Get-MgSiteDrive', 'Connect-MgGraph')
foreach ($cmdlet in $requiredCmdlets) {
    if (-not (Get-Command $cmdlet -ErrorAction SilentlyContinue)) {
        Write-Host "ERROR: Required cmdlet '$cmdlet' is not available. Please ensure all Microsoft Graph modules are properly installed." -ForegroundColor Red
        exit 1
    }
}
Write-Host "All required cmdlets are available." -ForegroundColor Green

# Script Parameters
Add-Type -AssemblyName System.Web
$startime = Get-Date -Format "yyyyMMdd_HHmmss"
$logFilePath = "$env:TEMP\Stale_Files_Report_$startime.txt"

# Set output file path based on format
if ($outputFormat -eq "CSV") {
    $outputFilePath = "$env:TEMP\Stale_Files_Report_$startime.csv"
    $fileExtension = "csv"
}
else {
    $outputFilePath = "$env:TEMP\Stale_Files_Report_$startime.xlsx"
    $fileExtension = "xlsx"
}

# Initialize collections for batch processing
$global:currentBatch = @()
$global:totalItemsProcessed = 0
$global:currentSheetNumber = 1
$global:summaryData = @()
$global:excelFileInitialized = $false

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
                    $errorMessage -match "Access denied" -or $errorMessage -match "403") {
                    $logLevel = "INFO" # Downgrade to INFO for these specific cases
                }
                Write-Log "General Error occurred During retrieval : $errorMessage" $logLevel
                throw $_
            }
        }
    }
    
    return $result
}

# Read site URLs from input file
function Read-SiteURLs {
    param (
        [string]$filePath
    )
    $urls = Get-Content -Path $filePath
    return $urls
}

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

# Helper function to get all document libraries and lists from a site
function Get-SiteLists {
    param (
        [string]$siteId
    )
    
    try {
        $lists = Invoke-WithRetry -ScriptBlock {
            return Get-MgSiteList -SiteId $siteId -All
        }
        
        if ($debug -and $lists) {
            Write-Log "DEBUG - Retrieved $($lists.Count) total lists from site" "INFO"
            # Avoid deep object serialization that causes JSON truncation warnings
            $sampleList = $lists | Select-Object -First 1
            if ($sampleList) {
                Write-Log "DEBUG - Sample list: Name='$($sampleList.DisplayName)', Template='$($sampleList.List.Template)', Hidden='$($sampleList.Hidden)'" "INFO"
            }
        }
        
        # Filter out hidden lists and ignored folders
        # Note: Graph API may use different property structures than PnP
        $filteredLists = $lists | Where-Object { 
            # Check if list is not hidden (some lists might not have Hidden property)
            ($null -eq $_.Hidden -or $_.Hidden -eq $false) -and 
            # Check display name is not in ignore list
            $_.DisplayName -notin $ignoreFolders -and
            # Check if it's a document library or generic list
            # Graph API might use different template names - be more flexible
            ($_.List.Template -eq "documentLibrary" -or 
            $_.List.Template -eq "genericList" -or
            $_.List.Template -eq "101" -or # Document Library template ID
            $_.List.Template -eq "100" -or # Generic List template ID
            $_.List.Template -eq 101 -or # Numeric template IDs
            $_.List.Template -eq 100 -or
            $_.Drive -ne $null -or # Has an associated drive (likely a document library)
            $_.List.BaseTemplate -eq 101 -or # Alternative property name
            $_.List.BaseTemplate -eq 100)   # Alternative property name
        }
        
        if ($debug) {
            Write-Log "DEBUG - After filtering: $($filteredLists.Count) lists remain" "INFO"
            if ($filteredLists.Count -gt 0) {
                Write-Log "DEBUG - Filtered list names: $(($filteredLists | ForEach-Object { $_.DisplayName }) -join ', ')" "INFO"
            }
            else {
                Write-Log "DEBUG - All lists were filtered out. Showing all list details for troubleshooting:" "INFO"
                foreach ($list in $lists) {
                    $templateInfo = "Template: $($list.List.Template)"
                    if ($list.List.BaseTemplate) {
                        $templateInfo += " | BaseTemplate: $($list.List.BaseTemplate)"
                    }
                    Write-Log "DEBUG - List: '$($list.DisplayName)' | Hidden: $($list.Hidden) | $templateInfo | Has Drive: $($null -ne $list.Drive)" "INFO"
                }
            }
        }
        
        return $filteredLists
    }
    catch {
        Write-Log "Failed to get lists for site ID $siteId : $($_.Exception.Message)" "ERROR"
        return @()
    }
}

# Helper function to get all items from a list
function Get-ListItems {
    param (
        [string]$siteId,
        [string]$listId
    )
    
    try {
        $items = Invoke-WithRetry -ScriptBlock {
            # Remove the invalid expand properties for list items
            return Get-MgSiteListItem -SiteId $siteId -ListId $listId -All -ExpandProperty "fields"
        }
        
        return $items
    }
    catch {
        Write-Log "Failed to get items for list $listId : $($_.Exception.Message)" "ERROR"
        return @()
    }
}

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

# Helper function to get drive items (for document libraries)
function Get-DriveItems {
    param (
        [string]$siteId,
        [string]$driveId
    )
    
    try {
        # Validate drive ID is not empty
        if ([string]::IsNullOrEmpty($driveId)) {
            Write-Log "Drive ID is empty or null, cannot retrieve drive items" "ERROR"
            return @()
        }
        
        $items = @()
        
        if ($debug) {
            Write-Log "DEBUG - Getting drive items for drive ID: $driveId" "INFO"
        }
        
        # Get root items first using the correct cmdlet
        $rootItems = Invoke-WithRetry -ScriptBlock {
            # For SharePoint sites, we need to get the root folder first, then its children
            try {
                # First try to get root folder children directly
                return Get-MgDriveItemChild -DriveId $driveId -DriveItemId "root" -All
            }
            catch {
                # If that fails, try getting the root item first
                $rootItem = Get-MgDriveItem -DriveId $driveId -DriveItemId "root"
                if ($rootItem) {
                    return Get-MgDriveItemChild -DriveId $driveId -DriveItemId $rootItem.Id -All
                }
                else {
                    throw $_
                }
            }
        }
        
        if ($debug) {
            Write-Log "DEBUG - Retrieved $($rootItems.Count) root items from drive" "INFO"
        }
        
        if ($rootItems) {
            foreach ($item in $rootItems) {
                $items += $item
                
                # If it's a folder, get its children recursively
                # BUT first validate it's actually a folder and not a file with incorrect folder flag
                if ($item.Folder -and $item.Id) {
                    # Check if this is actually a file masquerading as a folder
                    $isActuallyFile = $false
                    if ($item.Name) {
                        $fileName = $item.Name.ToLower()
                        # Check for common file extensions that shouldn't be folders
                        $fileExtensions = @('.docx', '.xlsx', '.pptx', '.pdf', '.txt', '.csv', '.doc', '.xls', '.ppt', '.zip', '.jpg', '.png', '.gif', '.mp4', '.mp3', '.agent')
                        foreach ($ext in $fileExtensions) {
                            if ($fileName.EndsWith($ext)) {
                                $isActuallyFile = $true
                                break
                            }
                        }
                    }
                    
                    if ($isActuallyFile) {
                        if ($debug) {
                            Write-Log "DEBUG - Item '$($item.Name)' appears to be a file with incorrect folder flag, skipping recursive processing" "INFO"
                        }6
                    }
                    else {
                        if ($debug) {
                            Write-Log "DEBUG - Processing folder: $($item.Name) with ID: $($item.Id)" "INFO"
                        }
                        try {
                            $childItems = Get-DriveItemsRecursive -driveId $driveId -itemId $item.Id
                            if ($childItems) {
                                $items += $childItems
                            }
                        }
                        catch {
                            Write-Log "Error getting children for folder '$($item.Name)' (ID: $($item.Id)): $($_.Exception.Message)" "WARNING"
                            # Continue processing other items even if this folder fails
                        }
                    }
                }
            }
        }
        
        return $items
    }
    catch {
        Write-Log "Failed to get drive items for drive $driveId : $($_.Exception.Message)" "ERROR"
        return @()
    }
}

# Helper function to recursively get drive items
function Get-DriveItemsRecursive {
    param (
        [string]$driveId,
        [string]$itemId
    )
    
    try {
        # Validate input parameters
        if ([string]::IsNullOrEmpty($driveId)) {
            Write-Log "Drive ID is empty in Get-DriveItemsRecursive, skipping" "WARNING"
            return @()
        }
        
        if ([string]::IsNullOrEmpty($itemId)) {
            Write-Log "Item ID is empty in Get-DriveItemsRecursive, skipping" "WARNING"
            return @()
        }
        
        # Additional validation for potentially problematic IDs
        if ($itemId.Contains(" ") -or $itemId.Length -lt 10) {
            Write-Log "Item ID appears invalid or too short: '$itemId', skipping recursive processing" "WARNING"
            return @()
        }
        
        $items = @()
        
        if ($debug) {
            Write-Log "DEBUG - Getting child items for drive ID: $driveId, item ID: $itemId" "INFO"
        }
        
        $childItems = Invoke-WithRetry -ScriptBlock {
            # Additional validation before making the API call
            if ([string]::IsNullOrWhiteSpace($driveId)) {
                throw "Drive ID is null or whitespace"
            }
            if ([string]::IsNullOrWhiteSpace($itemId)) {
                throw "Item ID is null or whitespace"
            }
            
            if ($debug) {
                Write-Log "DEBUG - About to call Get-MgDriveItemChild with DriveId='$driveId', ItemId='$itemId'" "INFO"
            }
            
            return Get-MgDriveItemChild -DriveId $driveId -DriveItemId $itemId -All
        }
        
        if ($debug) {
            $childCount = if ($childItems) { $childItems.Count } else { 0 }
            Write-Log "DEBUG - Retrieved $childCount child items" "INFO"
        }
        
        if ($childItems) {
            foreach ($childItem in $childItems) {
                $items += $childItem
                
                # If it's a folder, get its children recursively
                # BUT first validate it's actually a folder and not a file with incorrect folder flag
                if ($childItem.Folder -and $childItem.Id) {
                    # Check if this is actually a file masquerading as a folder
                    $isActuallyFile = $false
                    if ($childItem.Name) {
                        $fileName = $childItem.Name.ToLower()
                        # Check for common file extensions that shouldn't be folders
                        $fileExtensions = @('.docx', '.xlsx', '.pptx', '.pdf', '.txt', '.csv', '.doc', '.xls', '.ppt', '.zip', '.jpg', '.png', '.gif', '.mp4', '.mp3', '.agent')
                        foreach ($ext in $fileExtensions) {
                            if ($fileName.EndsWith($ext)) {
                                $isActuallyFile = $true
                                break
                            }
                        }
                    }
                    
                    if ($isActuallyFile) {
                        if ($debug) {
                            Write-Log "DEBUG - Child item '$($childItem.Name)' appears to be a file with incorrect folder flag, skipping recursive processing" "INFO"
                        }
                    }
                    else {
                        try {
                            $grandChildItems = Get-DriveItemsRecursive -driveId $driveId -itemId $childItem.Id
                            if ($grandChildItems) {
                                $items += $grandChildItems
                            }
                        }
                        catch {
                            Write-Log "Error getting grandchildren for folder '$($childItem.Name)' (ID: $($childItem.Id)): $($_.Exception.Message)" "WARNING"
                            # Continue processing other items even if this folder fails
                        }
                    }
                }
            }
        }
        
        return $items
    }
    catch {
        Write-Log "Failed to get child items for item $itemId : $($_.Exception.Message)" "WARNING"
        return @()
    }
}

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

# List of folders to ignore
$ignoreFolders = @(
    "_catalogs",
    "appdata",
    "forms",
    "Form Templates",
    "Site Assets",
    "List Template Gallery",
    "Master Page Gallery",
    "Solution Gallery",
    "Style Library",
    "Composed Looks",
    "Converted Forms",
    "Web Part Gallery",
    "Theme Gallery",
    "TaxonomyHiddenList",
    "Events",
    "_cts",
    "_private",
    "_vti_pvt",
    "Reference 778a30bb4f074ae3bec315889ee34b88",
    "Sharing Links",
    "Social",
    "FavoriteLists-e0157a47-72e4-43c1-bfd0-ed9f7040e894",
    "User Information List",
    "Web Template Extensions",
    "SmartCache-8189C6B3-4081-4F62-9015-35FDB7FDF042",
    "SharePointHomeCacheList",
    "RecentLists-56BAEAB4-E7AD-4E59-B92B-9290D871F5C3",
    "PersonalCacheLibrary",
    "microsoft.ListSync.Endpoints",
    "Maintenance Log Library",
    "DO_NOT_DELETE_ENTERPRISE_USER_CONTAINER_ENUM_LIST",
    "appfiles"
)

# Function to write batch data to file (CSV or Excel)
function Write-BatchToFile {
    param (
        [array]$Data,
        [string]$FilePath,
        [int]$SheetNumber = 1
    )
    
    if ($Data.Count -eq 0) { return }
    
    try {
        if ($outputFormat -eq "CSV") {
            # CSV Export
            if (-not (Test-Path $FilePath)) {
                # First write - include headers
                $Data | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8
                Write-Log "Created CSV file and wrote $($Data.Count) items to: $FilePath" "SUCCESS"
            }
            else {
                # Append mode - no headers
                $Data | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8 -Append
                Write-Log "Appended $($Data.Count) items to CSV file: $FilePath" "SUCCESS"
            }
        }
        else {
            # Excel Export (existing logic)
            $worksheetName = "Stale_Files_$SheetNumber"
            
            # Define Excel table style for better readability
            $excelParams = @{
                Path          = $FilePath
                WorksheetName = $worksheetName
                TableName     = "StaleFilesTable$SheetNumber"
                TableStyle    = 'Medium6'
                AutoSize      = $true
                FreezeTopRow  = $true
                BoldTopRow    = $true
            }
            
            # Add conditional formatting for item types
            $conditionalFormatting = @(
                New-ConditionalText -Text 'File' -BackgroundColor LightBlue -ConditionalTextColor Black
                New-ConditionalText -Text 'Folder' -BackgroundColor LightYellow -ConditionalTextColor Black
            )
            
            # Export data to Excel. Create the file on first write, append on subsequent writes.
            if (-not $global:excelFileInitialized) {
                $Data | Export-Excel @excelParams -ConditionalText $conditionalFormatting
                [void]($global:excelFileInitialized = $true)
            }
            else {
                $Data | Export-Excel @excelParams -ConditionalText $conditionalFormatting -Append
            }
            
            Write-Log "Successfully wrote $($Data.Count) items to worksheet: $worksheetName" "SUCCESS"
        }
    }
    catch {
        Write-Log "Failed to write batch to file: $($_.Exception.Message)" "ERROR"
        throw
    }
}

# Modified function to handle batch processing
function Add-ItemToBatch {
    param (
        [PSCustomObject]$Item
    )
    
    [void]($global:currentBatch += $Item)
    $global:totalItemsProcessed++
    
    # Check if we need to write the batch
    if ($global:currentBatch.Count -ge $batchSize) {
        Write-BatchToFile -Data $global:currentBatch -FilePath $outputFilePath -SheetNumber $global:currentSheetNumber
        $global:currentBatch = @()
        
        # Check if we need a new sheet (only relevant for Excel)
        if ($outputFormat -eq "XLSX") {
            $itemsInCurrentSheet = ($global:totalItemsProcessed % $maxItemsPerSheet)
            if ($itemsInCurrentSheet -eq 0) {
                $global:currentSheetNumber++
            }
        }
    }
    
    # Update progress every 10 items
    if ($global:totalItemsProcessed % 10 -eq 0) {
        Write-Host "Processed $global:totalItemsProcessed items..." -ForegroundColor Yellow
    }
}

# Process SharePoint Item (File or Folder) - Modified for Graph API and date filtering
function Get-SPItemInfo {
    param (
        $item,
        [string]$ItemSiteURL,
        [string]$ItemType, # "File" or "Folder"
        [string]$LibraryName,
        [string]$ItemSource = "Drive" # "Drive" for drive items, "List" for list items
    )
    try {
        # Handle different item types from Graph API
        if ($ItemSource -eq "Drive") {
            # Drive item from document library
            $itemName = $item.Name
            $itemPath = $item.WebUrl
            if ([string]::IsNullOrEmpty($itemPath)) {
                $itemPath = $item.ParentReference.Path + "/" + $item.Name
            }
            
            # Get modified date from drive item
            $modifiedDateTime = $null
            $modifiedDateStr = "Unknown"
            
            if ($item.LastModifiedDateTime) {
                $modifiedDateTime = [DateTime]$item.LastModifiedDateTime
                $modifiedDateStr = $modifiedDateTime.ToString("yyyy-MM-dd HH:mm:ss")
            }
            
            # Get created date and creator from drive item
            $createdDateTime = $null
            $creatorName = "Unknown"
            $creatorEmail = "Unknown"
            $creatorWithEmail = "Unknown"
            
            if ($item.CreatedDateTime) {
                $createdDateTime = [DateTime]$item.CreatedDateTime
            }
            
            if ($item.CreatedBy -and $item.CreatedBy.User) {
                $creatorName = $item.CreatedBy.User.DisplayName
                $creatorEmail = $item.CreatedBy.User.Email
                if ([string]::IsNullOrEmpty($creatorEmail)) {
                    $creatorWithEmail = $creatorName
                }
                else {
                    $creatorWithEmail = "$creatorName ($creatorEmail)"
                }
            }
        }
        else {
            # List item
            $itemName = if ($item.Fields.FileLeafRef) { $item.Fields.FileLeafRef } else { $item.Fields.Title }
            $itemPath = if ($item.Fields.FileRef) { $item.Fields.FileRef } else { $item.WebUrl }
            
            # Get modified date from list item
            $modifiedDateTime = $null
            $modifiedDateStr = "Unknown"
            
            if ($item.LastModifiedDateTime) {
                $modifiedDateTime = [DateTime]$item.LastModifiedDateTime
                $modifiedDateStr = $modifiedDateTime.ToString("yyyy-MM-dd HH:mm:ss")
            }
            elseif ($item.Fields.Modified) {
                try {
                    $modifiedDateTime = [DateTime]$item.Fields.Modified
                    $modifiedDateStr = $modifiedDateTime.ToString("yyyy-MM-dd HH:mm:ss")
                }
                catch {
                    if ($debug) {
                        Write-Log "DEBUG - Could not parse Modified field: $($item.Fields.Modified)" "INFO"
                    }
                }
            }
            
            # Get created date and creator from list item fields
            $createdDateTime = $null
            $creatorName = "Unknown"
            $creatorEmail = "Unknown"
            $creatorWithEmail = "Unknown"
            
            if ($item.CreatedDateTime) {
                $createdDateTime = [DateTime]$item.CreatedDateTime
            }
            elseif ($item.Fields.Created) {
                try {
                    $createdDateTime = [DateTime]$item.Fields.Created
                }
                catch {
                    if ($debug) {
                        Write-Log "DEBUG - Could not parse Created field: $($item.Fields.Created)" "INFO"
                    }
                }
            }
            
            # Try to get creator information from various field sources
            if ($item.Fields.Author) {
                $creatorName = $item.Fields.Author
                $creatorWithEmail = $creatorName
            }
            elseif ($item.CreatedBy -and $item.CreatedBy.User) {
                $creatorName = $item.CreatedBy.User.DisplayName
                $creatorEmail = $item.CreatedBy.User.Email
                if ([string]::IsNullOrEmpty($creatorEmail)) {
                    $creatorWithEmail = $creatorName
                }
                else {
                    $creatorWithEmail = "$creatorName ($creatorEmail)"
                }
            }
        }
        
        # Check if we have a valid modified date
        if ($null -eq $modifiedDateTime) {
            Write-Log "No modified date found for item $itemPath, skipping" "INFO"
            return $false
        }
        
        # Check if the modified date is BEFORE or EQUAL to the target date (files that are stale)
        $itemModifiedDate = $modifiedDateTime.Date
        if ($debug) {
            Write-Host "DEBUG: Checking file $itemName - Modified: $($itemModifiedDate.ToString('yyyy-MM-dd')) vs Target: $($targetDateParsed.Date.ToString('yyyy-MM-dd'))" -ForegroundColor Magenta
        }
        
        if ($itemModifiedDate -le $targetDateParsed.Date) {
            Write-Log "INCLUDING stale item $itemPath - Modified date ($($itemModifiedDate.ToString('yyyy-MM-dd'))) is on or before target date $targetDate" "SUCCESS"
            if ($debug) {
                Write-Host "INCLUDED: $itemName (Modified: $($itemModifiedDate.ToString('yyyy-MM-dd')))" -ForegroundColor Green
            }
        }
        else {
            if ($debug) {
                Write-Log "EXCLUDING item $itemPath - Modified date ($($itemModifiedDate.ToString('yyyy-MM-dd'))) is after target date $targetDate (not stale)" "INFO"
                Write-Host "EXCLUDED: $itemName (Modified: $($itemModifiedDate.ToString('yyyy-MM-dd'))) - too recent" -ForegroundColor Red
            }
            return $false
        }
        
        Write-Log "Processing item: $itemPath (Type: $ItemType, Modified: $modifiedDateStr)" "INFO"
        
        # Create an entry for the item
        $itemEntry = [PSCustomObject]@{
            SiteURL      = $ItemSiteURL
            ItemType     = $ItemType
            LibraryName  = $LibraryName
            ItemPath     = $itemPath 
            ItemName     = $itemName
            CreatedBy    = $creatorWithEmail
            CreatedDate  = $createdDateTime
            ModifiedDate = $modifiedDateTime
        }
        
        Add-ItemToBatch -Item $itemEntry
        Write-Log "Added item to batch: $itemName (Modified: $modifiedDateStr)" "INFO"
        return $true
    }
    catch {
        $itemId = try { $item.Id } catch { "Unknown" }
        Write-Log "Failed to process $ItemType (ID: $itemId): $($_.Exception.Message)" "ERROR"
        Write-Log "Stack trace: $($_.ScriptStackTrace)" "ERROR"
        return $false
    }
}

# Function to create summary file (CSV or Excel)
function New-SummaryFile {
    param (
        [string]$FilePath
    )
    
    try {
        # Create summary data
        $summary = [PSCustomObject]@{
            'Months Back'             = $monthsBack
            'Target Date'             = $targetDate
            'Total Sites Processed'   = $global:summaryData.Count
            'Total Stale Files Found' = $global:totalItemsProcessed
            'Processing Start Time'   = $script:startTime
            'Processing End Time'     = Get-Date
            'Processing Duration'     = (Get-Date) - $script:startTime
        }
        
        if ($outputFormat -eq "CSV") {
            # For CSV, create separate summary files
            $summaryPath = $FilePath -replace '\.csv$', '_Summary.csv'
            $siteSummaryPath = $FilePath -replace '\.csv$', '_SiteSummary.csv'
            
            # Export main summary
            $summary | Export-Csv -Path $summaryPath -NoTypeInformation -Encoding UTF8
            Write-Log "Summary CSV created: $summaryPath" "SUCCESS"
            
            # Export site-level summary
            if ($global:summaryData.Count -gt 0) {
                $global:summaryData | Export-Csv -Path $siteSummaryPath -NoTypeInformation -Encoding UTF8
                Write-Log "Site Summary CSV created: $siteSummaryPath" "SUCCESS"
            }
        }
        else {
            # Excel format (existing logic)
            # Export summary to first worksheet
            $summary | Export-Excel -Path $FilePath -WorksheetName "Summary" -TableName "SummaryTable" -TableStyle 'Medium2' -AutoSize -MoveToStart
            
            # Add site-level summary
            if ($global:summaryData.Count -gt 0) {
                $global:summaryData | Export-Excel -Path $FilePath -WorksheetName "Site Summary" -TableName "SiteSummaryTable" -TableStyle 'Medium4' -AutoSize -FreezeTopRow -BoldTopRow
            }
            
            Write-Log "Summary worksheet created successfully" "SUCCESS"
        }
    }
    catch {
        Write-Log "Failed to create summary file: $($_.Exception.Message)" "ERROR"
    }
}

# Main script execution
$script:startTime = Get-Date
Write-Log "Script started at $($script:startTime)"
Write-Log "Debug mode: $debug"
Write-Log "Output format: $outputFormat"
Write-Log "Months back parameter: $monthsBack (target date: $targetDate)"
Write-Log "Including files NOT modified since: $targetDate (stale files)"
Write-Log "Output will be saved to: $outputFilePath"

$siteURLs = Read-SiteURLs -filePath $inputFilePath
Write-Log "Found $($siteURLs.Count) sites to process"

# Connect to Microsoft Graph once for all sites
if (-not (Connect-GraphAPI)) {
    Write-Log "Failed to connect to Microsoft Graph. Exiting..." "ERROR"
    exit 1
}

foreach ($siteURL in $siteURLs) {
    $siteStartTime = Get-Date
    Write-Log "Starting processing for site: $siteURL" "INFO"
    
    $siteItemCount = 0
    
    # Get site ID using Graph API
    $siteId = Get-SiteIdFromUrl -siteUrl $siteURL
    if (-not $siteId) {
        Write-Log "Failed to get site ID for $siteURL, skipping..." "ERROR"
        continue
    }
    
    if ($debug) {
        Write-Log "DEBUG - Successfully got site ID: $siteId for URL: $siteURL" "INFO"
    }
    
    try {
        # Get all lists and document libraries from the site
        $lists = Get-SiteLists -siteId $siteId
        
        if ($null -eq $lists -or $lists.Count -eq 0) {
            Write-Log "No lists retrieved or all lists were ignored for site $siteURL." "WARNING"
            Write-Log "This could be due to permissions, site structure, or filtering criteria." "WARNING"
            
            # If debug is enabled, try to get more information
            if ($debug) {
                try {
                    $allLists = Get-MgSiteList -SiteId $siteId -All
                    Write-Log "DEBUG - Total lists found before filtering: $($allLists.Count)" "INFO"
                    if ($allLists.Count -gt 0) {
                        Write-Log "DEBUG - All list names: $(($allLists | ForEach-Object { $_.DisplayName }) -join ', ')" "INFO"
                    }
                }
                catch {
                    Write-Log "DEBUG - Error getting all lists for debugging: $($_.Exception.Message)" "ERROR"
                }
            }
        }
        else {
            Write-Log "Found $($lists.Count) lists to process in site $siteURL"
            if ($debug) {
                Write-Log "Lists to process: $($lists | ForEach-Object { $_.DisplayName } | Join-String -Separator ', ')" "INFO"
            }
            
            foreach ($list in $lists) { 
                try {
                    $listName = $list.DisplayName
                    Write-Log "Processing list/library: '$listName' on site: $siteURL"
                    
                    # Check if this is a document library (has associated drive or is template 101)
                    $isDocumentLibrary = ($list.Drive -ne $null) -or 
                    ($list.List.Template -eq "documentLibrary") -or 
                    ($list.List.Template -eq "101") -or 
                    ($list.List.Template -eq 101) -or
                    ($list.List.BaseTemplate -eq 101)
                    
                    if ($isDocumentLibrary) {
                        # Process as document library using Drive API
                        Write-Log "Processing '$listName' as document library using Drive API"
                        
                        try {
                            # Get the drive ID using our helper function
                            $driveId = Get-LibraryDriveId -siteId $siteId -list $list
                            
                            if ([string]::IsNullOrEmpty($driveId)) {
                                Write-Log "Could not get drive ID for document library '$listName', trying as regular list instead" "WARNING"
                                # Fall back to processing as regular list
                                $isDocumentLibrary = $false
                            }
                            else {
                                $driveItems = Get-DriveItems -siteId $siteId -driveId $driveId
                                
                                if ($null -eq $driveItems -or $driveItems.Count -eq 0) {
                                    Write-Log "No items retrieved from document library '$listName'" "WARNING"
                                    continue
                                }
                            }
                        }
                        catch {
                            Write-Log "Error getting drive ID for '$listName': $($_.Exception.Message)" "ERROR"
                            Write-Log "Falling back to processing '$listName' as regular list" "WARNING"
                            $isDocumentLibrary = $false
                        }
                    }
                    
                    if ($isDocumentLibrary -and $driveItems) {
                        # Continue with document library processing
                        try {
                            
                            Write-Log "Retrieved $($driveItems.Count) items from document library '$listName'"
                            $itemsProcessedInList = 0
                            
                            foreach ($currentItem in $driveItems) {
                                try {
                                    # Determine item type
                                    $itemTypeStr = if ($currentItem.Folder) { "Folder" } else { "File" }
                                    
                                    # Check if item should be ignored based on path
                                    $currentItemPath = $currentItem.WebUrl
                                    if ([string]::IsNullOrEmpty($currentItemPath)) {
                                        $currentItemPath = $currentItem.ParentReference.Path + "/" + $currentItem.Name
                                    }
                                    
                                    # Check if item should be ignored
                                    [void]($ignoreCurrentItem = $false)
                                    foreach ($ignoreFolderPattern in $ignoreFolders) {
                                        if ($currentItemPath -like "*/$ignoreFolderPattern/*" -or $currentItemPath -match "/$($ignoreFolderPattern)$") {
                                            [void]($ignoreCurrentItem = $true)
                                            break
                                        }
                                    }
                                    
                                    if ($ignoreCurrentItem) {
                                        if ($debug) {
                                            Write-Log "Ignoring item: $currentItemPath" "INFO"
                                        }
                                        continue
                                    }
                                    
                                    # Process the item using Graph API data structure
                                    $itemWasAdded = Get-SPItemInfo -item $currentItem -ItemSiteURL $siteURL -ItemType $itemTypeStr -LibraryName $listName -ItemSource "Drive"
                                    if ($itemWasAdded) {
                                        $siteItemCount++
                                        $itemsProcessedInList++
                                    }
                                }
                                catch {
                                    Write-Log "Error processing individual drive item: $($_.Exception.Message)" "WARNING"
                                }
                            }
                            
                            Write-Log "Completed processing document library '$listName'. Items processed: $itemsProcessedInList"
                        }
                        catch {
                            Write-Log "Error retrieving items from document library '$listName': $($_.Exception.Message)" "ERROR"
                        }
                    }
                    else {
                        # Process as regular list using Lists API
                        Write-Log "Processing '$listName' as regular list using Lists API"
                        
                        try {
                            $listItems = Get-ListItems -siteId $siteId -listId $list.Id
                            
                            if ($null -eq $listItems -or $listItems.Count -eq 0) {
                                Write-Log "No items retrieved from list '$listName'" "WARNING"
                                continue
                            }
                            
                            Write-Log "Retrieved $($listItems.Count) items from list '$listName'"
                            $itemsProcessedInList = 0
                            
                            foreach ($currentItem in $listItems) {
                                try {
                                    # For list items, determine type based on content type or other properties
                                    $itemTypeStr = "List Item"
                                    
                                    # Check if item should be ignored based on path or name
                                    $currentItemPath = if ($currentItem.Fields.FileRef) { $currentItem.Fields.FileRef } else { $currentItem.WebUrl }
                                    $currentItemName = if ($currentItem.Fields.FileLeafRef) { $currentItem.Fields.FileLeafRef } else { $currentItem.Fields.Title }
                                    
                                    # Check if item should be ignored
                                    [void]($ignoreCurrentItem = $false)
                                    foreach ($ignoreFolderPattern in $ignoreFolders) {
                                        if ($currentItemPath -like "*/$ignoreFolderPattern/*" -or $currentItemPath -match "/$($ignoreFolderPattern)$" -or $currentItemName -eq $ignoreFolderPattern) {
                                            [void]($ignoreCurrentItem = $true)
                                            break
                                        }
                                    }
                                    
                                    if ($ignoreCurrentItem) {
                                        if ($debug) {
                                            Write-Log "Ignoring item: $currentItemPath" "INFO"
                                        }
                                        continue
                                    }
                                    
                                    # Process the item using Graph API data structure
                                    $itemWasAdded = Get-SPItemInfo -item $currentItem -ItemSiteURL $siteURL -ItemType $itemTypeStr -LibraryName $listName -ItemSource "List"
                                    if ($itemWasAdded) {
                                        $siteItemCount++
                                        $itemsProcessedInList++
                                    }
                                }
                                catch {
                                    Write-Log "Error processing individual list item: $($_.Exception.Message)" "WARNING"
                                }
                            }
                            
                            Write-Log "Completed processing list '$listName'. Items processed: $itemsProcessedInList"
                        }
                        catch {
                            Write-Log "Error retrieving items from list '$listName': $($_.Exception.Message)" "ERROR"
                        }
                    }
                }
                catch {
                    Write-Log "Failed to process list '$($list.DisplayName)' on site '$siteURL'. Error: $($_.Exception.Message)" "ERROR"
                }
            }
        }
    }
    catch {
        Write-Log "Failed to get lists for site $siteURL. Error: $($_.Exception.Message)" "ERROR"
    }
    
    # Add site summary data
    $siteSummary = [PSCustomObject]@{
        SiteURL         = $siteURL
        StaleFilesFound = $siteItemCount
        ProcessingTime  = ((Get-Date) - $siteStartTime).ToString()
    }
    [void]($global:summaryData += $siteSummary)
    
    Write-Log "Completed processing for $siteURL. Stale files found: $siteItemCount" "SUCCESS"
}

# Write any remaining items in the batch - THIS IS CRITICAL
Write-Log "Writing final batch of $($global:currentBatch.Count) items"
if ($global:currentBatch.Count -gt 0) {
    Write-BatchToFile -Data $global:currentBatch -FilePath $outputFilePath -SheetNumber $global:currentSheetNumber
    Write-Log "Final batch written successfully"
}

# Create summary file and ensure main output file exists
if ($global:totalItemsProcessed -gt 0 -or $global:summaryData.Count -gt 0) {
    New-SummaryFile -FilePath $outputFilePath
}
else {
    Write-Log "No stale files found (files not modified since $targetDate). All files appear to be recently active." "WARNING"
}

# Always ensure the main output file exists, even if empty
if (-not (Test-Path $outputFilePath)) {
    if ($outputFormat -eq "CSV") {
        # Create empty CSV file with headers
        $emptyData = [PSCustomObject]@{
            SiteURL      = $null
            ItemType     = $null
            LibraryName  = $null
            ItemPath     = $null
            ItemName     = $null
            CreatedBy    = $null
            CreatedDate  = $null
            ModifiedDate = $null
        }
        $emptyData | Export-Csv -Path $outputFilePath -NoTypeInformation -Encoding UTF8
        Write-Log "Created empty CSV file with headers: $outputFilePath" "INFO"
    }
    else {
        # Create empty Excel file
        $emptyData = [PSCustomObject]@{
            SiteURL      = $null
            ItemType     = $null
            LibraryName  = $null
            ItemPath     = $null
            ItemName     = $null
            CreatedBy    = $null
            CreatedDate  = $null
            ModifiedDate = $null
        }
        $emptyData | Export-Excel -Path $outputFilePath -WorksheetName "Stale_Files_1" -TableName "StaleFilesTable1" -TableStyle 'Medium6' -AutoSize
        Write-Log "Created empty Excel file: $outputFilePath" "INFO"
    }
}

# Create summary files if they don't exist yet
if ($global:summaryData.Count -gt 0 -or $global:totalItemsProcessed -eq 0) {
    New-SummaryFile -FilePath $outputFilePath
}

# Final summary
$totalTime = (Get-Date) - $script:startTime
Write-Log "Stale files scan completed. Target date: $targetDate. Total stale files found: $global:totalItemsProcessed" "SUCCESS"
Write-Log "Total processing time: $totalTime"
Write-Log "Results available in: $outputFilePath" "SUCCESS"

# Disconnect from Microsoft Graph
try {
    Disconnect-MgGraph | Out-Null
    Write-Log "Disconnected from Microsoft Graph" "INFO"
}
catch {
    Write-Log "Error disconnecting from Microsoft Graph: $($_.Exception.Message)" "WARNING"
}

# Check if file exists before trying to open
if (Test-Path $outputFilePath) {
    Write-Log "$outputFormat file created successfully at: $outputFilePath" "SUCCESS"
    
    if ($outputFormat -eq "CSV") {
        # For CSV, also mention the summary files
        $summaryPath = $outputFilePath -replace '\.csv$', '_Summary.csv'
        $siteSummaryPath = $outputFilePath -replace '\.csv$', '_SiteSummary.csv'
        if (Test-Path $summaryPath) {
            Write-Log "Summary file: $summaryPath" "SUCCESS"
        }
        if (Test-Path $siteSummaryPath) {
            Write-Log "Site summary file: $siteSummaryPath" "SUCCESS"
        }
    }
    
    # Open the main output file
    try {
        Start-Process $outputFilePath
    }
    catch {
        Write-Log "Could not automatically open the $outputFormat file. Please open manually: $outputFilePath" "INFO"
    }
}
else {
    Write-Log "ERROR: $outputFormat file was not created. Check the log for errors." "ERROR"
}
