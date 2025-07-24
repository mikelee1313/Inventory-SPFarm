<#
.SYNOPSIS
    Scans SharePoint Online sites to identify all files and folders that have NOT been modified in the specified number of months (stale files).

.DESCRIPTION
    This script connects to SharePoint Online using provided tenant-level credentials and iterates through a list of 
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
    File Name      : Find-StaleSPOFiles.ps1
    Author         : Mike Lee
    Date Created   : 7/24/2025

    The script uses app-only authentication with a certificate thumbprint. Make sure the app has
    proper permissions in your tenant (Sites.FullControl.All is recommended).

    The script ignores several system folders and lists to improve performance and avoid errors.

    PREREQUISITES:
    - Install-Module ImportExcel -Scope CurrentUser
    - Install-Module PnP.PowerShell -Scope CurrentUser

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
$monthsBack = 6  # Number of months to look back from today (finds files NOT modified within this timeframe)

# --- Script Behavior Settings ---
$batchSize = 100  # How many items to process before writing to Excel
$maxItemsPerSheet = 5000 # Maximum items per sheet in Excel
$outputFormat = "xlsx"  # Output format: "CSV" or "XLSX"
$debug = $false  # Enable debug output: $true for detailed debug info, $false for informational only

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
$requiredModules = @('PnP.PowerShell')
if ($outputFormat -eq "XLSX") {
    $requiredModules += 'ImportExcel'
}

foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Host "Module '$module' is not installed. Installing..." -ForegroundColor Yellow
        Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber | Out-Null
    }
    Import-Module $module -Force | Out-Null
}

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

# Connect to SharePoint Online
function Connect-SharePoint {
    param (
        [string]$siteURL
    )
    try {
        Invoke-WithRetry -ScriptBlock {
            Connect-PnPOnline -Url $siteURL -ClientId $appID -Thumbprint $thumbprint -Tenant $tenant
        }
        Write-Log "Connected to SharePoint Online at $siteURL"
        
        # Validate connection by trying to get the web
        try {
            $web = Get-PnPWeb -ErrorAction Stop
            Write-Log "Successfully validated connection to: $($web.Title) ($($web.Url))"
        }
        catch {
            Write-Log "Connection validation failed: $($_.Exception.Message)" "ERROR"
            return $false
        }
        
        return $true # Connection successful
    }
    catch {
        Write-Log "Failed to connect to SharePoint Online at $siteURL : $($_.Exception.Message)" "ERROR"
        return $false # Connection failed
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

# Process SharePoint Item (File or Folder) - Modified for date filtering
function Get-SPItemInfo {
    param (
        $item,
        [string]$ItemSiteURL,
        [string]$ItemType, # "File" or "Folder"
        [string]$LibraryName
    )
    try {
        # Access field values using indexer
        $itemName = $item["FileLeafRef"]
        $itemPath = $item["FileRef"]
        
        # Get modified date
        $modifiedDateTime = $null
        $modifiedDateStr = "Unknown"
        
        try {
            $modifiedField = $item["Modified"]
            if ($modifiedField) {
                $modifiedDateTime = $modifiedField
                $modifiedDateStr = $modifiedDateTime.ToString("yyyy-MM-dd HH:mm:ss")
                
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
            }
            else {
                Write-Log "No modified date found for item $itemPath, skipping" "INFO"
                return $false
            }
        }
        catch {
            Write-Log "Error retrieving modified date for item $itemPath : $($_.Exception.Message)" "INFO"
            return $false
        }
        
        Write-Log "Processing item: $itemPath (Type: $ItemType, Modified: $modifiedDateStr)" "INFO"

        # Creator and Created Date
        $creatorName = "Unknown"
        $creatorEmail = "Unknown"
        $creatorWithEmail = "Unknown"
        $createdDateTime = $null

        try {
            $authorField = $item["Author"]
            if ($null -ne $authorField) {
                if ($null -ne $authorField.LookupId) {
                    $creatorInfo = Get-PnPUser -Identity $authorField.LookupId -ErrorAction SilentlyContinue
                    if ($null -ne $creatorInfo) {
                        $creatorName = $creatorInfo.Title
                        $creatorEmail = $creatorInfo.Email
                        if ([string]::IsNullOrEmpty($creatorEmail)) {
                            $creatorWithEmail = $creatorName
                        }
                        else {
                            $creatorWithEmail = "$creatorName ($creatorEmail)"
                        }
                    }
                }
                elseif ($null -ne $authorField.LookupValue) {
                    $creatorName = $authorField.LookupValue
                    $creatorWithEmail = $creatorName
                }
            }
            
            $createdField = $item["Created"]
            if ($createdField) {
                $createdDateTime = $createdField
            }
        }
        catch {
            Write-Log "Error retrieving creator/date for item $itemPath : $($_.Exception.Message)" "INFO"
        }
        
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

foreach ($siteURL in $siteURLs) {
    $siteStartTime = Get-Date
    Write-Log "Starting processing for site: $siteURL" "INFO"
    
    $siteItemCount = 0
    
    if (Connect-SharePoint -siteURL $siteURL) {
        try {
            # Get all lists including document libraries
            $lists = Get-PnPList -Includes BaseType, Hidden, Title, ItemCount | Where-Object { 
                $_.Hidden -eq $false -and 
                $_.Title -notin $ignoreFolders -and
                ($_.BaseType -eq "DocumentLibrary" -or $_.BaseType -eq "GenericList")
            }
            
            if ($null -eq $lists -or $lists.Count -eq 0) {
                Write-Log "No lists retrieved or all lists were ignored for site $siteURL." "WARNING"
                
                # Debug: Show all lists for troubleshooting (only if debug is enabled)
                if ($debug) {
                    $allLists = Get-PnPList -Includes Title, Hidden, BaseType
                    Write-Log "Debug - All lists in site: $($allLists | ForEach-Object { "$($_.Title) (Hidden: $($_.Hidden), BaseType: $($_.BaseType))" } | Out-String)" "INFO"
                }
            }
            else {
                Write-Log "Found $($lists.Count) lists to process in site $siteURL"
                if ($debug) {
                    Write-Log "Lists to process: $($lists | ForEach-Object { $_.Title } | Join-String -Separator ', ')" "INFO"
                }
                
                foreach ($list in $lists) { 
                    try {
                        $listName = $list.Title
                        Write-Log "Processing list/library: '$listName' on site: $siteURL"
                        
                        # Get item count first
                        $itemCount = $list.ItemCount
                        Write-Log "List '$listName' contains $itemCount items"
                        
                        if ($itemCount -eq 0) {
                            Write-Log "Skipping empty list: $listName"
                            continue
                        }
                        
                        # Get all items at once with required fields
                        try {
                            Write-Log "Retrieving all items from list '$listName'..."
                            
                            $items = @(Get-PnPListItem -List $list -PageSize 2000)
                            
                            if ($null -eq $items -or $items.Count -eq 0) {
                                Write-Log "No items retrieved from list '$listName'" "WARNING"
                                continue
                            }
                            
                            Write-Log "Retrieved $($items.Count) items from list '$listName'"
                            $itemsProcessedInList = 0
                            
                            foreach ($currentItem in $items) {
                                try {
                                    # Get field values
                                    $fsObjType = $currentItem["FSObjType"]
                                    $itemTypeStr = ""
                                    
                                    if ($fsObjType -eq 0) {
                                        $itemTypeStr = "File"
                                    }
                                    elseif ($fsObjType -eq 1) {
                                        $itemTypeStr = "Folder"
                                    }
                                    else {
                                        Write-Log "Skipping item with unknown FSObjType: $fsObjType" "INFO"
                                        continue
                                    }
                                    
                                    $currentItemPath = $currentItem["FileRef"]
                                    
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
                                    
                                    # Get-SPItemInfo will only add to batch if item meets stale file criteria
                                    $itemWasAdded = Get-SPItemInfo -item $currentItem -ItemSiteURL $siteURL -ItemType $itemTypeStr -LibraryName $listName
                                    if ($itemWasAdded) {
                                        $siteItemCount++
                                        $itemsProcessedInList++
                                    }
                                }
                                catch {
                                    Write-Log "Error processing individual item: $($_.Exception.Message)" "WARNING"
                                }
                            }
                            
                            Write-Log "Completed processing list '$listName'. Items processed: $itemsProcessedInList"
                        }
                        catch {
                            Write-Log "Error retrieving items from list '$listName': $($_.Exception.Message)" "ERROR"
                        }
                    }
                    catch {
                        Write-Log "Failed to process list '$($list.Title)' on site '$siteURL'. Error: $($_.Exception.Message)" "ERROR"
                    }
                }
            }
        }
        catch {
            Write-Log "Failed to get lists for site $siteURL. Error: $($_.Exception.Message)" "ERROR"
        }
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
