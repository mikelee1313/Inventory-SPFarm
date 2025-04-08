<#
.SYNOPSIS
    Generates a comprehensive report of SharePoint sites, lists, and libraries with owner information.

.DESCRIPTION
    This script scans an entire SharePoint farm to collect detailed information about all sites, lists, and libraries.
    It records item counts, sizes, last modified dates, and owner information for each component.
    The script implements batch processing to handle large lists efficiently and exports the results to a CSV file.

.PARAMETER None
    This script does not accept parameters directly. Configuration is handled within the script.

.OUTPUTS
    - CSV file with detailed information about all SharePoint sites, lists, and libraries
    - Log file with timestamps of processing activities and any errors encountered

.NOTES
    File Name: SharePoint_Farm_Inventory_Report.ps1
    Authors: Mike Lee
    Date: 4/8/2025
    
.EXAMPLE
    .\SharePoint_Farm_Inventory_Report.ps1
    
    Runs the script and generates a report in the user's temp directory.

.FUNCTIONALITY
    - Automatically loads required SharePoint PowerShell snap-ins
    - Processes all site collections in a SharePoint farm (excluding "sitemaster" sites)
    - Collects list/library details including size, item count, and last modified date
    - Identifies site owners and their contact information
    - Processes items in batches to handle large lists efficiently
    - Generates comprehensive logs with timestamps
    - Exports all data to a single CSV file
    - Properly disposes of SharePoint objects to prevent memory leaks
#>
# Function to check if the PSSnapin is already loaded
function Test-PSSnapinLoaded {
    param (
        [string]$snapinName
    )

    if (Get-PSSnapin -Registered -Name $snapinName -ErrorAction SilentlyContinue) {
        if (-not (Get-PSSnapin -Name $snapinName -ErrorAction SilentlyContinue)) {
            Add-PSSnapin $snapinName -ErrorAction SilentlyContinue
        }
    }
    else {
        Write-Host "PSSnapin $snapinName is not registered on this system."
        exit
    }
}

# Ensure the SharePoint PowerShell snap-in is loaded
Test-PSSnapinLoaded -snapinName "Microsoft.SharePoint.PowerShell"

# Function to write log messages with timestamps
function Write-Log {
    param (
        [string]$message,       # The message to log
        [string]$logFilePath    # The path to the log file
    )

    # Get the current timestamp
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    # Create the log entry with the timestamp and message
    $logEntry = "$timestamp - $message"
    # Append the log entry to the log file
    Add-Content -Path $logFilePath -Value $logEntry
}

# Function to get details of all lists and libraries in a site
function Get-ListDetails {
    param (
        [Microsoft.SharePoint.SPWeb]$web,  # The SharePoint site (web) to process
        [string]$logFilePath               # The path to the log file
    )

    # Get all visible lists and libraries in the site
    $lists = $web.Lists | Where-Object { $_.Hidden -eq $false }
    foreach ($list in $lists) {
        # Write the processing of the current list/library to the log
        Write-Log -message "Processing list/library: $($list.Title) in site: $($web.Url)" -logFilePath $logFilePath

        # Initialize variables for total size and item count
        $totalSize = 0
        $itemCount = $list.ItemCount
        $lastModified = $list.LastItemModifiedDate

        # Function to get items in batches
        function Get-ListItemsInBatches {
            param (
                [Microsoft.SharePoint.SPList]$list,
                [int]$batchSize,
                [ref]$totalSize
            )

            $query = New-Object Microsoft.SharePoint.SPQuery
            $query.RowLimit = $batchSize
            $position = $null

            try {
                do {
                    $query.ListItemCollectionPosition = $position
                    $items = $list.GetItems($query)
                    $position = $items.ListItemCollectionPosition

                    foreach ($item in $items) {
                        if ($null -ne $item.File) {
                            $totalSize.Value += $item.File.Length
                        }
                    }
                } while ($null -ne $position)
            }
            catch {
                Write-Log -message "Error retrieving items in batches for list: $($list.Title) - $_" -logFilePath $logFilePath
            }
        }

        # Get items in batches of 500
        $totalSizeRef = [ref]$totalSize
        Get-ListItemsInBatches -list $list -batchSize 500 -totalSize $totalSizeRef

        # Create a custom object with the list/library details and site owner details in a single cell
        $ownerDetailsList = @()
        foreach ($owner in GetSiteOwner($web.Url)) {
            $ownerDetailsList += "$($owner.UserName) ($($owner.UserEmail))"
        }
        
        [PSCustomObject]@{
            SiteUrl      = $web.Url
            ListName     = $list.Title
            ItemCount    = $itemCount
            TotalSizeMB  = [math]::Round($totalSize / 1MB, 2)
            LastModified = $lastModified
            FullUrl      = $list.DefaultViewUrl
            SiteTitle    = $owner.SiteTitle 
            OwnerDetails = [string]::Join(", ", $ownerDetailsList)
        }
    }
}

# Function to process each site and its sub-sites
function Invoke-SiteProcessing {
    param (
        [Microsoft.SharePoint.SPWeb]$web,  # The SharePoint site (web) to process
        [string]$logFilePath               # The path to the log file
    )

    try {
        # Write the processing of the current site to the log
        Write-Log -message "Processing site: $($web.Url)" -logFilePath $logFilePath
        
        # Get details of all lists and libraries in the site including site owner details 
        Get-ListDetails -web $web -logFilePath $logFilePath

        # Recursively process each sub-site 
        foreach ($subWeb in $web.Webs) { 
            Invoke-SiteProcessing -web $subWeb -logFilePath $logFilePath 
            $subWeb.Dispose() 
        } 
    }
    catch { 
        # Write any errors that occur during processing to the log 
        Write-Log -message "Error processing site: $($web.Url) - $_" -logFilePath $logFilePath 
    }
}

# Function to get site owners list across the farm 
function GetSiteOwner {
    param (
        [string]$webURL          # The URL of the web to process 
    )
    
    try {
        # Get the web object 
        $web = Get-SPWeb -Identity $webURL
        
        # AssociatedOwnerGroup will give the details of the owner group exist in web 
        if ($web.AssociatedOwnerGroup) {
            foreach ($owner in $web.AssociatedOwnerGroup) {
                foreach ($user in $owner.Users) {
                    [PSCustomObject]@{
                        SiteUrl      = $web.Url 
                        SiteTitle    = $web.Title 
                        UserName     = $user.Name 
                        UserEmail    = $user.Email 
                        LastModified = $web.LastItemModifiedDate 
                    }
                }
            }
        }
        else {            
            foreach ($grp in ($web.Groups | Where-Object { $_.Name -match 'Owner' })) {             
                foreach ($user in $grp.Users) {
                    [PSCustomObject]@{
                        SiteUrl      = $web.Url 
                        SiteTitle    = $web.Title 
                        UserName     = $user.Name 
                        UserEmail    = $user.Email 
                        LastModified = $web.LastItemModifiedDate 
                    }
                }
            }
        }
        
    }
    catch {
        Write-Log -message "Error retrieving site owners for web: $webURL - $_" -logFilePath $logFilePath
    }
    finally {
        if ($null -ne $web) { $web.Dispose() }
    }
}

# Main script execution 
try { 
    # Get all site collections in the SharePoint farm 
    $allSiteCollections = Get-SPWebApplication | Get-SPSite -Limit All 
    $output = @() 
    $totalSiteCollections = $allSiteCollections.Count 
    $currentSiteCollection = 0 
    
    # Generate a timestamp for the output and log file names 
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss" 
    $outputFilePath = "$env:TEMP\SharePoint_Farm_Inventory_Report_$timestamp.csv" 
    $logFilePath = "$env:TEMP\SharePoint_Farm_Inventory_Log_$timestamp.txt"

    # Process each site collection 
    foreach ($siteCollection in $allSiteCollections) { 
        $currentSiteCollection++ 
        # Update the progress indicator 
        Write-Progress -Activity "Processing Site Collections" -Status "Processing $currentSiteCollection of $totalSiteCollections" -PercentComplete (($currentSiteCollection / $totalSiteCollections) * 100) 

        # Skip site collections that contain "sitemaster" in the URL 
        if ($siteCollection.Url -notmatch "sitemaster") { 
            # Process each site (web) in the site collection 
            foreach ($web in $siteCollection.AllWebs) { 
                $output += Invoke-SiteProcessing -web $web -logFilePath $logFilePath 
                $web.Dispose() 
            } 
            $siteCollection.Dispose() 
        } 
    } 

    # Export the collected data to a single CSV file 
    $output | Export-Csv -Path $outputFilePath -NoTypeInformation
    Write-Host "The output file has been saved to $outputFilePath" 
    Write-Host "The log file has been saved to $logFilePath" 
}
catch { 
    # Write any errors that occur during the main script execution to the log 
    Write-Log -message "Error in main script execution - $_" -logFilePath $logFilePath 
}
finally { 
    # Dispose of all site collections to free up resources 
    foreach ($siteCollection in $allSiteCollections) { 
        $siteCollection.Dispose() 
    } 
}
