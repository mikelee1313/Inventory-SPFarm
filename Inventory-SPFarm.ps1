<#
.SYNOPSIS
    Script to gather details of all lists and libraries in SharePoint site collections and log the process.

.DESCRIPTION
    This script ensures the SharePoint PowerShell snap-in is loaded, processes each site collection in the SharePoint farm,
    gathers details of all lists and libraries, and logs the process. The collected data is exported to a CSV file.

.PARAMETER snapinName
    The name of the PowerShell snap-in to check and load if not already loaded.

.PARAMETER message
    The message to log.

.PARAMETER logFilePath
    The path to the log file.

.PARAMETER web
    The SharePoint site (web) to process.

.PARAMETER list
    The SharePoint list to process.

.PARAMETER batchSize
    The number of items to retrieve in each batch.

.PARAMETER siteCollection
    The SharePoint site collection to process.

.PARAMETER outputFilePath
    The path to the output CSV file.

.PARAMETER allSiteCollections
    All site collections in the SharePoint farm.

.PARAMETER totalSiteCollections
    The total number of site collections.

.PARAMETER currentSiteCollection
    The current site collection being processed.

.PARAMETER timestamp
    The timestamp used for generating output and log file names.

.FUNCTION Test-PSSnapinLoaded
    Checks if the specified PowerShell snap-in is loaded and loads it if not.

.FUNCTION Write-Log
    Writes log messages with timestamps to a specified log file.

.FUNCTION Get-ListDetails
    Retrieves details of all visible lists and libraries in a specified SharePoint site (web).

.FUNCTION Get-ListItemsInBatches
    Retrieves items from a SharePoint list in batches to calculate the total size.

.FUNCTION Invoke-SiteProcessing
    Processes each site and its sub-sites, logging the process and retrieving list details.

.EXAMPLE
    .\Inventory-SPFarm
    Executes the script to gather details of all lists and libraries in SharePoint site collections and logs the process.

.NOTES

Authors: Mike Lee
Date: 1/14/2025

Disclaimer: The sample scripts are provided AS IS without warranty of any kind. 

Microsoft further disclaims all implied warranties including, without limitation, 
any implied warranties of merchantability or of fitness for a particular purpose. 
The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. 
In no event shall Microsoft, its authors, or anyone else involved in the creation, 
production, or delivery of the scripts be liable for any damages whatsoever 
(including, without limitation, damages for loss of business profits, business interruption, 
loss of business information, or other pecuniary loss) arising out of the use of or inability 
to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages.

Ensure you have the necessary permissions to run SharePoint PowerShell commands and access the SharePoint farm.

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
        [string]$message, # The message to log
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
        [Microsoft.SharePoint.SPWeb]$web, # The SharePoint site (web) to process
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
                [int]$batchSize
            )

            $query = New-Object Microsoft.SharePoint.SPQuery
            $query.RowLimit = $batchSize
            $position = $null

            do {
                $query.ListItemCollectionPosition = $position
                $items = $list.GetItems($query)
                $position = $items.ListItemCollectionPosition

                foreach ($item in $items) {
                    if ($null -ne $item.File) {
                        $totalSize += $item.File.Length
                    }
                }
            } while ($null -ne $position)
        }

        # Get items in batches of 5000
        Get-ListItemsInBatches -list $list -batchSize 5000

        # Create a custom object with the list/library details
        [PSCustomObject]@{
            SiteUrl      = $web.Url
            ListName     = $list.Title
            ItemCount    = $itemCount
            TotalSizeMB  = [math]::Round($totalSize / 1MB, 2)
            LastModified = $lastModified
            FullUrl      = $list.DefaultViewUrl
        }
    }
}

# Function to process each site and its sub-sites
function Invoke-SiteProcessing {
    param (
        [Microsoft.SharePoint.SPWeb]$web, # The SharePoint site (web) to process
        [string]$logFilePath               # The path to the log file
    )

    try {
        # Write the processing of the current site to the log
        Write-Log -message "Processing site: $($web.Url)" -logFilePath $logFilePath
        # Get details of all lists and libraries in the site
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

# Main script execution
try {
    # Get all site collections in the SharePoint farm
    $allSiteCollections = Get-SPWebApplication | Get-SPSite -Limit All
    $output = @()
    $totalSiteCollections = $allSiteCollections.Count
    $currentSiteCollection = 0
    # Generate a timestamp for the output and log file names
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $outputFilePath = "$env:TEMP\SharePointListsReport_$timestamp.csv"
    $logFilePath = "$env:TEMP\SharePointScriptLog_$timestamp.txt"

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

    # Export the collected data to a CSV file
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
