<#
.SYNOPSIS
    Scans a SharePoint Web Application to retrieve 2010 workflow details and their last run status.

.DESCRIPTION
    This script iterates through all site collections, webs, and lists within a specified SharePoint Web Application.
    It identifies workflows associated with each list, retrieves workflow details such as name, GUID, running instances,
    enabled status, author, creation and modification dates, and the last run timestamp. The results are exported to a CSV file.
    Any sites that cannot be accessed or cause exceptions are logged separately.

.PARAMETER webapplication
    The URL of the SharePoint Web Application to scan. Update this variable before running the script.

.PARAMETER path
    The directory path where the output CSV files (workflow details and blocked sites) will be saved. Update this variable before running the script.

.OUTPUTS
    CSV files containing workflow details and blocked sites:
        - workflow_<timestamp>.csv: Contains workflow details and their last run status.
        - blockedsites_<timestamp>.csv: Contains URLs or error messages of sites that could not be processed.

.NOTES

Authors: Mike Lee / Sean Gerlinger 
Date: 4/2/2025

    Requirements:
        - SharePoint PowerShell snap-in (Microsoft.SharePoint.PowerShell) must be available.
        - Execute the script with appropriate permissions to access SharePoint objects.

    The script excludes site collections matching patterns "sitemaster", "/my", or starting with "app-".

.EXAMPLE
    Update the variables $webapplication and $path, then execute the script:
    .\Scan-SharePoint2010Workflows.ps1

    This will generate CSV files with workflow details and blocked sites in the specified directory.

#>

# Variables that need updating

# Web Application to scan
$webapplication = "http://YOURWEBAPPLICATION/"

# Log file location and name
$path = "c:\\temp\\"

# Setup log paths - no changes required
$dateTime = Get-Date -Format yyyy-MM-dd_HH-mm-ss
$WorkFlowList = $path + "workflow" + "_" + $dateTime + ".csv"
$BlockedFileName = $path + "blockedsites" + "_" + $dateTime + ".csv"

# Add SharePoint snapin if not already loaded
if ($null -eq (Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue)) {
    Add-PSSnapin Microsoft.SharePoint.PowerShell
}

Clear-Host

# Function to look up a workflow's last run status
function Get-WorkFlowLastRan {
    param (
        [string]$WorkFlowWeb,
        [string]$HistoryListName,
        [string]$WorkFlowGUID        
    )  
           
    # Get the SPWeb and SPList objects
    $site = Get-SPWeb $WorkFlowWeb
    $web = $site
    $list = $web.Lists[$HistoryListName]

    # Fix $WorkflowGUID by adding missing {}
    $WorkFlowGUID = '{' + $WorkFlowGUID + '}'
   
    # Get all items from the list with paging support
    if ($list.ItemCount -eq 0) {
        return $null   
    }
    
    $query = New-Object Microsoft.SharePoint.SPQuery
    $query.Query = "<Where><Eq><FieldRef Name='WorkflowAssociation' /><Value Type='Text'>$WorkFlowGUID</Value></Eq></Where>"
    $query.RowLimit = 300
    $listItems = $list.GetItems($query)
    
    if ($listItems.Count -eq 0) {
        return $null
    }
    
    $results = @()
   
    # Loop through each item and find the matching item
    do {
        foreach ($item in $listItems) {
            $RowDetails = @{                            
                "Last Run" = $item['Occurred']
            }
            $results += New-Object PSObject -Property $RowDetails                             
        }
        # Get the next batch of items
        $query.ListItemCollectionPosition = $listItems.ListItemCollectionPosition
        if ($null -ne $query.ListItemCollectionPosition) {
            $listItems = $list.GetItems($query)
        }
        else {
            break
        }
    } while ($true)

    if ($results.Count -eq 0) {
        return $null
    }
    else {
        $sortedResults = $results | Sort-Object { $_.'Last Run' } -Descending
        $newestItem = $sortedResults
        return $newestItem.'Last Run'
    }
   
    # Dispose SharePoint objects
    $web.Dispose()
    $site.Dispose()
}

# Recursive function to iterate through all webs and their subsites
function Get-AllWebs {
    param (
        [Microsoft.SharePoint.SPWeb]$web
    )

    # Process the current web
    Write-Host "Reading Lists in: " $web.Url -ForegroundColor Magenta

    # Look at all the lists in a web
    foreach ($list in $web.Lists) {  
        if ($list.WorkflowAssociations) {  
            foreach ($wflowAssociation in $list.WorkflowAssociations) {  
                # Get the last run status of the workflow
                $LastRans = Get-WorkFlowLastRan -WorkFlowWeb $web.Url -HistoryListName $wflowAssociation.HistoryListTitle -WorkFlowGUID $wflowAssociation.Id
                # Construct the list URL
                $ListURL = "$($web.Url)/$($list.Title)"
                # Collect row details
                $RowDetails = @{            
                    "Workflow Name"     = $wflowAssociation.InternalName
                    "Workflow GUID"     = $wflowAssociation.Id
                    "RunningInstances"  = $wflowAssociation.RunningInstances
                    "Is Enabled"        = $wflowAssociation.Enabled  
                    "List URL"          = $ListURL
                    "Author"            = (Get-SPUser -Identity $wflowAssociation.Author -Web $web.Url).DisplayName
                    "Created On"        = $wflowAssociation.Created  
                    "Modified On"       = $wflowAssociation.Modified  
                    "Parent Web"        = $web.Url
                    "History List Name" = $wflowAssociation.HistoryListTitle  
                    "Last Ran"          = ($LastRans | Sort-Object { $_.'Last Run' } -Descending | Select-Object -First 1)                                                           
                }  
                # Add row details to results array
                [System.Collections.ArrayList]$script:results += New-Object PSObject -Property $RowDetails  
            }            
        }  
    }

    # Recursively process all subsites of the current web
    foreach ($subweb in $web.Webs) {
        Get-AllWebs -web $subweb
        # Dispose subweb object to free up memory
        $subweb.Dispose()
    }
}

# Start of scanning
$results = @()
$WebApp = Get-SPWebApplication $webapplication
Write-Host "Scanning Web Application:" $WebApp.Name -ForegroundColor Green
   
# Get All site collections and iterate through
$SitesColl = $WebApp.Sites

foreach ($Site in $SitesColl) {
    try {

        # Skip sites that match "sitemaster" or "/my/" or start with "app-"
        if ($Site.Url -match "sitemaster" -or $Site.Url -match "/my" -or $Site.Url -match "app-*") {
        
            continue
        }

        # Look in all webs in a site collection recursively using the new function Get-AllWebs
        Get-AllWebs -web $Site.RootWeb
          
    }
    catch {  
        # Log any exceptions to the blocked sites file
        $_.Exception.Message | Out-File -FilePath $BlockedFileName -Append
    } 
}

# Dump the results in the log file after processing all sites and webs.
if ($results.Count -gt 0) {
    [System.Collections.ArrayList]$results | Export-Csv -Path "$WorkFlowList" -NoTypeInformation   
}

Write-Host " === === === === === Completed! === === === === === === == "
Write-Host "Log files saved to: `n$WorkFlowList`n$BlockedFileName" 
