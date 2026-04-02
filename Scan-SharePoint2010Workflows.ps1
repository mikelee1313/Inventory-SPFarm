<#
.SYNOPSIS
    Scans a SharePoint Web Application to retrieve 2010 workflow details and their last run status.

.DESCRIPTION
    This script iterates through all site collections, webs, and lists within a specified SharePoint Web Application.
    It identifies workflows associated with each list, retrieves workflow details such as name, GUID, running instances,
    enabled status, author, creation and modification dates, and the last run timestamp. The results are exported to a CSV file.
    Any sites that cannot be accessed or cause exceptions are reported as warnings and skipped.

.PARAMETER webapplication
    The URL of the SharePoint Web Application to scan. Update this variable before running the script.

.PARAMETER path
    The directory path where the output CSV file will be saved. Update this variable before running the script.

.OUTPUTS
    CSV file containing workflow details:
        - workflow_<timestamp>.csv: Contains workflow details and their last run status.
    Any sites that cannot be accessed are reported as warnings to the console.

.NOTES

Authors: Mike Lee / Sean Gerlinger 
Date: 4/2/2025
Updated: 4/2/2026 - added last run status retrieval and improved error handling. Corrected some variable names and added progress indication.

    Requirements:
        - SharePoint PowerShell snap-in (Microsoft.SharePoint.PowerShell) must be available.
        - Must be run from a SharePoint Management Shell session.
        - Executing account must be a Farm Administrator or have Full Read on the target web application.

    The script excludes site collections matching patterns "sitemaster", "/my", or starting with "app-".

.EXAMPLE
    Update the variables $webapplication and $path, then execute the script:
    .\Scan-SharePoint2010Workflows.ps1

    This will generate a CSV file with workflow details in the specified directory.

#>

[CmdletBinding()]
param()

#region Configuration

# Variables that need updating

# Web Application to scan
$webapplication = "http://YOURWEBAPPLICATION/"

# Log file location and name
$path = "c:\temp\"

# Setup log paths - no changes required
$dateTime = Get-Date -Format yyyy-MM-dd_HH-mm-ss
$WorkFlowList = Join-Path $path ("workflow_" + $dateTime + ".csv")

#endregion Configuration

#region Initialization

# Ensure output directory exists
if (-not (Test-Path -Path $path)) {
    New-Item -ItemType Directory -Path $path -Force | Out-Null
    Write-Host "Created output directory: $path" -ForegroundColor Yellow
}

# Add SharePoint snapin if not already loaded
if ($null -eq (Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue)) {
    Add-PSSnapin Microsoft.SharePoint.PowerShell
}

#endregion Initialization

#region Functions

# Function to look up a workflow's last run status
function Get-WorkFlowLastRan {
    param (
        [Microsoft.SharePoint.SPWeb]$web,
        [string]$HistoryListName,
        [System.Guid]$WorkFlowGUID        
    )  

    $list = $web.Lists[$HistoryListName]

    # Guard against missing or empty history list
    if ($null -eq $list -or $list.ItemCount -eq 0) {
        return $null   
    }

    # Format GUID with braces as required by CAML text comparison
    $guidFormatted = $WorkFlowGUID.ToString("B")
    
    # Fetch only the single most recent history entry using OrderBy + RowLimit=1
    $query = New-Object Microsoft.SharePoint.SPQuery
    $query.Query = "<Where><Eq><FieldRef Name='WorkflowAssociation' /><Value Type='Text'>$guidFormatted</Value></Eq></Where><OrderBy><FieldRef Name='Occurred' Ascending='FALSE'/></OrderBy>"
    $query.RowLimit = 1
    $listItems = $list.GetItems($query)
    
    if ($listItems.Count -eq 0) {
        return $null
    }

    return $listItems[0]['Occurred']
}

# Recursive function to iterate through all webs and their subsites
function Get-AllWebs {
    param (
        [Microsoft.SharePoint.SPWeb]$web
    )

    # Process the current web
    Write-Verbose "Reading Lists in: $($web.Url)"

    # Look at all the lists in a web
    foreach ($list in $web.Lists) {  
        if ($list.WorkflowAssociations.Count -gt 0) {  
            foreach ($wflowAssociation in $list.WorkflowAssociations) {  
                # Get the last run status of the workflow
                $lastRan = Get-WorkFlowLastRan -web $web -HistoryListName $wflowAssociation.HistoryListTitle -WorkFlowGUID $wflowAssociation.Id
                # Construct the list URL
                $ListURL = "$($web.Url)/$($list.Title)"
                # Collect row details
                $authorDisplayName = try {
                    (Get-SPUser -Identity $wflowAssociation.Author -Web $web.Url -ErrorAction Stop).DisplayName
                }
                catch {
                    $wflowAssociation.Author
                }
                $script:results.Add([PSCustomObject]@{            
                        "Workflow Name"     = $wflowAssociation.InternalName
                        "Workflow GUID"     = $wflowAssociation.Id
                        "RunningInstances"  = $wflowAssociation.RunningInstances
                        "Is Enabled"        = $wflowAssociation.Enabled  
                        "List URL"          = $ListURL
                        "Author"            = $authorDisplayName
                        "Created On"        = $wflowAssociation.Created  
                        "Modified On"       = $wflowAssociation.Modified  
                        "Parent Web"        = $web.Url
                        "History List Name" = $wflowAssociation.HistoryListTitle  
                        "Last Ran"          = $lastRan                                                           
                    }) | Out-Null  
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

#endregion Functions

#region Main
[System.Collections.ArrayList]$results = @()
$WebApp = Get-SPWebApplication $webapplication
Write-Host "Scanning Web Application:" $WebApp.Name -ForegroundColor Green

# Get All site collections and iterate through
$SitesColl = $WebApp.Sites
$siteCount = $SitesColl.Count
$siteIndex = 0

foreach ($Site in $SitesColl) {
    $siteIndex++
    $rootWeb = $null
    try {

        # Skip sites that match "sitemaster" or "/my/" or start with "app-"
        if ($Site.Url -match "sitemaster" -or $Site.Url -match "/my" -or $Site.Url -match "app-") {
            continue
        }

        Write-Progress -Activity "Scanning Site Collections" -Status "[$siteIndex/$siteCount] $($Site.Url)" -PercentComplete (($siteIndex / $siteCount) * 100)

        # Look in all webs in a site collection recursively using the new function Get-AllWebs
        $rootWeb = $Site.RootWeb
        Get-AllWebs -web $rootWeb
          
    }
    catch {  
        # Log any exceptions as warnings and continue
        Write-Warning "Skipped $($Site.Url): $($_.Exception.Message)"
    }
    finally {
        if ($null -ne $rootWeb) { $rootWeb.Dispose() }
        $Site.Dispose()
    }
}

Write-Progress -Activity "Scanning Site Collections" -Completed

# Dump the results in the log file after processing all sites and webs.
if ($results.Count -gt 0) {
    $results | Export-Csv -Path "$WorkFlowList" -NoTypeInformation   
}

Write-Host " === === === === === Completed! === === === === === === == "
Write-Host "Log file saved to: `n$WorkFlowList"

#endregion Main 
