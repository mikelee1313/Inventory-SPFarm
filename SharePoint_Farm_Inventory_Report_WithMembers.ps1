<#
.SYNOPSIS
    Generates a comprehensive inventory report of SharePoint farm sites, lists, and permissions.

.DESCRIPTION
    This script creates a detailed inventory of all SharePoint site collections, sites, lists, and libraries
    in a SharePoint farm. It collects information about item counts, sizes, last modified dates, 
    owners, and user permissions. The results are exported to a CSV file, and all processing 
    activities are logged to a text file.

    The script handles large lists efficiently by retrieving items in batches and properly disposes
    of SharePoint objects to manage memory usage.

.PARAMETER None
    This script does not accept parameters directly. All paths are generated dynamically.

.OUTPUTS
    - CSV file with inventory data saved to %TEMP% directory
    - Log file with processing information saved to %TEMP% directory

.EXAMPLE
    .\SharePoint_Farm_Inventory_Report_With_Members.ps1
    Runs the script and generates inventory report files with timestamped names in the %TEMP% directory.

.NOTES
    File Name      : SharePoint_Farm_Inventory_Report_With_Members.ps1
    Author         : Mike Lee
    Date Created   : 6/18/25
    Prerequisite   : SharePoint PowerShell Snap-in (Microsoft.SharePoint.PowerShell)


.FUNCTIONALITY
    - Verifies and loads required SharePoint PowerShell snap-in
    - Processes all site collections in the farm (excluding sitemaster)
    - Collects detailed information about lists and libraries
    - Identifies site and list owners
    - Maps user permissions while excluding owners and limited access
    - Calculates storage usage for each list/library
    - Exports data to CSV and logs all activities
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

# Function to get all users with permissions (excluding owners and limited access)
function Get-UserPermissions {
    param (
        [Microsoft.SharePoint.SPWeb]$web,
        [Microsoft.SharePoint.SPList]$list = $null,
        [string]$logFilePath
    )
    
    try {
        $userPermissionsList = @()
        $ownerUsers = @()
        
        # Get owner users to exclude them from the permissions list
        if ($web.AssociatedOwnerGroup) {
            foreach ($user in $web.AssociatedOwnerGroup.Users) {
                $ownerUsers += $user.LoginName
            }
        }
        
        # Get all owner groups
        foreach ($grp in ($web.Groups | Where-Object { $_.Name -match 'Owner' })) {
            foreach ($user in $grp.Users) {
                $ownerUsers += $user.LoginName
            }
        }
        
        # Check if we're getting permissions for a list or the web
        $securable = if ($list) { $list } else { $web }
        
        # Get all role assignments
        foreach ($roleAssignment in $securable.RoleAssignments) {
            $member = $roleAssignment.Member
            
            # Process groups
            if ($member -is [Microsoft.SharePoint.SPGroup]) {
                foreach ($user in $member.Users) {
                    # Skip if user is an owner or has only limited access
                    if ($ownerUsers -notcontains $user.LoginName) {
                        $hasNonLimitedAccess = $false
                        foreach ($roleDefinition in $roleAssignment.RoleDefinitionBindings) {
                            if ($roleDefinition.Name -ne "Limited Access") {
                                $hasNonLimitedAccess = $true
                                break
                            }
                        }
                        
                        if ($hasNonLimitedAccess) {
                            $permissions = ($roleAssignment.RoleDefinitionBindings | Where-Object { $_.Name -ne "Limited Access" } | ForEach-Object { $_.Name }) -join ";"
                            $userPermissionsList += "$($user.Name) ($($user.Email)) [$permissions]"
                        }
                    }
                }
            }
            # Process individual users
            elseif ($member -is [Microsoft.SharePoint.SPUser]) {
                # Skip if user is an owner or has only limited access
                if ($ownerUsers -notcontains $member.LoginName) {
                    $hasNonLimitedAccess = $false
                    foreach ($roleDefinition in $roleAssignment.RoleDefinitionBindings) {
                        if ($roleDefinition.Name -ne "Limited Access") {
                            $hasNonLimitedAccess = $true
                            break
                        }
                    }
                    
                    if ($hasNonLimitedAccess) {
                        $permissions = ($roleAssignment.RoleDefinitionBindings | Where-Object { $_.Name -ne "Limited Access" } | ForEach-Object { $_.Name }) -join ";"
                        $userPermissionsList += "$($member.Name) ($($member.Email)) [$permissions]"
                    }
                }
            }
        }
        
        return $userPermissionsList | Select-Object -Unique
    }
    catch {
        Write-Log -message "Error retrieving user permissions - $_" -logFilePath $logFilePath
        return @()
    }
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

        # Get owner details
        $ownerDetailsList = @()
        $siteTitle = $web.Title
        $owners = @(GetSiteOwner($web.Url))
        
        if ($owners.Count -gt 0) {
            foreach ($owner in $owners) {
                $ownerDetailsList += "$($owner.UserName) ($($owner.UserEmail))"
            }
            # Use the first owner's site title if available
            if ($owners[0].SiteTitle) {
                $siteTitle = $owners[0].SiteTitle
            }
        }
        
        # Get user permissions for the list (if it has unique permissions) or the web
        $userPermissions = @()
        if ($list.HasUniqueRoleAssignments) {
            $userPermissions = @(Get-UserPermissions -web $web -list $list -logFilePath $logFilePath)
        }
        else {
            $userPermissions = @(Get-UserPermissions -web $web -logFilePath $logFilePath)
        }
        
        # Ensure we have non-null arrays for joining
        if ($ownerDetailsList.Count -eq 0) {
            $ownerDetailsList = @("No owners found")
        }
        if ($userPermissions.Count -eq 0) {
            $userPermissions = @("No additional permissions")
        }
        
        [PSCustomObject]@{
            SiteUrl            = $web.Url
            ListName           = $list.Title
            ItemCount          = $itemCount
            TotalSizeMB        = [math]::Round($totalSize / 1MB, 2)
            LastModified       = $lastModified
            FullUrl            = $list.DefaultViewUrl
            SiteTitle          = $siteTitle 
            OwnerDetails       = [string]::Join(", ", $ownerDetailsList)
            'User Permissions' = [string]::Join(", ", $userPermissions)
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
    
    $results = @()
    
    try {
        # Get the web object 
        $web = Get-SPWeb -Identity $webURL
        
        # AssociatedOwnerGroup will give the details of the owner group exist in web 
        if ($web.AssociatedOwnerGroup) {
            foreach ($user in $web.AssociatedOwnerGroup.Users) {
                $results += [PSCustomObject]@{
                    SiteUrl      = $web.Url 
                    SiteTitle    = $web.Title 
                    UserName     = $user.Name 
                    UserEmail    = $user.Email 
                    LastModified = $web.LastItemModifiedDate 
                }
            }
        }
        else {            
            foreach ($grp in ($web.Groups | Where-Object { $_.Name -match 'Owner' })) {             
                foreach ($user in $grp.Users) {
                    $results += [PSCustomObject]@{
                        SiteUrl      = $web.Url 
                        SiteTitle    = $web.Title 
                        UserName     = $user.Name 
                        UserEmail    = $user.Email 
                        LastModified = $web.LastItemModifiedDate 
                    }
                }
            }
        }
        
        return $results
    }
    catch {
        Write-Log -message "Error retrieving site owners for web: $webURL - $_" -logFilePath $logFilePath
        return @()
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
