<#
.SYNOPSIS
    Retrieves and exports Site Collection Administrators from SharePoint Online sites.

.DESCRIPTION
    This script connects to SharePoint Online using app-only authentication (certificate) and retrieves
    all site collection administrators from non-personal sites. It handles throttling with
    automatic retries and exports the results to a CSV file. The script also creates a log file
    to track execution.

.NOTES
    File Name      : Get-SCA-AllSites-noEntraMembers.ps1
    Prerequisite   : PnP PowerShell module installed
    Author         : Mike Lee | Vijay Kumar
    Date           : 4/10/2025

.EXAMPLE
    .\Get-SCA-AllSites-noEntraMembers.ps1

.OUTPUTS
    - CSV file with site collection admins (in %TEMP% folder)
    - Log file with execution details (in %TEMP% folder)

.FUNCTIONALITY
    - Uses certificate-based authentication to connect to SharePoint admin center
    - Retrieves all non-personal site collections
    - Connects to each site and gets its administrators
    - Special handling for site Owners groups (retrieves group members)
    - Implements retry mechanism for throttling
    - Exports results to CSV and logs activities
#>

# Set Variables
$tenantname = "contoso" #This is your tenant name
$appID = "1e488dc4-1977-48ef-8d4d-9856f4e04536"  #This is your Entra App ID
$thumbprint = "5EAD7303A5C7E27DB4245878AD554642940BA082" #This is certificate thumbprint
$tenant = "9cfc42cb-51da-4055-87e9-b20a170b6ba3" #This is your Tenant ID

#Define Log path
$startime = Get-Date -Format "yyyyMMdd_HHmmss"
$ouputpath = "$env:TEMP\" + 'SiteCollectionAdmins_' + $startime + ".csv"
$logFilePath = "$env:TEMP\" + 'SiteCollectionAdmins_' + $startime + ".log"

#This is the logging function
function Write-Log {
    param (
        [string]$message,
        [string]$level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $level - $message"
    Add-Content -Path $logFilePath -Value $logMessage
}

Write-Host "Starting script to get Site Collection Admins at $startime" -ForegroundColor Yellow
Write-Log "Starting script to get Site Collection Admins" -level "INFO"
$SiteURL = "https://$tenantname-admin.sharepoint.com"
Connect-PnPOnline -Url $SiteURL -ClientId $appID -Thumbprint $thumbprint -Tenant $tenant

# Function to handle throttling
function Invoke-WithRetry {
    param (
        [scriptblock]$ScriptBlock
    )
    $retryCount = 0
    $maxRetries = 5
    while ($retryCount -lt $maxRetries) {
        try {
            $result = & $ScriptBlock
            return $result
        }
        catch {
            if ($_.Exception.Response.StatusCode -eq 429) {
                $retryAfter = $_.Exception.Response.Headers["Retry-After"]
                if (-not $retryAfter) {
                    $retryAfter = 30 # Default retry interval in seconds
                    Write-Warning "Throttled. 'Retry-After' header missing. Using default retry interval of $retryAfter seconds."
                }
                else {
                    Write-Warning "Throttled. Retrying after $retryAfter seconds."
                }
                Start-Sleep -Seconds $retryAfter
                $retryCount++
            }
            else {
                throw $_
            }
        }
    }
    throw "Max retries reached. Exiting."
}

# Get all site collections
$sites = Invoke-WithRetry { Get-PnPTenantSite | Where-Object { $_.Url -notlike "*-my.sharepoint.com*" } }

# Prepare an array to store results
$results = @()

foreach ($site in $sites) {
    # Connect to each site collection
    Write-Host "Getting Site Collection Admins from: $($site.Url)" -ForegroundColor Green
    Write-Log "Getting Site Collection Admins from: $($site.Url)"
    Invoke-WithRetry { Connect-PnPOnline -Url $site.Url -ClientId $appID -Thumbprint $thumbprint -Tenant $tenant }

    # Get site collection administrators
    $admins = Invoke-WithRetry { Get-PnPSiteCollectionAdmin }


    foreach ($admin in $admins) {
        if ($admin.PrincipalType -eq "User") {
            # Add user admin details directly
            $results += [PSCustomObject]@{
                SiteUrl       = $site.Url
                Title         = $admin.Title
                User          = $admin.Email
                PrincipalType = $admin.PrincipalType
            }
        }
        elseif ($admin.PrincipalType -eq "SecurityGroup" -and $admin.Title.ToLower().Contains("owners")) {
            try {
                # Retrieve group members explicitly
                try {
                    $groupMembers = Invoke-WithRetry { Get-PnPGroupMember -Identity $admin.Title }
                    
                    # Create a comma-separated list of members
                    $membersList = ($groupMembers | ForEach-Object { "$($_.Title) ($($_.Email))" }) -join "; "
                    
                    # Add a single row with all members in one field
                    $results += [PSCustomObject]@{
                        SiteUrl       = $site.Url
                        Title         = $admin.Title
                        User          = "Owners Group: $membersList"
                        PrincipalType = "SharePoint Group"
                    }
                }
                catch {
                    Write-Log "Group '$($admin.Title)' in site '$($site.Url)' is deleted or inaccessible: $_" -level "WARNING"
                }
            }
            catch {
                Write-Log "Failed to retrieve members for group '$($admin.Title)' in site '$($site.Url)': $_" -level "WARNING"
            }
        }

        else {
            # Add user admin details directly
            $results += [PSCustomObject]@{
                SiteUrl       = $site.Url
                Title         = $admin.Title
                User          = $admin.Email
                PrincipalType = $admin.PrincipalType
            }
        }
    }
}

# Export results to CSV
$results | Export-Csv -Path $ouputpath -NoTypeInformation -Encoding UTF8
Write-Host "Operations completed successfully and results exported to $ouputpath" -ForegroundColor Yellow
write-host "Check Log file any issues: $logFilePath" -ForegroundColor Cyan
Write-Log "Operations completed successfully and results exported to $ouputpath" -level "INFO"
