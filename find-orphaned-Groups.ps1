# ============================================================================
# Script: Find SharePoint Sites with Orphaned Groups
# Description: Scans all SharePoint sites and identifies those with orphaned
#              Microsoft 365 Groups (GroupID exists but group is deleted)
# ============================================================================

# Tenant and authentication configuration
$tenantname = "m365x61250205" #This is your tenant name
$appID = "5baa1427-1e90-4501-831d-a8e67465f0d9"  #This is your Entra App ID
$thumbprint = "B696FDCFE1453F3FBC6031F54DE988DA0ED905A9" #This is certificate thumbprint
$tenant = "85612ccb-4c28-4a34-88df-a538cc139a51" #This is your Tenant ID
$tenantUrl = "https://$tenantname-admin.sharepoint.com"

# Connection parameters for reuse
$connectionParams = @{
    ClientId   = $appID
    Thumbprint = $thumbprint
    Tenant     = $tenant
}

# Connect to Exchange Online
Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
Connect-ExchangeOnline

# Connect to SharePoint Online Admin Center
Write-Host "Connecting to SharePoint Online..." -ForegroundColor Cyan
Connect-PnPOnline -Url $tenantUrl @connectionParams -ErrorAction Stop

# Specify the output file path
$outputFilePath = ".\SharePoint-Orphaned-Groups-Report-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"

# Create CSV file with headers
$csvHeaders = "Title,URL,Owner,IBMode,IBSegment,HasOrphanedGroup,IsTeamsConnected,IsTeamsChannelConnected,TeamsChannelType,GroupID,RelatedGroupId,LastContentModifiedDate,StorageQuota,StorageUsageCurrent,IsHubSite,HubSiteId,Status"
$csvHeaders | Out-File -FilePath $outputFilePath -Encoding UTF8

# Get all SharePoint sites and filter for group-connected sites
Write-Host "Retrieving all SharePoint sites..." -ForegroundColor Cyan
$allSites = Get-PnPTenantSite
$sites = $allSites | Where-Object { $_.GroupId -and $_.GroupId -ne "00000000-0000-0000-0000-000000000000" }

Write-Host "Found $($allSites.Count) total sites, $($sites.Count) are group-connected. Processing..." -ForegroundColor Green

# Process each site
$counter = 0
foreach ($site in $sites) {
    $counter++
    Write-Progress -Activity "Processing Sites" -Status "Site $counter of $($sites.Count): $($site.Title)" -PercentComplete (($counter / $sites.Count) * 100)
    
    Write-Host "Processing: $($site.Url)" -ForegroundColor Gray
    
    # Initialize result object
    $siteReport = [PSCustomObject]@{
        Title                   = $site.Title
        URL                     = $site.Url
        Owner                   = $site.Owner
        IBMode                  = ""  # Not available via Get-PnPTenantSite
        IBSegment               = $site.InformationBarrierSegments -join "; "
        HasOrphanedGroup        = "N/A"
        IsTeamsConnected        = ""  # Not reliably available via Get-PnPTenantSite
        IsTeamsChannelConnected = $site.IsTeamsChannelConnected
        TeamsChannelType        = $site.TeamsChannelType
        GroupID                 = $site.GroupId
        RelatedGroupId          = $site.RelatedGroupId
        LastContentModifiedDate = $site.LastContentModifiedDate
        StorageQuota            = $site.StorageQuota
        StorageUsageCurrent     = $site.StorageUsageCurrent
        IsHubSite               = $site.IsHubSite
        HubSiteId               = $site.HubSiteId
        Status                  = $site.Status
    }
    
    # Check if the site has a GroupID and it's not the empty GUID
    if ($site.GroupId -and $site.GroupId -ne "00000000-0000-0000-0000-000000000000") {
        try {
            # Try to get the unified group
            $group = Get-UnifiedGroup -Identity $site.GroupId -ErrorAction SilentlyContinue
            
            if ($null -eq $group) {
                # Group doesn't exist - it's orphaned
                $siteReport.HasOrphanedGroup = "YES"
                Write-Host "  [ORPHANED] $($site.Title) - GroupID: $($site.GroupId)" -ForegroundColor Yellow
            }
            else {
                # Group exists
                $siteReport.HasOrphanedGroup = "NO"
            }
        }
        catch {
            # Error checking group - likely orphaned
            $siteReport.HasOrphanedGroup = "YES (Error: $($_.Exception.Message))"
            Write-Host "  [ERROR] $($site.Title) - $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    else {
        # No GroupID or empty GUID - not group-connected
        $siteReport.HasOrphanedGroup = "N/A (No Group)"
    }
    
    # Write to CSV immediately after processing each site
    $siteReport | Export-Csv -Path $outputFilePath -NoTypeInformation -Encoding UTF8 -Append
}

Write-Progress -Activity "Processing Sites" -Completed

# Read back the results for summary
Write-Host "`nGenerating summary..." -ForegroundColor Cyan
$results = Import-Csv -Path $outputFilePath

# Display summary
$orphanedCount = ($results | Where-Object { $_.HasOrphanedGroup -like "YES*" }).Count
$totalWithGroups = ($results | Where-Object { $_.GroupID -and $_.GroupID -ne "00000000-0000-0000-0000-000000000000" }).Count

Write-Host "`n============================================================================" -ForegroundColor Green
Write-Host "Report completed successfully!" -ForegroundColor Green
Write-Host "============================================================================" -ForegroundColor Green
Write-Host "Total sites processed:        $($sites.Count)" -ForegroundColor White
Write-Host "Sites with GroupID:           $totalWithGroups" -ForegroundColor White
Write-Host "Sites with orphaned groups:   $orphanedCount" -ForegroundColor Yellow
Write-Host "Report saved to:              $outputFilePath" -ForegroundColor Cyan
Write-Host "============================================================================`n" -ForegroundColor Green

# Disconnect
Disconnect-PnPOnline
Disconnect-ExchangeOnline -Confirm:$false
