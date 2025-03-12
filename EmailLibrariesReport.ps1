<#

.SYNOPSIS
    Generates a CSV report of SharePoint document libraries configured with incoming email.

.DESCRIPTION
    This script iterates through all site collections and webs within a specified SharePoint web application, identifies document libraries configured with incoming email, and exports their details (Site Collection URL, Library Name, Owner, and Email Alias) to a CSV file.

.PARAMETER webAppUrl
    The URL of the SharePoint web application to scan.

.PARAMETER pageLimit
    The number of document libraries to process per page during enumeration. Adjust this value based on performance considerations.

.OUTPUTS
    CSV file containing details of document libraries configured with incoming email. The file is saved to "C:\temp\" with a timestamped filename.

.NOTES

Authors: Mike Lee
Date: 3/12/2025

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
    .\EmailLibrariesReport.ps1

    Runs the script with default parameters and generates a CSV report in the "C:\temp\" directory.
#>
# Define variables
$webAppUrl = "http://spwfe"
$pageLimit = 100
$starttime = Get-Date -Format "yyyyMMdd_HHmmss"
$csvFile = "C:\\temp\\Email_Libraries_Report_$starttime.csv"

# Function to check if the PSSnapin is already loaded
function Ensure-PSSnapinLoaded {
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

# Function to get document libraries with paging
function Get-DocumentLibraries {
    param (
        [string]$webUrl,
        [int]$pageLimit
    )
    $web = Get-SPWeb $webUrl
    $libraries = @()
    $position = 0

    while ($true) {
        $pagedLibraries = $web.Lists | Where-Object { $_.BaseTemplate -eq "DocumentLibrary" } | Select-Object -Skip $position -First $pageLimit
        if ($pagedLibraries.Count -eq 0) {
            break
        }
        $libraries += $pagedLibraries
        $position += $pageLimit
    }

    $web.Dispose() # Dispose of the web object after use
    return $libraries
}

# Function to process document libraries
function Get-DocumentLibrariesInfo {
    param (
        [string]$webUrl,
        [string]$siteUrl,
        [int]$pageLimit
    )
    $libraries = Get-DocumentLibraries -webUrl $webUrl -pageLimit $pageLimit
    $totalLibraries = $libraries.Count
    $processedLibraries = 0

    foreach ($library in $libraries) {
        $email = $library.EmailAlias
        if ($null -ne $email) {
            $owner = $library.Author
            $csvData = [PSCustomObject]@{
                SiteCollectionUrl = $siteUrl
                LibraryName       = $library.Title
                Owner             = $owner
                Email             = $email
            }
            $csvData | Export-Csv -Path $csvFile -Append -NoTypeInformation
        }
        
        $processedLibraries++
        Write-Progress -Activity "Processing Libraries" -Status "Processing $processedLibraries of $totalLibraries" -PercentComplete (($processedLibraries / $totalLibraries) * 100)
    }
}

# Main script
$webApp = Get-SPWebApplication $webAppUrl
$totalSites = $webApp.Sites.Count
$processedSites = 0

foreach ($site in $webApp.Sites) {
    $totalWebs = $site.AllWebs.Count
    $processedWebs = 0

    foreach ($web in $site.AllWebs) {
        try {
            Get-DocumentLibrariesInfo -webUrl $web.Url -siteUrl $site.Url -pageLimit $pageLimit
        }
        finally {
            $web.Dispose() # Dispose of the web object after use
        }
        
        $processedWebs++
        Write-Progress -Activity "Processing Webs" -Status "Processing $processedWebs of $totalWebs in site $($processedSites + 1) of $totalSites" -PercentComplete (($processedWebs / $totalWebs) * 100)
    }

    try {
        $site.Dispose() # Dispose of the site object after use
    }
    finally {
        $processedSites++
        Write-Progress -Activity "Processing Sites" -Status "Processing $processedSites of $totalSites" -PercentComplete (($processedSites / $totalSites) * 100)
    }
}

Write-Host "The output file has been saved to $csvFile" -ForegroundColor Green
