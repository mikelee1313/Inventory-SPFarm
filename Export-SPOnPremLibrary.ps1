<#
.SYNOPSIS
    Exports all files from SharePoint On-premises document libraries to a local folder structure.

.DESCRIPTION
    This script connects to a SharePoint On-premises site and downloads all files from specified document libraries
    while preserving the folder structure. It handles various edge cases including long file names, long paths,
    special characters, and different file types. The script uses multiple download methods as fallbacks to ensure
    maximum compatibility and success rate.

.PARAMETER SiteUrl
    The URL of the SharePoint On-premises site to connect to.
    Default: "https://spwfe.contoso.local"

.PARAMETER LibraryNames
    An array of document library names to export from.
    Default: @("Shared Documents", "wikipages")

.PARAMETER DownloadPath
    The local directory path where exported files will be saved.
    Default: "C:\ExportedLibrary"

.PARAMETER Username
    The username for SharePoint authentication in domain\username format.
    Default: "contoso\spfarm"

.PARAMETER Password
    The password for SharePoint authentication. Should be provided securely.
    Default: "" (empty string)

.EXAMPLE
    .\Export-SPOnPremLibrary.ps1
    Runs the script with default parameters to export from the default libraries.

.EXAMPLE
    .\Export-SPOnPremLibrary.ps1 -SiteUrl "https://mysharepoint.company.com" -LibraryNames @("Documents", "Shared Documents") -DownloadPath "C:\MyExport" -Username "domain\myuser" -Password "mypassword"
    Exports specific libraries from a custom SharePoint site to a custom local path.

.NOTES
    - Requires SharePointPnPPowerShell2019 module (will auto-install if missing)
    - Handles long file names by truncating them to avoid filesystem limitations
    - Creates simplified folder structures for extremely long paths
    - Uses multiple download methods for maximum compatibility
    - Preserves original folder structure from SharePoint
    - Skips files that already exist locally
    - Provides detailed progress information and error handling

    Authored by: Mike Lee
    Date: 10/3/2025

.FUNCTIONALITY
    - Auto-installs required PowerShell modules
    - Connects to SharePoint On-premises using credentials
    - Downloads files while preserving folder structure
    - Handles file and path length limitations
    - Provides multiple fallback download methods
    - Creates comprehensive logging output
    - Disconnects cleanly after completion
#>
# Export-SPOnPremLibrary.ps1
# Export all files from SharePoint On-premises document libraries to a local folder
# Requires: SharePointPnPPowerShell2019 module

param(
    [string]$SiteUrl = "https://spwfe.contoso.local",
    [string[]]$LibraryNames = @("Shared Documents", "wikipages"),
    [string]$DownloadPath = "C:\ExportedLibrary",
    [string]$Username = "contoso\spfarm",
    [string]$Password = ""
)

# Check if SharePointPnPPowerShell2019 module is installed
if (-not (Get-Module -ListAvailable -Name SharePointPnPPowerShell2019)) {
    Write-Host "SharePointPnPPowerShell2019 module not found. Installing..." -ForegroundColor Yellow
    try {
        Install-Module SharePointPnPPowerShell2019 -Force -AllowClobber -Scope CurrentUser
        Write-Host "SharePointPnPPowerShell2019 module installed successfully." -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to install SharePointPnPPowerShell2019 module: $($_.Exception.Message)"
        exit 1
    }
}
else {
    Write-Host "SharePointPnPPowerShell2019 module is already installed." -ForegroundColor Green
}

# Import module
Import-Module SharePointPnPPowerShell2019 -Force -Verbose

# Create credential object
$SecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential($Username, $SecurePassword)

# Function to download files from a library with folder structure
function Get-LibraryFiles {
    param(
        [string]$LibraryName,
        [string]$BasePath
    )
    
    Write-Host "Starting download from library: $LibraryName" -ForegroundColor Cyan
    
    # Create library-specific folder
    $LibraryPath = Join-Path $BasePath $LibraryName
    if (-not (Test-Path $LibraryPath)) {
        New-Item -ItemType Directory -Path $LibraryPath -Force | Out-Null
        Write-Host "Created folder: $LibraryPath" -ForegroundColor Green
    }
    
    try {
        # Get all items in the library including folders
        $items = Get-PnPListItem -List $LibraryName -PageSize 2000
        $downloadCount = 0
        
        foreach ($item in $items) {
            $fileType = $item.FieldValues["FSObjType"]
            
            # Check if it's a file (FSObjType = 0) or folder (FSObjType = 1)
            if ($fileType -eq 0) {
                # It's a file
                try {
                    $fileRef = $item.FieldValues["FileRef"]
                    $fileName = $item.FieldValues["FileLeafRef"]
                    $fileDirRef = $item.FieldValues["FileDirRef"]
                    
                    Write-Host "Processing file: $fileName from $fileDirRef" -ForegroundColor DarkGray
                    
                    # Create the local folder structure based on SharePoint folder path
                    # Remove the site and library portion from the path
                    $siteRelativePath = $fileDirRef
                    if ($siteRelativePath -match "/Lists/$LibraryName") {
                        $relativePath = $siteRelativePath -replace "^.*/Lists/$LibraryName/?", ""
                    }
                    elseif ($siteRelativePath -match "/$LibraryName") {
                        $relativePath = $siteRelativePath -replace "^.*/$LibraryName/?", ""
                    }
                    elseif ($siteRelativePath -match "/SiteAssets") {
                        # Handle Site Assets library specifically
                        $relativePath = $siteRelativePath -replace "^.*/SiteAssets/?", ""
                    }
                    else {
                        $relativePath = ""
                    }
                    
                    # Handle long paths by truncating folder names if needed
                    if ($relativePath) {
                        # Split path into parts and truncate if necessary
                        $pathParts = $relativePath -split '/'
                        $truncatedParts = @()
                        $pathTooLong = $false
                        
                        foreach ($part in $pathParts) {
                            if ($part.Length -gt 50) {
                                # Truncate long folder names but use safe characters only
                                $truncatedPart = $part.Substring(0, 47) + "_TR"
                                $truncatedParts += $truncatedPart
                                $pathTooLong = $true
                            }
                            else {
                                $truncatedParts += $part
                            }
                        }
                        
                        $relativePath = $truncatedParts -join '/'
                        $localFolderPath = Join-Path $LibraryPath $relativePath
                        
                        # If we had to truncate, try to create the path, but fall back if it's still too complex
                        if ($pathTooLong) {
                            # Test if the full path would be too long
                            $testPath = Join-Path $localFolderPath "test.txt"
                            if ($testPath.Length -gt 240) {
                                Write-Host "Path still too long after truncation, using simplified structure" -ForegroundColor Yellow
                                $simplePath = Join-Path $LibraryPath "LongPaths"
                                if (-not (Test-Path $simplePath)) {
                                    New-Item -ItemType Directory -Path $simplePath -Force | Out-Null
                                }
                                $localFolderPath = $simplePath
                            }
                        }
                        
                        # Ensure the folder path exists
                        if (-not (Test-Path $localFolderPath)) {
                            try {
                                New-Item -ItemType Directory -Path $localFolderPath -Force | Out-Null
                                Write-Host "Created folder path: $localFolderPath" -ForegroundColor Yellow
                            }
                            catch {
                                Write-Warning "Failed to create folder path: $localFolderPath - $($_.Exception.Message)"
                                # Fall back to a simpler path structure
                                $simplePath = Join-Path $LibraryPath "ComplexPaths"
                                if (-not (Test-Path $simplePath)) {
                                    New-Item -ItemType Directory -Path $simplePath -Force | Out-Null
                                }
                                $localFolderPath = $simplePath
                                Write-Host "Using fallback folder: $localFolderPath" -ForegroundColor Yellow
                            }
                        }
                    }
                    else {
                        $localFolderPath = $LibraryPath
                    }
                    
                    # Handle long file names
                    $originalFileName = $fileName
                    if ($fileName.Length -gt 100) {
                        $extension = [System.IO.Path]::GetExtension($fileName)
                        $nameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
                        $truncatedName = $nameWithoutExt.Substring(0, 95) + "..." + $extension
                        $fileName = $truncatedName
                        Write-Host "Truncated long filename: $originalFileName -> $fileName" -ForegroundColor Yellow
                    }
                    
                    # Check total path length and use alternative approach if too long
                    $localFilePath = Join-Path $localFolderPath $fileName
                    $useShortPath = $false
                    
                    if ($localFilePath.Length -gt 250) {
                        Write-Host "Path too long ($($localFilePath.Length) chars), using short path approach" -ForegroundColor Yellow
                        # Create a shorter path structure
                        $shortLibraryPath = Join-Path $BasePath ($LibraryName.Substring(0, [Math]::Min(10, $LibraryName.Length)))
                        if (-not (Test-Path $shortLibraryPath)) {
                            New-Item -ItemType Directory -Path $shortLibraryPath -Force | Out-Null
                        }
                        $localFolderPath = $shortLibraryPath
                        $localFilePath = Join-Path $localFolderPath $fileName
                        $useShortPath = $true
                    }
                    
                    # Create local directory if it doesn't exist (with additional safety check)
                    if (-not (Test-Path $localFolderPath)) {
                        try {
                            New-Item -ItemType Directory -Path $localFolderPath -Force | Out-Null
                            Write-Host "Created local folder: $localFolderPath" -ForegroundColor Yellow
                        }
                        catch {
                            Write-Warning "Failed to create directory: $localFolderPath - $($_.Exception.Message)"
                            # Use library root as fallback
                            $localFolderPath = $LibraryPath
                            Write-Host "Using library root as fallback: $localFolderPath" -ForegroundColor Yellow
                        }
                    }
                    
                    # Update the file path with the final folder path
                    $localFilePath = Join-Path $localFolderPath $fileName
                    
                    # Try different approaches to download the file
                    $downloadSuccess = $false
                    
                    # Method 1: Use file stream approach (most reliable for path control)
                    try {
                        # Check if file already exists and skip if so
                        if (Test-Path $localFilePath) {
                            Write-Host "File already exists, skipping: $fileName" -ForegroundColor Yellow
                            $downloadSuccess = $true
                            $downloadCount++
                        }
                        else {
                            # Use the file stream method for better control
                            $web = Get-PnPWeb
                            $file = $web.GetFileByServerRelativeUrl($fileRef)
                            $web.Context.Load($file)
                            $web.Context.ExecuteQuery()
                            
                            # Get file stream
                            $fileStream = $file.OpenBinaryStream()
                            $web.Context.ExecuteQuery()
                            
                            # Create memory stream and copy data
                            $memoryStream = New-Object System.IO.MemoryStream
                            $fileStream.Value.CopyTo($memoryStream)
                            $fileBytes = $memoryStream.ToArray()
                            $memoryStream.Dispose()
                            
                            if ($fileBytes.Length -gt 0) {
                                [System.IO.File]::WriteAllBytes($localFilePath, $fileBytes)
                                Write-Host "Downloaded (Method 1 - Stream): $fileName" -ForegroundColor Gray
                                $downloadSuccess = $true
                                $downloadCount++
                            }
                            else {
                                Write-Warning "File appears to be empty: $fileName"
                            }
                        }
                    }
                    catch {
                        Write-Host "Method 1 failed for $fileName`: $($_.Exception.Message)" -ForegroundColor DarkRed
                    }
                    
                    # Method 2: Try with URL encoding for special characters
                    if (-not $downloadSuccess) {
                        try {
                            # URL encode the file reference to handle special characters
                            $encodedFileRef = [System.Web.HttpUtility]::UrlPathEncode($fileRef)
                            
                            $web = Get-PnPWeb
                            $file = $web.GetFileByServerRelativeUrl($encodedFileRef)
                            $web.Context.Load($file)
                            $web.Context.ExecuteQuery()
                            
                            $fileStream = $file.OpenBinaryStream()
                            $web.Context.ExecuteQuery()
                            
                            $memoryStream = New-Object System.IO.MemoryStream
                            $fileStream.Value.CopyTo($memoryStream)
                            $fileBytes = $memoryStream.ToArray()
                            $memoryStream.Dispose()
                            
                            if ($fileBytes.Length -gt 0) {
                                [System.IO.File]::WriteAllBytes($localFilePath, $fileBytes)
                                Write-Host "Downloaded (Method 2 - URL Encoded): $fileName" -ForegroundColor Gray
                                $downloadSuccess = $true
                                $downloadCount++
                            }
                            else {
                                Write-Warning "Encoded file appears to be empty: $fileName"
                            }
                        }
                        catch {
                            Write-Host "Method 2 failed for $fileName`: $($_.Exception.Message)" -ForegroundColor DarkRed
                        }
                    }
                    
                    # Method 3: Alternative file stream approach
                    if (-not $downloadSuccess) {
                        try {
                            # Try alternative CSOM approach
                            $web = Get-PnPWeb
                            $file = $web.GetFileByServerRelativeUrl($fileRef)
                            $web.Context.Load($file)
                            $web.Context.ExecuteQuery()
                            
                            $fileStream = $file.OpenBinaryStream()
                            $web.Context.ExecuteQuery()
                            
                            $memoryStream = New-Object System.IO.MemoryStream
                            $fileStream.Value.CopyTo($memoryStream)
                            $fileBytes = $memoryStream.ToArray()
                            $memoryStream.Dispose()
                            
                            if ($fileBytes.Length -gt 0) {
                                [System.IO.File]::WriteAllBytes($localFilePath, $fileBytes)
                                Write-Host "Downloaded (Method 3 - Alternative Stream): $fileName" -ForegroundColor Gray
                                $downloadCount++
                                $downloadSuccess = $true
                            }
                            else {
                                Write-Warning "File appears to be empty: $fileName"
                            }
                        }
                        catch {
                            Write-Host "Method 3 failed for $fileName`: $($_.Exception.Message)" -ForegroundColor DarkRed
                        }
                    }
                    
                    # Method 4: Alternative approach using relative URL without site collection
                    if (-not $downloadSuccess) {
                        try {
                            # Try with just the relative part of the URL
                            $relativeUrl = $fileRef
                            if ($relativeUrl.Contains('/sites/')) {
                                $relativeUrl = $relativeUrl.Substring($relativeUrl.IndexOf('/sites/'))
                            }
                            
                            $web = Get-PnPWeb
                            $file = $web.GetFileByServerRelativeUrl($relativeUrl)
                            $web.Context.Load($file)
                            $web.Context.ExecuteQuery()
                            
                            $fileStream = $file.OpenBinaryStream()
                            $web.Context.ExecuteQuery()
                            
                            $memoryStream = New-Object System.IO.MemoryStream
                            $fileStream.Value.CopyTo($memoryStream)
                            $fileBytes = $memoryStream.ToArray()
                            $memoryStream.Dispose()
                            
                            if ($fileBytes.Length -gt 0) {
                                [System.IO.File]::WriteAllBytes($localFilePath, $fileBytes)
                                Write-Host "Downloaded (Method 4 - Relative): $fileName" -ForegroundColor Gray
                                $downloadSuccess = $true
                                $downloadCount++
                            }
                            else {
                                Write-Warning "Relative URL file appears to be empty: $fileName"
                            }
                        }
                        catch {
                            Write-Host "Method 4 failed for $fileName`: $($_.Exception.Message)" -ForegroundColor DarkRed
                        }
                    }
                    
                    # Final attempt: Skip certain file types that might cause issues (but allow ASPX files)
                    if (-not $downloadSuccess) {
                        $fileExtension = [System.IO.Path]::GetExtension($fileName).ToLower()
                        if ($fileExtension -in @('.master', '.ascx')) {
                            Write-Warning "Skipping potentially problematic file type: $fileName ($fileExtension)"
                        }
                        elseif ($useShortPath) {
                            Write-Warning "Failed to download even with short path: $fileName (Original: $originalFileName)"
                        }
                        else {
                            Write-Warning "All download methods failed for file: $fileName"
                        }
                    }
                }
                catch {
                    Write-Warning "Failed to process file item: $($_.Exception.Message)"
                }
            }
        }
        
        Write-Host "Completed download from '$LibraryName': $downloadCount files downloaded" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to access library '$LibraryName': $($_.Exception.Message)"
    }
}

# Connect to SharePoint On-premises
Connect-PnPOnline -Url $SiteUrl -Credentials $Cred
Write-Host "Connected to SharePoint site: $SiteUrl" -ForegroundColor Green

# Create base download directory
if (-not (Test-Path $DownloadPath)) {
    New-Item -ItemType Directory -Path $DownloadPath -Force | Out-Null
    Write-Host "Created base download folder: $DownloadPath" -ForegroundColor Green
}

# Download files from each library
foreach ($LibraryName in $LibraryNames) {
    Get-LibraryFiles -LibraryName $LibraryName -BasePath $DownloadPath
}

Disconnect-PnPOnline
Write-Host "Export complete. All files saved to $DownloadPath" -ForegroundColor Yellow
