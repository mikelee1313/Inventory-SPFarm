<#
.SYNOPSIS
    Exports files from SharePoint On-premises document libraries or data/attachments from SharePoint lists with automatic detection of list type.

.DESCRIPTION
    This script automatically detects whether the target is a Document Library or a regular SharePoint List and processes accordingly:
    - Document Libraries: Downloads all files maintaining folder structure
    - SharePoint Lists: Exports list data to CSV and downloads any item attachments
    - InfoPath Forms Lists: Special handling for InfoPath XML forms with readable text extraction

    InfoPath Forms Support:
    - Automatically detects InfoPath forms by content type
    - Downloads the form XML data as form-data.xml
    - Creates a readable text version as form-data-readable.txt
    - Preserves all form field data in JSON metadata

    The script creates individual folders for each list item containing:
    - item-metadata.json (complete item data and field values)
    - attachments-info.json (attachment details, if any)
    - actual attachment files
    - form-data.xml (InfoPath form XML, if applicable)
    - form-data-readable.txt (readable XML text, if applicable)

.PARAMETER SiteUrl
    The URL of the SharePoint On-premises site. Default: "https://spwfe.contoso.local"

.PARAMETER LibraryNames
    Array of library or list names to export. Default: @("List1")

.PARAMETER DownloadPath
    Local path where exported files will be saved. Default: "C:\ExportedLibrary"

.PARAMETER Username
    Username for SharePoint authentication in domain\username format. Default: "contoso\spfarm"

.PARAMETER Password
    Password for SharePoint authentication. Default: "LS1setup!"

.EXAMPLE
    .\Export-SPOnPremList.ps1 -SiteUrl "https://sharepoint.contoso.com" -LibraryNames @("Documents", "CustomList") -DownloadPath "C:\Export"

.EXAMPLE
    .\Export-SPOnPremList.ps1 -LibraryNames @("InfoPathForms") -Username "domain\user" -Password "password123"

.NOTES
    Requires: SharePointPnPPowerShell2019 module (will auto-install if missing)
    
    The script handles:
    - Long file and folder names (truncation)
    - Special characters in file names
    - Multiple download methods for reliability
    - Path length limitations
    - Empty files and attachments
    - InfoPath form processing
    - Comprehensive error handling and logging

    Author: Mike Lee
    Date: 10/16/25

.LINK
    https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/
#>



param(
    [string]$SiteUrl = "https://spwfe.contoso.local",
    [string[]]$LibraryNames = @("List1"),
    [string]$DownloadPath = "C:\ExportedLibrary",
    [string]$Username = "contoso\spfarm",
    [string]$Password = "LS1setup!"
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

# Function to export list data as CSV
function Export-ListData {
    param(
        [string]$ListName,
        [string]$BasePath
    )
    
    Write-Host "Exporting list data for: $ListName" -ForegroundColor Cyan
    
    try {
        $csvPath = Join-Path $BasePath "$ListName-Data.csv"
        
        # Get all list items and their field values
        $items = Get-PnPListItem -List $ListName -PageSize 2000
        Write-Host "Found $($items.Count) items in list" -ForegroundColor Cyan
        
        if ($items.Count -gt 0) {
            # Get the list fields to determine columns
            $list = Get-PnPList -Identity $ListName
            $fields = Get-PnPField -List $list | Where-Object { $_.Hidden -eq $false -and $_.ReadOnlyField -eq $false }
            
            # Create array to hold export data
            $exportData = @()
            
            foreach ($item in $items) {
                $row = [PSCustomObject]@{
                    ID = $item.Id
                }
                
                foreach ($field in $fields) {
                    $fieldName = $field.InternalName
                    $value = $item.FieldValues[$fieldName]
                    
                    # Handle different field types
                    if ($null -ne $value) {
                        if ($value -is [System.Array]) {
                            $value = $value -join "; "
                        }
                        elseif ($value.GetType().Name -eq "FieldLookupValue") {
                            $value = $value.LookupValue
                        }
                        elseif ($value.GetType().Name -eq "FieldUserValue") {
                            $value = $value.LookupValue
                        }
                    }
                    
                    $row | Add-Member -MemberType NoteProperty -Name $field.Title -Value $value
                }
                
                $exportData += $row
            }
            
            # Export to CSV
            $exportData | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
            Write-Host "List data exported to: $csvPath" -ForegroundColor Green
        }
        else {
            Write-Warning "No items found in list $ListName"
        }
        
        # Process individual list items with their attachments and metadata
        Write-Host "Processing individual list items with attachments and metadata..." -ForegroundColor Cyan
        $attachmentCount = 0
        $itemsWithAttachmentsCount = 0
        
        $itemsFolder = Join-Path $BasePath "$ListName-Items"
        if (-not (Test-Path $itemsFolder)) {
            New-Item -ItemType Directory -Path $itemsFolder -Force | Out-Null
        }
        
        foreach ($item in $items) {
            # Create a folder for each list item (whether it has attachments or not)
            $itemTitle = $item.FieldValues["Title"]
            if ([string]::IsNullOrWhiteSpace($itemTitle)) {
                $itemTitle = "Item"
            }
            
            # Clean the title for use as folder name
            $safeItemTitle = $itemTitle -replace '[<>:"/\\|?*]', '_'
            $safeItemTitle = $safeItemTitle.Trim()
            if ($safeItemTitle.Length -gt 50) {
                $safeItemTitle = $safeItemTitle.Substring(0, 47) + "..."
            }
            
            $itemFolderName = "Item$($item.Id)-$safeItemTitle"
            $itemFolderPath = Join-Path $itemsFolder $itemFolderName
            
            if (-not (Test-Path $itemFolderPath)) {
                New-Item -ItemType Directory -Path $itemFolderPath -Force | Out-Null
            }
            
            # Create item metadata JSON file
            $itemMetadata = @{
                ID              = $item.Id
                Created         = $item.FieldValues["Created"]
                Modified        = $item.FieldValues["Modified"]
                Author          = if ($item.FieldValues["Author"]) { $item.FieldValues["Author"].LookupValue } else { $null }
                Editor          = if ($item.FieldValues["Editor"]) { $item.FieldValues["Editor"].LookupValue } else { $null }
                Title           = $itemTitle
                HasAttachments  = $item.FieldValues["Attachments"]
                ContentType     = $item.FieldValues["ContentType"]
                IsInfoPathForm  = $false
                InfoPathFormUrl = $null
                AllFields       = @{}
            }
            
            # Check if this is an InfoPath form
            $contentType = $item.FieldValues["ContentType"]
            if ($contentType -and ($contentType.ToString().Contains("Form") -or $contentType.ToString().Contains("InfoPath"))) {
                $itemMetadata.IsInfoPathForm = $true
                Write-Host "  This appears to be an InfoPath form" -ForegroundColor Cyan
            }
            
            # Add all field values to metadata
            foreach ($field in $fields) {
                $fieldName = $field.InternalName
                $value = $item.FieldValues[$fieldName]
                
                # Handle different field types for JSON serialization
                if ($null -ne $value) {
                    if ($value -is [System.Array]) {
                        $value = $value -join "; "
                    }
                    elseif ($value.GetType().Name -eq "FieldLookupValue") {
                        $value = @{
                            LookupId    = $value.LookupId
                            LookupValue = $value.LookupValue
                        }
                    }
                    elseif ($value.GetType().Name -eq "FieldUserValue") {
                        $value = @{
                            LookupId    = $value.LookupId
                            LookupValue = $value.LookupValue
                            Email       = $value.Email
                        }
                    }
                    elseif ($value -is [DateTime]) {
                        $value = $value.ToString("yyyy-MM-dd HH:mm:ss")
                    }
                }
                
                $itemMetadata.AllFields[$field.Title] = $value
            }
            
            # Try to download InfoPath form XML if this is an InfoPath form
            if ($itemMetadata.IsInfoPathForm) {
                try {
                    # Try to get the form XML file
                    $fileRef = $item.FieldValues["FileRef"]
                    if ($fileRef) {
                        Write-Host "  Attempting to download InfoPath form XML..." -ForegroundColor Cyan
                        
                        $formXmlPath = Join-Path $itemFolderPath "form-data.xml"
                        
                        try {
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
                                [System.IO.File]::WriteAllBytes($formXmlPath, $fileBytes)
                                Write-Host "  Downloaded InfoPath XML: form-data.xml ($($fileBytes.Length) bytes)" -ForegroundColor Gray
                                $itemMetadata.InfoPathFormUrl = $fileRef
                                
                                # Try to extract readable data from XML
                                try {
                                    $xmlContent = [System.Text.Encoding]::UTF8.GetString($fileBytes)
                                    $formReadablePath = Join-Path $itemFolderPath "form-data-readable.txt"
                                    $xmlContent | Set-Content -Path $formReadablePath -Encoding UTF8
                                    Write-Host "  Created readable XML text file: form-data-readable.txt" -ForegroundColor Gray
                                }
                                catch {
                                    Write-Host "  Could not create readable XML file: $($_.Exception.Message)" -ForegroundColor Yellow
                                }
                            }
                            else {
                                Write-Warning "  InfoPath form XML appears to be empty"
                            }
                        }
                        catch {
                            Write-Warning "  Failed to download InfoPath XML: $($_.Exception.Message)"
                        }
                    }
                }
                catch {
                    Write-Warning "  Failed to process InfoPath form: $($_.Exception.Message)"
                }
            }
            
            # Save item metadata as JSON
            $metadataPath = Join-Path $itemFolderPath "item-metadata.json"
            $itemMetadata | ConvertTo-Json -Depth 10 | Set-Content -Path $metadataPath -Encoding UTF8
            
            # Process attachments if they exist
            if ($item.FieldValues["Attachments"] -eq $true) {
                Write-Host "Processing item $($item.Id): '$itemTitle' with attachments..." -ForegroundColor DarkGray
                $itemsWithAttachmentsCount++
                
                try {
                    $attachments = Get-PnPProperty -ClientObject $item -Property "AttachmentFiles"
                    $itemAttachments = @()
                    
                    foreach ($attachment in $attachments) {
                        $attachmentUrl = $attachment.ServerRelativeUrl
                        $attachmentName = $attachment.FileName
                        
                        # Clean filename for Windows compatibility
                        $safeAttachmentName = $attachmentName -replace '[<>:"/\\|?*]', '_'
                        
                        $localAttachmentPath = Join-Path $itemFolderPath $safeAttachmentName
                        
                        try {
                            # Download the attachment
                            $web = Get-PnPWeb
                            $file = $web.GetFileByServerRelativeUrl($attachmentUrl)
                            $web.Context.Load($file)
                            $web.Context.ExecuteQuery()
                            
                            $fileStream = $file.OpenBinaryStream()
                            $web.Context.ExecuteQuery()
                            
                            $memoryStream = New-Object System.IO.MemoryStream
                            $fileStream.Value.CopyTo($memoryStream)
                            $fileBytes = $memoryStream.ToArray()
                            $memoryStream.Dispose()
                            
                            if ($fileBytes.Length -gt 0) {
                                [System.IO.File]::WriteAllBytes($localAttachmentPath, $fileBytes)
                                Write-Host "  Downloaded: $attachmentName ($($fileBytes.Length) bytes)" -ForegroundColor Gray
                                $attachmentCount++
                                
                                # Track attachment info for metadata
                                $itemAttachments += @{
                                    OriginalName      = $attachmentName
                                    SavedAs           = $safeAttachmentName
                                    Size              = $fileBytes.Length
                                    Downloaded        = $true
                                    ServerRelativeUrl = $attachmentUrl
                                }
                            }
                            else {
                                Write-Warning "  Attachment appears to be empty: $attachmentName"
                                $itemAttachments += @{
                                    OriginalName      = $attachmentName
                                    SavedAs           = $null
                                    Size              = 0
                                    Downloaded        = $false
                                    Error             = "File appears to be empty"
                                    ServerRelativeUrl = $attachmentUrl
                                }
                            }
                        }
                        catch {
                            Write-Warning "  Failed to download: $attachmentName - $($_.Exception.Message)"
                            $itemAttachments += @{
                                OriginalName      = $attachmentName
                                SavedAs           = $null
                                Size              = $null
                                Downloaded        = $false
                                Error             = $_.Exception.Message
                                ServerRelativeUrl = $attachmentUrl
                            }
                        }
                    }
                    
                    # Save attachment details to a separate file
                    if ($itemAttachments.Count -gt 0) {
                        $attachmentDetailsPath = Join-Path $itemFolderPath "attachments-info.json"
                        $itemAttachments | ConvertTo-Json -Depth 5 | Set-Content -Path $attachmentDetailsPath -Encoding UTF8
                    }
                }
                catch {
                    Write-Warning "Failed to process attachments for item $($item.Id): $($_.Exception.Message)"
                }
            }
            else {
                Write-Host "Processing item $($item.Id): '$itemTitle' (no attachments)" -ForegroundColor DarkGray
            }
        }
        
        # Count InfoPath forms
        $infoPathFormsCount = 0
        foreach ($item in $items) {
            $contentType = $item.FieldValues["ContentType"]
            if ($contentType -and ($contentType.ToString().Contains("Form") -or $contentType.ToString().Contains("InfoPath"))) {
                $infoPathFormsCount++
            }
        }
        
        Write-Host "`nExport Summary:" -ForegroundColor Green
        Write-Host "- Created individual folders for $($items.Count) list items" -ForegroundColor Green
        Write-Host "- $itemsWithAttachmentsCount items had attachments" -ForegroundColor Green
        Write-Host "- Downloaded $attachmentCount attachment files" -ForegroundColor Green
        if ($infoPathFormsCount -gt 0) {
            Write-Host "- $infoPathFormsCount InfoPath forms detected and processed" -ForegroundColor Cyan
        }
        Write-Host "- Each item includes:" -ForegroundColor Yellow
        Write-Host "  • item-metadata.json (complete item data)" -ForegroundColor Yellow
        Write-Host "  • attachments-info.json (attachment details, if any)" -ForegroundColor Yellow
        Write-Host "  • actual attachment files" -ForegroundColor Yellow
        if ($infoPathFormsCount -gt 0) {
            Write-Host "  • form-data.xml (InfoPath form XML, if applicable)" -ForegroundColor Yellow
            Write-Host "  • form-data-readable.txt (readable XML text, if applicable)" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Error "Failed to export list data: $($_.Exception.Message)"
    }
}

# Function to download files from a library with folder structure
function Download-LibraryFiles {
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
        # First, check if this is actually a document library or a regular list
        $list = Get-PnPList -Identity $LibraryName
        Write-Host "List Type: $($list.BaseType)" -ForegroundColor Cyan
        Write-Host "List Template: $($list.BaseTemplate)" -ForegroundColor Cyan
        
        if ($list.BaseType -ne "DocumentLibrary") {
            Write-Warning "Warning: '$LibraryName' is not a Document Library (BaseType: $($list.BaseType)). It's a regular SharePoint list."
            Write-Host "Switching to list data export mode instead of file download mode." -ForegroundColor Yellow
            
            # Export list data and attachments instead
            Export-ListData -ListName $LibraryName -BasePath $BasePath
            return
        }
        
        # Get all items in the library including folders
        $items = Get-PnPListItem -List $LibraryName -PageSize 2000
        $downloadCount = 0
        
        foreach ($item in $items) {
            $fileType = $item.FieldValues["FSObjType"]
            
            # Check if it's a file (FSObjType = 0) or folder (FSObjType = 1)
            if ($fileType -eq 0) {
                # It's a file - but let's verify it's actually a document, not just a list item
                try {
                    $fileRef = $item.FieldValues["FileRef"]
                    $fileName = $item.FieldValues["FileLeafRef"]
                    $fileDirRef = $item.FieldValues["FileDirRef"]
                    
                    # Check if this is actually a document with file content
                    # List items might show as files but not have actual file content
                    if ($null -eq $fileRef -or $fileRef -eq "") {
                        Write-Warning "Skipping item with no file reference: $($item.Id)"
                        continue
                    }
                    
                    # Check if the item has attachments or is an actual file
                    $hasAttachments = $item.FieldValues["Attachments"]
                    if ($hasAttachments -eq $false -and $fileName -match '^\d+_\.\d+$') {
                        Write-Warning "Skipping list item (not a file): $fileName - This appears to be a list item ID, not a file"
                        continue
                    }
                    
                    Write-Host "Processing file: $fileName from $fileDirRef" -ForegroundColor DarkGray
                    
                    # Add diagnostic information
                    Write-Host "  File Reference: $fileRef" -ForegroundColor DarkCyan
                    Write-Host "  File Size: $($item.FieldValues['File_x0020_Size']) bytes" -ForegroundColor DarkCyan
                    Write-Host "  Content Type: $($item.FieldValues['ContentType'])" -ForegroundColor DarkCyan
                    Write-Host "  Has File: $($item.FieldValues.ContainsKey('File'))" -ForegroundColor DarkCyan
                    
                    # Skip if file size is 0 or null
                    $fileSize = $item.FieldValues['File_x0020_Size']
                    if ($null -eq $fileSize -or $fileSize -eq 0) {
                        Write-Warning "Skipping zero-byte or null file: $fileName"
                        continue
                    }
                    
                    # Create the local folder structure based on SharePoint folder path
                    # Remove the site and library portion from the path
                    $siteRelativePath = $fileDirRef
                    if ($siteRelativePath -match "/Lists/$LibraryName") {
                        $relativePath = $siteRelativePath -replace "^.*/Lists/$LibraryName/?", ""
                    }
                    elseif ($siteRelativePath -match "/$LibraryName") {
                        $relativePath = $siteRelativePath -replace "^.*/$LibraryName/?", ""
                    }
                    else {
                        $relativePath = ""
                    }
                    
                    # Handle long paths by truncating folder names if needed
                    if ($relativePath) {
                        # Split path into parts and truncate if necessary
                        $pathParts = $relativePath -split '/'
                        $truncatedParts = @()
                        foreach ($part in $pathParts) {
                            if ($part.Length -gt 50) {
                                # Truncate long folder names but keep meaningful parts
                                $truncatedPart = $part.Substring(0, 47) + "..."
                                $truncatedParts += $truncatedPart
                            }
                            else {
                                $truncatedParts += $part
                            }
                        }
                        $relativePath = $truncatedParts -join '/'
                        $localFolderPath = Join-Path $LibraryPath $relativePath
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
                    
                    # Create local directory if it doesn't exist
                    if (-not (Test-Path $localFolderPath)) {
                        New-Item -ItemType Directory -Path $localFolderPath -Force | Out-Null
                        Write-Host "Created local folder: $localFolderPath" -ForegroundColor Yellow
                    }
                    
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
                    
                    # Final attempt: Skip certain file types that might cause issues
                    if (-not $downloadSuccess) {
                        $fileExtension = [System.IO.Path]::GetExtension($fileName).ToLower()
                        if ($fileExtension -in @('.aspx', '.master', '.ascx')) {
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

# Process each library/list
foreach ($LibraryName in $LibraryNames) {
    Write-Host "`nProcessing: $LibraryName" -ForegroundColor Magenta
    
    try {
        # Check if the list exists
        $list = Get-PnPList -Identity $LibraryName -ErrorAction Stop
        
        if ($list.BaseType -eq "DocumentLibrary") {
            Write-Host "Processing as Document Library..." -ForegroundColor Green
            Download-LibraryFiles -LibraryName $LibraryName -BasePath $DownloadPath
        }
        else {
            Write-Host "Processing as SharePoint List..." -ForegroundColor Green
            Export-ListData -ListName $LibraryName -BasePath $DownloadPath
        }
    }
    catch {
        Write-Error "Failed to process '$LibraryName': $($_.Exception.Message)"
        Write-Host "Available lists on this site:" -ForegroundColor Yellow
        try {
            Get-PnPList | Select-Object Title, BaseType, ItemCount | Format-Table
        }
        catch {
            Write-Warning "Could not retrieve list of available lists"
        }
    }
}

Disconnect-PnPOnline
Write-Host "Export complete. All files saved to $DownloadPath" -ForegroundColor Yellow
