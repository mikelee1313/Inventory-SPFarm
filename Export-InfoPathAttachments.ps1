<#
.SYNOPSIS
  Extracts file attachments from InfoPath XML forms and saves them to the local file system.

.DESCRIPTION
  This script processes InfoPath XML files to extract embedded base64-encoded file attachments.
  It can process individual XML files or all XML files in a specified folder (with optional recursion).
  The script handles both InfoPath attachments with headers and basic base64-encoded content without headers.
  
  The script automatically detects base64-encoded content by analyzing text nodes for:
  - Minimum length requirements
  - Base64 format validation
  - InfoPath attachment header presence
  
  Extracted attachments are saved with descriptive names that include the XML node name and original filename.
  If multiple attachments with the same name exist, the script automatically generates unique filenames.

.PARAMETER InputFolder
  Specifies the folder containing InfoPath XML files to process. 
  Overrides the default configuration when specified.
  The folder must exist or the script will throw an error.

.PARAMETER InfoPathForm
  Specifies one or more individual InfoPath XML files to process.
  Accepts pipeline input and overrides the default configuration when specified.
  Each file must exist or the script will throw an error.

.PARAMETER OutputFolder
  Specifies the base folder where extracted attachments will be saved.
  Overrides the default configuration when specified.
  If not specified, attachments are saved in subfolders next to each XML file.

.PARAMETER BasicFileName
  Specifies the default filename for attachments that don't have InfoPath headers.
  Overrides the default configuration when specified.
  Default is 'uploadedImage.jpg'.

.PARAMETER Recurse
  When specified with InputFolder, includes subdirectories in the processing.
  Only applies when processing a folder of XML files.

.INPUTS
  System.String[]
  You can pipe file paths to this script via the InfoPathForm parameter.

.OUTPUTS
  None
  The script outputs verbose information about processing and saves attachments to disk.
  A summary report is displayed at the end showing processing statistics.

.EXAMPLE
  .\Export-InfoPathAttachments.ps1
  
  Processes XML files using the default configuration settings defined in the script.

.EXAMPLE
  .\Export-InfoPathAttachments.ps1 -InputFolder "C:\InfoPathForms" -OutputFolder "C:\Attachments"
  
  Processes all XML files in C:\InfoPathForms and saves attachments to C:\Attachments.

.EXAMPLE
  .\Export-InfoPathAttachments.ps1 -InfoPathForm "C:\form1.xml", "C:\form2.xml"
  
  Processes specific XML files and extracts their attachments.

.EXAMPLE
  Get-ChildItem "C:\Forms\*.xml" | .\Export-InfoPathAttachments.ps1 -OutputFolder "C:\Extracted"
  
  Uses pipeline input to process XML files and save attachments to a specific folder.

.EXAMPLE
  .\Export-InfoPathAttachments.ps1 -InputFolder "C:\InfoPathForms" -Recurse -BasicFileName "unknown_file.bin"
  
  Processes XML files recursively and uses a custom filename for attachments without headers.

.NOTES
  File Name      : Export-InfoPathAttachments.ps1
  Prerequisite   : PowerShell 5x or later
  
  This script uses the System.IO namespace for efficient file operations.
  
  Configuration can be customized by modifying variables in the CONFIGURATION region:
  - $DefaultInputFolder: Default folder to process
  - $ProcessSubdirectories: Whether to include subdirectories
  - $DefaultInfoPathForms: Individual files to process if no folder is specified
  - $DefaultAttachmentFolder: Where to save extracted attachments
  - $DefaultBasicFileName: Default name for headerless attachments
  - $EnableVerboseOutput: Whether to show detailed processing information
  
  The script maintains statistics during processing and provides a comprehensive summary
  including the number of files processed, attachments extracted, and any errors encountered.

.LINK
  https://docs.microsoft.com/en-us/powershell/
#>

using namespace System.IO

[CmdletBinding()]
Param(
  # Folder containing InfoPath XML files to process (overrides configuration if specified)
  [Parameter()]
  [ValidateScript({
      if (-not (Test-Path -Path $_ -PathType Container)) { throw "Folder '$_' does not exist" }
      $true
    })]
  [string]$InputFolder,

  # The InfoPath form to extract attachments from (overrides configuration if specified)
  [Parameter(ValueFromPipeline)]
  [ValidateScript({
      if (-not (Test-Path -Path $_ -PathType Leaf)) { throw "File '$_' does not exist" }
      $true
    })]
  [string[]]$InfoPathForm,

  # A base folder to store attachments in (overrides configuration if specified)
  [Parameter()]
  [string]$OutputFolder,

  # File name for attachments that don't have the InfoPath attachment header (overrides configuration if specified)
  [Parameter()]
  [string]$BasicFileName,

  # Include subdirectories when processing InputFolder
  [Parameter()]
  [switch]$Recurse
)

#region CONFIGURATION
# =====================================================================================
# Modify these variables to customize the script behavior without using command-line switches
# =====================================================================================

# Default folder containing InfoPath XML files to process
# Set to a folder path to process all XML files in that folder
# Examples:
#   $DefaultInputFolder = "C:\path\to\forms"  # Process all XML files in this folder
#   $DefaultInputFolder = ""                  # Use individual file specification instead
$DefaultInputFolder = "C:\xmlfiles\infopath_xml"

# Include subdirectories when processing a folder
$ProcessSubdirectories = $true

# Alternative: Individual InfoPath form file(s) to process (used if $DefaultInputFolder is empty)
# Can be a single file path or an array of file paths
# Examples:
#   $DefaultInfoPathForms = "C:\path\to\single-file.xml"
#   $DefaultInfoPathForms = @("C:\path\to\file1.xml", "C:\path\to\file2.xml")
$DefaultInfoPathForms = @()

# Base folder to store extracted attachments
# If left empty, attachments will be saved in a folder next to each XML file
$DefaultAttachmentFolder = "C:\xmlfiles\infopath_xml\attachments"

# Default filename for attachments that don't have InfoPath headers
$DefaultBasicFileName = 'uploadedImage.jpg'

# Enable verbose output (shows detailed processing information)
$EnableVerboseOutput = $true

#endregion CONFIGURATION

#region SCRIPT INITIALIZATION
# Initialize error tracking
$script:erroredFiles = @{}
$script:processedFiles = 0
$script:extractedAttachments = 0

# Set verbose preference based on configuration
if ($EnableVerboseOutput) {
  $VerbosePreference = 'Continue'
}

# Determine input source: folder or individual files
if ($PSBoundParameters.ContainsKey('InputFolder')) {
  # Use provided input folder
  $sourceFolder = $InputFolder
  $recurseFiles = $Recurse.IsPresent
}
elseif ($DefaultInputFolder -and (Test-Path $DefaultInputFolder)) {
  # Use configured input folder
  $sourceFolder = $DefaultInputFolder
  $recurseFiles = $ProcessSubdirectories
}
else {
  # Use individual files
  $sourceFolder = $null
}

# Get list of files to process
if ($sourceFolder) {
  Write-Verbose "Processing XML files from folder: $sourceFolder"
  if ($recurseFiles) {
    $FilesToProcess = Get-ChildItem -Path $sourceFolder -Filter "*.xml" -File -Recurse | ForEach-Object { $_.FullName }
  }
  else {
    $FilesToProcess = Get-ChildItem -Path $sourceFolder -Filter "*.xml" -File | ForEach-Object { $_.FullName }
  }
  Write-Verbose "Found $($FilesToProcess.Count) XML files to process"
}
elseif ($PSBoundParameters.ContainsKey('InfoPathForm')) {
  # Use provided individual files
  $FilesToProcess = $InfoPathForm
}
else {
  # Use configured individual files
  $FilesToProcess = $DefaultInfoPathForms
}

# Set output folder
if ($PSBoundParameters.ContainsKey('OutputFolder')) {
  $CreateAttachmentsFolderIn = $OutputFolder
}
else {
  $CreateAttachmentsFolderIn = $DefaultAttachmentFolder
}

# Set basic filename
if ($PSBoundParameters.ContainsKey('BasicFileName')) {
  $BasicFileNameToUse = $BasicFileName
}
else {
  $BasicFileNameToUse = $DefaultBasicFileName
}

# Create output folder if it doesn't exist
if ($CreateAttachmentsFolderIn -and -not (Test-Path $CreateAttachmentsFolderIn)) {
  New-Item -Path $CreateAttachmentsFolderIn -ItemType Directory -Force | Out-Null
  Write-Verbose "Created output folder: $CreateAttachmentsFolderIn"
}

# Validate configuration
Write-Verbose "Configuration loaded:"
Write-Verbose "  Source: $(if($sourceFolder){"Folder: $sourceFolder"}else{"Individual files"})"
Write-Verbose "  Files to Process: $($FilesToProcess.Count) files"
Write-Verbose "  Output Folder: $CreateAttachmentsFolderIn"
Write-Verbose "  Basic File Name: $BasicFileNameToUse"
Write-Verbose "  Recurse Subdirectories: $recurseFiles"
#endregion SCRIPT INITIALIZATION

#region MAIN PROCESSING
# Process each InfoPath XML file
foreach ($formFile in $FilesToProcess) {
  $script:processedFiles++
  
  Write-Verbose "[$script:processedFiles/$($FilesToProcess.Count)] Processing: $formFile"
  
  # Validate file exists
  if (-not (Test-Path -Path $formFile -PathType Leaf)) {
    Write-Warning "File not found: $formFile"
    continue
  }
  
  $formPath = (Resolve-Path -Path $formFile).Path
  $formName = Split-Path -Leaf -Path $formPath

  Write-Verbose "Checking if $formName is valid XML"
  try { 
    [xml]$xml = (Get-Content -Path $formPath).Replace("§", "") 
  }
  catch {
    Write-Warning "$formFile isn't valid XML: $($_.Exception.Message)"
    $script:erroredFiles[$formName] = "Invalid XML: $($_.Exception.Message)"
    continue
  }

  Write-Verbose 'Resetting the attachment folder variable'
  $attachmentFolder = $null

  # Fastest way to get through this file (without using XMLStreamReader) is to filter only text nodes
  foreach ($textNode in $xml.SelectNodes("//*[text()]")) {

    $text = $textNode.InnerText

    # Several easy qualifiers for confirming that a text node isn't a base64 encoded string
    if ($text.length -le 100) { continue }
    if (($text.length % 4) -ne 0) { continue }
    if ($text.indexOf(" ") -ne -1) { continue }
    if ($text -match "http(s?)\:\/\/.*") { continue }

    try { 
      $bytes = [Convert]::FromBase64String($text) 
    }
    catch {
      # Not a valid base64 string, continue to next node
      continue
    }

    if ($bytes.length -eq 0) { continue }

    # When the attachment is broken into byte strings, the 20th byte tells you how many bytes are
    # used for the filename. Multiply by 2 for Unicode encoding
    $fileNameByteLen = $bytes[20] * 2

    # Handle attachments *without* an InfoPath attachment header
    if ($bytes[0] -ne 199 -or $bytes[1] -ne 73 -or $bytes[2] -ne 70 -or $bytes[3] -ne 65) {
      Write-Verbose "[$formName] Found an attachment without an InfoPath header, saving as $BasicFileNameToUse"
      $fileName = $BasicFileNameToUse
      $arrFileContentBytes = $bytes
    }
    # Handle attachments *with* an InfoPath attachment header
    else {
      # The header is 24 bytes long for InfoPath attachments
      $fileByteHeader = 24

      # Extract the bytes containing the filename
      $arrFileNameBytes = for ($i = 0; $i -lt $fileNameByteLen; $i++) {
        $bytes[$fileByteHeader + $i]
      }

      # Convert the filename bytes to a string
      try { 
        $fileName = [System.Text.Encoding]::Unicode.GetString($arrFileNameBytes) 
      }
      catch {
        $script:erroredFiles[$formName] = "Failed to decode filename: $($_.Exception.Message)"
        continue
      }
      $fileName = $fileName.substring(0, $fileName.length - 1)

      # Determine content length by total - header - filename
      $fileContentByteLen = $bytes.length - $fileByteHeader - $fileNameByteLen
      $fileContentBytesStart = $fileByteHeader + $fileNameByteLen
      $fileContentBytesEnd = $fileContentBytesStart + $fileContentByteLen

      # Create new array by cloning the content bytes into new array
      $arrFileContentBytes = $bytes[($fileContentBytesStart)..($fileContentBytesEnd)]
    }

    # Clean up filename to get rid of spaces and illegal characters
    $fileName = $fileName.Trim() -replace '[^\p{L}\p{Nd}/(/_/)/./@/,/-]', ''
    Write-Verbose "[$formName] Attachment $fileName is $([Math]::Round($arrFileContentBytes.Length/1MB,2)) MB"

    # Establish the base file name for the attachment
    $nodeName = $textNode.LocalName
    $fileInfo = [FileInfo]$fileName

    # Create the attachment folder if it doesn't exist yet
    if ($null -eq $attachmentFolder) {
      # Use configured attachment folder or create one next to the form file
      if ($CreateAttachmentsFolderIn) {
        $attachmentFolder = Join-Path -Path $CreateAttachmentsFolderIn -ChildPath ([System.IO.Path]::GetFileNameWithoutExtension($formName))
      }
      else {
        $attachmentFolder = $formPath.Substring(0, $formPath.LastIndexOf('.'))
      }
      
      if (-not (Test-Path -Path $attachmentFolder -PathType Container)) {
        New-Item -Path $attachmentFolder -ItemType Directory -Force | Out-Null
        Write-Verbose "[$formName] Created attachment folder: $attachmentFolder"
      }
      else {
        Write-Verbose "[$formName] Using existing attachment folder: $attachmentFolder"
      }
    }

    # Check for existing attachments with the same name
    Write-Verbose 'Checking for existing attachments with the same name'
    $attachmentFilter = '{0}*-{1}{2}' -f $nodeName, $fileInfo.BaseName, $fileInfo.Extension
    $existingAttachments = Get-ChildItem -Path $attachmentFolder -Filter $attachmentFilter -ErrorAction SilentlyContinue

    # Generate unique filename if needed
    $attachmentName = if ($null -ne $existingAttachments -and $existingAttachments.Count -gt 0) {
      Write-Verbose "[$formName] Found existing attachment(s) with similar name: $fileName"
      $last = $existingAttachments | Sort-Object -Property Name -Descending | Select-Object -First 1
      
      $lastNum = $last.Name -replace "^$nodeName-copy(\d+)-.*", '$1'
      if ([string]::IsNullOrEmpty($lastNum) -or $lastNum -eq $last.Name) { $lastNum = 0 }
      $nextNum = [int]$lastNum + 1
      '{0}-copy{1}-{2}{3}' -f $nodeName, $nextNum, $fileInfo.BaseName, $fileInfo.Extension
    }
    else {
      # If there are no existing files with the same name, use the base name
      Write-Verbose "[$formName] No existing attachments found with name: $fileName"
      '{0}-{1}{2}' -f $nodeName, $fileInfo.BaseName, $fileInfo.Extension
    }

    # Combine the directory and the attachment name to get the full path
    $attachmentPath = Join-Path -Path $attachmentFolder -ChildPath $attachmentName

    # Final step - save the document to the local computer
    try { 
      [File]::WriteAllBytes($attachmentPath, $arrFileContentBytes) 
      $script:extractedAttachments++
      Write-Verbose "[$formName] Saved attachment: $attachmentName ($([Math]::Round($arrFileContentBytes.Length/1KB,1)) KB)"
    }
    catch {
      Write-Warning "Can't save attachment from $formFile to $attachmentPath : $($_.Exception.Message)"
      $script:erroredFiles[$formName] = "Failed to save attachment: $($_.Exception.Message)"
      continue
    }
  }
}
#endregion MAIN PROCESSING

#region RESULTS SUMMARY
# Display summary of processing results
Write-Host "`n" -NoNewline
Write-Host "InfoPath Attachment Extraction Complete" -ForegroundColor Green
Write-Host "=======================================" -ForegroundColor Green

Write-Host "`nProcessing Summary:" -ForegroundColor Cyan
Write-Host "  Files Processed: $script:processedFiles" -ForegroundColor White
Write-Host "  Attachments Extracted: $script:extractedAttachments" -ForegroundColor White
Write-Host "  Files with Errors: $($script:erroredFiles.Count)" -ForegroundColor $(if ($script:erroredFiles.Count -gt 0) { 'Yellow' }else { 'Green' })

if ($script:erroredFiles.Count -gt 0) {
  Write-Host "`nFiles with errors:" -ForegroundColor Yellow
  foreach ($file in $script:erroredFiles.Keys) {
    Write-Host "  - $file : $($script:erroredFiles[$file])" -ForegroundColor Yellow
  }
}

if ($CreateAttachmentsFolderIn) {
  Write-Host "`nAttachments saved to: $CreateAttachmentsFolderIn" -ForegroundColor Cyan
}

if ($script:extractedAttachments -eq 0 -and $script:processedFiles -gt 0) {
  Write-Host "`nNote: No attachments found in the processed XML files." -ForegroundColor Yellow
  Write-Host "This could mean the files don't contain embedded attachments or they're in a different format." -ForegroundColor Yellow
}
#endregion RESULTS SUMMARY
