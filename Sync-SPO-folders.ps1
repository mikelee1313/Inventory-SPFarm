<#
.SYNOPSIS
  Moves all items from a source subfolder to a destination subfolder within a SharePoint
  document library, automatically renaming items that would otherwise collide.

.DESCRIPTION
  Connects to a SharePoint site using PnP.PowerShell and moves every file and folder found
  directly inside a source folder into a destination folder in the same document library.
  Use -IncludeSourceFolder to preserve the selected source folder as a folder beneath the
  destination (for example, moving Sagebrush Software into clients/s/Sagebrush Software).
  If an item with the same name already exists at the destination, behaviour depends on
  the -MoveDuplicateFileandFolders parameter. When $true (default) the moved item is renamed by
  appending an incrementing number (e.g. Report.docx -> Report1.docx) so that no existing
  content is overwritten, and same-named folders are merged. When $false the duplicate file
  or folder is skipped entirely and left in the source location.

.PREREQUISITES
  - PnP.PowerShell installed: Install-Module PnP.PowerShell -Scope CurrentUser
  - Entra ID app registration with SharePoint application permissions
  - App auth configured with either certificate thumbprint or client secret
  - Admin consent granted for required permissions

.Version 10 - added support for moving duplicate files and folders with the -MoveDuplicateFileandFolders parameter
#>

[CmdletBinding(SupportsShouldProcess)]
#region Parameters
param(
  [Parameter()]
  [string]$SiteUrl = 'https://m365cpi13246019.sharepoint.com/sites/SPSite1/',

  [Parameter()]
  [string]$DocumentLibrary = 'Shared Documents',

  [Parameter()]
  [string]$SourceFolderPath = 'general/clients/u/Urban Vale Networks',

  [Parameter()]
  [string]$DestinationFolderPath = 'clients/u',

  [Parameter()]
  [string]$TenantId = '9cfc42cb-51da-4055-87e9-b20a170b6ba3',

  [Parameter()]
  [string]$ClientId = 'abc64618-283f-47ba-a185-50d935d51d57',

  [Parameter()]
  [ValidateSet('Certificate', 'ClientSecret')]
  [string]$AuthType = 'Certificate',

  [Parameter()]
  [string]$Thumbprint = 'B696FDCFE1453F3FBC6031F54DE988DA0ED905A9',

  [Parameter()]
  [ValidateSet('LocalMachine', 'CurrentUser')]
  [string]$CertStore = 'LocalMachine',

  [Parameter()]
  [string]$ClientSecret = $env:PNP_CLIENT_SECRET,

  [Parameter()]
  [bool]$MoveDuplicateFileandFolders = $false,

  [Parameter()]
  [switch]$IncludeSourceFolder = $true,

  [Parameter()]
  [int]$ThrottleDelayMs = 0,

  [Parameter()]
  [int]$MaxRetries = 10,

  [Parameter()]
  [int]$InitialBackoffSeconds = 2,

  [Parameter()]
  [string]$LogPath = (Join-Path -Path (Get-Location) -ChildPath (
      "MoveLog_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss')
    ))
)
#endregion Parameters

#region Input Validation
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$requiredParams = @{
  SiteUrl               = $SiteUrl
  DocumentLibrary       = $DocumentLibrary
  SourceFolderPath      = $SourceFolderPath
  DestinationFolderPath = $DestinationFolderPath
  TenantId              = $TenantId
  ClientId              = $ClientId
}
foreach ($entry in $requiredParams.GetEnumerator()) {
  if ([string]::IsNullOrWhiteSpace([string]$entry.Value)) {
    throw "Required parameter '$($entry.Key)' is empty. Provide a value in the param block or at runtime."
  }
}

if ($AuthType -eq 'Certificate' -and [string]::IsNullOrWhiteSpace($Thumbprint)) {
  throw "AuthType 'Certificate' requires a non-empty Thumbprint."
}

if ($AuthType -eq 'ClientSecret' -and [string]::IsNullOrWhiteSpace($ClientSecret)) {
  throw "AuthType 'ClientSecret' requires -ClientSecret or env var PNP_CLIENT_SECRET."
}
#endregion Input Validation

#region Logging Helpers
function Write-Info {
  [CmdletBinding()]
  param([Parameter(Mandatory)] [string]$Message)

  Write-Host $Message -ForegroundColor Cyan
}

function Write-Success {
  [CmdletBinding()]
  param([Parameter(Mandatory)] [string]$Message)

  Write-Host $Message -ForegroundColor Green
}

function Write-Warn {
  [CmdletBinding()]
  param([Parameter(Mandatory)] [string]$Message)

  Write-Host $Message -ForegroundColor Yellow
}
#endregion Logging Helpers

#region Throttling and Retry
function Get-HeaderValue {
  [CmdletBinding()]
  param(
    [Parameter()] $Headers,
    [Parameter(Mandatory)] [string]$Name
  )

  if ($null -eq $Headers) { return $null }

  try {
    $value = $Headers[$Name]
    if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace([string]$value)) {
      return [string]$value
    }
  }
  catch {
  }

  return $null
}

function Get-ThrottleWaitSecondsFromHeaders {
  [CmdletBinding()]
  param(
    [Parameter()] $Headers,
    [Parameter()] [int]$DefaultSeconds = 1
  )

  $retryAfterSec = $null
  $rateResetSec = $null

  $retryAfter = Get-HeaderValue -Headers $Headers -Name 'Retry-After'
  if (-not [string]::IsNullOrWhiteSpace($retryAfter)) {
    $intVal = 0
    if ([int]::TryParse($retryAfter, [ref]$intVal)) {
      $retryAfterSec = [Math]::Max($intVal, 0)
    }
    else {
      $dtVal = [datetime]::MinValue
      if ([datetime]::TryParse($retryAfter, [ref]$dtVal)) {
        $retryAfterSec = [Math]::Max([int][Math]::Ceiling(($dtVal.ToUniversalTime() - [datetime]::UtcNow).TotalSeconds), 0)
      }
    }
  }

  $rateReset = Get-HeaderValue -Headers $Headers -Name 'RateLimit-Reset'
  if (-not [string]::IsNullOrWhiteSpace($rateReset)) {
    $resetVal = 0
    if ([int]::TryParse($rateReset, [ref]$resetVal)) {
      $rateResetSec = [Math]::Max($resetVal, 0)
    }
  }

  $candidates = @()
  if ($null -ne $retryAfterSec) { $candidates += $retryAfterSec }
  if ($null -ne $rateResetSec) { $candidates += $rateResetSec }

  if ($candidates.Count -gt 0) {
    return ($candidates | Measure-Object -Maximum).Maximum
  }

  return [Math]::Max($DefaultSeconds, 1)
}

function Invoke-PnPWithRetry {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] [scriptblock]$ScriptBlock,
    [Parameter()] [int]$MaxRetries = $script:MaxRetries,
    [Parameter()] [int]$InitialBackoffSeconds = $script:InitialBackoffSeconds
  )

  $retryCount = 0
  $backoffSec = $InitialBackoffSeconds

  while ($retryCount -le $MaxRetries) {
    try {
      return & $ScriptBlock
    }
    catch {
      $statusCode = $null
      $headers = $null

      $responseProp = $_.Exception.PSObject.Properties['Response']
      if ($null -ne $responseProp -and $null -ne $responseProp.Value) {
        try { $statusCode = [int]$responseProp.Value.StatusCode } catch { }
        try { $headers = $responseProp.Value.Headers } catch { }
      }

      if (-not $statusCode -and $_.Exception.Message -match '(429|502|503|504)') {
        $statusCode = [int]$Matches[1]
      }

      $looksThrottleLike = $_.Exception.Message -match '(throttl|too many requests|server too busy|try again)'
      $isRetryable = ($statusCode -in @(429, 502, 503, 504)) -or ($statusCode -eq 403 -and $looksThrottleLike)
      if (-not $isRetryable) { throw }

      if ($retryCount -ge $MaxRetries) { throw }

      $headerWaitSec = Get-ThrottleWaitSecondsFromHeaders -Headers $headers -DefaultSeconds $backoffSec
      $waitSec = [Math]::Max($backoffSec, $headerWaitSec)
      $jitterMs = Get-Random -Minimum 200 -Maximum 1200
      $waitSec = [Math]::Min($waitSec + ($jitterMs / 1000.0), 900)

      $retryCount++
      Write-Warn ("Throttled (HTTP {0}). Waiting {1:n1}s (attempt {2}/{3})." -f $statusCode, $waitSec, $retryCount, $MaxRetries)

      Start-Sleep -Milliseconds ([int][Math]::Ceiling($waitSec * 1000.0))
      $backoffSec = [Math]::Min($backoffSec * 2, 300)
    }
  }
}
#endregion Throttling and Retry

#region Library and Path Resolution
function Resolve-DocumentLibraryContext {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] [string]$LibraryInput,
    [Parameter(Mandatory)] [string]$FallbackSiteUrl
  )

  $trimmed = $LibraryInput.Trim()
  if ($trimmed -notmatch '^https?://') {
    return [pscustomobject]@{
      SiteUrl                 = $FallbackSiteUrl
      LibrarySiteRelativePath = $trimmed.Trim('/')
      Original                = $LibraryInput
    }
  }

  $uri = [Uri]$trimmed
  $path = [Uri]::UnescapeDataString($uri.AbsolutePath)
  $sitePath = ''
  $remaining = ''

  if ($path -match '^/(sites|teams)/([^/]+)(/.*)?$') {
    $sitePath = "/$($Matches[1])/$($Matches[2])"
    $remaining = if ($Matches[3]) { $Matches[3].Trim('/') } else { '' }
  }
  else {
    $remaining = $path.Trim('/')
  }

  if ([string]::IsNullOrWhiteSpace($remaining)) {
    throw "Could not derive document library path from '$LibraryInput'."
  }

  return [pscustomobject]@{
    SiteUrl                 = "{0}://{1}{2}" -f $uri.Scheme, $uri.Host, $sitePath
    LibrarySiteRelativePath = $remaining
    Original                = $LibraryInput
  }
}
#endregion Library and Path Resolution

#region SharePoint Connection
function Connect-ToPnPSite {
  [CmdletBinding()]
  param([Parameter(Mandatory)] [string]$Url)

  if ($AuthType -eq 'Certificate') {
    if (-not (Test-Path "Cert:\$CertStore\My\$Thumbprint")) {
      throw "Certificate $Thumbprint not found in Cert:\$CertStore\My"
    }

    Invoke-PnPWithRetry {
      Connect-PnPOnline -Url $Url -Tenant $TenantId -ClientId $ClientId -Thumbprint $Thumbprint -ErrorAction Stop
    }
  }
  else {
    Invoke-PnPWithRetry {
      Connect-PnPOnline -Url $Url -ClientId $ClientId -ClientSecret $ClientSecret -ErrorAction Stop
    }
  }

  $web = Invoke-PnPWithRetry { Get-PnPWeb }
  Write-Info "Connected to: $($web.Url)"
}
#endregion SharePoint Connection

#region Item Metadata Helpers
function Get-ListNameFromLibraryUrl {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] [string]$LibraryInput
  )

  $candidate = $LibraryInput.Trim()
  if ($candidate -match '^https?://') {
    $uri = [Uri]$candidate
    $path = [Uri]::UnescapeDataString($uri.AbsolutePath)
    if ($path -match '^/(sites|teams)/[^/]+/(.*)$') {
      $candidate = $Matches[2]
    }
    else {
      $candidate = $path.Trim('/')
    }
  }

  $segments = @($candidate.Split('/') | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

  if ($segments.Count -eq 0) {
    throw "Could not identify the library name from '$LibraryInput'."
  }

  return $segments[0]
}

function Get-ResolvedLibraryName {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] [string]$LibraryInput
  )

  $candidateName = Get-ListNameFromLibraryUrl -LibraryInput $LibraryInput
  $libraryList = $null

  foreach ($identity in @($LibraryInput.Trim().Trim('/'), $candidateName) | Select-Object -Unique) {
    try {
      $libraryList = Get-PnPList -Identity $identity -Includes RootFolder -ErrorAction Stop
      if ($null -ne $libraryList) { break }
    }
    catch {
      $libraryList = $null
    }
  }

  if ($null -eq $libraryList) {
    throw "Could not resolve the document library from '$LibraryInput'."
  }

  return $libraryList.Title
}
#endregion Item Metadata Helpers

#region Move Operations
function Get-PnPFolderChildNames {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] [string]$FolderServerRelativeUrl
  )

  $contents = Invoke-PnPWithRetry {
    $folder = Get-PnPFolder -Url $FolderServerRelativeUrl -ErrorAction Stop
    Get-PnPProperty -ClientObject $folder -Property Files, Folders -ErrorAction Stop
    return [pscustomobject]@{
      Files   = @($folder.Files)
      Folders = @($folder.Folders)
    }
  }

  $names = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
  foreach ($file in $contents.Files) { [void]$names.Add($file.Name) }
  foreach ($folder in $contents.Folders) { [void]$names.Add($folder.Name) }

  return [pscustomobject]@{
    Names   = $names
    Files   = $contents.Files
    Folders = $contents.Folders
  }
}

function Get-UniqueItemName {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] [string]$Name,
    [Parameter(Mandatory)] [bool]$IsFolder,
    [Parameter(Mandatory)] [System.Collections.Generic.HashSet[string]]$ExistingNames
  )

  if (-not $ExistingNames.Contains($Name)) { return $Name }

  if ($IsFolder) {
    $baseName = $Name
    $extension = ''
  }
  else {
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($Name)
    $extension = [System.IO.Path]::GetExtension($Name)
  }

  $counter = 1
  do {
    $candidate = "{0}{1}{2}" -f $baseName, $counter, $extension
    $counter++
  } while ($ExistingNames.Contains($candidate))

  return $candidate
}

function Move-FolderContentsRecursive {
  [CmdletBinding(SupportsShouldProcess)]
  param(
    [Parameter(Mandatory)] [string]$SourceFolderServerRelativeUrl,
    [Parameter(Mandatory)] [string]$DestinationFolderServerRelativeUrl,
    [Parameter(Mandatory)] [hashtable]$Stats,
    [Parameter(Mandatory)] [AllowEmptyCollection()] [System.Collections.Generic.List[object]]$LogRows,
    [Parameter()] [int]$ThrottleDelayMs = 0,
    [Parameter()] [bool]$MoveDuplicateFileandFolders = $true
  )

  $sourceContents = Get-PnPFolderChildNames -FolderServerRelativeUrl $SourceFolderServerRelativeUrl
  $destinationContents = Get-PnPFolderChildNames -FolderServerRelativeUrl $DestinationFolderServerRelativeUrl
  $destinationNames = $destinationContents.Names
  $destinationFolderNames = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
  foreach ($folder in $destinationContents.Folders) { [void]$destinationFolderNames.Add($folder.Name) }

  foreach ($folder in $sourceContents.Folders) {
    $Stats.Total++
    $name = $folder.Name
    $sourceUrl = "$SourceFolderServerRelativeUrl/$name"

    if ($destinationFolderNames.Contains($name)) {
      if (-not $MoveDuplicateFileandFolders) {
        $Stats.Skipped++
        Write-Warn "Duplicate folder found. Skipped '$name'; left in source folder."
        $LogRows.Add([pscustomobject]@{ OriginalName = $name; MovedAsName = ''; ItemType = 'Folder'; Renamed = $false; Status = 'Skipped'; Error = 'Duplicate name at destination; MoveDuplicateFileandFolders is false.' })
        continue
      }

      # Same-named folder exists at destination: merge contents instead of renaming the folder.
      Write-Info "Folder '$name' already exists at destination. Merging contents."
      $destChildUrl = "$DestinationFolderServerRelativeUrl/$name"
      $errorsBeforeMerge = $Stats.Errors
      $skippedBeforeMerge = $Stats.Skipped
      Move-FolderContentsRecursive -SourceFolderServerRelativeUrl $sourceUrl -DestinationFolderServerRelativeUrl $destChildUrl -Stats $Stats -LogRows $LogRows -ThrottleDelayMs $ThrottleDelayMs -MoveDuplicateFileandFolders $MoveDuplicateFileandFolders

      # Only delete the source folder if nothing beneath it failed to move or was intentionally
      # left behind; otherwise deleting the folder would send those items to the recycle bin.
      if ($Stats.Errors -gt $errorsBeforeMerge) {
        Write-Warn "Skipping removal of source folder '$name': one or more nested items failed to move."
        $LogRows.Add([pscustomobject]@{ OriginalName = $name; MovedAsName = $name; ItemType = 'Folder (merged)'; Renamed = $false; Status = 'Skipped'; Error = 'Nested item(s) failed to move; source folder retained.' })
      }
      elseif ($Stats.Skipped -gt $skippedBeforeMerge) {
        Write-Warn "Skipping removal of source folder '$name': duplicate item(s) were retained in the source."
        $LogRows.Add([pscustomobject]@{ OriginalName = $name; MovedAsName = $name; ItemType = 'Folder (merged)'; Renamed = $false; Status = 'Skipped'; Error = 'Duplicate item(s) retained in source; source folder retained.' })
      }
      elseif ($PSCmdlet.ShouldProcess($sourceUrl, "Remove now-empty source folder after merge")) {
        try {
          Invoke-PnPWithRetry { Remove-PnPFolder -Name $name -Folder $SourceFolderServerRelativeUrl -Recycle -Force -ErrorAction Stop } | Out-Null
          $Stats.Merged++
          $LogRows.Add([pscustomobject]@{ OriginalName = $name; MovedAsName = $name; ItemType = 'Folder (merged)'; Renamed = $false; Status = 'Success'; Error = '' })
        }
        catch {
          $Stats.Errors++
          Write-Warn "Failed to remove empty source folder '$name' after merge: $($_.Exception.Message)"
          $LogRows.Add([pscustomobject]@{ OriginalName = $name; MovedAsName = $name; ItemType = 'Folder (merged)'; Renamed = $false; Status = 'Failed'; Error = $_.Exception.Message })
        }
      }
    }
    else {
      $targetUrl = "$DestinationFolderServerRelativeUrl/$name"
      if (-not $PSCmdlet.ShouldProcess($DestinationFolderServerRelativeUrl, "Move folder '$name'")) { continue }

      try {
        Invoke-PnPWithRetry { Move-PnPFile -SourceUrl $sourceUrl -TargetUrl $targetUrl -Force -ErrorAction Stop } | Out-Null
        [void]$destinationNames.Add($name)
        [void]$destinationFolderNames.Add($name)
        $Stats.Moved++
        Write-Info "Moved folder '$name'."
        $LogRows.Add([pscustomobject]@{ OriginalName = $name; MovedAsName = $name; ItemType = 'Folder'; Renamed = $false; Status = 'Success'; Error = '' })
      }
      catch {
        $Stats.Errors++
        Write-Warn "Failed to move folder '$name': $($_.Exception.Message)"
        $LogRows.Add([pscustomobject]@{ OriginalName = $name; MovedAsName = $name; ItemType = 'Folder'; Renamed = $false; Status = 'Failed'; Error = $_.Exception.Message })
      }
    }

    if ($ThrottleDelayMs -gt 0) {
      Start-Sleep -Milliseconds $ThrottleDelayMs
    }
  }

  foreach ($file in $sourceContents.Files) {
    $Stats.Total++
    $originalName = $file.Name

    if (-not $MoveDuplicateFileandFolders -and $destinationNames.Contains($originalName)) {
      $Stats.Skipped++
      Write-Warn "Duplicate found. Skipped '$originalName'; left in source folder."
      $LogRows.Add([pscustomobject]@{ OriginalName = $originalName; MovedAsName = ''; ItemType = 'File'; Renamed = $false; Status = 'Skipped'; Error = 'Duplicate name at destination; MoveDuplicateFileandFolders is false.' })
      continue
    }

    $uniqueName = Get-UniqueItemName -Name $originalName -IsFolder $false -ExistingNames $destinationNames
    $wasRenamed = $uniqueName -ne $originalName

    $sourceUrl = "$SourceFolderServerRelativeUrl/$originalName"
    $targetUrl = "$DestinationFolderServerRelativeUrl/$uniqueName"

    $actionDescription = if ($wasRenamed) { "Move file '$originalName' to destination as '$uniqueName' (duplicate name)" } else { "Move file '$originalName'" }
    if (-not $PSCmdlet.ShouldProcess($DestinationFolderServerRelativeUrl, $actionDescription)) { continue }

    try {
      Invoke-PnPWithRetry { Move-PnPFile -SourceUrl $sourceUrl -TargetUrl $targetUrl -Force -ErrorAction Stop } | Out-Null
      [void]$destinationNames.Add($uniqueName)
      $Stats.Moved++

      if ($wasRenamed) {
        $Stats.Renamed++
        Write-Warn "Duplicate found. Moved '$originalName' as '$uniqueName'."
      }
      else {
        Write-Info "Moved '$originalName'."
      }

      $LogRows.Add([pscustomobject]@{ OriginalName = $originalName; MovedAsName = $uniqueName; ItemType = 'File'; Renamed = $wasRenamed; Status = 'Success'; Error = '' })
    }
    catch {
      $Stats.Errors++
      Write-Warn "Failed to move '$originalName': $($_.Exception.Message)"
      $LogRows.Add([pscustomobject]@{ OriginalName = $originalName; MovedAsName = $uniqueName; ItemType = 'File'; Renamed = $wasRenamed; Status = 'Failed'; Error = $_.Exception.Message })
    }

    if ($ThrottleDelayMs -gt 0) {
      Start-Sleep -Milliseconds $ThrottleDelayMs
    }
  }
}

function Move-LibraryFolderItems {
  [CmdletBinding(SupportsShouldProcess)]
  param(
    [Parameter(Mandatory)] [string]$SourceFolderServerRelativeUrl,
    [Parameter(Mandatory)] [string]$DestinationFolderServerRelativeUrl,
    [Parameter()] [int]$ThrottleDelayMs = 0,
    [Parameter()] [bool]$MoveDuplicateFileandFolders = $true,
    [Parameter()] [switch]$RemoveSourceRootAfterMove,
    [Parameter()] [string]$LogPath
  )

  $stats = @{ Total = 0; Moved = 0; Renamed = 0; Merged = 0; Skipped = 0; Errors = 0 }
  $logRows = [System.Collections.Generic.List[object]]::new()

  Move-FolderContentsRecursive -SourceFolderServerRelativeUrl $SourceFolderServerRelativeUrl -DestinationFolderServerRelativeUrl $DestinationFolderServerRelativeUrl -Stats $stats -LogRows $logRows -ThrottleDelayMs $ThrottleDelayMs -MoveDuplicateFileandFolders $MoveDuplicateFileandFolders

  if ($RemoveSourceRootAfterMove -and $stats.Errors -eq 0 -and $stats.Skipped -eq 0) {
    $sourceFolderUrl = $SourceFolderServerRelativeUrl.TrimEnd('/')
    $sourceFolderName = $sourceFolderUrl.Substring($sourceFolderUrl.LastIndexOf('/') + 1)
    $sourceParentUrl = $sourceFolderUrl.Substring(0, $sourceFolderUrl.LastIndexOf('/'))

    if ($PSCmdlet.ShouldProcess($sourceFolderUrl, 'Remove now-empty source root folder')) {
      try {
        Invoke-PnPWithRetry { Remove-PnPFolder -Name $sourceFolderName -Folder $sourceParentUrl -Recycle -Force -ErrorAction Stop } | Out-Null
        $logRows.Add([pscustomobject]@{ OriginalName = $sourceFolderName; MovedAsName = $sourceFolderName; ItemType = 'Folder (root)'; Renamed = $false; Status = 'Success'; Error = '' })
      }
      catch {
        $stats.Errors++
        Write-Warn "Failed to remove empty source root folder '$sourceFolderName': $($_.Exception.Message)"
        $logRows.Add([pscustomobject]@{ OriginalName = $sourceFolderName; MovedAsName = $sourceFolderName; ItemType = 'Folder (root)'; Renamed = $false; Status = 'Failed'; Error = $_.Exception.Message })
      }
    }
  }

  if (-not [string]::IsNullOrWhiteSpace($LogPath) -and $logRows.Count -gt 0) {
    $logRows | Export-Csv -Path $LogPath -NoTypeInformation
  }

  return [pscustomobject]@{
    TotalItems = $stats.Total
    Moved      = $stats.Moved
    Renamed    = $stats.Renamed
    Merged     = $stats.Merged
    Skipped    = $stats.Skipped
    Errors     = $stats.Errors
  }
}
#endregion Move Operations

#region Main Execution
try {
  if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    throw 'PnP.PowerShell module not found. Install with: Install-Module PnP.PowerShell -Scope CurrentUser'
  }

  Import-Module PnP.PowerShell -ErrorAction Stop

  $libraryContext = Resolve-DocumentLibraryContext -LibraryInput $DocumentLibrary -FallbackSiteUrl $SiteUrl
  Connect-ToPnPSite -Url $libraryContext.SiteUrl

  [void](Get-ResolvedLibraryName -LibraryInput $DocumentLibrary)
  $libraryRelativeUrl = $libraryContext.LibrarySiteRelativePath.Trim('/')

  $web = Invoke-PnPWithRetry { Get-PnPWeb -Includes ServerRelativeUrl }
  $siteServerRelativeUrl = $web.ServerRelativeUrl.TrimEnd('/')

  $sourceSiteRelativeUrl = ("{0}/{1}" -f $libraryRelativeUrl, $SourceFolderPath.Trim('/')).Trim('/')
  $destinationBaseSiteRelativeUrl = ("{0}/{1}" -f $libraryRelativeUrl, $DestinationFolderPath.Trim('/')).Trim('/')
  $sourceFolderName = ($SourceFolderPath.Trim('/') -split '/')[-1]
  $destinationFolderName = ($DestinationFolderPath.Trim('/') -split '/')[-1]
  $destinationAlreadyIncludesSourceFolder = $destinationFolderName.Equals($sourceFolderName, [StringComparison]::OrdinalIgnoreCase)
  $destinationSiteRelativeUrl = if ($IncludeSourceFolder -and -not $destinationAlreadyIncludesSourceFolder) {
    "$destinationBaseSiteRelativeUrl/$sourceFolderName"
  }
  else {
    $destinationBaseSiteRelativeUrl
  }

  # Guard against the destination being the same as, or nested inside, the source (would corrupt/duplicate content).
  if ($destinationSiteRelativeUrl -eq $sourceSiteRelativeUrl -or
      $destinationSiteRelativeUrl.StartsWith("$sourceSiteRelativeUrl/", [StringComparison]::OrdinalIgnoreCase) -or
      $sourceSiteRelativeUrl.StartsWith("$destinationSiteRelativeUrl/", [StringComparison]::OrdinalIgnoreCase)) {
    throw "DestinationFolderPath ('$DestinationFolderPath') must not be the same as, or nested within, SourceFolderPath ('$SourceFolderPath')."
  }

  Write-Info "Source folder: $sourceSiteRelativeUrl"
  Write-Info "Destination folder: $destinationSiteRelativeUrl"

  # Fail fast if the source doesn't exist; create the destination if it's missing.
  Invoke-PnPWithRetry { Get-PnPFolder -Url $sourceSiteRelativeUrl -ErrorAction Stop } | Out-Null
  Invoke-PnPWithRetry { Resolve-PnPFolder -SiteRelativePath $destinationSiteRelativeUrl -ErrorAction Stop } | Out-Null

  $sourceFolderServerRelativeUrl = "$siteServerRelativeUrl/$sourceSiteRelativeUrl"
  $destinationFolderServerRelativeUrl = "$siteServerRelativeUrl/$destinationSiteRelativeUrl"

  $logDirectory = Split-Path -Path $LogPath -Parent
  if (-not [string]::IsNullOrWhiteSpace($logDirectory) -and -not (Test-Path -Path $logDirectory)) {
    New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
  }

  $summary = Move-LibraryFolderItems -SourceFolderServerRelativeUrl $sourceFolderServerRelativeUrl -DestinationFolderServerRelativeUrl $destinationFolderServerRelativeUrl -ThrottleDelayMs $ThrottleDelayMs -MoveDuplicateFileandFolders $MoveDuplicateFileandFolders -RemoveSourceRootAfterMove:$IncludeSourceFolder -LogPath $LogPath

  if ($summary.TotalItems -eq 0) {
    Write-Warn "No items found in the source folder: $sourceSiteRelativeUrl"
  }
  elseif ($summary.Errors -gt 0) {
    Write-Warn ("Move completed with errors. Items: {0} | Moved: {1} (Renamed for duplicates: {2}, Merged folders: {3}) | Skipped duplicates: {4} | Errors: {5} | Log: {6}" -f $summary.TotalItems, $summary.Moved, $summary.Renamed, $summary.Merged, $summary.Skipped, $summary.Errors, $LogPath)
  }
  else {
    Write-Success ("Move completed. Items: {0} | Moved: {1} (Renamed for duplicates: {2}, Merged folders: {3}) | Skipped duplicates: {4} | Log: {5}" -f $summary.TotalItems, $summary.Moved, $summary.Renamed, $summary.Merged, $summary.Skipped, $LogPath)
  }

  # Surface a non-zero exit code so automation/CI can detect partial failures.
  if ($summary.Errors -gt 0) {
    try { $host.SetShouldExit(1) } catch { }
  }
}
catch {
  Write-Warn "Script failed: $($_.Exception.Message)"
  try { $host.SetShouldExit(1) } catch { }
  throw
}
finally {
  try { Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch { }
}
#endregion Main Execution
