<#
.SYNOPSIS
  Generates an inventory report for folders and files in a SharePoint document library.

.DESCRIPTION
  Connects to a SharePoint site using PnP.PowerShell and exports a CSV report of every
  folder and file in the target library, including folder path, item name, created by,
  modified by, modified date, and whether the item is a folder or file.

.PREREQUISITES
  - PnP.PowerShell installed: Install-Module PnP.PowerShell -Scope CurrentUser
  - Entra ID app registration with SharePoint application permissions
  - App auth configured with either certificate thumbprint or client secret
  - Admin consent granted for required permissions
#>

[CmdletBinding()]
#region Parameters
param(
  [Parameter()]
  [string]$SiteUrl = 'https://m365cpi13246019.sharepoint.com/sites/SPSite1/',

  [Parameter()]
  [string]$DocumentLibrary = 'Shared Documents',

  [Parameter()]
  [string]$FolderPath = 'general/clients',

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
  [int]$PageSize = 2000,

  [Parameter()]
  [int]$ThrottleDelayMs = 0,

  [Parameter()]
  [string]$OutputPath = (Join-Path -Path (Get-Location) -ChildPath (
    "{0}_{1}.csv" -f (
      (($FolderPath -split '[\\/]') | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object {
        (Get-Culture).TextInfo.ToTitleCase($_.ToLowerInvariant()) -replace '[\\/:*?"<>|]', '_'
      }) -join '_'
    ), (Get-Date -Format 'yyyyMMdd_HHmmss')
  ))
)
#endregion Parameters

#region Input Validation
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$requiredParams = @{
  SiteUrl = $SiteUrl
  DocumentLibrary = $DocumentLibrary
  TenantId = $TenantId
  ClientId = $ClientId
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
    [Parameter()] [int]$MaxRetries = 10,
    [Parameter()] [int]$InitialBackoffSeconds = 2
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
function Split-LibraryAndFolderPath {
  [CmdletBinding()]
  param(
    [Parameter()] [string]$LibraryInput,
    [Parameter()] [string]$FolderPath
  )

  $libraryValue = if ($null -ne $LibraryInput) { $LibraryInput.Trim() } else { '' }
  $folderValue = if ($null -ne $FolderPath) { $FolderPath.Trim() } else { '' }

  $segments = @()
  if (-not [string]::IsNullOrWhiteSpace($libraryValue)) {
    if ($libraryValue -match '^https?://') {
      $uri = [Uri]$libraryValue
      $path = [Uri]::UnescapeDataString($uri.AbsolutePath).Trim('/')
      $segments = @($path.Split('/') | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }
    else {
      $segments = @($libraryValue.Split('/') | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }
  }

  $segmentCount = @($segments).Count
  $libraryName = if ($segmentCount -gt 0) { $segments[0] } else { $libraryValue }

  $folderSegments = @()
  if ($segmentCount -gt 1) {
    $folderSegments += @($segments)[1..($segmentCount - 1)]
  }
  if (-not [string]::IsNullOrWhiteSpace($folderValue)) {
    $folderSegments += @($folderValue.Split('/') | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
  }

  $folderRelative = if (@($folderSegments).Count -gt 0) { ($folderSegments -join '/') } else { '' }

  return [pscustomobject]@{
    LibraryName = $libraryName
    FolderRelative = $folderRelative
  }
}

function Resolve-DocumentLibraryContext {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] [string]$LibraryInput,
    [Parameter(Mandatory)] [string]$FallbackSiteUrl
  )

  $trimmed = $LibraryInput.Trim()
  if ($trimmed -notmatch '^https?://') {
    return [pscustomobject]@{
      SiteUrl = $FallbackSiteUrl
      LibrarySiteRelativePath = $trimmed.Trim('/')
      Original = $LibraryInput
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
    SiteUrl = "{0}://{1}{2}" -f $uri.Scheme, $uri.Host, $sitePath
    LibrarySiteRelativePath = $remaining
    Original = $LibraryInput
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
function Get-PersonDisplayName {
  [CmdletBinding()]
  param(
    [Parameter()] $Person
  )

  if ($null -eq $Person) { return '' }

  foreach ($propertyName in @('LookupValue', 'Title', 'LoginName', 'Email', 'Name')) {
    if ($Person.PSObject.Properties.Match($propertyName).Count -gt 0) {
      $candidate = $Person.PSObject.Properties[$propertyName].Value
      if (-not [string]::IsNullOrWhiteSpace([string]$candidate)) {
        return [string]$candidate
      }
    }
  }

  return [string]$Person
}

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

  $segments = $candidate.Split('/') | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

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

  $libraryList = Get-PnPList -Includes RootFolder | Where-Object {
    $rootUrl = $_.RootFolder.ServerRelativeUrl
    $normalizedRoot = $rootUrl.Trim('/').ToLowerInvariant()
    $normalizedInput = $LibraryInput.Trim().Trim('/').ToLowerInvariant()

    if ($normalizedInput -match '^https?://') {
      $uri = [Uri]$normalizedInput
      $normalizedInput = [Uri]::UnescapeDataString($uri.AbsolutePath).Trim('/').ToLowerInvariant()
    }

    $normalizedInput = $normalizedInput.Replace('%20', ' ')
    $normalizedRoot = $normalizedRoot.Replace('%20', ' ')

    $normalizedInput -eq $normalizedRoot -or
    $normalizedInput.EndsWith('/' + $normalizedRoot) -or
    $normalizedRoot.EndsWith('/' + $normalizedInput)
  } | Select-Object -First 1

  if ($null -eq $libraryList) {
    $candidateName = Get-ListNameFromLibraryUrl -LibraryInput $LibraryInput
    $libraryList = Get-PnPList -Includes RootFolder | Where-Object { $_.Title -eq $candidateName } | Select-Object -First 1
  }

  if ($null -eq $libraryList) {
    throw "Could not resolve the document library from '$LibraryInput'."
  }

  return $libraryList.Title
}
#endregion Item Metadata Helpers

#region Reporting
function ConvertTo-InventoryRow {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] $Item
  )

  $fieldValues = $Item.FieldValues

  $folderPath = if ($fieldValues.ContainsKey('FileDirRef')) { [string]$fieldValues['FileDirRef'] } else { '' }

  $itemType = 'Unknown'
  if ($fieldValues.ContainsKey('FSObjType')) {
    $itemType = switch ([int]$fieldValues['FSObjType']) {
      1 { 'Folder' }
      0 { 'File' }
      default { 'Unknown' }
    }
  }

  $name = if ($fieldValues.ContainsKey('FileLeafRef')) { [string]$fieldValues['FileLeafRef'] } else { '' }
  $modifiedDate = if ($fieldValues.ContainsKey('Modified')) { $fieldValues['Modified'] } else { $null }

  $createdBy = ''
  if ($fieldValues.ContainsKey('Author') -and $null -ne $fieldValues['Author']) {
    $createdBy = Get-PersonDisplayName -Person $fieldValues['Author']
  }

  $modifiedBy = ''
  if ($fieldValues.ContainsKey('Editor') -and $null -ne $fieldValues['Editor']) {
    $modifiedBy = Get-PersonDisplayName -Person $fieldValues['Editor']
  }

  $itemCount = $null
  if ($itemType -eq 'Folder') {
    $fileCount = 0
    $subfolderCount = 0
    if ($fieldValues.ContainsKey('ItemChildCount') -and -not [string]::IsNullOrWhiteSpace([string]$fieldValues['ItemChildCount'])) {
      $fileCount = [int]$fieldValues['ItemChildCount']
    }
    if ($fieldValues.ContainsKey('FolderChildCount') -and -not [string]::IsNullOrWhiteSpace([string]$fieldValues['FolderChildCount'])) {
      $subfolderCount = [int]$fieldValues['FolderChildCount']
    }
    $itemCount = $fileCount + $subfolderCount
  }

  return [pscustomobject]@{
    FolderPath = $folderPath
    Name = $name
    CreatedBy = $createdBy
    ModifiedBy = $modifiedBy
    ModifiedDate = $modifiedDate
    ItemType = $itemType
    ItemCount = $itemCount
  }
}

function Export-LibraryInventoryReport {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] [string]$ListName,
    [Parameter(Mandatory)] [string]$FolderServerRelativeUrl,
    [Parameter(Mandatory)] [string]$OutputPath,
    [Parameter()] [int]$PageSize = 2000,
    [Parameter()] [int]$ThrottleDelayMs = 0
  )

  $fields = @('FSObjType', 'FileLeafRef', 'FileDirRef', 'Created', 'Modified', 'Author', 'Editor', 'ItemChildCount', 'FolderChildCount')

  $csvHeader = 'FolderPath,Name,CreatedBy,ModifiedBy,ModifiedDate,ItemType,ItemCount'
  [System.IO.File]::WriteAllText($OutputPath, $csvHeader + [Environment]::NewLine)

  $stats = @{ TotalItems = 0; FolderCount = 0; FileCount = 0 }
  $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

  # -ScriptBlock streams each page to disk immediately instead of buffering
  # the whole (potentially multi-million row) result set in memory.
  # Get-PnPListItem invokes this in a scope where scalar variables don't
  # close over the caller, so counters are tracked via the $stats hashtable.
  $pageHandler = {
    param($pageItems)

    $rows = foreach ($item in $pageItems) {
      ConvertTo-InventoryRow -Item $item
    }

    $rows | Export-Csv -Path $OutputPath -NoTypeInformation -Append

    foreach ($row in $rows) {
      $stats.TotalItems++
      if ($row.ItemType -eq 'Folder') { $stats.FolderCount++ }
      elseif ($row.ItemType -eq 'File') { $stats.FileCount++ }
    }

    Write-Info ("Processed {0} items so far ({1:n1}s elapsed)..." -f $stats.TotalItems, $stopwatch.Elapsed.TotalSeconds)

    if ($ThrottleDelayMs -gt 0) {
      Start-Sleep -Milliseconds $ThrottleDelayMs
    }
  }.GetNewClosure()

  Invoke-PnPWithRetry {
    Get-PnPListItem -List $ListName -FolderServerRelativeUrl $FolderServerRelativeUrl -PageSize $PageSize -Fields $fields -ScriptBlock $pageHandler -ErrorAction Stop | Out-Null
  }

  $stopwatch.Stop()

  return [pscustomobject]@{
    TotalItems = $stats.TotalItems
    FolderCount = $stats.FolderCount
    FileCount = $stats.FileCount
    ElapsedSeconds = $stopwatch.Elapsed.TotalSeconds
  }
}
#endregion Reporting

#region Main Execution
try {
  if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    throw 'PnP.PowerShell module not found. Install with: Install-Module PnP.PowerShell -Scope CurrentUser'
  }

  Import-Module PnP.PowerShell -ErrorAction Stop

  $libraryParts = Split-LibraryAndFolderPath -LibraryInput $DocumentLibrary -FolderPath $FolderPath
  $libraryName = $libraryParts.LibraryName
  $folderRelativeUrl = $libraryParts.FolderRelative

  if (-not [string]::IsNullOrWhiteSpace($folderRelativeUrl)) {
    $targetLibraryInput = "{0}/{1}" -f $libraryName, $folderRelativeUrl
  }
  else {
    $targetLibraryInput = $libraryName
  }

  $libraryContext = Resolve-DocumentLibraryContext -LibraryInput $targetLibraryInput -FallbackSiteUrl $SiteUrl
  Write-Info "Target context -> Site: $($libraryContext.SiteUrl) | Library path: $($libraryContext.LibrarySiteRelativePath)"

  Connect-ToPnPSite -Url $libraryContext.SiteUrl

  $libraryName = Get-ResolvedLibraryName -LibraryInput $libraryName
  $targetFolderRelativeUrl = $libraryContext.LibrarySiteRelativePath.Trim('/')
  Write-Info "Generating inventory report for list/library: $libraryName | target folder: $targetFolderRelativeUrl"

  $web = Invoke-PnPWithRetry { Get-PnPWeb -Includes ServerRelativeUrl }
  $siteServerRelativeUrl = $web.ServerRelativeUrl.TrimEnd('/')
  $folderServerRelativeUrl = "$siteServerRelativeUrl/$targetFolderRelativeUrl".TrimEnd('/')

  $directory = Split-Path -Path $OutputPath -Parent
  if (-not [string]::IsNullOrWhiteSpace($directory) -and -not (Test-Path -Path $directory)) {
    New-Item -ItemType Directory -Path $directory -Force | Out-Null
  }

  $summary = Export-LibraryInventoryReport -ListName $libraryName -FolderServerRelativeUrl $folderServerRelativeUrl -OutputPath $OutputPath -PageSize $PageSize -ThrottleDelayMs $ThrottleDelayMs

  if ($summary.TotalItems -eq 0) {
    Write-Warn "No items found in the target library. A header-only CSV was written to: $OutputPath"
  }
  else {
    Write-Success ("Inventory report exported to: {0} | Items: {1} (Folders: {2}, Files: {3}) in {4:n1}s" -f $OutputPath, $summary.TotalItems, $summary.FolderCount, $summary.FileCount, $summary.ElapsedSeconds)
  }
}
finally {
  try { Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch { }
}
#endregion Main Execution
