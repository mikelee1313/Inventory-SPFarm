<#
.SYNOPSIS
    Deletes all files and folders from a SharePoint Online document library, then removes the library.

.DESCRIPTION
    Enumerates every file and folder in the configured library, deletes files first, then deletes
    folders deepest-first. Items blocked by retention/hold/checkout are logged and skipped so the
    run continues. After the content passes complete, the script attempts to remove the library.

.NOTES
    Requires PowerShell 5.1+ and the PnP.PowerShell module.
#>

#region ---------------- Configuration ----------------

$tenantId   = '9cfc42cb-51da-4055-87e9-b20a170b6ba3'
$clientId   = 'abc64618-283f-47ba-a185-50d935d51d57'

# 'Application' = app-only auth using the certificate thumbprint below
# 'Delegate'    = interactive sign-in as the current user (no thumbprint)
$AuthMethod = 'Application'

# Certificate thumbprint in the local certificate store. Required only when $AuthMethod = 'Application'.
$CertificateThumbprint = 'B696FDCFE1453F3FBC6031F54DE988DA0ED905A9'

# Full URL of the site collection / web that contains the library
$SiteUrl    = 'https://m365cpi13246019.sharepoint.com/sites/SPSite2'

# Library to empty. Accepts the library title, the site-relative URL (e.g. 'Shared Documents'),
# the server-relative URL (e.g. '/sites/YourSite/Shared Documents'), or the full library URL.
$LibraryUrl = 'Shared Documents'

# Optional. Limit the cleanup to one folder inside the library instead of the whole library.
# Leave empty to process everything. Accepts a path relative to the library root (e.g. 'Q' or
# 'Q/Archive'), a server-relative URL, or the full folder URL.
$FolderUrl = 'folder1'

# $true  = send items to the recycle bin (faster to recover, still counts against storage)
# $false = permanently delete items
$UseRecycleBin = $true

# Attempt to delete the library itself once the content passes finish.
# Ignored when $FolderUrl is set.
$RemoveLibraryWhenEmpty = $true

# When $FolderUrl is set, also delete that folder once its contents are gone
$RemoveTargetFolderWhenEmpty = $true

# Maximum number of full enumerate/delete passes over the library
$MaxPasses = 5

# Retry behavior for throttling / transient failures
$MaxRetries        = 5
$RetryDelaySeconds = 5

# Logging
$EnableDebugLogging = $true
$date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$log  = Join-Path $env:TEMP ("SPOLibrary-Cleanup_" + $date + "_logfile.log")

#endregion

#region ---------------- Prerequisites ----------------

if ($PSVersionTable.PSVersion.Major -lt 5 -or
    ($PSVersionTable.PSVersion.Major -eq 5 -and $PSVersionTable.PSVersion.Minor -lt 1)) {
    Write-Host "ERROR: This script requires PowerShell 5.1 or later." -ForegroundColor Red
    Write-Host "  Current version: $($PSVersionTable.PSVersion)" -ForegroundColor Red
    exit
}

if (-not (Get-Module -ListAvailable -Name 'PnP.PowerShell')) {
    Write-Host "ERROR: The PnP.PowerShell module is not installed." -ForegroundColor Red
    Write-Host "  Install it by running:  Install-Module PnP.PowerShell -Scope CurrentUser" -ForegroundColor Yellow
    exit
}

if ($AuthMethod -notin @('Application', 'Delegate')) {
    Write-Host "ERROR: `$AuthMethod must be 'Application' or 'Delegate'. Current value: '$AuthMethod'" -ForegroundColor Red
    exit
}

if ($AuthMethod -eq 'Application' -and [string]::IsNullOrWhiteSpace($CertificateThumbprint)) {
    Write-Host "ERROR: `$CertificateThumbprint is required when `$AuthMethod is 'Application'." -ForegroundColor Red
    exit
}

#endregion

#region ---------------- Helper functions ----------------

Function Write-LogEntry {
    param(
        [string] $LogName,
        [string] $LogEntryText,
        [string] $LogLevel = "INFO"
    )
    if ($null -ne $LogName) {
        if ($LogLevel -eq "DEBUG" -and $EnableDebugLogging -eq $false) { return }
        "$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToShortTimeString()) : [$LogLevel] $LogEntryText" |
            Out-File -FilePath $LogName -Append
    }
}

# Only throttling and transient transport failures are worth retrying. Every other error is
# reported with whatever SharePoint returned rather than being matched against message text.
Function Test-Throttled {
    param([System.Management.Automation.ErrorRecord] $ErrorRecord)

    $message = $ErrorRecord.Exception.Message
    return ($message -like '*429*' -or
            $message -like '*throttl*' -or
            $message -like '*503*' -or
            $message -like '*timed out*')
}

Function Invoke-PnPWithRetry {
    param(
        [Parameter(Mandatory = $true)][scriptblock] $ScriptBlock,
        [Parameter(Mandatory = $true)][string] $OperationDescription
    )

    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            return & $ScriptBlock
        }
        catch {
            if ($attempt -eq $MaxRetries -or -not (Test-Throttled -ErrorRecord $_)) { throw }

            $wait = $RetryDelaySeconds * $attempt
            Write-Host "    Retry $attempt/$MaxRetries in ${wait}s - failed to $OperationDescription" -ForegroundColor DarkYellow
            Write-LogEntry -LogName $log -LogEntryText "Retry $attempt/$MaxRetries for '$OperationDescription': $_" -LogLevel "WARNING"
            Start-Sleep -Seconds $wait
        }
    }
}

# Resolve the configured library value to a PnP list object.
Function Resolve-TargetLibrary {
    param(
        [Parameter(Mandatory = $true)][string] $Library,
        [Parameter(Mandatory = $true)][string] $WebUrl
    )

    $identity = $Library.Trim()

    if ($identity -match '^https?://') {
        $identity = ([Uri]$identity).AbsolutePath
    }

    $identity = [Uri]::UnescapeDataString($identity).Trim('/')

    # Try the value as-is (title or site-relative URL)
    $list = Get-PnPList -Identity $identity -Includes RootFolder -ErrorAction SilentlyContinue
    if ($null -ne $list) { return $list }

    # Fall back to matching on the root folder server-relative URL
    $webPath = ([Uri]$WebUrl).AbsolutePath.TrimEnd('/')
    $candidates = @(
        "$webPath/$identity".Replace('//', '/'),
        "/$identity"
    )

    $allLists = Get-PnPList -Includes RootFolder -ErrorAction Stop
    foreach ($candidate in $candidates) {
        $match = $allLists | Where-Object { $_.RootFolder.ServerRelativeUrl -eq $candidate }
        if ($match) { return $match }
    }

    # Last resort: match on the final path segment
    $leaf = ($identity -split '/')[-1]
    return ($allLists | Where-Object {
        $_.Title -eq $leaf -or $_.RootFolder.ServerRelativeUrl -like "*/$leaf"
    } | Select-Object -First 1)
}

Function Get-LibraryContent {
    param(
        [Parameter(Mandatory = $true)] $List,
        [Parameter(Mandatory = $false)][string] $FolderRef
    )

    $listTitle = $List.Title
    $items = Invoke-PnPWithRetry -OperationDescription "enumerate items in '$listTitle'" -ScriptBlock ({
        @(Get-PnPListItem -List $listTitle -PageSize 2000 -Fields "FileRef", "FileLeafRef", "FileDirRef" -ErrorAction Stop)
    }.GetNewClosure())

    $scope = if ($FolderRef) { $FolderRef.TrimEnd('/') + '/' } else { $null }

    $files   = New-Object System.Collections.Generic.List[object]
    $folders = New-Object System.Collections.Generic.List[object]

    foreach ($item in $items) {
        $ref = [string]$item["FileRef"]

        # Keep only descendants of the target folder; the folder itself is removed separately.
        if ($scope -and $ref -notlike "$scope*") { continue }

        $entry = [pscustomobject]@{
            Id      = $item.Id
            Name    = $item["FileLeafRef"]
            FileRef = $ref
            Depth   = $ref.Split('/').Count
        }
        if ($item.FileSystemObjectType -eq "Folder") { $folders.Add($entry) } else { $files.Add($entry) }
    }

    return [pscustomobject]@{
        Files   = @($files   | Sort-Object -Property Depth -Descending)
        Folders = @($folders | Sort-Object -Property Depth -Descending)
    }
}

# Resolve the configured folder value to a server-relative URL inside the library.
Function Resolve-TargetFolder {
    param(
        [Parameter(Mandatory = $true)][string] $Folder,
        [Parameter(Mandatory = $true)] $List
    )

    $path = $Folder.Trim()
    if ($path -match '^https?://') { $path = ([Uri]$path).AbsolutePath }
    $path = [Uri]::UnescapeDataString($path).Trim('/')

    $root = $List.RootFolder.ServerRelativeUrl.Trim('/')
    $serverRelative = if ($path -eq $root -or $path -like "$root/*") { "/$path" } else { "/$root/$path" }

    $resolved = Get-PnPFolder -Url $serverRelative -ErrorAction SilentlyContinue
    if ($null -eq $resolved) { return $null }

    return $serverRelative
}

Function Remove-LibraryItem {
    param(
        [Parameter(Mandatory = $true)][string] $ListTitle,
        [Parameter(Mandatory = $true)] $Item,
        [Parameter(Mandatory = $true)][string] $Kind
    )

    $itemId   = $Item.Id
    $fileRef  = $Item.FileRef
    $recycle  = $UseRecycleBin

    try {
        # Use the file/folder APIs the UI uses; ListItem.Recycle() returns E_ACCESSDENIED on
        # document library folders even when the caller can delete them in the browser.
        Invoke-PnPWithRetry -OperationDescription "delete $Kind '$fileRef'" -ScriptBlock ({
            if ($Kind -eq "folder") {
                $parent = $fileRef.Substring(0, $fileRef.LastIndexOf('/'))
                $name   = $fileRef.Substring($fileRef.LastIndexOf('/') + 1)
                if ($recycle) {
                    Remove-PnPFolder -Name $name -Folder $parent -Recycle -Force -ErrorAction Stop
                }
                else {
                    Remove-PnPFolder -Name $name -Folder $parent -Force -ErrorAction Stop
                }
            }
            else {
                if ($recycle) {
                    Remove-PnPFile -ServerRelativeUrl $fileRef -Recycle -Force -ErrorAction Stop
                }
                else {
                    Remove-PnPFile -ServerRelativeUrl $fileRef -Force -ErrorAction Stop
                }
            }
        }.GetNewClosure())

        Write-Host "    Deleted $Kind : $fileRef" -ForegroundColor Green
        Write-LogEntry -LogName $log -LogEntryText "Deleted $Kind '$fileRef'" -LogLevel "INFO"
        return "Deleted"
    }
    catch {
        $firstError = $_.Exception.Message

        # Fall back to the list item API before giving up.
        try {
            if ($recycle) {
                Remove-PnPListItem -List $ListTitle -Identity $itemId -Recycle -Force -ErrorAction Stop | Out-Null
            }
            else {
                Remove-PnPListItem -List $ListTitle -Identity $itemId -Force -ErrorAction Stop | Out-Null
            }
            Write-Host "    Deleted $Kind : $fileRef" -ForegroundColor Green
            Write-LogEntry -LogName $log -LogEntryText "Deleted $Kind '$fileRef' via list item fallback (first error: $firstError)" -LogLevel "INFO"
            return "Deleted"
        }
        catch {
            $secondError = $_.Exception.Message
            Write-Host "    Failed ($Kind): $fileRef" -ForegroundColor Red
            Write-Host "      Server response: $secondError" -ForegroundColor Red
            if ($secondError -ne $firstError) {
                Write-Host "      First attempt  : $firstError" -ForegroundColor DarkYellow
            }
            Write-LogEntry -LogName $log -LogEntryText "Failed to delete $Kind '$fileRef'. File/folder API: $firstError | List item API: $secondError" -LogLevel "ERROR"
            return "Failed"
        }
    }
}

Function Write-DeletePermissionPreflight {
    param([Parameter(Mandatory = $true)] $List)

    $writeOk = $false

    try {
        $web = Get-PnPWeb -Includes EffectiveBasePermissions, CurrentUser -ErrorAction Stop
        $listPerms = (Get-PnPList -Identity $List.Title -Includes EffectiveBasePermissions -ErrorAction Stop).EffectiveBasePermissions

        $identity = if ($AuthMethod -eq 'Application') { "app $clientId" } else { $web.CurrentUser.LoginName }
        $canDelete = $listPerms.Has([Microsoft.SharePoint.Client.PermissionKind]::DeleteListItems)
        $canManage = $listPerms.Has([Microsoft.SharePoint.Client.PermissionKind]::ManageLists)

        Write-Host "  Identity        : $identity" -ForegroundColor Gray
        Write-Host "  DeleteListItems : $canDelete" -ForegroundColor $(if ($canDelete) { 'Gray' } else { 'Red' })
        Write-Host "  ManageLists     : $canManage" -ForegroundColor $(if ($canManage) { 'Gray' } else { 'Red' })
        Write-LogEntry -LogName $log -LogEntryText "Preflight identity=$identity DeleteListItems=$canDelete ManageLists=$canManage" -LogLevel "INFO"
    }
    catch {
        Write-Host "  Could not read effective permissions: $($_.Exception.Message)" -ForegroundColor DarkYellow
        Write-LogEntry -LogName $log -LogEntryText "Preflight permission check failed: $($_.Exception.Message)" -LogLevel "WARNING"
    }

    # Delegated access is the intersection of the app registration's granted scopes and the user's
    # rights, so a read-only app registration produces access denied even for a site admin.
    try {
        $token  = Get-PnPAccessToken -Decoded -ErrorAction Stop
        $scopes = ($token.Claims | Where-Object { $_.Type -eq 'scp' -or $_.Type -eq 'roles' } |
                   Select-Object -ExpandProperty Value) -join ' '
        Write-Host "  Token scopes    : $scopes" -ForegroundColor Gray
        Write-LogEntry -LogName $log -LogEntryText "Preflight token scopes: $scopes" -LogLevel "INFO"

        if ($scopes -and $scopes -notmatch 'Manage|Write|FullControl|AllSites\.') {
            Write-Host "  WARNING: Token carries no write scope. Grant Sites.ReadWrite.All or" -ForegroundColor Red
            Write-Host "           AllSites.Manage to app $clientId, or use the PnP Management Shell app." -ForegroundColor Red
        }
    }
    catch {
        Write-LogEntry -LogName $log -LogEntryText "Could not decode access token: $($_.Exception.Message)" -LogLevel "DEBUG"
    }

    try {
        $site = Get-PnPSite -Includes ReadOnly -ErrorAction Stop
        if ($site.ReadOnly) {
            Write-Host "  Site is LOCKED READ-ONLY. No deletion is possible until the lock is cleared." -ForegroundColor Red
            Write-LogEntry -LogName $log -LogEntryText "Site $SiteUrl is read-only locked" -LogLevel "ERROR"
        }
    }
    catch {
        Write-LogEntry -LogName $log -LogEntryText "Could not read site lock state: $($_.Exception.Message)" -LogLevel "DEBUG"
    }

    # Live round trip: create a throwaway folder and delete it.
    $probeName = "_cleanuptool_probe_" + [Guid]::NewGuid().ToString('N').Substring(0, 8)
    $probeRoot = $List.RootFolder.ServerRelativeUrl
    try {
        Add-PnPFolder -Name $probeName -Folder $probeRoot -ErrorAction Stop | Out-Null
        Remove-PnPFolder -Name $probeName -Folder $probeRoot -Force -ErrorAction Stop
        Write-Host "  Write probe     : OK (created and deleted a test folder)" -ForegroundColor Green
        Write-LogEntry -LogName $log -LogEntryText "Preflight write probe succeeded" -LogLevel "INFO"
        $writeOk = $true
    }
    catch {
        Write-Host "  Write probe     : FAILED - $($_.Exception.Message)" -ForegroundColor Red
        Write-LogEntry -LogName $log -LogEntryText "Preflight write probe failed: $($_.Exception.Message)" -LogLevel "ERROR"
    }

    return $writeOk
}

#endregion

#region ---------------- Main ----------------

Write-Host "`n==== SharePoint Online Library Cleanup ====" -ForegroundColor Cyan
Write-Host "  Site    : $SiteUrl"
Write-Host "  Library : $LibraryUrl"
Write-Host "  Scope   : $(if ($FolderUrl) { "Folder '$FolderUrl'" } else { 'Entire library' })"
Write-Host "  Auth    : $AuthMethod"
Write-Host "  Mode    : $(if ($UseRecycleBin) { 'Recycle bin' } else { 'Permanent delete' })" -ForegroundColor Yellow
Write-Host "  Log     : $log`n"
Write-LogEntry -LogName $log -LogEntryText "Starting cleanup. Site=$SiteUrl Library=$LibraryUrl Folder='$FolderUrl' Recycle=$UseRecycleBin" -LogLevel "INFO"

$confirm = Read-Host "This permanently removes content from the $(if ($FolderUrl) { 'folder' } else { 'library' }). Type DELETE to continue"
if ($confirm -cne "DELETE") {
    Write-Host "Operation cancelled by user." -ForegroundColor Yellow
    Write-LogEntry -LogName $log -LogEntryText "Cleanup cancelled by user" -LogLevel "INFO"
    exit
}

try {
    if ($AuthMethod -eq 'Application') {
        Connect-PnPOnline -Url $SiteUrl -ClientId $clientId -Tenant $tenantId -Thumbprint $CertificateThumbprint -ErrorAction Stop
    }
    else {
        Connect-PnPOnline -Url $SiteUrl -ClientId $clientId -Tenant $tenantId -Interactive -ErrorAction Stop
    }
    Write-LogEntry -LogName $log -LogEntryText "Connected to $SiteUrl using $AuthMethod authentication" -LogLevel "INFO"
}
catch {
    Write-Host "ERROR: Failed to connect to the site." -ForegroundColor Red
    Write-Host "  URL    : $SiteUrl" -ForegroundColor Red
    Write-Host "  Detail : $_" -ForegroundColor Red
    Write-LogEntry -LogName $log -LogEntryText "Failed to connect to ${SiteUrl}: $_" -LogLevel "ERROR"
    exit
}

$list = Resolve-TargetLibrary -Library $LibraryUrl -WebUrl $SiteUrl
if ($null -eq $list) {
    Write-Host "ERROR: Library '$LibraryUrl' was not found on $SiteUrl." -ForegroundColor Red
    Write-LogEntry -LogName $log -LogEntryText "Library '$LibraryUrl' not found on $SiteUrl" -LogLevel "ERROR"
    exit
}

$listTitle = $list.Title
Write-Host "Resolved library: '$listTitle' ($($list.RootFolder.ServerRelativeUrl))" -ForegroundColor Cyan
Write-LogEntry -LogName $log -LogEntryText "Resolved library '$listTitle' at $($list.RootFolder.ServerRelativeUrl)" -LogLevel "INFO"

$targetFolderRef = $null
if (-not [string]::IsNullOrWhiteSpace($FolderUrl)) {
    $targetFolderRef = Resolve-TargetFolder -Folder $FolderUrl -List $list
    if ($null -eq $targetFolderRef) {
        Write-Host "ERROR: Folder '$FolderUrl' was not found in library '$listTitle'." -ForegroundColor Red
        Write-LogEntry -LogName $log -LogEntryText "Folder '$FolderUrl' not found in library '$listTitle'" -LogLevel "ERROR"
        exit
    }
    Write-Host "Scoped to folder: $targetFolderRef" -ForegroundColor Cyan
    Write-LogEntry -LogName $log -LogEntryText "Scoped cleanup to folder $targetFolderRef" -LogLevel "INFO"
}

$writeAccessOk = Write-DeletePermissionPreflight -List $list
if (-not $writeAccessOk) {
    Write-Host "`nThe write probe failed, so every delete in this library will return access denied." -ForegroundColor Red
    Write-Host "Resolve access before re-running. Most common causes:" -ForegroundColor Yellow
    Write-Host "  - App registration '$clientId' lacks a write scope (needs Sites.ReadWrite.All / AllSites.Manage)" -ForegroundColor Yellow
    Write-Host "  - Signed-in account is not a site collection admin / site owner" -ForegroundColor Yellow
    Write-Host "  - Site is locked read-only, or the library is under a preservation hold" -ForegroundColor Yellow
    Write-LogEntry -LogName $log -LogEntryText "Aborting: write probe failed" -LogLevel "ERROR"

    $continueAnyway = Read-Host "`nContinue anyway? (y/N)"
    if ($continueAnyway -ne 'y') { exit }
}

$totalDeleted = 0
$totalFailed  = 0
$remaining    = -1

for ($pass = 1; $pass -le $MaxPasses; $pass++) {
    Write-Host "`n-- Pass $pass of $MaxPasses --" -ForegroundColor Cyan

    $content = Get-LibraryContent -List $list -FolderRef $targetFolderRef
    $fileCount   = $content.Files.Count
    $folderCount = $content.Folders.Count

    if ($fileCount -eq 0 -and $folderCount -eq 0) {
        Write-Host "  $(if ($targetFolderRef) { 'Folder' } else { 'Library' }) is empty." -ForegroundColor Green
        $remaining = 0
        break
    }

    Write-Host "  Found $fileCount file(s) and $folderCount folder(s)." -ForegroundColor Gray
    Write-LogEntry -LogName $log -LogEntryText "Pass ${pass}: $fileCount files, $folderCount folders" -LogLevel "INFO"

    $passDeleted = 0
    $passFailed  = 0
    $failedPaths = New-Object System.Collections.Generic.List[string]

    # Files first, deepest paths first
    foreach ($file in $content.Files) {
        switch (Remove-LibraryItem -ListTitle $listTitle -Item $file -Kind "file") {
            "Deleted" { $passDeleted++ }
            "Failed"  { $passFailed++; $failedPaths.Add($file.FileRef) }
        }
    }

    # Folders deepest-first; skip any folder that still contains undeleted content
    foreach ($folder in $content.Folders) {
        $folderPath = $folder.FileRef.TrimEnd('/') + '/'
        $hasFailedChild = $failedPaths | Where-Object { $_ -like "$folderPath*" } | Select-Object -First 1

        if ($hasFailedChild) {
            Write-Host "    Skipped folder (undeleted content inside): $($folder.FileRef)" -ForegroundColor Yellow
            Write-LogEntry -LogName $log -LogEntryText "Skipped folder '$($folder.FileRef)' because it still contains undeleted items" -LogLevel "WARNING"
            $passFailed++
            $failedPaths.Add($folder.FileRef)
            continue
        }

        switch (Remove-LibraryItem -ListTitle $listTitle -Item $folder -Kind "folder") {
            "Deleted" { $passDeleted++ }
            "Failed"  { $passFailed++; $failedPaths.Add($folder.FileRef) }
        }
    }

    $totalDeleted += $passDeleted
    $totalFailed  += $passFailed
    $remaining     = $passFailed

    Write-Host "  Pass $pass complete. Deleted: $passDeleted; Failed: $passFailed" -ForegroundColor Cyan
    Write-LogEntry -LogName $log -LogEntryText "Pass ${pass} complete. Deleted=$passDeleted, Failed=$passFailed" -LogLevel "INFO"

    if ($passDeleted -eq 0) {
        Write-Host "  No progress made this pass; stopping content deletion." -ForegroundColor Yellow
        break
    }
}

Write-Host "`nContent cleanup summary - Deleted: $totalDeleted; Failed: $totalFailed" -ForegroundColor Cyan
Write-LogEntry -LogName $log -LogEntryText "Content cleanup summary. Deleted=$totalDeleted, Failed=$totalFailed" -LogLevel "INFO"

if ($targetFolderRef) {
    if ($RemoveTargetFolderWhenEmpty) {
        if ($remaining -gt 0) {
            Write-Host "Folder still contains $remaining item(s) that could not be deleted; skipping folder removal." -ForegroundColor Yellow
            Write-LogEntry -LogName $log -LogEntryText "Skipped folder removal; $remaining items remain" -LogLevel "WARNING"
        }
        else {
            $parent = $targetFolderRef.Substring(0, $targetFolderRef.LastIndexOf('/'))
            $name   = $targetFolderRef.Substring($targetFolderRef.LastIndexOf('/') + 1)
            try {
                Invoke-PnPWithRetry -OperationDescription "remove folder '$targetFolderRef'" -ScriptBlock ({
                    if ($UseRecycleBin) {
                        Remove-PnPFolder -Name $name -Folder $parent -Recycle -Force -ErrorAction Stop
                    }
                    else {
                        Remove-PnPFolder -Name $name -Folder $parent -Force -ErrorAction Stop
                    }
                }.GetNewClosure())
                Write-Host "Folder '$targetFolderRef' removed." -ForegroundColor Green
                Write-LogEntry -LogName $log -LogEntryText "Removed folder '$targetFolderRef'" -LogLevel "INFO"
            }
            catch {
                $removalError = $_.Exception.Message
                Write-Host "`nERROR: Folder '$targetFolderRef' could not be deleted." -ForegroundColor Red
                Write-Host "  Server response : $removalError" -ForegroundColor Red
                Write-Error "Folder '$targetFolderRef' could not be deleted: $removalError"
                Write-LogEntry -LogName $log -LogEntryText "Folder removal failed. Folder='$targetFolderRef' ServerResponse=$removalError" -LogLevel "ERROR"
            }
        }
    }
}
elseif ($RemoveLibraryWhenEmpty) {
    if ($remaining -gt 0) {
        Write-Host "Library still contains $remaining item(s) that could not be deleted; skipping library removal." -ForegroundColor Yellow
        Write-LogEntry -LogName $log -LogEntryText "Skipped library removal; $remaining items remain" -LogLevel "WARNING"
    }
    else {
        try {
            Invoke-PnPWithRetry -OperationDescription "remove library '$listTitle'" -ScriptBlock ({
                Remove-PnPList -Identity $listTitle -Force -ErrorAction Stop
            }.GetNewClosure())
            Write-Host "Library '$listTitle' removed." -ForegroundColor Green
            Write-LogEntry -LogName $log -LogEntryText "Removed library '$listTitle' from $SiteUrl" -LogLevel "INFO"
        }
        catch {
            $removalError = $_.Exception.Message

            $allowDeletion = $null
            try {
                $allowDeletion = (Get-PnPList -Identity $listTitle -Includes AllowDeletion -ErrorAction Stop).AllowDeletion
            }
            catch { }

            Write-Host "`nERROR: Library '$listTitle' could not be deleted." -ForegroundColor Red
            Write-Host "  Site            : $SiteUrl" -ForegroundColor Red
            Write-Host "  Library         : $listTitle ($($list.RootFolder.ServerRelativeUrl))" -ForegroundColor Red
            Write-Host "  AllowDeletion   : $allowDeletion" -ForegroundColor Red
            Write-Host "  Server response : $removalError" -ForegroundColor Red
            Write-Error "Library '$listTitle' on $SiteUrl could not be deleted (AllowDeletion=$allowDeletion): $removalError"
            Write-LogEntry -LogName $log -LogEntryText "Library removal failed. Site=$SiteUrl Library='$listTitle' AllowDeletion=$allowDeletion ServerResponse=$removalError" -LogLevel "ERROR"
        }
    }
}

Write-Host "`nDone. Log file: $log" -ForegroundColor Cyan
Write-LogEntry -LogName $log -LogEntryText "Cleanup run finished" -LogLevel "INFO"

#endregion
