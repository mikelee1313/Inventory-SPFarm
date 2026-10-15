#Requires -Version 7.2
#Requires -Modules PnP.PowerShell

<#
.SYNOPSIS
    Clones the structure and content of a SharePoint Online site
    into a new modern site collection.

.DESCRIPTION
    1. Creates the destination modern site collection if requested.
    2. Extracts a PnP provisioning template from the source site.
    3. Normalizes null modern-page web-part maps and applies the template.
    4. Copies folders and files from document libraries.
    5. Copies standard items from non-document lists.
    6. Writes an incremental CSV log, console transcript, and PnP debug trace.

.PARAMETER Tenant
    Entra ID tenant ID or tenant domain, such as contoso.onmicrosoft.com.

.PARAMETER Thumbprint
    Thumbprint of a certificate in the Windows certificate store, accessible
    to the account running the script. The certificate must have a private key.
    Upload its matching public certificate to the Entra ID app registration.

.PARAMETER MaxItemsPerList
    Required positive maximum number of source items selected per in-scope list
    or library per run, ordered by ascending source ID. Folders do not count.
    Failed items and existing files skipped without OverwriteFiles still consume
    a selection slot; later items are not substituted. Library folders are
    copied without a limit. Existing destination content is not removed.
    Does not limit template-managed pages or branding assets.
    CopyListItems must still be enabled to copy non-document list items.

.EXAMPLE

    .\SPO-SiteCloner.ps1 -ClientId 'abc64618-283f-47ba-a185-50d935d51d57' -Tenant '9cfc42cb-51da-4055-87e9-b20a170b6ba3' 
    -TenantAdminUrl 'https://m365cpi13246019-admin.sharepoint.com' -Thumbprint  'B696FDCFE1453F3FBC6031F54DE988DA0ED905A9' 
    -SourceSiteUrl 'https://m365cpi13246019.sharepoint.com/sites/SalesandMarketing' 
    -TargetSiteUrl 'https://m365cpi13246019.sharepoint.com/sites/sales3' 
    -TargetSiteTitle "Sales and Marketing3" 
    -TargetSiteOwner 'admin@M365CPI13246019.onmicrosoft.com' -MaxItemsPerList 10 

.NOTES
    Site-header provisioning is skipped by the internal SkipSiteHeader setting,
    which defaults to true. Change it in the script to false to clone the site
    header layout. No command-line switch is required or accepted for this setting.
    Modern-page headers, navigation, and other branding handlers are unaffected.
    Before provisioning, an unused legacy Design File content type definition
    (0x010100A2E0023C4C216242892749CE60EC3D07) is omitted from the normalized
    package only if its name conflicts with a different destination ID.
    Referenced definitions and other same-name/different-ID conflicts stop the
    run for review. Destination content types are never deleted or remapped.
    Provisioning failures stop the run. Required destination lists and libraries
    are validated before any content copying starts.

    Custom script is not enabled by this script and is not required for ordinary
    lists and document libraries. NoScript can cause classic branding assets such
    as the Style Library to be skipped. If these are required, a SharePoint
    administrator must review and temporarily allow custom script on the target
    under the customer's policy; SharePoint normally resets it within 24 hours.
    Rerunning against a partial destination can duplicate standard list items;
    use a clean test destination when validating a corrected run.

    Each run writes CloneLog.csv, CloneTranscript.txt, and PnPTrace.log to its
    timestamped working folder. Logs include site URLs, object names, and error
    details; protect them as customer data. The CSV is written after each event.
    Null serverProcessedContent maps are replaced with empty JSON objects in
    NormalizedTemplate\SourceSiteTemplate.pnp. The original export remains unchanged.
    Logs describe attempted operations, not a full SharePoint before/after audit.
    A failed provisioning run can leave a partially modified destination.

    All PnP connections use certificate-based application authentication.
    No interactive or delegated sign-in is used.

    Grant the app SharePoint APPLICATION permissions with tenant admin consent.
    Sites.FullControl.All supports provisioning, permissions, and tenant site
    creation in this script. Do not substitute Microsoft Graph permissions or
    delegated permissions for these SharePoint application permissions.
    For cloning existing sites with Sites.Selected instead, explicitly grant
    the app full control on both sites and omit CreateTargetSite; tenant admin
    operations require broader permissions. Some provisioning handlers or
    SharePoint features may still require additional permissions.
    TargetSiteOwner sets the new site's owner; it is not the sign-in identity.
    Protect the certificate's private key and grant access only to the account
    running the script. No PFX path or password is required by this script.

    This is not a full-fidelity migration tool.

    Not preserved:
      - Version history
      - Sharing links
      - Audit history
      - Retention record state
      - Power Automate flows
      - Teams or Planner content
      - Original item IDs
      - Original Created By / Modified By in every scenario
      - Every complex field type
      - All unique permissions
#>

#region Script Parameters
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$SourceSiteUrl,

    [Parameter(Mandatory)]
    [string]$TargetSiteUrl,

    [Parameter(Mandatory)]
    [string]$TargetSiteTitle,

    [Parameter(Mandatory)]
    [string]$TargetSiteOwner,

    [Parameter(Mandatory)]
    [string]$TenantAdminUrl,

    [Parameter(Mandatory)]
    [string]$ClientId,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$Tenant,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$Thumbprint,

    [ValidateSet(
        "CommunicationSite",
        "TeamSiteWithoutMicrosoft365Group"
    )]
    [string]$TargetSiteType = "TeamSiteWithoutMicrosoft365Group",

    [string]$WorkingFolder = "C:\Temp\PnPSiteClone",

    [switch]$CreateTargetSite,

    [switch]$CopyDocumentLibraries = $true,

    [switch]$CopyListItems = $true,

    [Parameter(Mandatory)]
    [ValidateRange(1, [int]::MaxValue)]
    [int]$MaxItemsPerList,

    [switch]$OverwriteFiles,

    [switch]$IncludeHiddenLists
)
#endregion Script Parameters

#region Initialization and Run Configuration
$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

#------------------------------------------------------------
# Initialization
#------------------------------------------------------------

$Timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$RunFolder = Join-Path $WorkingFolder $Timestamp
$TemplatePath = Join-Path $RunFolder "SourceSiteTemplate.pnp"
$DownloadFolder = Join-Path $RunFolder "DownloadedContent"
$LogPath = Join-Path $RunFolder "CloneLog.csv"
$TranscriptPath = Join-Path $RunFolder "CloneTranscript.txt"
$TracePath = Join-Path $RunFolder "PnPTrace.log"
$ProvisioningTemplatePath = $TemplatePath
$TranscriptStarted = $false
$TraceStarted = $false
$RunTimer = [System.Diagnostics.Stopwatch]::StartNew()
$SkipSiteHeader = $true

New-Item -Path $RunFolder -ItemType Directory -Force | Out-Null
New-Item -Path $DownloadFolder -ItemType Directory -Force | Out-Null

$script:LogEntries = [System.Collections.Generic.List[object]]::new()
$script:SubfolderPagesToRelocate = [System.Collections.Generic.List[object]]::new()
#endregion Initialization and Run Configuration

#region Helper Functions
#region Incremental Logging
function Write-CloneLog {
    param(
        [Parameter(Mandatory)]
        [ValidateSet("Information", "Warning", "Error")]
        [string]$Level,

        [Parameter(Mandatory)]
        [string]$Operation,

        [string]$Object = "",

        [Parameter(Mandatory)]
        [string]$Message,

        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    $Details = ""
    if ($null -ne $ErrorRecord) {
        $Details = @(
            $ErrorRecord.Exception.ToString()
            "Error ID: $($ErrorRecord.FullyQualifiedErrorId)"
            "Category: $($ErrorRecord.CategoryInfo)"
            $ErrorRecord.InvocationInfo.PositionMessage
            "PowerShell stack: $($ErrorRecord.ScriptStackTrace)"
        ) -join [Environment]::NewLine
    }

    $Entry = [pscustomobject]@{
        Time      = (Get-Date).ToString("o")
        Sequence  = $script:LogEntries.Count + 1
        ElapsedSeconds = [math]::Round($RunTimer.Elapsed.TotalSeconds, 3)
        Level     = $Level
        Operation = $Operation
        Object    = $Object
        Message   = $Message
        Details   = $Details
    }

    $Entry | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8 -Append
    $script:LogEntries.Add($Entry)

    $Color = switch ($Level) {
        "Information" { "Green" }
        "Warning"     { "Yellow" }
        "Error"       { "Red" }
    }

    Write-Host "[$Level] [$Operation] $Object $Message" -ForegroundColor $Color
}
#endregion Incremental Logging

#region Template Validation and Repair
function Repair-PnPPageWebPartJson {
        param(
            [Parameter(Mandatory)]
            [string]$Path,

            [Parameter(Mandatory)]
            [AllowEmptyCollection()]
            [object[]]$DestinationContentTypes
        )

        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $Updates = @{}
        $Documents = @{}
        $RepairCount = 0
        $script:SubfolderPagesToRelocate.Clear()
        $MapNames = @(
            "htmlStrings", "searchablePlainTexts", "imageSources",
            "links", "componentDependencies", "customMetadata"
        )

        $Archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
        try {
            foreach ($Entry in $Archive.Entries) {
                if (-not $Entry.FullName.EndsWith(".xml", [System.StringComparison]::OrdinalIgnoreCase)) {
                    continue
                }

                $Reader = [System.IO.StreamReader]::new($Entry.Open())
                try {
                    $Document = [System.Xml.XmlDocument]::new()
                    $Document.PreserveWhitespace = $true
                    $Document.LoadXml($Reader.ReadToEnd())
                }
                finally {
                    $Reader.Dispose()
                }

                $Documents[$Entry.FullName] = $Document
            }

            foreach ($EntryName in $Documents.Keys) {
                $Document = $Documents[$EntryName]
                $Changed = $false

                foreach ($ContentType in @($Document.SelectNodes(
                    "//*[local-name()='ContentTypes']/*[local-name()='ContentType']"
                ))) {
                    $Id = $ContentType.GetAttribute("ID")
                    $Name = $ContentType.GetAttribute("Name")
                    if (@($DestinationContentTypes | Where-Object { [string]$_.Id -eq $Id }).Count -gt 0) {
                        continue
                    }
                    $Conflicts = @($DestinationContentTypes | Where-Object { $_.Name -eq $Name })
                    if ($Conflicts.Count -eq 0) {
                        continue
                    }

                    $ConflictIds = ($Conflicts | ForEach-Object { [string]$_.Id }) -join ", "
                    if ($Name -ne "Design File" -or $Id -ne "0x010100A2E0023C4C216242892749CE60EC3D07") {
                        throw "Content type '$Name' has template ID '$Id' but destination ID(s) '$ConflictIds'. Resolve this conflict before provisioning."
                    }

                    # Include descendant IDs and references in other XML parts, not just list bindings.
                    foreach ($Part in $Documents.Values) {
                        foreach ($Reference in $Part.SelectNodes("//@* | //text()")) {
                            if ([object]::ReferenceEquals($Reference, $ContentType.GetAttributeNode("ID"))) {
                                continue
                            }
                            if ($Reference.Value.IndexOf($Id, [System.StringComparison]::OrdinalIgnoreCase) -ge 0 -or
                                $Reference.Value -match '\{contenttypeid:\s*Design File\s*\}') {
                                throw "Conflicting content type '$Name' ($Id; destination: $ConflictIds) is referenced in the package. Resolve its dependencies before provisioning."
                            }
                        }
                    }

                    $null = $ContentType.ParentNode.RemoveChild($ContentType)
                    $Changed = $true
                    Write-CloneLog -Level Warning -Operation "Template repair" -Object $Name `
                        -Message "Omitted unused legacy content type '$Id'; destination already has '$Name' with ID(s) '$ConflictIds'. No destination content type was changed."
                }

                # Normalize modern pages defined in subfolders (e.g. folder1/Folder2/Page.aspx).
                # PnP Framework has a Windows-specific bug in ObjectClientSidePages.DeterminePageName
                # where System.IO.Path.GetDirectoryName() injects backslashes into the page name,
                # causing PnP Core's LoadClientSidePage to fail to find the page and throw NullReferenceException.
                # We flatten the page name in the template so PnP provisions it cleanly into SitePages root,
                # then relocate it to its destination subfolder post-provisioning.
                foreach ($PageNode in $Document.SelectNodes("//*[local-name()='ClientSidePage']")) {
                    $RawPageName = $PageNode.GetAttribute("PageName")
                    if ($RawPageName -match "[\\/]") {
                        $FlatPageName = [System.IO.Path]::GetFileName($RawPageName)
                        $Subfolder = ([System.IO.Path]::GetDirectoryName($RawPageName)).Replace("\", "/")
                        $script:SubfolderPagesToRelocate.Add([pscustomobject]@{
                            OriginalRelativePath = $RawPageName.Replace("\", "/")
                            FlatPageName         = $FlatPageName
                            Subfolder            = $Subfolder
                        })
                        $PageNode.SetAttribute("PageName", $FlatPageName)
                        $Changed = $true
                        Write-CloneLog -Level Information -Operation "Template repair" -Object $RawPageName `
                            -Message "Flattened subfolder page name to '$FlatPageName'. Will relocate to '$Subfolder/' post-provisioning."
                    }
                }

                foreach ($Control in $Document.SelectNodes("//*[@JsonControlData]")) {
                    $Page = $Control.SelectSingleNode("ancestor::*[local-name()='ClientSidePage'][1]")
                    $ObjectName = "$EntryName/$($Control.GetAttribute('ControlId'))"
                    if ($null -ne $Page) {
                        $ObjectName = "$($Page.GetAttribute('PageName'))/$($Control.GetAttribute('ControlId'))"
                    }

                    try {
                        $Json = ConvertFrom-Json -InputObject $Control.GetAttribute("JsonControlData") -AsHashtable
                        if ($Json -isnot [System.Collections.IDictionary]) {
                            throw "Web-part JSON must be an object."
                        }
                        if (-not $Json.ContainsKey("serverProcessedContent")) {
                            continue
                        }

                        $RepairedMaps = [System.Collections.Generic.List[string]]::new()
                        if ($null -eq $Json["serverProcessedContent"]) {
                            $Json["serverProcessedContent"] = @{}
                            $RepairedMaps.Add("serverProcessedContent")
                        }
                        else {
                            $Content = $Json["serverProcessedContent"]
                            if ($Content -isnot [System.Collections.IDictionary]) {
                                throw "serverProcessedContent must be an object or null."
                            }
                            foreach ($Name in $MapNames) {
                                if ($Content.ContainsKey($Name)) {
                                    if ($null -eq $Content[$Name]) {
                                        $Content[$Name] = @{}
                                        $RepairedMaps.Add("serverProcessedContent.$Name")
                                    }
                                    elseif ($Content[$Name] -isnot [System.Collections.IDictionary]) {
                                        throw "serverProcessedContent.$Name must be an object or null."
                                    }
                                }
                            }
                        }

                        if ($RepairedMaps.Count -gt 0) {
                            $Control.SetAttribute("JsonControlData", (
                                ConvertTo-Json -InputObject $Json -Depth 100 -Compress -WarningAction Stop
                            ))
                            $Changed = $true
                            $RepairCount += $RepairedMaps.Count
                            Write-CloneLog -Level Warning -Operation "Template repair" -Object $ObjectName `
                                -Message "Normalized null maps to empty objects: $($RepairedMaps -join ', ')."
                        }
                    }
                    catch {
                        Write-CloneLog -Level Error -Operation "Template validation" -Object $ObjectName `
                            -Message $_.Exception.Message -ErrorRecord $_
                        throw
                    }
                }

                if ($Changed) {
                    $Updates[$EntryName] = $Document
                }
            }
        }
        finally {
            $Archive.Dispose()
        }

        if ($Updates.Count -eq 0) {
            Write-CloneLog -Level Information -Operation "Template validation" -Object $Path `
                -Message "No template repairs required (no content-type conflicts, null maps, or subfolder pages)."
            return $Path
        }

        # PnP resolves the template XML using the package basename when TemplateFileName is empty.
        $NormalizedFolder = Join-Path (Split-Path $Path -Parent) "NormalizedTemplate"
        New-Item -Path $NormalizedFolder -ItemType Directory -Force | Out-Null
        $NormalizedPath = Join-Path $NormalizedFolder ([System.IO.Path]::GetFileName($Path))
        Copy-Item -LiteralPath $Path -Destination $NormalizedPath -ErrorAction Stop
        $Archive = [System.IO.Compression.ZipFile]::Open(
            $NormalizedPath, [System.IO.Compression.ZipArchiveMode]::Update
        )
        try {
            foreach ($Name in $Updates.Keys) {
                $Archive.GetEntry($Name).Delete()
                $Entry = $Archive.CreateEntry($Name)
                $Stream = $Entry.Open()
                try {
                    $Updates[$Name].Save($Stream)
                }
                finally {
                    $Stream.Dispose()
                }
            }
        }
        finally {
            $Archive.Dispose()
        }

        Write-CloneLog -Level Information -Operation "Template repair" -Object $NormalizedPath `
            -Message "Saved normalized package including content-type repairs, $RepairCount null-map repair(s), and $($script:SubfolderPagesToRelocate.Count) flattened subfolder page(s). Original package preserved at '$Path'."
        return $NormalizedPath
}
    #endregion Template Validation and Repair

    #region Log Export
function Save-CloneLog {
    $script:LogEntries |
        Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8
}
    #endregion Log Export

    #region File Names and Site-Relative Paths
function Convert-ToSafeFileName {
    param(
        [Parameter(Mandatory)]
        [string]$Name
    )

    $InvalidCharacters = [System.IO.Path]::GetInvalidFileNameChars()

    foreach ($Character in $InvalidCharacters) {
        $Name = $Name.Replace($Character, "_")
    }

    return $Name
}

function Get-SiteRelativePath {
    param(
        [Parameter(Mandatory)]
        [string]$ServerRelativeUrl,

        [Parameter(Mandatory)]
        [string]$WebServerRelativeUrl
    )

    $WebRoot = $WebServerRelativeUrl.TrimEnd("/")

    if ([string]::IsNullOrWhiteSpace($WebRoot)) {
        return $ServerRelativeUrl.TrimStart("/")
    }

    if ($ServerRelativeUrl.StartsWith(
        $WebRoot,
        [System.StringComparison]::OrdinalIgnoreCase
    )) {
        return $ServerRelativeUrl.Substring($WebRoot.Length).TrimStart("/")
    }

    return $ServerRelativeUrl.TrimStart("/")
}
#endregion File Names and Site-Relative Paths

#region System Library Filtering
function Test-IsSystemLibrary {
    param(
        [Parameter(Mandatory)]
        $List
    )

    $ExcludedTitles = @(
        "Form Templates",
        "Converted Forms",
        "Master Page Gallery",
        "Site Collection Documents",
        "Site Collection Images",
        "Style Library",
        "Preservation Hold Library",
        "User Information List",
        "Site Assets",
        "Site Pages",
        "Web Part Gallery",
        "List Template Gallery",
        "Theme Gallery",
        "Solution Gallery"
    )

    if ($ExcludedTitles -contains $List.Title) {
        return $true
    }

    if (
        $List.RootFolder.ServerRelativeUrl -match
        "/(_catalogs|_private|Style Library|PreservationHoldLibrary)(/|$)"
    ) {
        return $true
    }

    return $false
}
#endregion System Library Filtering

function Get-CloneSourceLists {
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$Lists
    )

    $ExcludedListTemplates = @(110, 111, 112, 113, 114, 117, 140, 150, 544)
    foreach ($List in $Lists) {
        if ($List.Hidden -and -not $IncludeHiddenLists) {
            continue
        }
        if ([int]$List.BaseType -eq 1) {
            if (-not $CopyDocumentLibraries) {
                continue
            }
            if (Test-IsSystemLibrary -List $List) {
                Write-CloneLog -Level Warning -Operation "Library" -Object $List.Title `
                    -Message "Skipped system or template-managed library."
                continue
            }
        }
        else {
            if (-not $CopyListItems) {
                continue
            }
            if ($ExcludedListTemplates -contains [int]$List.BaseTemplate) {
                Write-CloneLog -Level Warning -Operation "List" -Object $List.Title `
                    -Message "Skipped system list template."
                continue
            }
        }
        $List
    }
}

function Assert-CloneDestinationLists {
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$SourceLists,

        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$TargetLists,

        [Parameter(Mandatory)]
        [string]$SourceWebServerRelativeUrl,

        [Parameter(Mandatory)]
        [string]$TargetWebServerRelativeUrl
    )

    $Failures = 0
    foreach ($List in $SourceLists) {
        $RelativePath = Get-SiteRelativePath -ServerRelativeUrl $List.RootFolder.ServerRelativeUrl `
            -WebServerRelativeUrl $SourceWebServerRelativeUrl
        $ExpectedPath = "$($TargetWebServerRelativeUrl.TrimEnd('/'))/$RelativePath"
        $Matches = @($TargetLists | Where-Object {
            [uri]::UnescapeDataString($_.RootFolder.ServerRelativeUrl).TrimEnd("/") -eq
                [uri]::UnescapeDataString($ExpectedPath).TrimEnd("/")
        })
        if ($Matches.Count -ne 1 -or $Matches[0].Title -ne $List.Title -or
            [int]$Matches[0].BaseTemplate -ne [int]$List.BaseTemplate) {
            $Failures++
            Write-CloneLog -Level Error -Operation "Structure validation" -Object $List.Title `
                -Message "Expected destination list at '$ExpectedPath' with title '$($List.Title)' and template '$($List.BaseTemplate)' is missing or incompatible. Content copying has not started."
        }
    }
    if ($Failures -gt 0) {
        throw "Destination structure validation failed for $Failures list(s) or library/libraries. Review the log before retrying."
    }
    Write-CloneLog -Level Information -Operation "Structure validation" `
        -Message "Validated all $($SourceLists.Count) in-scope destination lists and libraries."
}

#region Destination Folder Creation
#------------------------------------------------------------
# Folder creation
#------------------------------------------------------------

function Ensure-TargetFolder {
    param(
        [Parameter(Mandatory)]
        [string]$SiteRelativeFolder,

        [Parameter(Mandatory)]
        $Connection
    )

    $NormalizedPath = $SiteRelativeFolder.Trim("/")

    if ([string]::IsNullOrWhiteSpace($NormalizedPath)) {
        return
    }

    Resolve-PnPFolder `
        -SiteRelativePath $NormalizedPath `
        -Connection $Connection | Out-Null
}
    #endregion Destination Folder Creation

    #region List Field-Value Conversion
#------------------------------------------------------------
# Field-value conversion for standard SharePoint lists
#------------------------------------------------------------

function Convert-ListItemValues {
    param(
        [Parameter(Mandatory)]
        $Item,

        [Parameter(Mandatory)]
        [array]$Fields
    )

    $Values = @{}

    $SystemFieldNames = @(
        "ID",
        "GUID",
        "Author",
        "Editor",
        "Created",
        "Modified",
        "Attachments",
        "ContentType",
        "ContentTypeId",
        "_UIVersionString",
        "_ModerationStatus",
        "FileRef",
        "FileDirRef",
        "FileLeafRef",
        "FSObjType",
        "UniqueId",
        "owshiddenversion",
        "AppAuthor",
        "AppEditor",
        "ComplianceAssetId",
        "_ComplianceFlags",
        "_ComplianceTag",
        "_ComplianceTagWrittenTime",
        "_ComplianceTagUserId",
        "_IsRecord",
        "MetaInfo",
        "Order"
    )

    foreach ($Field in $Fields) {
        $InternalName = $Field.InternalName

        if ($SystemFieldNames -contains $InternalName) {
            continue
        }

        if ($Field.Hidden -or $Field.ReadOnlyField -or $Field.Sealed) {
            continue
        }

        if (-not $Item.FieldValues.ContainsKey($InternalName)) {
            continue
        }

        $Value = $Item.FieldValues[$InternalName]

        if ($null -eq $Value) {
            continue
        }

        try {
            switch ($Field.TypeAsString) {
                { $_ -in @("Computed", "Lookup", "LookupMulti") } {
                    Write-CloneLog `
                        -Level Warning `
                        -Operation "List field" `
                        -Object $InternalName `
                        -Message "Skipped computed or lookup field."
                    continue
                }

                { $_ -in @("User", "UserMulti") } {
                    if ($Value -is [System.Array]) {
                        $Emails = @(
                            $Value |
                            ForEach-Object {
                                if ($_.Email) {
                                    $_.Email
                                }
                                elseif ($_.LookupValue) {
                                    $_.LookupValue
                                }
                            }
                        )

                        if ($Emails.Count -gt 0) {
                            $Values[$InternalName] = $Emails
                        }
                    }
                    elseif ($Value.Email) {
                        $Values[$InternalName] = $Value.Email
                    }

                    continue
                }

                "TaxonomyFieldType" {
                    if ($Value.TermGuid) {
                        $Values[$InternalName] =
                            "$($Value.Label)|$($Value.TermGuid)"
                    }

                    continue
                }

                "TaxonomyFieldTypeMulti" {
                    $TaxonomyValues = @(
                        $Value |
                        ForEach-Object {
                            "$($_.Label)|$($_.TermGuid)"
                        }
                    )

                    if ($TaxonomyValues.Count -gt 0) {
                        $Values[$InternalName] =
                            $TaxonomyValues -join ";"
                    }

                    continue
                }

                "URL" {
                    if ($Value.Url) {
                        $Values[$InternalName] =
                            "$($Value.Url), $($Value.Description)"
                    }

                    continue
                }

                default {
                    $Values[$InternalName] = $Value
                }
            }
        }
        catch {
            Write-CloneLog `
                -Level Warning `
                -Operation "List field" `
                -Object $InternalName `
                -Message "Could not convert value: $($_.Exception.Message)" `
                -ErrorRecord $_
        }
    }

    return $Values
}
#endregion List Field-Value Conversion

#region Limited Source Item Selection
#------------------------------------------------------------
# Copy regular SharePoint list items
#------------------------------------------------------------

function Get-LimitedSourceItems {
    param(
        [Parameter(Mandatory)]
        $SourceList,

        [Parameter(Mandatory)]
        $SourceConnection
    )

    $PageSize = [math]::Min(500, $MaxItemsPerList)
    $Query = @"
<View Scope='RecursiveAll'>
  <Query>
    <Where><Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>0</Value></Eq></Where>
    <OrderBy><FieldRef Name='ID' Ascending='TRUE'/></OrderBy>
  </Query>
  <RowLimit Paged='TRUE'>$PageSize</RowLimit>
</View>
"@

    # RowLimit controls page size, not the total returned by PnP's paging loop.
    Get-PnPListItem -List $SourceList.Title -Query $Query -PageSize $PageSize `
        -Connection $SourceConnection |
        Select-Object -First $MaxItemsPerList
}
    #endregion Limited Source Item Selection

    #region Standard List Item Copying
function Copy-StandardListItems {
    param(
        [Parameter(Mandatory)]
        $SourceList,

        [Parameter(Mandatory)]
        $SourceConnection,

        [Parameter(Mandatory)]
        $TargetConnection
    )

    Write-CloneLog `
        -Level Information `
        -Operation "List items" `
        -Object $SourceList.Title `
        -Message "Starting list-item copy."

    try {
        $TargetList = Get-PnPList `
            -Identity $SourceList.Title `
            -Includes BaseTemplate,RootFolder `
            -Connection $TargetConnection `
            -ErrorAction Stop

        $Fields = Get-PnPField `
            -List $SourceList.Title `
            -Connection $SourceConnection

        $Items = @(Get-LimitedSourceItems -SourceList $SourceList -SourceConnection $SourceConnection)
        Write-CloneLog -Level Information -Operation "List items" -Object $SourceList.Title `
            -Message "Selected $($Items.Count) items in ascending source ID order; limit=$MaxItemsPerList (folders excluded)."

        $ItemCount = 0

        foreach ($Item in $Items) {
            try {
                $Values = Convert-ListItemValues `
                    -Item $Item `
                    -Fields $Fields

                $Folder = $null

                if (
                    $Item.FieldValues.ContainsKey("FileDirRef") -and
                    $Item.FieldValues["FileDirRef"]
                ) {
                    $SourceFolder = [string]$Item.FieldValues["FileDirRef"]
                    $ListRoot = $SourceList.RootFolder.ServerRelativeUrl

                    if (
                        $SourceFolder.StartsWith(
                            $ListRoot,
                            [System.StringComparison]::OrdinalIgnoreCase
                        )
                    ) {
                        $Folder = $SourceFolder.Substring(
                            $ListRoot.Length
                        ).Trim("/")
                    }
                }

                $AddParameters = @{
                    List       = $TargetList.Title
                    Values     = $Values
                    Connection = $TargetConnection
                }

                if (-not [string]::IsNullOrWhiteSpace($Folder)) {
                    $AddParameters["Folder"] = $Folder
                }

                Add-PnPListItem @AddParameters | Out-Null
                $ItemCount++

                Write-CloneLog `
                    -Level Information `
                    -Operation "List item" `
                    -Object "$($SourceList.Title)/$($Item.Id)" `
                    -Message "Copied."
            }
            catch {
                Write-CloneLog `
                    -Level Error `
                    -Operation "List item" `
                    -Object "$($SourceList.Title)/$($Item.Id)" `
                    -Message $_.Exception.Message `
                    -ErrorRecord $_
            }
        }

        Write-CloneLog `
            -Level Information `
            -Operation "List items" `
            -Object $SourceList.Title `
            -Message "Completed. Copied $ItemCount of $($Items.Count) selected items; limit=$MaxItemsPerList."
    }
    catch {
        Write-CloneLog `
            -Level Error `
            -Operation "List items" `
            -Object $SourceList.Title `
            -Message $_.Exception.Message `
            -ErrorRecord $_
    }
}

#endregion Standard List Item Copying

#region Document Library Content Copying
#------------------------------------------------------------
# Copy document-library files
#------------------------------------------------------------

function Copy-DocumentLibraryContent {
    param(
        [Parameter(Mandatory)]
        $SourceList,

        [Parameter(Mandatory)]
        [string]$SourceWebServerRelativeUrl,

        [Parameter(Mandatory)]
        [string]$TargetWebServerRelativeUrl,

        [Parameter(Mandatory)]
        $SourceConnection,

        [Parameter(Mandatory)]
        $TargetConnection
    )

    Write-CloneLog `
        -Level Information `
        -Operation "Library" `
        -Object $SourceList.Title `
        -Message "Starting content copy."

    try {
        $TargetList = Get-PnPList `
            -Identity $SourceList.Title `
            -Includes RootFolder `
            -Connection $TargetConnection `
            -ErrorAction Stop

        $SourceLibraryRoot = $SourceList.RootFolder.ServerRelativeUrl
        $TargetLibraryRoot = $TargetList.RootFolder.ServerRelativeUrl

        #region Library Folder Creation
        $FolderQuery = @"
<View Scope='RecursiveAll'>
  <Query>
    <Where><Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>1</Value></Eq></Where>
  </Query>
  <ViewFields><FieldRef Name='FileRef'/></ViewFields>
</View>
"@
        $FolderItems = Get-PnPListItem `
            -List $SourceList.Title `
            -PageSize 500 `
            -Query $FolderQuery `
            -Connection $SourceConnection

        # Create folders first
        $Folders = $FolderItems |
            Sort-Object {
                ([string]$_.FieldValues["FileRef"]).Split("/").Count
            }

        foreach ($FolderItem in $Folders) {
            try {
                $SourceFolderUrl =
                    [string]$FolderItem.FieldValues["FileRef"]

                $RelativeFolder = $SourceFolderUrl.Substring(
                    $SourceLibraryRoot.Length
                ).Trim("/")

                if ([string]::IsNullOrWhiteSpace($RelativeFolder)) {
                    continue
                }

                $TargetFolderPath = (
                    Get-SiteRelativePath `
                        -ServerRelativeUrl $TargetLibraryRoot `
                        -WebServerRelativeUrl $TargetWebServerRelativeUrl
                ).TrimEnd("/")

                $TargetFolderPath =
                    "$TargetFolderPath/$RelativeFolder"

                Ensure-TargetFolder `
                    -SiteRelativeFolder $TargetFolderPath `
                    -Connection $TargetConnection

                Write-CloneLog `
                    -Level Information `
                    -Operation "Folder" `
                    -Object $TargetFolderPath `
                    -Message "Created or verified."
            }
            catch {
                Write-CloneLog `
                    -Level Error `
                    -Operation "Folder" `
                    -Object $FolderItem.FieldValues["FileRef"] `
                    -Message $_.Exception.Message `
                    -ErrorRecord $_
            }
        }

        #endregion Library Folder Creation

        #region Library File Copying
        # Copy files
        $Files = @(Get-LimitedSourceItems -SourceList $SourceList -SourceConnection $SourceConnection)
        Write-CloneLog -Level Information -Operation "Library" -Object $SourceList.Title `
            -Message "Selected $($Files.Count) files in ascending source ID order; limit=$MaxItemsPerList (folders excluded)."

        $FileCount = 0

        foreach ($FileItem in $Files) {
            $MemoryStream = $null

            try {
                $SourceFileUrl =
                    [string]$FileItem.FieldValues["FileRef"]

                $FileName =
                    [string]$FileItem.FieldValues["FileLeafRef"]

                $SourceParentFolder =
                    [string]$FileItem.FieldValues["FileDirRef"]

                $RelativeFolder = $SourceParentFolder.Substring(
                    $SourceLibraryRoot.Length
                ).Trim("/")

                $TargetLibrarySiteRelative = Get-SiteRelativePath `
                    -ServerRelativeUrl $TargetLibraryRoot `
                    -WebServerRelativeUrl $TargetWebServerRelativeUrl

                $TargetFolder = $TargetLibrarySiteRelative.TrimEnd("/")

                if (-not [string]::IsNullOrWhiteSpace($RelativeFolder)) {
                    $TargetFolder = "$TargetFolder/$RelativeFolder"
                }

                Ensure-TargetFolder `
                    -SiteRelativeFolder $TargetFolder `
                    -Connection $TargetConnection

                $TargetServerRelativeFileUrl = (
                    "$TargetWebServerRelativeUrl/$TargetFolder/$FileName"
                ) -replace "//+", "/"

                if (-not $OverwriteFiles) {
                    $ExistingFile = Get-PnPFile `
                        -Url $TargetServerRelativeFileUrl `
                        -AsFileObject `
                        -Connection $TargetConnection `
                        -ErrorAction SilentlyContinue

                    if ($null -ne $ExistingFile) {
                        Write-CloneLog `
                            -Level Warning `
                            -Operation "File" `
                            -Object $SourceFileUrl `
                            -Message "Skipped because the destination file exists."
                        continue
                    }
                }

                $MemoryStream = Get-PnPFile `
                    -Url $SourceFileUrl `
                    -AsMemoryStream `
                    -Connection $SourceConnection

                Add-PnPFile `
                    -Folder $TargetFolder `
                    -FileName $FileName `
                    -Stream $MemoryStream `
                    -Connection $TargetConnection | Out-Null

                $FileCount++

                Write-CloneLog `
                    -Level Information `
                    -Operation "File" `
                    -Object $SourceFileUrl `
                    -Message "Copied to $TargetFolder/$FileName"
            }
            catch {
                Write-CloneLog `
                    -Level Error `
                    -Operation "File" `
                    -Object $FileItem.FieldValues["FileRef"] `
                    -Message $_.Exception.Message `
                    -ErrorRecord $_
            }
            finally {
                if ($null -ne $MemoryStream) {
                    $MemoryStream.Dispose()
                }
            }
        }

        #endregion Library File Copying

        Write-CloneLog `
            -Level Information `
            -Operation "Library" `
            -Object $SourceList.Title `
            -Message "Completed. Copied $FileCount of $($Files.Count) selected files; limit=$MaxItemsPerList."
    }
    catch {
        Write-CloneLog `
            -Level Error `
            -Operation "Library" `
            -Object $SourceList.Title `
            -Message $_.Exception.Message `
            -ErrorRecord $_
    }
}

#endregion Document Library Content Copying
#endregion Helper Functions

#region Main Execution
#------------------------------------------------------------
# Main execution
#------------------------------------------------------------

try {
    #region Diagnostics and Authentication Setup
    Start-Transcript -Path $TranscriptPath -ErrorAction Stop | Out-Null
    $TranscriptStarted = $true
    Set-PnPTraceLog -On -LogFile $TracePath -Level Debug -AutoFlush $true
    $TraceStarted = $true

    Write-CloneLog `
        -Level Information `
        -Operation "Initialize" `
        -Message "Starting SharePoint site clone."

    $PnPModule = Get-Module PnP.PowerShell
    Write-CloneLog -Level Information -Operation "Configuration" `
        -Message ("PowerShell=$($PSVersionTable.PSVersion); PnP.PowerShell=$($PnPModule.Version); " +
            "Source=$SourceSiteUrl; Target=$TargetSiteUrl; CreateTargetSite=$CreateTargetSite; " +
            "CopyDocumentLibraries=$CopyDocumentLibraries; CopyListItems=$CopyListItems; " +
            "MaxItemsPerList=$MaxItemsPerList; " +
            "SkipSiteHeader=$SkipSiteHeader; " +
            "OverwriteFiles=$OverwriteFiles; IncludeHiddenLists=$IncludeHiddenLists")

    $AuthenticationParameters = @{
        ClientId         = $ClientId
        Tenant           = $Tenant
        Thumbprint       = $Thumbprint
        ReturnConnection = $true
    }

    #endregion Diagnostics and Authentication Setup

    #region Source Site Connection
    # Connect to source
    $SourceConnection = Connect-PnPOnline `
        -Url $SourceSiteUrl `
        @AuthenticationParameters

    $SourceWeb = Get-PnPWeb `
        -Includes Title,Url,ServerRelativeUrl,WebTemplate `
        -Connection $SourceConnection

    Write-CloneLog `
        -Level Information `
        -Operation "Connect" `
        -Object $SourceSiteUrl `
        -Message "Connected to source site '$($SourceWeb.Title)'."

    #endregion Source Site Connection

    #region Optional Destination Site Creation
    # Create destination site if requested
    if ($CreateTargetSite) {
        $AdminConnection = Connect-PnPOnline `
            -Url $TenantAdminUrl `
            @AuthenticationParameters

        try {
            $ExistingTarget = Get-PnPTenantSite `
                -Url $TargetSiteUrl `
                -Connection $AdminConnection `
                -ErrorAction SilentlyContinue
        }
        catch {
            $ExistingTarget = $null
        }

        if ($null -eq $ExistingTarget) {
            Write-CloneLog `
                -Level Information `
                -Operation "Create site" `
                -Object $TargetSiteUrl `
                -Message "Creating destination site collection."

            if ($TargetSiteType -eq "CommunicationSite") {
                New-PnPSite `
                    -Type CommunicationSite `
                    -Title $TargetSiteTitle `
                    -Url $TargetSiteUrl `
                    -Owner $TargetSiteOwner `
                    -Wait `
                    -Connection $AdminConnection | Out-Null
            }
            else {
                New-PnPSite `
                    -Type TeamSiteWithoutMicrosoft365Group `
                    -Title $TargetSiteTitle `
                    -Url $TargetSiteUrl `
                    -Owner $TargetSiteOwner `
                    -Wait `
                    -Connection $AdminConnection | Out-Null
            }

            Write-CloneLog `
                -Level Information `
                -Operation "Create site" `
                -Object $TargetSiteUrl `
                -Message "Destination site collection created."
        }
        else {
            Write-CloneLog `
                -Level Warning `
                -Operation "Create site" `
                -Object $TargetSiteUrl `
                -Message "Destination already exists; creation skipped."
        }
    }

    #endregion Optional Destination Site Creation

    #region Destination Site Connection
    # Connect to destination
    $TargetConnection = Connect-PnPOnline `
        -Url $TargetSiteUrl `
        @AuthenticationParameters

    $TargetWeb = Get-PnPWeb `
        -Includes Title,Url,ServerRelativeUrl `
        -Connection $TargetConnection

    Write-CloneLog `
        -Level Information `
        -Operation "Connect" `
        -Object $TargetSiteUrl `
        -Message "Connected to destination site '$($TargetWeb.Title)'."

    #endregion Destination Site Connection

    #region Template Export and Repair
    # Export provisioning package
    Write-CloneLog `
        -Level Information `
        -Operation "Template" `
        -Object $TemplatePath `
        -Message "Extracting source structure."

    $TemplateParameters = @{
        Out                    = $TemplatePath
        IncludeAllPages        = $true
        IncludeSiteGroups      = $true
        PersistBrandingFiles   = $true
        PersistPublishingFiles = $true
        Force                  = $true
        Connection             = $SourceConnection
    }

    if ($IncludeHiddenLists) {
        $TemplateParameters["IncludeHiddenLists"] = $true
    }

    Get-PnPSiteTemplate @TemplateParameters

    Write-CloneLog `
        -Level Information `
        -Operation "Template" `
        -Object $TemplatePath `
        -Message "Source structure extracted."

    $DestinationContentTypes = @(Get-PnPContentType -Connection $TargetConnection -ErrorAction Stop)
    $ProvisioningTemplatePath = Repair-PnPPageWebPartJson -Path $TemplatePath `
        -DestinationContentTypes $DestinationContentTypes

    #endregion Template Export and Repair

    #region Template Application
    # Apply provisioning package
    Write-CloneLog `
        -Level Information `
        -Operation "Template" `
        -Object $TargetSiteUrl `
        -Message "Applying source structure to destination."

    $ApplyParameters = @{
        Path                         = $ProvisioningTemplatePath
        Connection                   = $TargetConnection
        ClearNavigation              = $true
        IgnoreDuplicateDataRowErrors  = $true
    }
    if ($SkipSiteHeader) {
        $ApplyParameters["ExcludeHandlers"] = [PnP.Framework.Provisioning.Model.Handlers]::SiteHeader
        Write-CloneLog -Level Warning -Operation "Template" -Object $TargetSiteUrl `
            -Message "SiteHeader provisioning explicitly skipped; destination site-header layout retained."
    }

    $TemplateProvisioned = $false
    try {
        Invoke-PnPSiteTemplate @ApplyParameters -ErrorAction Stop
        $TemplateProvisioned = $true
    }
    catch {
        Write-CloneLog -Level Error -Operation "Template" -Object $TargetSiteUrl `
            -Message "Destination provisioning failed. Content copying will not start. $($_.Exception.Message)" `
            -ErrorRecord $_
        throw
    }

    if ($TemplateProvisioned) {
        Write-CloneLog `
            -Level Information `
            -Operation "Template" `
            -Object $TargetSiteUrl `
            -Message "Destination structure provisioned."
    }

    #endregion Template Application

    $SourceLists = @(Get-PnPList `
        -Includes Title,Hidden,BaseTemplate,BaseType,RootFolder,ItemCount `
        -Connection $SourceConnection -ErrorAction Stop)
    $SourceLists = @(Get-CloneSourceLists -Lists $SourceLists)
    $TargetLists = @(Get-PnPList -Includes Title,BaseTemplate,RootFolder `
        -Connection $TargetConnection -ErrorAction Stop)
    Assert-CloneDestinationLists -SourceLists $SourceLists -TargetLists $TargetLists `
        -SourceWebServerRelativeUrl $SourceWeb.ServerRelativeUrl `
        -TargetWebServerRelativeUrl $TargetWeb.ServerRelativeUrl

    #region Modern Page Relocation
    # Relocate any modern pages defined in subfolders that were flattened during template preparation
    if ($null -ne $script:SubfolderPagesToRelocate -and $script:SubfolderPagesToRelocate.Count -gt 0) {
        Write-CloneLog `
            -Level Information `
            -Operation "Page relocation" `
            -Object $TargetSiteUrl `
            -Message "Relocating $($script:SubfolderPagesToRelocate.Count) modern page(s) to destination subfolder(s)."

        $TargetServerRelUrl = $TargetWeb.ServerRelativeUrl.TrimEnd("/")

        foreach ($PageReloc in $script:SubfolderPagesToRelocate) {
            try {
                $SubfolderSiteRel = "SitePages/$($PageReloc.Subfolder)"
                Resolve-PnPFolder -SiteRelativePath $SubfolderSiteRel -Connection $TargetConnection | Out-Null

                $SourceUrl = "$TargetServerRelUrl/SitePages/$($PageReloc.FlatPageName)"
                $DestUrl = "$TargetServerRelUrl/SitePages/$($PageReloc.OriginalRelativePath)"

                Move-PnPFile -SourceUrl $SourceUrl -TargetUrl $DestUrl -Overwrite -Force -Connection $TargetConnection

                Write-CloneLog `
                    -Level Information `
                    -Operation "Page relocation" `
                    -Object $PageReloc.OriginalRelativePath `
                    -Message "Relocated '$($PageReloc.FlatPageName)' -> '$($PageReloc.OriginalRelativePath)' successfully."
            }
            catch {
                Write-CloneLog `
                    -Level Warning `
                    -Operation "Page relocation" `
                    -Object $PageReloc.OriginalRelativePath `
                    -Message "Failed to relocate page '$($PageReloc.FlatPageName)': $($_.Exception.Message)" `
                    -ErrorRecord $_
            }
        }
    }

    #endregion Modern Page Relocation

    #region Source List and Library Content Copying
    foreach ($SourceList in $SourceLists) {
        # Document libraries use BaseType 1
        if ([int]$SourceList.BaseType -eq 1) {
            Copy-DocumentLibraryContent `
                -SourceList $SourceList `
                -SourceWebServerRelativeUrl $SourceWeb.ServerRelativeUrl `
                -TargetWebServerRelativeUrl $TargetWeb.ServerRelativeUrl `
                -SourceConnection $SourceConnection `
                -TargetConnection $TargetConnection

            continue
        }

        Copy-StandardListItems `
            -SourceList $SourceList `
            -SourceConnection $SourceConnection `
            -TargetConnection $TargetConnection
    }

    #endregion Source List and Library Content Copying

    #region Completion Status
    $CopyErrors = @($script:LogEntries | Where-Object Level -eq "Error").Count
    if ($CopyErrors -gt 0) {
        Write-CloneLog -Level Warning -Operation "Complete" -Object $TargetSiteUrl `
            -Message "SharePoint site clone finished with $CopyErrors error events. Review the logs for items that were not copied."
    }
    else {
        Write-CloneLog -Level Information -Operation "Complete" -Object $TargetSiteUrl `
            -Message "SharePoint site clone completed."
    }
    #endregion Completion Status
}
catch {
    #region Fatal Error Diagnostics
    $FatalError = $_
    try {
        $PnPException = Get-PnPException
        if ($null -ne $PnPException) {
            Write-CloneLog -Level Error -Operation "PnP exception" `
                -Message ($PnPException | Format-List * -Force | Out-String -Width 240)
        }
    }
    catch {
        Write-CloneLog -Level Warning -Operation "Diagnostics" `
            -Message "Could not retrieve PnP exception: $($_.Exception.Message)" -ErrorRecord $_
    }

    Write-CloneLog `
        -Level Error `
        -Operation "Fatal" `
        -Message $FatalError.Exception.Message `
        -ErrorRecord $FatalError

    throw
    #endregion Fatal Error Diagnostics
}
finally {
    #region Run Summary and Cleanup
    try {
        $Errors = @($script:LogEntries | Where-Object Level -eq "Error").Count
        $Warnings = @($script:LogEntries | Where-Object Level -eq "Warning").Count
        Write-CloneLog -Level Information -Operation "Run summary" `
            -Message "Elapsed=$($RunTimer.Elapsed); Error events=$Errors; Warning events=$Warnings."
        Save-CloneLog
    }
    finally {
        try {
            if ($TraceStarted) {
                Set-PnPTraceLog -Off
            }
        }
        finally {
            Write-Host ""
            Write-Host "Clone log: $LogPath" -ForegroundColor Cyan
            Write-Host "PnP trace: $TracePath" -ForegroundColor Cyan
            Write-Host "Transcript: $TranscriptPath" -ForegroundColor Cyan
            Write-Host "Original template: $TemplatePath" -ForegroundColor Cyan
            Write-Host "Applied template:  $ProvisioningTemplatePath" -ForegroundColor Cyan
            if ($TranscriptStarted) {
                Stop-Transcript | Out-Null
            }
            $RunTimer.Stop()
        }
    }
    #endregion Run Summary and Cleanup
}
#endregion Main Execution
