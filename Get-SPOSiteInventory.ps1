#requires -Version 7.2
#requires -Modules PnP.PowerShell

<#
.SYNOPSIS
Creates a configuration inventory for one SharePoint Online site collection.

.DESCRIPTION
Uses PnP PowerShell with Entra ID application certificate authentication to
inventory the root web and all subwebs. The script does not enumerate list
items or file contents. It produces separate CSV files for site/web settings,
lists and libraries, fields, views, content types, pages, page components,
features, event receivers, navigation, and optionally role assignments.

.PARAMETER SiteUrl
URL of the SharePoint Online site collection to inventory.

.PARAMETER ClientId
Application (client) ID of the Entra ID app registration.

.PARAMETER Tenant
Tenant name (contoso.onmicrosoft.com) or tenant ID.

.PARAMETER Thumbprint
Thumbprint of a certificate installed in the current user's certificate store.

.PARAMETER CertificatePath
Path to a PFX certificate. Use CertificatePassword when the PFX is protected.

.PARAMETER OutputPath
Parent folder for the timestamped report folder.

.PARAMETER IncludeHiddenLists
Includes hidden lists and libraries. Hidden lists are excluded by default.

.PARAMETER IncludePermissions
Inventories web- and list-level role assignments. Item-level permissions are
not enumerated.

.EXAMPLE
.\Get-SPOSiteInventory.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/HR" `
    -ClientId "00000000-0000-0000-0000-000000000000" `
    -Tenant "contoso.onmicrosoft.com" -Thumbprint "CERTIFICATE_THUMBPRINT"

.NOTES
The app normally needs SharePoint application permission Sites.FullControl.All,
or an equivalent Sites.Selected grant on the target site. Tenant-wide site
enumeration is not required because this script accepts one source SiteUrl.
#>

#region Parameters
[CmdletBinding(DefaultParameterSetName = 'CertificateStore')]
param(
    [Parameter()]
    [ValidatePattern('^https://')]
    [uri]$SiteUrl = 'https://CONTOSO.sharepoint.com/sites/PCAHome',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$ClientId = 'abc12345-283f-47ba-a185-50d935d51d57',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$Tenant = '9cfc42cb-51da-9654-87e9-b20a170b6ba3',

    [Parameter(ParameterSetName = 'CertificateStore')]
    [ValidateNotNullOrEmpty()]
    [string]$Thumbprint = 'C696FDCFE1453F3FBC6031F54DE988DA0FD905A9',

    [Parameter(ParameterSetName = 'CertificateFile')]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]$CertificatePath,

    [Parameter(ParameterSetName = 'CertificateFile')]
    [securestring]$CertificatePassword,

    [string]$OutputPath = (Join-Path -Path $PWD -ChildPath 'SPOInventory'),

    [switch]$IncludeHiddenLists = $true,

    [switch]$IncludePermissions = $true,

    [ValidateRange(0, 10)]
    [int]$MaxRetries = 5,

    [ValidateRange(1, 300)]
    [int]$BaseRetryDelaySeconds = 2,

    [ValidateRange(0, 60000)]
    [int]$DelayBetweenWebsMilliseconds = 250
)
#endregion

#region Initialization
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:AuthenticationParameterSet = $PSCmdlet.ParameterSetName
$script:Errors = [System.Collections.Generic.List[object]]::new()
$script:ThrottleCount = 0
#endregion

#region CSV Schema Definitions
$schemas = [ordered]@{
    Sites = @('SiteUrl', 'SiteId', 'CompatibilityLevel', 'ReadOnly', 'WriteLocked', 'InventoryStartedUtc')
    Webs = @('SiteUrl', 'WebUrl', 'WebId', 'Title', 'Description', 'ServerRelativeUrl', 'WebTemplate', 'Configuration', 'Language', 'Created', 'LastItemModifiedDate', 'HasUniqueRoleAssignments', 'MasterUrl', 'CustomMasterUrl')
    Lists = @('SiteUrl', 'WebUrl', 'ListId', 'Title', 'Description', 'BaseType', 'BaseTemplate', 'ItemCount', 'Hidden', 'ContentTypesEnabled', 'EnableVersioning', 'EnableMinorVersions', 'MajorVersionLimit', 'ForceCheckout', 'NoCrawl', 'OnQuickLaunch', 'Created', 'LastItemModifiedDate', 'DefaultViewUrl', 'RootFolderUrl', 'HasUniqueRoleAssignments')
    Fields = @('SiteUrl', 'WebUrl', 'Scope', 'ListTitle', 'FieldId', 'Title', 'InternalName', 'TypeAsString', 'Required', 'Hidden', 'ReadOnlyField', 'Indexed', 'EnforceUniqueValues', 'Group')
    Views = @('SiteUrl', 'WebUrl', 'ListTitle', 'ViewId', 'Title', 'ViewType', 'DefaultView', 'PersonalView', 'Paged', 'RowLimit', 'ViewUrl')
    ContentTypes = @('SiteUrl', 'WebUrl', 'Scope', 'ListTitle', 'ContentTypeId', 'Name', 'Group', 'Description', 'Hidden', 'ReadOnly', 'Sealed')
    Pages = @('SiteUrl', 'WebUrl', 'LibraryTitle', 'PageName', 'PageUrl', 'ContentTypeId', 'PromotedState', 'Created', 'Modified', 'Author', 'Editor')
    WebParts = @('SiteUrl', 'WebUrl', 'PageUrl', 'PageName', 'Title', 'WebPartId', 'InstanceId', 'ComponentType', 'Section', 'Column')
    Features = @('SiteUrl', 'WebUrl', 'Scope', 'DefinitionId', 'DisplayName')
    EventReceivers = @('SiteUrl', 'WebUrl', 'Scope', 'ListTitle', 'ReceiverName', 'ReceiverClass', 'ReceiverAssembly', 'ReceiverUrl', 'EventType', 'Synchronization')
    Navigation = @('SiteUrl', 'WebUrl', 'Location', 'Title', 'Url', 'IsExternal', 'Id', 'ParentId')
    Permissions = @('SiteUrl', 'WebUrl', 'Scope', 'ObjectTitle', 'ObjectUrl', 'PrincipalTitle', 'PrincipalType', 'LoginName', 'Roles')
    Errors = @('TimestampUtc', 'SiteUrl', 'WebUrl', 'Operation', 'Attempt', 'Message')
}

$results = @{}
foreach ($name in $schemas.Keys) {
    if ($name -ne 'Errors') {
        $results[$name] = [System.Collections.Generic.List[object]]::new()
    }
}
#endregion

#region Helper Functions
function Get-Value {
    param(
        [AllowNull()][object]$InputObject,
        [Parameter(Mandatory)][string]$PropertyName,
        [AllowNull()][object]$DefaultValue = $null
    )

    if ($null -eq $InputObject) {
        return $DefaultValue
    }

    $property = $InputObject.PSObject.Properties[$PropertyName]
    if ($null -eq $property -or $null -eq $property.Value) {
        return $DefaultValue
    }

    return $property.Value
}

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO'
    )

    $line = '{0} [{1}] {2}' -f ([datetime]::UtcNow.ToString('yyyy-MM-dd HH:mm:ss.fff')), $Level, $Message

    if ($script:LogPath) {
        Add-Content -LiteralPath $script:LogPath -Value $line -Encoding utf8
    }

    switch ($Level) {
        'WARN' { Write-Warning $Message }
        'ERROR' { Write-Host $line -ForegroundColor Red }
        default { Write-Host $line }
    }
}

function Add-InventoryError {
    param(
        [Parameter(Mandatory)][string]$Operation,
        [string]$WebUrl,
        [Parameter(Mandatory)][int]$Attempt,
        [Parameter(Mandatory)][string]$Message
    )

    $script:Errors.Add([pscustomobject][ordered]@{
        TimestampUtc = [datetime]::UtcNow.ToString('o')
        SiteUrl = $SiteUrl.AbsoluteUri.TrimEnd('/')
        WebUrl = $WebUrl
        Operation = $Operation
        Attempt = $Attempt
        Message = $Message
    })

    Write-Log -Level ERROR -Message "$Operation (attempt $Attempt) [$WebUrl]: $Message"
}

function Get-HttpStatusCode {
    param([Parameter(Mandatory)][System.Management.Automation.ErrorRecord]$ErrorRecord)

    $response = Get-Value -InputObject $ErrorRecord.Exception -PropertyName 'Response'
    $statusCode = Get-Value -InputObject $response -PropertyName 'StatusCode'
    if ($null -ne $statusCode) {
        return [int]$statusCode
    }

    if ($ErrorRecord.Exception.Message -match '(?<!\d)(429|503)(?!\d)') {
        return [int]$Matches[1]
    }

    return $null
}

function Invoke-InventoryOperation {
    param(
        [Parameter(Mandatory)][string]$Operation,
        [Parameter(Mandatory)][scriptblock]$ScriptBlock,
        [string]$WebUrl,
        [switch]$ThrowOnFailure
    )

    for ($attempt = 0; $attempt -le $MaxRetries; $attempt++) {
        try {
            return & $ScriptBlock
        }
        catch {
            $statusCode = Get-HttpStatusCode -ErrorRecord $_
            $isRetryable = $statusCode -in 429, 503

            if ($isRetryable -and $attempt -lt $MaxRetries) {
                $script:ThrottleCount++
                $response = Get-Value -InputObject $_.Exception -PropertyName 'Response'
                $headers = Get-Value -InputObject $response -PropertyName 'Headers'
                $retryAfter = if ($null -ne $headers) { $headers['Retry-After'] } else { $null }
                $delay = if ($retryAfter -as [int]) {
                    [int]$retryAfter
                }
                else {
                    [math]::Min(300, [math]::Pow(2, $attempt + 1) * $BaseRetryDelaySeconds)
                }
                Write-Log -Level WARN -Message "$Operation was throttled (HTTP $statusCode). Retrying in $delay second(s)."
                Start-Sleep -Seconds $delay
                continue
            }

            Add-InventoryError -Operation $Operation -WebUrl $WebUrl -Attempt ($attempt + 1) -Message $_.Exception.Message
            if ($ThrowOnFailure) {
                throw
            }
            Write-Log -Level WARN -Message "$Operation failed: $($_.Exception.Message)"
            return $null
        }
    }
}

function New-SiteConnection {
    param([Parameter(Mandatory)][string]$Url)

    $parameters = @{
        Url = $Url
        ClientId = $ClientId
        Tenant = $Tenant
        ReturnConnection = $true
        WarningAction = 'SilentlyContinue'
        ErrorAction = 'Stop'
    }

    if ($script:AuthenticationParameterSet -eq 'CertificateFile') {
        $parameters.CertificatePath = $CertificatePath
        if ($null -ne $CertificatePassword) {
            $parameters.CertificatePassword = $CertificatePassword
        }
    }
    else {
        $parameters.Thumbprint = $Thumbprint
    }

    return Connect-PnPOnline @parameters
}

function Add-RoleAssignments {
    param(
        [Parameter(Mandatory)][object]$SecurableObject,
        [Parameter(Mandatory)][string]$Scope,
        [Parameter(Mandatory)][string]$ObjectTitle,
        [string]$ObjectUrl,
        [Parameter(Mandatory)][string]$CurrentWebUrl,
        [Parameter(Mandatory)][object]$Connection
    )

    $operation = "Read $Scope permissions for $ObjectTitle"

    # Build rows locally and add them once after success; adding directly inside a retried
    # scriptblock would re-insert already-recorded rows into $results.Permissions on every retry.
    $rows = Invoke-InventoryOperation -Operation $operation -WebUrl $CurrentWebUrl -ScriptBlock {
        Get-PnPProperty -ClientObject $SecurableObject -Property RoleAssignments -Connection $Connection | Out-Null
        $rowsLocal = [System.Collections.Generic.List[object]]::new()
        foreach ($assignment in $SecurableObject.RoleAssignments) {
            Get-PnPProperty -ClientObject $assignment -Property Member, RoleDefinitionBindings -Connection $Connection | Out-Null
            $roles = @(
                $assignment.RoleDefinitionBindings |
                    Where-Object { $_.Name -ne 'Limited Access' } |
                    ForEach-Object { $_.Name }
            )
            if ($roles.Count -eq 0) {
                continue
            }

            $rowsLocal.Add([pscustomobject][ordered]@{
                SiteUrl = $SiteUrl.AbsoluteUri.TrimEnd('/')
                WebUrl = $CurrentWebUrl
                Scope = $Scope
                ObjectTitle = $ObjectTitle
                ObjectUrl = $ObjectUrl
                PrincipalTitle = Get-Value $assignment.Member 'Title'
                PrincipalType = Get-Value $assignment.Member 'PrincipalType'
                LoginName = Get-Value $assignment.Member 'LoginName'
                Roles = $roles -join '; '
            })
        }
        return , $rowsLocal
    }

    foreach ($row in @($rows)) {
        if ($null -ne $row) {
            $results.Permissions.Add($row)
        }
    }
}

function Add-NavigationNodes {
    param(
        [Parameter(Mandatory)][object[]]$Nodes,
        [Parameter(Mandatory)][string]$Location,
        [Parameter(Mandatory)][string]$CurrentWebUrl,
        [string]$ParentId = ''
    )

    foreach ($node in $Nodes) {
        $nodeId = [string](Get-Value $node 'Id')
        $results.Navigation.Add([pscustomobject][ordered]@{
            SiteUrl = $SiteUrl.AbsoluteUri.TrimEnd('/')
            WebUrl = $CurrentWebUrl
            Location = $Location
            Title = Get-Value $node 'Title'
            Url = Get-Value $node 'Url'
            IsExternal = Get-Value $node 'IsExternal'
            Id = $nodeId
            ParentId = $ParentId
        })

        $children = @(Get-Value $node 'Children' @())
        if ($children.Count -gt 0) {
            Add-NavigationNodes -Nodes $children -Location $Location -CurrentWebUrl $CurrentWebUrl -ParentId $nodeId
        }
    }
}

function Export-InventoryData {
    param(
        [Parameter(Mandatory)][string]$Name,
        [AllowNull()][object[]]$Rows,
        [Parameter(Mandatory)][string[]]$Columns,
        [Parameter(Mandatory)][string]$Folder
    )

    # Rows can arrive as $null (e.g. an empty typed collection collapsed by the caller); normalize before use.
    $safeRows = @($Rows)

    $path = Join-Path -Path $Folder -ChildPath "$Name.csv"
    if ($safeRows.Count -gt 0) {
        $safeRows | Select-Object -Property $Columns | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding utf8
    }
    else {
        $header = '"' + ($Columns -join '","') + '"'
        Set-Content -LiteralPath $path -Value $header -Encoding utf8
    }
}
#endregion

#region Report Folder and Logging Setup
$inventoryStarted = [datetime]::UtcNow
$siteUrlText = $SiteUrl.AbsoluteUri.TrimEnd('/')
$safeSiteName = ($SiteUrl.Host + $SiteUrl.AbsolutePath).TrimEnd('/') -replace '[^a-zA-Z0-9.-]', '_'
$reportFolder = Join-Path -Path $OutputPath -ChildPath ("{0}_{1}" -f $safeSiteName, (Get-Date -Format 'yyyyMMdd_HHmmss'))
New-Item -ItemType Directory -Path $reportFolder -Force | Out-Null
$script:LogPath = Join-Path -Path $reportFolder -ChildPath 'Inventory.log'
New-Item -ItemType File -Path $script:LogPath -Force | Out-Null
#endregion

#region Main Inventory Execution
$script:FatalError = $null

try {
#region Connect and Inventory Site Collection
Write-Log -Message "Inventory started for $siteUrlText"
Write-Log -Message "Log file: $script:LogPath"
Write-Log -Message "Connecting to $siteUrlText"
$rootConnection = Invoke-InventoryOperation -Operation 'Connect to root web' -WebUrl $siteUrlText -ThrowOnFailure -ScriptBlock {
    New-SiteConnection -Url $siteUrlText
}

$site = Invoke-InventoryOperation -Operation 'Get site collection properties' -WebUrl $siteUrlText -ThrowOnFailure -ScriptBlock {
    Get-PnPSite -Includes Id, Url, CompatibilityLevel, ReadOnly, WriteLocked -Connection $rootConnection
}
$results.Sites.Add([pscustomobject][ordered]@{
    SiteUrl = $siteUrlText
    SiteId = Get-Value $site 'Id'
    CompatibilityLevel = Get-Value $site 'CompatibilityLevel'
    ReadOnly = Get-Value $site 'ReadOnly'
    WriteLocked = Get-Value $site 'WriteLocked'
    InventoryStartedUtc = $inventoryStarted.ToString('o')
})

$rootWeb = Invoke-InventoryOperation -Operation 'Get root web' -WebUrl $siteUrlText -ThrowOnFailure -ScriptBlock {
    Get-PnPWeb -Includes Id, Title, Url, ServerRelativeUrl -Connection $rootConnection
}
$subWebs = @(
    Invoke-InventoryOperation -Operation 'Get subwebs' -WebUrl $siteUrlText -ScriptBlock {
        Get-PnPSubWeb -Recurse -Includes Id, Title, Url, ServerRelativeUrl -Connection $rootConnection
    }
)
$webUrls = @($rootWeb.Url) + @($subWebs | ForEach-Object { $_.Url })
Write-Log -Message "Discovered $($webUrls.Count) web(s) to inventory"

$siteFeatures = @(
    Invoke-InventoryOperation -Operation 'Get site-scoped features' -WebUrl $siteUrlText -ScriptBlock {
        Get-PnPFeature -Scope Site -Connection $rootConnection
    }
)
foreach ($feature in $siteFeatures) {
    $results.Features.Add([pscustomobject][ordered]@{
        SiteUrl = $siteUrlText
        WebUrl = $siteUrlText
        Scope = 'Site'
        DefinitionId = Get-Value $feature 'DefinitionId'
        DisplayName = Get-Value $feature 'DisplayName'
    })
}

$siteReceivers = @(
    Invoke-InventoryOperation -Operation 'Get site event receivers' -WebUrl $siteUrlText -ScriptBlock {
        Get-PnPEventReceiver -Scope Site -Connection $rootConnection
    }
)
foreach ($receiver in $siteReceivers) {
    $results.EventReceivers.Add([pscustomobject][ordered]@{
        SiteUrl = $siteUrlText
        WebUrl = $siteUrlText
        Scope = 'Site'
        ListTitle = ''
        ReceiverName = Get-Value $receiver 'ReceiverName'
        ReceiverClass = Get-Value $receiver 'ReceiverClass'
        ReceiverAssembly = Get-Value $receiver 'ReceiverAssembly'
        ReceiverUrl = Get-Value $receiver 'ReceiverUrl'
        EventType = Get-Value $receiver 'EventType'
        Synchronization = Get-Value $receiver 'Synchronization'
    })
}
#endregion

#region Per-Web Inventory Loop
$webIndex = 0
foreach ($webUrl in $webUrls) {
    $webIndex++
    Write-Progress -Activity 'Creating SharePoint Online site inventory' -Status "$webIndex of $($webUrls.Count): $webUrl" -PercentComplete (($webIndex / $webUrls.Count) * 100)
    Write-Log -Message "Processing web $webIndex of $($webUrls.Count): $webUrl"

    try {
    $connection = Invoke-InventoryOperation -Operation 'Connect to web' -WebUrl $webUrl -ScriptBlock {
        New-SiteConnection -Url $webUrl
    }
    if ($null -eq $connection) {
        continue
    }

    $web = Invoke-InventoryOperation -Operation 'Get web properties' -WebUrl $webUrl -ScriptBlock {
        Get-PnPWeb -Includes Id, Title, Description, Url, ServerRelativeUrl, WebTemplate, Configuration, Language, Created, LastItemModifiedDate, HasUniqueRoleAssignments, MasterUrl, CustomMasterUrl -Connection $connection
    }
    if ($null -eq $web) {
        continue
    }

    $results.Webs.Add([pscustomobject][ordered]@{
        SiteUrl = $siteUrlText
        WebUrl = $webUrl
        WebId = Get-Value $web 'Id'
        Title = Get-Value $web 'Title'
        Description = Get-Value $web 'Description'
        ServerRelativeUrl = Get-Value $web 'ServerRelativeUrl'
        WebTemplate = Get-Value $web 'WebTemplate'
        Configuration = Get-Value $web 'Configuration'
        Language = Get-Value $web 'Language'
        Created = Get-Value $web 'Created'
        LastItemModifiedDate = Get-Value $web 'LastItemModifiedDate'
        HasUniqueRoleAssignments = Get-Value $web 'HasUniqueRoleAssignments'
        MasterUrl = Get-Value $web 'MasterUrl'
        CustomMasterUrl = Get-Value $web 'CustomMasterUrl'
    })

    if ($IncludePermissions) {
        Add-RoleAssignments -SecurableObject $web -Scope 'Web' -ObjectTitle $web.Title -ObjectUrl $webUrl -CurrentWebUrl $webUrl -Connection $connection
    }

    $webFields = @(
        Invoke-InventoryOperation -Operation 'Get web fields' -WebUrl $webUrl -ScriptBlock {
            Get-PnPField -Connection $connection
        }
    )
    foreach ($field in $webFields) {
        $results.Fields.Add([pscustomobject][ordered]@{
            SiteUrl = $siteUrlText
            WebUrl = $webUrl
            Scope = 'Web'
            ListTitle = ''
            FieldId = Get-Value $field 'Id'
            Title = Get-Value $field 'Title'
            InternalName = Get-Value $field 'InternalName'
            TypeAsString = Get-Value $field 'TypeAsString'
            Required = Get-Value $field 'Required'
            Hidden = Get-Value $field 'Hidden'
            ReadOnlyField = Get-Value $field 'ReadOnlyField'
            Indexed = Get-Value $field 'Indexed'
            EnforceUniqueValues = Get-Value $field 'EnforceUniqueValues'
            Group = Get-Value $field 'Group'
        })
    }

    $webContentTypes = @(
        Invoke-InventoryOperation -Operation 'Get web content types' -WebUrl $webUrl -ScriptBlock {
            Get-PnPContentType -Connection $connection
        }
    )
    foreach ($contentType in $webContentTypes) {
        $results.ContentTypes.Add([pscustomobject][ordered]@{
            SiteUrl = $siteUrlText
            WebUrl = $webUrl
            Scope = 'Web'
            ListTitle = ''
            ContentTypeId = Get-Value $contentType 'StringId' (Get-Value $contentType 'Id')
            Name = Get-Value $contentType 'Name'
            Group = Get-Value $contentType 'Group'
            Description = Get-Value $contentType 'Description'
            Hidden = Get-Value $contentType 'Hidden'
            ReadOnly = Get-Value $contentType 'ReadOnly'
            Sealed = Get-Value $contentType 'Sealed'
        })
    }

    $webFeatures = @(
        Invoke-InventoryOperation -Operation 'Get web-scoped features' -WebUrl $webUrl -ScriptBlock {
            Get-PnPFeature -Scope Web -Connection $connection
        }
    )
    foreach ($feature in $webFeatures) {
        $results.Features.Add([pscustomobject][ordered]@{
            SiteUrl = $siteUrlText
            WebUrl = $webUrl
            Scope = 'Web'
            DefinitionId = Get-Value $feature 'DefinitionId'
            DisplayName = Get-Value $feature 'DisplayName'
        })
    }

    $webReceivers = @(
        Invoke-InventoryOperation -Operation 'Get web event receivers' -WebUrl $webUrl -ScriptBlock {
            Get-PnPEventReceiver -Scope Web -Connection $connection
        }
    )
    foreach ($receiver in $webReceivers) {
        $results.EventReceivers.Add([pscustomobject][ordered]@{
            SiteUrl = $siteUrlText
            WebUrl = $webUrl
            Scope = 'Web'
            ListTitle = ''
            ReceiverName = Get-Value $receiver 'ReceiverName'
            ReceiverClass = Get-Value $receiver 'ReceiverClass'
            ReceiverAssembly = Get-Value $receiver 'ReceiverAssembly'
            ReceiverUrl = Get-Value $receiver 'ReceiverUrl'
            EventType = Get-Value $receiver 'EventType'
            Synchronization = Get-Value $receiver 'Synchronization'
        })
    }

    foreach ($location in @('QuickLaunch', 'TopNavigationBar')) {
        $nodes = @(
            Invoke-InventoryOperation -Operation "Get $location navigation" -WebUrl $webUrl -ScriptBlock {
                Get-PnPNavigationNode -Location $location -Tree -Connection $connection
            }
        )
        Add-NavigationNodes -Nodes $nodes -Location $location -CurrentWebUrl $webUrl
    }

    $lists = @(
        Invoke-InventoryOperation -Operation 'Get lists and libraries' -WebUrl $webUrl -ScriptBlock {
            Get-PnPList -Includes Id, Title, Description, BaseType, BaseTemplate, ItemCount, Hidden, ContentTypesEnabled, EnableVersioning, EnableMinorVersions, MajorVersionLimit, ForceCheckout, NoCrawl, OnQuickLaunch, Created, LastItemModifiedDate, DefaultViewUrl, 'RootFolder.ServerRelativeUrl', HasUniqueRoleAssignments -Connection $connection
        }
    )
    if (-not $IncludeHiddenLists) {
        $lists = @($lists | Where-Object { -not $_.Hidden })
    }

    foreach ($list in $lists) {
      try {
        $rootFolderUrl = Get-Value (Get-Value $list 'RootFolder') 'ServerRelativeUrl'
        $results.Lists.Add([pscustomobject][ordered]@{
            SiteUrl = $siteUrlText
            WebUrl = $webUrl
            ListId = Get-Value $list 'Id'
            Title = Get-Value $list 'Title'
            Description = Get-Value $list 'Description'
            BaseType = Get-Value $list 'BaseType'
            BaseTemplate = Get-Value $list 'BaseTemplate'
            ItemCount = Get-Value $list 'ItemCount'
            Hidden = Get-Value $list 'Hidden'
            ContentTypesEnabled = Get-Value $list 'ContentTypesEnabled'
            EnableVersioning = Get-Value $list 'EnableVersioning'
            EnableMinorVersions = Get-Value $list 'EnableMinorVersions'
            MajorVersionLimit = Get-Value $list 'MajorVersionLimit'
            ForceCheckout = Get-Value $list 'ForceCheckout'
            NoCrawl = Get-Value $list 'NoCrawl'
            OnQuickLaunch = Get-Value $list 'OnQuickLaunch'
            Created = Get-Value $list 'Created'
            LastItemModifiedDate = Get-Value $list 'LastItemModifiedDate'
            DefaultViewUrl = Get-Value $list 'DefaultViewUrl'
            RootFolderUrl = $rootFolderUrl
            HasUniqueRoleAssignments = Get-Value $list 'HasUniqueRoleAssignments'
        })

        if ($IncludePermissions -and $list.HasUniqueRoleAssignments) {
            Add-RoleAssignments -SecurableObject $list -Scope 'List' -ObjectTitle $list.Title -ObjectUrl $rootFolderUrl -CurrentWebUrl $webUrl -Connection $connection
        }

        $listFields = @(
            Invoke-InventoryOperation -Operation "Get fields for list '$($list.Title)'" -WebUrl $webUrl -ScriptBlock {
                Get-PnPField -List $list.Id -Connection $connection
            }
        )
        foreach ($field in $listFields) {
            $results.Fields.Add([pscustomobject][ordered]@{
                SiteUrl = $siteUrlText
                WebUrl = $webUrl
                Scope = 'List'
                ListTitle = $list.Title
                FieldId = Get-Value $field 'Id'
                Title = Get-Value $field 'Title'
                InternalName = Get-Value $field 'InternalName'
                TypeAsString = Get-Value $field 'TypeAsString'
                Required = Get-Value $field 'Required'
                Hidden = Get-Value $field 'Hidden'
                ReadOnlyField = Get-Value $field 'ReadOnlyField'
                Indexed = Get-Value $field 'Indexed'
                EnforceUniqueValues = Get-Value $field 'EnforceUniqueValues'
                Group = Get-Value $field 'Group'
            })
        }

        $views = @(
            Invoke-InventoryOperation -Operation "Get views for list '$($list.Title)'" -WebUrl $webUrl -ScriptBlock {
                Get-PnPView -List $list.Id -Includes Id, Title, ViewType, DefaultView, PersonalView, Paged, RowLimit, ServerRelativeUrl -Connection $connection
            }
        )
        foreach ($view in $views) {
            $results.Views.Add([pscustomobject][ordered]@{
                SiteUrl = $siteUrlText
                WebUrl = $webUrl
                ListTitle = $list.Title
                ViewId = Get-Value $view 'Id'
                Title = Get-Value $view 'Title'
                ViewType = Get-Value $view 'ViewType'
                DefaultView = Get-Value $view 'DefaultView'
                PersonalView = Get-Value $view 'PersonalView'
                Paged = Get-Value $view 'Paged'
                RowLimit = Get-Value $view 'RowLimit'
                ViewUrl = Get-Value $view 'ServerRelativeUrl'
            })
        }

        $listContentTypes = @(
            Invoke-InventoryOperation -Operation "Get content types for list '$($list.Title)'" -WebUrl $webUrl -ScriptBlock {
                Get-PnPContentType -List $list.Id -Connection $connection
            }
        )
        foreach ($contentType in $listContentTypes) {
            $results.ContentTypes.Add([pscustomobject][ordered]@{
                SiteUrl = $siteUrlText
                WebUrl = $webUrl
                Scope = 'List'
                ListTitle = $list.Title
                ContentTypeId = Get-Value $contentType 'StringId' (Get-Value $contentType 'Id')
                Name = Get-Value $contentType 'Name'
                Group = Get-Value $contentType 'Group'
                Description = Get-Value $contentType 'Description'
                Hidden = Get-Value $contentType 'Hidden'
                ReadOnly = Get-Value $contentType 'ReadOnly'
                Sealed = Get-Value $contentType 'Sealed'
            })
        }

        $listReceivers = @(
            Invoke-InventoryOperation -Operation "Get event receivers for list '$($list.Title)'" -WebUrl $webUrl -ScriptBlock {
                Get-PnPEventReceiver -List $list.Id -Connection $connection
            }
        )
        foreach ($receiver in $listReceivers) {
            $results.EventReceivers.Add([pscustomobject][ordered]@{
                SiteUrl = $siteUrlText
                WebUrl = $webUrl
                Scope = 'List'
                ListTitle = $list.Title
                ReceiverName = Get-Value $receiver 'ReceiverName'
                ReceiverClass = Get-Value $receiver 'ReceiverClass'
                ReceiverAssembly = Get-Value $receiver 'ReceiverAssembly'
                ReceiverUrl = Get-Value $receiver 'ReceiverUrl'
                EventType = Get-Value $receiver 'EventType'
                Synchronization = Get-Value $receiver 'Synchronization'
            })
        }

        $isPageLibrary = $list.BaseTemplate -eq 119 -or $list.Title -in @('Site Pages', 'Pages')
        if (-not $isPageLibrary) {
            continue
        }

        $pageItems = @(
            Invoke-InventoryOperation -Operation "Get pages from '$($list.Title)'" -WebUrl $webUrl -ScriptBlock {
                Get-PnPListItem -List $list.Id -PageSize 500 -Fields 'FileLeafRef', 'FileRef', 'ContentTypeId', 'PromotedState', 'Created', 'Modified', 'Author', 'Editor' -Connection $connection
            }
        )
        foreach ($pageItem in $pageItems) {
            $pageName = [string]$pageItem.FieldValues['FileLeafRef']
            if ($pageName -notlike '*.aspx') {
                continue
            }

            $pageUrl = [string]$pageItem.FieldValues['FileRef']
            $author = Get-Value $pageItem.FieldValues['Author'] 'Email' (Get-Value $pageItem.FieldValues['Author'] 'LookupValue')
            $editor = Get-Value $pageItem.FieldValues['Editor'] 'Email' (Get-Value $pageItem.FieldValues['Editor'] 'LookupValue')
            $results.Pages.Add([pscustomobject][ordered]@{
                SiteUrl = $siteUrlText
                WebUrl = $webUrl
                LibraryTitle = $list.Title
                PageName = $pageName
                PageUrl = $pageUrl
                ContentTypeId = $pageItem.FieldValues['ContentTypeId']
                PromotedState = $pageItem.FieldValues['PromotedState']
                Created = $pageItem.FieldValues['Created']
                Modified = $pageItem.FieldValues['Modified']
                Author = $author
                Editor = $editor
            })

            $components = @(
                Invoke-InventoryOperation -Operation "Get page components for '$pageName'" -WebUrl $webUrl -ScriptBlock {
                    Get-PnPPageComponent -Page $pageName -Connection $connection
                }
            )
            foreach ($component in $components) {
                $results.WebParts.Add([pscustomobject][ordered]@{
                    SiteUrl = $siteUrlText
                    WebUrl = $webUrl
                    PageUrl = $pageUrl
                    PageName = $pageName
                    Title = Get-Value $component 'Title'
                    WebPartId = Get-Value $component 'WebPartId'
                    InstanceId = Get-Value $component 'InstanceId'
                    ComponentType = Get-Value $component 'Type'
                    Section = Get-Value $component 'Section'
                    Column = Get-Value $component 'Column'
                })
            }
        }
      }
      catch {
        Add-InventoryError -Operation "Process list '$($list.Title)'" -WebUrl $webUrl -Attempt 1 -Message $_.Exception.Message
      }
    }

    if ($DelayBetweenWebsMilliseconds -gt 0) {
        Start-Sleep -Milliseconds $DelayBetweenWebsMilliseconds
    }
    }
    catch {
        Add-InventoryError -Operation 'Process web' -WebUrl $webUrl -Attempt 1 -Message $_.Exception.Message
    }
}
#endregion

}
catch {
    $script:FatalError = $_
    Write-Log -Level ERROR -Message ("Unhandled error: {0}`nAt: {1}`n{2}" -f $_.Exception.Message, ($_.InvocationInfo.PositionMessage -replace "`r?`n", ' '), $_.ScriptStackTrace)
}
finally {
    #region Export CSV Files and Manifest
    Write-Progress -Activity 'Creating SharePoint Online site inventory' -Completed
    Write-Log -Message 'Exporting CSV files'

    foreach ($name in $schemas.Keys) {
        $rows = if ($name -eq 'Errors') { @($script:Errors) } else { @($results[$name]) }
        Export-InventoryData -Name $name -Rows $rows -Columns $schemas[$name] -Folder $reportFolder
    }

    $counts = [ordered]@{}
    foreach ($name in $schemas.Keys) {
        $counts[$name] = if ($name -eq 'Errors') { $script:Errors.Count } else { $results[$name].Count }
    }

    $manifest = [ordered]@{
        SiteUrl = $siteUrlText
        StartedUtc = $inventoryStarted.ToString('o')
        CompletedUtc = [datetime]::UtcNow.ToString('o')
        PnPPowerShellVersion = (Get-Module PnP.PowerShell | Sort-Object Version -Descending | Select-Object -First 1 -ExpandProperty Version).ToString()
        IncludeHiddenLists = [bool]$IncludeHiddenLists
        IncludePermissions = [bool]$IncludePermissions
        ThrottleEvents = $script:ThrottleCount
        Counts = $counts
        Completed = ($null -eq $script:FatalError)
    }
    $manifest | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $reportFolder 'Manifest.json') -Encoding utf8

    Write-Log -Message "Inventory complete: $reportFolder"
    Write-Log -Message "Errors recorded: $($script:Errors.Count); throttle events: $script:ThrottleCount"
    #endregion
}
#endregion

if ($script:FatalError) {
    throw $script:FatalError
}
