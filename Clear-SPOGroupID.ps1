#Requires -Version 7.4
<#
.SYNOPSIS
    Clears Microsoft 365 group IDs from orphaned SharePoint sites in a findings CSV.

.DESCRIPTION
    Accepts the GroupSiteFindings.csv output from Find-OphanedSiteGroups.ps1.
    Only rows with Status OrphanedGroupSite are processed; all other rows are
    reported as skipped. Review the findings and directory replication delays
    before running remediation.

    Uses the same certificate authentication as the discovery script and runs
    Set-PnPSite -Identity <site URL> -ClearGroupId against the tenant admin
    connection. SharePoint requires the group to be permanently deleted.
    If SharePoint requires permanent group deletion, the site is logged and
    reported as PermanentDeletionRequired, separately from unexpected failures.
    Sites whose group ID is already empty are reported as AlreadyClear, not
    as failures; rerunning previously remediated findings is safe.
    This script NEVER deletes or purges any group.

    Writes a timestamped results CSV and diagnostic log. Individual failures
    do not stop other sites, but unexpected failures cause a terminating error
    after export. Sites awaiting permanent group deletion produce an action
    summary instead; their IDs remain unchanged. Duplicate site URLs are
    processed only once.
    -WhatIf writes a preview report without importing PnP or connecting.
    Confirmation is requested for each site by default. Set ConfirmBeforeClear
    to $false in the parameter defaults to skip prompts. Explicit -Confirm
    or -Confirm:$false overrides this setting.
    Clearing GroupId does not promise to clear RelatedGroupId or change
    information barrier settings.

.PARAMETER InputPath
    Path to the reviewed GroupSiteFindings.csv produced by the discovery script.
    Set the default in the configuration below, or override it with -InputPath.
.PARAMETER TenantName
    SharePoint tenant prefix, without .sharepoint.com.
.PARAMETER AppId
    Entra application (client) ID.
.PARAMETER Thumbprint
    Thumbprint of the certificate with private key installed for PnP authentication.
.PARAMETER Tenant
    Entra tenant ID or verified domain.
.PARAMETER OutputDirectory
    Existing or new report directory. Defaults to the current user's TEMP folder.
.PARAMETER AdminUrl
    Optional SharePoint admin URL. Defaults to https://TenantName-admin.sharepoint.com.
    For Multi-Geo, use the geography's admin URL and its corresponding findings CSV.
.PARAMETER ConfirmBeforeClear
    Controls confirmation prompts. Set $true to prompt; hard-code $false below
    or supply -ConfirmBeforeClear:$false to skip prompts.
    Explicit -Confirm or -Confirm:$false takes precedence. -WhatIf is unchanged.

.NOTES
    Requires PnP.PowerShell 3.2.0 or later with Set-PnPSite -ClearGroupId support and application
    permission SharePoint Sites.FullControl.All with admin consent.
    No Microsoft Graph module or Graph write permission is required.
    Connection defaults match the discovery script; update them as appropriate.
    Install a compatible release with:
      Install-Module PnP.PowerShell -MinimumVersion 3.2.0 -Scope CurrentUser
    Restart PowerShell after upgrading if an older PnP module is already loaded.
    Compatibility is checked before connecting or modifying sites. -WhatIf
    remains module-independent and does not perform this check.
    Documentation: https://pnp.github.io/powershell/cmdlets/Set-PnPSite.html
    Release notes: https://github.com/pnp/powershell/blob/v3.2.0/CHANGELOG.md

Author: Mike Lee
Date: 11/15/26

.EXAMPLE
    # Use the InputPath configured in the script:
    .\Clear-SPOGroupID.ps1 -WhatIf

.EXAMPLE
    .\Clear-SPOGroupID.ps1 -InputPath 'C:\Reports\GroupSites_GroupSiteFindings.csv' -WhatIf

.EXAMPLE
    .\Clear-SPOGroupID.ps1 -InputPath 'C:\Reports\GroupSites_GroupSiteFindings.csv'

.EXAMPLE
    .\Clear-SPOGroupID.ps1 -InputPath 'C:\Reports\GroupSites_GroupSiteFindings.csv' -Confirm:$false

.EXAMPLE
    .\Clear-SPOGroupID.ps1 -ConfirmBeforeClear:$false
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    #region Parameters and configuration
    [ValidateNotNullOrEmpty()][string] $InputPath = 'C:\Temp\GroupSites_M365CPI13246019_2026-11-05_12-56-23_114_GroupSiteFindings.csv',
    [ValidateNotNullOrEmpty()][string] $TenantName = 'M365CPI13246019',
    [guid] $AppId = 'abc64618-283f-47ba-a185-50d935d51d57',
    [ValidateNotNullOrEmpty()][string] $Thumbprint = 'B696FDCFE1453F3FBC6031F54DE988DA0ED905A9',
    [ValidateNotNullOrEmpty()][string] $Tenant = '9cfc42cb-51da-4055-87e9-b20a170b6ba3',
    [ValidateNotNullOrEmpty()][string] $OutputDirectory = [System.IO.Path]::GetTempPath(),
    [uri] $AdminUrl,
    [bool] $ConfirmBeforeClear = $false
    #endregion
)

#region Runtime preferences
$ErrorActionPreference = 'Stop'
if (-not $ConfirmBeforeClear -and -not $PSBoundParameters.ContainsKey('Confirm')) {
    $ConfirmPreference = 'None'
}
# Keep preview report writes enabled while ShouldProcess gates tenant changes.
$previewOnly = [bool]$WhatIfPreference
$WhatIfPreference = $false
#endregion

#region Configuration and input validation
$missingSettings = @(
    foreach ($name in 'InputPath', 'TenantName', 'Thumbprint', 'Tenant') {
        if ([string]::IsNullOrWhiteSpace((Get-Variable -Name $name -ValueOnly))) { $name }
    }
    if ($AppId -eq [guid]::Empty) { 'AppId' }
)
if ($missingSettings.Count -gt 0) {
    throw "Configure $($missingSettings -join ', ') in the parameter defaults, or supply them as command-line parameters."
}
if (-not $AdminUrl) {
    $AdminUrl = [uri]"https://$TenantName-admin.sharepoint.com"
}
if (-not $AdminUrl.IsAbsoluteUri -or $AdminUrl.Scheme -ne 'https' -or
    $AdminUrl.Host -notlike '*-admin.sharepoint.com' -or $AdminUrl.UserInfo) {
    throw 'AdminUrl must be an HTTPS SharePoint tenant admin URL.'
}
$siteHost = $AdminUrl.Host -replace '-admin\.sharepoint\.com$', '.sharepoint.com'

$rows = @(Import-Csv -LiteralPath $InputPath -ErrorAction Stop)
if ($rows.Count -gt 0) {
    foreach ($column in 'URL', 'Status', 'Group ID', 'RelatedGroupId') {
        if ($column -notin $rows[0].PSObject.Properties.Name) {
            throw "Input CSV is missing required column '$column'. Use the discovery script's GroupSiteFindings.csv."
        }
    }
}

#endregion

#region Report paths and run state
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$stamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss_fff'
$prefix = Join-Path $OutputDirectory "ClearSPOGroupID_${TenantName}_$stamp"
$reportPath = "${prefix}_Results.csv"
$logPath = "${prefix}.log"
$connection = $null
$seenUrls = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$cleared = 0
$alreadyClear = 0
$failed = 0
$pendingDeletion = 0
$skipped = 0
$planned = 0
#endregion

#region Logging helper
function Write-LogEntry {
    param(
        [string] $Message,
        [ValidateSet('INFO', 'WARNING', 'ERROR')][string] $Level = 'INFO'
    )
    Add-Content -LiteralPath $logPath -Value "$(Get-Date -Format o) [$Level] $Message"
}
#endregion

#region Results report initialization
$schema = [pscustomobject][ordered]@{
    URL = ''
    'Group ID' = ''
    RelatedGroupId = ''
    InputStatus = ''
    ClearGroupIdStatus = ''
    Details = ''
    ProcessedAtUtc = ''
}
Set-Content -LiteralPath $reportPath -Value ($schema | ConvertTo-Csv -NoTypeInformation)[0] -Encoding utf8BOM
#endregion

try {
    #region PnP compatibility and connection
    Write-LogEntry "Loaded $($rows.Count) findings from $InputPath. Preview only: $previewOnly."
    if (-not $previewOnly -and @($rows | Where-Object Status -eq 'OrphanedGroupSite').Count -gt 0) {
        $minimumPnPVersion = [version]'3.2.0'
        $installInstruction = 'Install-Module PnP.PowerShell -MinimumVersion 3.2.0 -Scope CurrentUser'
        $loadedModules = @(Get-Module -Name PnP.PowerShell)
        if (@($loadedModules | Where-Object Version -lt $minimumPnPVersion).Count -gt 0) {
            $versions = ($loadedModules.Version | Sort-Object -Unique) -join ', '
            throw "PnP.PowerShell $versions is already loaded. Version $minimumPnPVersion or later is required. Run '$installInstruction' if needed, then restart PowerShell to avoid mixing module assemblies."
        }
        $availableModules = @(Get-Module -ListAvailable -Name PnP.PowerShell | Sort-Object Version -Descending)
        $compatibleModule = $availableModules | Where-Object Version -ge $minimumPnPVersion | Select-Object -First 1
        if ($null -eq $compatibleModule) {
            $versions = if ($availableModules.Count -gt 0) { ($availableModules.Version | Sort-Object -Unique) -join ', ' } else { 'none' }
            throw "PnP.PowerShell $minimumPnPVersion or later is required for Set-PnPSite -ClearGroupId. Installed versions: $versions. Run '$installInstruction', then restart PowerShell."
        }
        Import-Module -Name $compatibleModule.Path -ErrorAction Stop
        $setSiteCommand = Get-Command -Name 'PnP.PowerShell\Set-PnPSite' -ErrorAction Stop
        if ($setSiteCommand.Module.Version -lt $minimumPnPVersion -or
            -not $setSiteCommand.Parameters.ContainsKey('ClearGroupId')) {
            throw "Loaded PnP.PowerShell $($setSiteCommand.Module.Version) is incompatible: Set-PnPSite -ClearGroupId requires version $minimumPnPVersion or later and parameter support. Run '$installInstruction', then restart PowerShell."
        }
        Write-LogEntry "Compatibility verified: PnP.PowerShell $($setSiteCommand.Module.Version), Set-PnPSite -ClearGroupId supported."
        Write-LogEntry "Connecting to $AdminUrl."
        $connection = PnP.PowerShell\Connect-PnPOnline -Url $AdminUrl.AbsoluteUri -ClientId $AppId.ToString() `
            -Thumbprint $Thumbprint -Tenant $Tenant -ReturnConnection -ErrorAction Stop
    }

    #endregion

    #region Site processing
    foreach ($row in $rows) {
        $result = [pscustomobject][ordered]@{
            URL = $row.URL
            'Group ID' = $row.'Group ID'
            RelatedGroupId = $row.RelatedGroupId
            InputStatus = $row.Status
            ClearGroupIdStatus = ''
            Details = ''
            ProcessedAtUtc = ''
        }
        try {
            if ($row.Status -ne 'OrphanedGroupSite') {
                $result.ClearGroupIdStatus = 'SkippedNotOrphaned'
                $result.Details = 'Only OrphanedGroupSite findings are eligible.'
                $skipped++
            }
            else {
                $siteUrl = $null
                if (-not [uri]::TryCreate([string]$row.URL, [UriKind]::Absolute, [ref]$siteUrl) -or
                    $siteUrl.Scheme -ne 'https' -or $siteUrl.Host -ne $siteHost -or
                    $siteUrl.UserInfo -or $siteUrl.Query -or $siteUrl.Fragment -or -not $siteUrl.IsDefaultPort) {
                    throw "Invalid site URL or URL outside the selected SharePoint geography: $($row.URL)"
                }
                $url = $siteUrl.AbsoluteUri.TrimEnd('/')
                if (-not $seenUrls.Add($url)) {
                    $result.ClearGroupIdStatus = 'SkippedDuplicate'
                    $result.Details = 'This site URL already appeared as an orphan candidate in the input.'
                    $skipped++
                }
                elseif ($previewOnly) {
                    $result.ClearGroupIdStatus = 'WhatIf'
                    $result.Details = 'Would run Set-PnPSite -ClearGroupId for this site; no tenant changes made.'
                    Write-Host "What if: Clear Microsoft 365 group ID on $url"
                    $planned++
                }
                elseif ($PSCmdlet.ShouldProcess($url, 'Clear Microsoft 365 group ID')) {
                    PnP.PowerShell\Set-PnPSite -Identity $url -ClearGroupId -Connection $connection -ErrorAction Stop
                    $result.ClearGroupIdStatus = 'Cleared'
                    $result.Details = 'Set-PnPSite -ClearGroupId completed successfully.'
                    $cleared++
                    Write-Host "Cleared: $url" -ForegroundColor Green
                }
                else {
                    $result.ClearGroupIdStatus = 'SkippedNotConfirmed'
                    $result.Details = 'The site change was not confirmed.'
                    $skipped++
                }
            }
        }
        #region Site outcome handling
        catch {
            $message = $_.Exception.Message
            $result.ClearGroupIdStatus = 'Failed'
            $result.Details = $message
            if ($message.Trim() -eq 'The group ID is already empty.') {
                $result.ClearGroupIdStatus = 'AlreadyClear'
                $result.Details = 'The group ID is already empty; no change needed.'
                $alreadyClear++
                Write-Host "Already cleared: $($row.URL)" -ForegroundColor Yellow
            }
            elseif ($message -like "*Couldn't clear group ID because the group is in the *" -and
                $message -like '*first permanently delete the group*') {
                $result.ClearGroupIdStatus = 'PermanentDeletionRequired'
                $result.Details = "$message Review and permanently delete the group in Azure Groups Management Portal, then rerun for this site. This script did not delete the group."
                $pendingDeletion++
                Write-Host "Pending permanent group deletion (unchanged): $($row.URL)" -ForegroundColor Red
            }
            else {
                $failed++
                Write-Warning "$($row.URL): $($result.Details)"
            }
        }
        #endregion

        #region Site result export and logging
        $result.ProcessedAtUtc = [DateTimeOffset]::UtcNow.ToString('o')
        $result | Export-Csv -LiteralPath $reportPath -Append -NoTypeInformation -Encoding utf8BOM
        $level = if ($result.ClearGroupIdStatus -eq 'Failed') { 'ERROR' }
            elseif ($result.ClearGroupIdStatus -eq 'PermanentDeletionRequired') { 'WARNING' }
            else { 'INFO' }
        Write-LogEntry "$($row.URL): $($result.ClearGroupIdStatus). $($result.Details)" $level
        #endregion
    }
    #endregion

    #region Run summary and next steps
    Write-LogEntry "Finished: $cleared cleared; $alreadyClear already clear; $pendingDeletion pending permanent group deletion; $planned previewed; $skipped skipped; $failed failed."
    Write-Host ''
    Write-Host 'Run summary' -ForegroundColor Cyan
    $summaryRows = @(
        [pscustomobject]@{ Status = 'Group IDs cleared successfully'; Count = $cleared; Color = 'Green' }
        [pscustomobject]@{ Status = 'Already clear - no changes needed'; Count = $alreadyClear; Color = 'Yellow' }
        [pscustomobject]@{ Status = 'Pending permanent group deletion'; Count = $pendingDeletion; Color = 'Red' }
        [pscustomobject]@{ Status = 'Preview only - sites that would be cleared'; Count = $planned; Color = 'Cyan' }
        [pscustomobject]@{ Status = 'Sites skipped'; Count = $skipped; Color = 'Gray' }
        [pscustomobject]@{ Status = 'Unexpected errors'; Count = $failed; Color = if ($failed -gt 0) { 'Red' } else { 'Green' } }
    )
    Write-Host ('  {0,-44} {1,10}' -f 'Status', 'Sites')
    Write-Host ('  {0,-44} {1,10}' -f ('-' * 44), ('-' * 10))
    foreach ($summaryRow in $summaryRows) {
        if ($summaryRow.Count -gt 0 -or $summaryRow.Status -eq 'Unexpected errors') {
            Write-Host ('  {0,-44} {1,10}' -f $summaryRow.Status, $summaryRow.Count) -ForegroundColor $summaryRow.Color
        }
    }
    if ($pendingDeletion -gt 0) {
        Write-Host ''
        Write-Host "Next step: $pendingDeletion sites still need attention." -ForegroundColor Red
        Write-Host '  Their group IDs could not be cleared because SharePoint requires the groups to be permanently deleted.'
        Write-Host '  Review those groups in Azure Groups Management Portal. If appropriate, permanently delete them, then rerun this script.'
        Write-Host '  Those sites were left unchanged. This script did not delete any groups.'
    }
    Write-Host ''
    Write-Host "Results report: $reportPath"
    Write-Host "Log: $logPath"
    #endregion

    #region Unexpected failure reporting
    if ($failed -gt 0) {
        throw "$failed site(s) failed. Review the results report and log; remediation is incomplete."
    }
    #endregion
}
#region Run error handling and cleanup
catch {
    Write-LogEntry $_.Exception.Message 'ERROR'
    throw
}
finally {
    if ($null -ne $connection -and $null -ne $connection.Context) {
        $connection.Context.Dispose()
    }
}
#endregion
