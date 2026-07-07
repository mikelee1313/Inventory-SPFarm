<#
.SYNOPSIS
Locates SharePoint Remote Event Receivers across the tenant in preparation for MC1411726.

.DESCRIPTION
This script connects to SharePoint Online by using app-only certificate authentication,
enumerates tenant sites, inspects non-hidden lists, and records Remote Event Receivers
that should be reviewed for the MC1411726 retirement effort.

The script excludes personal OneDrive sites and excludes
Microsoft.SharePoint.Webhooks.SPWebhookItemEventReceiver because webhook receivers are
not part of this retirement scope.

.OUTPUTS
Creates two files in the current user's temp folder:
- A log file that records scan progress, discovered receivers, and errors.
- A CSV file that contains the receiver inventory.

The CSV includes these columns:
SiteUrl, ListName, ReceiverName, ReceiverClass, ReceiverAssembly, ReceiverUrl, EventType, Type.

.NOTES
The app registration used by this script must be able to enumerate tenant sites and
read SharePoint list event receiver information.

Required API permission:
- SharePoint application permission `Sites.FullControl.All` with admin consent.

This script uses PnP PowerShell against SharePoint Online, so Microsoft Graph
permissions are not required for this specific inventory.

.AUTHOR
Mike Lee
Date: 7/7/2026
#>

$appID = "1e892341-f9cd-4c54-82d6-0fc3287954cf"  #This is your Entra App ID
$thumbprint = "B696FDCFE1453F3FBC6031F54DE988DA0ED905A9" #This is certificate thumbprint
$tenant = "9cfc42cb-51da-4055-87e9-b20a170b6ba3" #This is your Tenant ID
$adminUrl = "https://m365cpi13246019-admin.sharepoint.com"
$logPath = "$env:TEMP\MC1411726-RemoteEventReceivers-{0}.log" -f (Get-Date -Format "yyyyMMdd-HHmmss")
$csvPath = "$env:TEMP\MC1411726-RemoteEventReceivers-{0}.csv" -f (Get-Date -Format "yyyyMMdd-HHmmss")

# Define the connection parameters for reuse across PnP cmdlets
$connectionParams = @{
    ClientId      = $appID         # Azure AD App ID for authentication
    Thumbprint    = $thumbprint    # Certificate thumbprint for app-based authentication
    Tenant        = $tenant         # Tenant ID (GUID)
    WarningAction = 'SilentlyContinue' # Suppress PnP warnings that are not errors
}

Connect-PnPOnline -Url $adminUrl @connectionParams

$sites = Get-PnPTenantSite |
Where-Object { $_.Url -notmatch '-my\.sharepoint\.com(/|$)' }
$siteCount = $sites.Count
$siteIndex = 0
$totalReceiverCount = 0

Add-Content -Path $logPath -Value ("[{0}] Starting scan across {1} sites" -f (Get-Date -Format s), $siteCount)

foreach ($site in $sites)
{
    $siteIndex++
    Write-Progress -Activity "Scanning SharePoint sites for remote event receivers" -Status $site.Url -PercentComplete (($siteIndex / $siteCount) * 100)
    Add-Content -Path $logPath -Value ("[{0}] Processing site {1} ({2}/{3})" -f (Get-Date -Format s), $site.Url, $siteIndex, $siteCount)

    try
    {
        Connect-PnPOnline -Url $site.Url @connectionParams
        $siteResults = New-Object System.Collections.Generic.List[object]

        $lists = Get-PnPList |
        Where-Object { -not $_.Hidden }

        foreach ($list in $lists)
        {
            $remoteReceivers = Get-PnPEventReceiver -List $list |
            Where-Object {
                -not [string]::IsNullOrWhiteSpace($_.ReceiverUrl) -and
                $_.ReceiverClass -ne 'Microsoft.SharePoint.Webhooks.SPWebhookItemEventReceiver'
            }

            foreach ($receiver in $remoteReceivers)
            {
                $siteResults.Add([pscustomobject]@{
                    SiteUrl          = $site.Url
                    ListName         = $list.Title
                    ReceiverName     = $receiver.ReceiverName
                    ReceiverClass    = $receiver.ReceiverClass
                    ReceiverAssembly = $receiver.ReceiverAssembly
                    ReceiverUrl      = $receiver.ReceiverUrl
                    EventType        = $receiver.EventType
                    Type             = $receiver.Type
                })

                Add-Content -Path $logPath -Value ("[{0}] Found receiver {1} on list {2} in site {3}" -f (Get-Date -Format s), $receiver.ReceiverName, $list.Title, $site.Url)
            }
        }

        if ($siteResults.Count -gt 0)
        {
            $siteResults | Export-Csv -Path $csvPath -NoTypeInformation -Append
            $totalReceiverCount += $siteResults.Count
            Add-Content -Path $logPath -Value ("[{0}] Wrote {1} receiver records for site {2} to CSV" -f (Get-Date -Format s), $siteResults.Count, $site.Url)
            Write-Host ("Completed site {0}: found {1} Remote Event Receiver(s) affected by MC1411726" -f $site.Url, $siteResults.Count) -ForegroundColor Red
        }
        else
        {
            Add-Content -Path $logPath -Value ("[{0}] No Remote Event Receivers affected by MC1411726 found in site {1}" -f (Get-Date -Format s), $site.Url)
            Write-Host ("Completed site {0}: found 0 Remote Event Receivers affected by MC1411726" -f $site.Url) -ForegroundColor Green
        }
    }
    catch
    {
        Add-Content -Path $logPath -Value ("[{0}] Failed to process site {1}: {2}" -f (Get-Date -Format s), $site.Url, $_.Exception.Message)
    }
}

Write-Progress -Activity "Scanning SharePoint sites for remote event receivers" -Completed

if ($totalReceiverCount -eq 0)
{
    Add-Content -Path $logPath -Value ("[{0}] Scan complete. No Remote Event Receivers affected by MC1411726 were found." -f (Get-Date -Format s))
    Write-Host "No Remote Event Receivers affected by MC1411726 were found." -ForegroundColor Green
    Write-Host ("Remote event receiver log written to {0}" -f $logPath) -ForegroundColor Green
}
else
{
    Add-Content -Path $logPath -Value ("[{0}] Scan complete. Found {1} Remote Event Receivers affected by MC1411726. CSV output written to {2}" -f (Get-Date -Format s), $totalReceiverCount, $csvPath)
    Write-Host ("Remote event receiver log written to {0}" -f $logPath) -ForegroundColor Yellow
    Write-Host ("Remote event receiver CSV written to {0}" -f $csvPath) -ForegroundColor Yellow
}
