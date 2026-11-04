# Inventory-SPFarm PowerShell Scripts

This repository contains PowerShell scripts for inventorying, auditing, reporting, cleanup, and migration across SharePoint on-premises, SharePoint Online, and related Microsoft 365 workloads.

## Table of Contents

- [Scripts Overview](#scripts-overview)
- [SharePoint On-Premises Scripts](#sharepoint-on-premises-scripts)
- [SharePoint Online (SPO) Scripts](#sharepoint-online-spo-scripts)
- [Migration, Cleanup, and Utility Scripts](#migration-cleanup-and-utility-scripts)
- [Power Platform Scripts](#power-platform-scripts)
- [Disclaimer](#disclaimer)
- [License](#license)

## Scripts Overview

The scripts are grouped below by platform so you can easily distinguish SharePoint on-premises scripts from SharePoint Online (SPO) scripts. The list includes the scripts currently documented in this README.

## SharePoint On-Premises Scripts

| Script | Description |
|---|---|
| [Inventory-SPFarm.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Inventory-SPFarm.ps1) | Generates a report of SharePoint on-premises farm lists and libraries, including item counts and sizes. |
| [SharePoint_Farm_Inventory_Report.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/SharePoint_Farm_Inventory_Report.ps1) | Scans the SharePoint farm and collects detailed inventory information. |
| [SharePoint_Farm_Inventory_Report_WithMembers.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/SharePoint_Farm_Inventory_Report_WithMembers.ps1) | Generates a farm inventory report that also includes membership information. |
| [EmailLibrariesReport.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/EmailLibrariesReport.ps1) | Identifies document libraries configured for incoming email and exports key details. |
| [Get-PreservationHoldLibraryReport.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Get-PreservationHoldLibraryReport.ps1) | Reports on preservation hold libraries and related disposition details. |
| [Scan-SharePoint2010Workflows.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Scan-SharePoint2010Workflows.ps1) | Scans web applications for SharePoint 2010 workflow details and produces a report. |
| [Export-InfoPathAttachments.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Export-InfoPathAttachments.ps1) | Exports InfoPath form attachments from SharePoint content for archival or migration. |

## SharePoint Online (SPO) Scripts

| Script | Description |
|---|---|
| [Get-SPOSiteInventory.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Get-SPOSiteInventory.ps1) | Creates an inventory of SharePoint Online sites and related site metadata. |
| [SPO-Document-Inventory-Report.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/SPO-Document-Inventory-Report.ps1) | Produces a document inventory report for SharePoint Online libraries. |
| [Export-SPOnPremLibrary.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Export-SPOnPremLibrary.ps1) | Exports a SharePoint document library to a local directory structure or CSV and is used for SharePoint Online content management scenarios. |
| [Get-EventReceivers.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Get-EventReceivers.ps1) | Lists event receivers configured across SharePoint Online sites, lists, or libraries. |
| [Get-AllDocLibrariesReport.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Get-AllDocLibrariesReport.ps1) | Exports a broad inventory of document libraries across the SharePoint Online environment. |
| [Find-StaleSPOFiles.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Find-StaleSPOFiles.ps1) | Identifies stale SharePoint Online files based on last-modified activity and age thresholds. |
| [Find-StaleSPOFiles-Graph.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Find-StaleSPOFiles-Graph.ps1) | Uses Microsoft Graph to detect stale SharePoint Online files and produce a report. |
| [Find-SPOClassicWebparts.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Find-SPOClassicWebparts.ps1) | Scans SharePoint Online pages or sites for classic web parts. |
| [Find-SPOFeatures.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Find-SPOFeatures.ps1) | Finds feature usage and configuration details across SharePoint Online content. |
| [Find-SPOWebparts.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Find-SPOWebparts.ps1) | Detects SharePoint Online web parts in use across sites and pages. |
| [Find-OphanedSiteGroups.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Find-OphanedSiteGroups.ps1) | Identifies SharePoint site groups that may no longer have valid owners. |
| [Find-OrphanedPrivateChannels.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Find-OrphanedPrivateChannels.ps1) | Finds orphaned Microsoft Teams private channels or associated SharePoint sites. |
| [SPO-IRMDetector.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/SPO-IRMDetector.ps1) | Detects Information Rights Management settings and related document protection state. |
| [SPOLibrary-Cleanup.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/SPOLibrary-Cleanup.ps1) | Performs cleanup tasks for SharePoint Online libraries and related content. |
| [Move-SPO-folders.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Move-SPO-folders.ps1) | Moves SharePoint Online folders for content reorganization or migration. |
| [SPO-SiteCloner.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/SPO-SiteCloner.ps1) | Clones or copies SharePoint Online site structures and related configuration. |
| [find-orphaned-Groups.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/find-orphaned-Groups.ps1) | Finds orphaned SharePoint groups that may no longer have valid ownership. |
| [get-onedriveinfo-graph.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/get-onedriveinfo-graph.ps1) | Queries Microsoft Graph for OneDrive usage and metadata details. |

## Migration, Cleanup, and Utility Scripts

| Script | Description |
|---|---|
| [Calculate_CumulativeNumber_SPO_ApprovedNotDeleted.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Calculate_CumulativeNumber_SPO_ApprovedNotDeleted.ps1) | Calculates cumulative counts of approved but not deleted SharePoint Online items. |
| [Calculate_CumulativeNumber_SPO_ApprovedNotDeleted-Automated.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Calculate_CumulativeNumber_SPO_ApprovedNotDeleted-Automated.ps1) | Automates the cumulative approved-but-not-deleted count calculation. |
| [Calculate_MonthlyNumber_SPO_PermDeletions.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Calculate_MonthlyNumber_SPO_PermDeletions.ps1) | Aggregates monthly permanent deletion counts. |
| [Calculate_MonthlyNumber_SPO_PermDeletions-Automated.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Calculate_MonthlyNumber_SPO_PermDeletions-Automated.ps1) | Automates monthly permanent deletion count reporting. |

## Power Platform Scripts

| Script | Description |
|---|---|
| [Get-PP-Info.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Get-PP-Info.ps1) | Inventories Power Platform environments, Power Automate flows, Power Apps, connectors, and endpoint URLs. |

## Disclaimer

The sample scripts are provided AS IS without warranty of any kind. Use them at your own risk. See individual script headers for details.

## License

MIT License.
