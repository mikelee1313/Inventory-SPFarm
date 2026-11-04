# Inventory-SPFarm PowerShell Scripts

This repository contains PowerShell scripts for inventorying, auditing, reporting, cleanup, and migration across SharePoint on-premises, SharePoint Online, and related Microsoft 365 workloads. The scripts help administrators gather farm and site metadata, review content and workflow configuration, identify stale or orphaned objects, and export information for analysis.

## Table of Contents

- [Scripts Overview](#scripts-overview)
- [SharePoint Farm and Inventory Scripts](#sharepoint-farm-and-inventory-scripts)
- [SharePoint Online and Governance Scripts](#sharepoint-online-and-governance-scripts)
- [Migration, Cleanup, and Utility Scripts](#migration-cleanup-and-utility-scripts)
- [Power Platform Scripts](#power-platform-scripts)
- [Disclaimer](#disclaimer)
- [License](#license)

## Scripts Overview

This repository includes the following PowerShell scripts. The list below is intended to reflect the current content of the repository and includes all script files currently present in the root of the project.

## SharePoint Farm and Inventory Scripts

| Script | Description |
|---|---|
| [Inventory-SPFarm.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Inventory-SPFarm.ps1) | Generates a report of SharePoint on-prem farm lists and libraries, including item counts, sizes, modified dates, and URLs. |
| [SharePoint_Farm_Inventory_Report.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/SharePoint_Farm_Inventory_Report.ps1) | Scans the SharePoint farm to collect detailed information about sites, lists, libraries, and owners. |
| [SharePoint_Farm_Inventory_Report_WithMembers.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/SharePoint_Farm_Inventory_Report_WithMembers.ps1) | Similar to the farm inventory report, but also includes SharePoint group membership details for each site. |
| [EmailLibrariesReport.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/EmailLibrariesReport.ps1) | Identifies document libraries configured for incoming email and exports key details such as URL, owner, and email alias. |
| [Get-AllDocLibrariesReport.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Get-AllDocLibrariesReport.ps1) | Exports a broad inventory of document libraries across the environment for reporting and review. |
| [Get-EventReceivers.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Get-EventReceivers.ps1) | Lists SharePoint event receivers configured on sites, lists, or libraries. |
| [Get-PreservationHoldLibraryReport.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Get-PreservationHoldLibraryReport.ps1) | Reports on preservation hold libraries and related disposition or retention metadata. |
| [Export-SPOnPremLibrary.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Export-SPOnPremLibrary.ps1) | Exports a SharePoint on-prem document library to a local directory structure or CSV-based inventory format. |
| [Scan-SharePoint2010Workflows.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Scan-SharePoint2010Workflows.ps1) | Scans web applications for SharePoint 2010 workflow details and produces workflow-related reports. |
| [Get-SPOSiteInventory.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Get-SPOSiteInventory.ps1) | Creates an inventory of SharePoint Online sites and related site metadata. |
| [SPO-Document-Inventory-Report.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/SPO-Document-Inventory-Report.ps1) | Produces a document inventory report for SharePoint Online libraries and files. |
| [Export-InfoPathAttachments.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Export-InfoPathAttachments.ps1) | Exports InfoPath form attachments from SharePoint content for archival or review purposes. |

## SharePoint Online and Governance Scripts

| Script | Description |
|---|---|
| [Find-StaleSPOFiles.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Find-StaleSPOFiles.ps1) | Identifies stale SharePoint Online files based on last-modified activity and age thresholds. |
| [Find-StaleSPOFiles-Graph.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Find-StaleSPOFiles-Graph.ps1) | Uses Microsoft Graph to detect stale SharePoint Online files and produce an audit report. |
| [Find-SPOClassicWebparts.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Find-SPOClassicWebparts.ps1) | Scans SharePoint Online pages or sites for classic web parts and reports where they are used. |
| [Find-SPOFeatures.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Find-SPOFeatures.ps1) | Finds feature usage and configuration details across SharePoint Online content. |
| [Find-SPOWebparts.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Find-SPOWebparts.ps1) | Detects SharePoint Online web parts in use across sites and pages. |
| [Find-OphanedSiteGroups.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Find-OphanedSiteGroups.ps1) | Identifies site groups that no longer have valid owners or are otherwise orphaned. |
| [Find-OrphanedPrivateChannels.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Find-OrphanedPrivateChannels.ps1) | Finds orphaned Microsoft Teams private channels or associated SharePoint resources. |
| [SPO-IRMDetector.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/SPO-IRMDetector.ps1) | Detects Information Rights Management (IRM) settings and related document protection state. |
| [SPOLibrary-Cleanup.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/SPOLibrary-Cleanup.ps1) | Performs cleanup tasks for SharePoint Online libraries and related content. |
| [Move-SPO-folders.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Move-SPO-folders.ps1) | Moves SharePoint Online folders as part of content reorganization or migration tasks. |
| [SPO-SiteCloner.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/SPO-SiteCloner.ps1) | Clones or copies SharePoint Online site structures and related configuration for testing or migration. |
| [find-orphaned-Groups.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/find-orphaned-Groups.ps1) | Finds orphaned SharePoint groups that may no longer be connected to valid ownership or membership models. |
| [get-onedriveinfo-graph.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/get-onedriveinfo-graph.ps1) | Queries Microsoft Graph for OneDrive usage and metadata details. |

## Migration, Cleanup, and Utility Scripts

| Script | Description |
|---|---|
| [Calculate_CumulativeNumber_SPO_ApprovedNotDeleted.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Calculate_CumulativeNumber_SPO_ApprovedNotDeleted.ps1) | Calculates cumulative counts for approved, non-deleted files or records in a SharePoint Online context. |
| [Calculate_CumulativeNumber_SPO_ApprovedNotDeleted-Automated.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Calculate_CumulativeNumber_SPO_ApprovedNotDeleted-Automated.ps1) | Automates the cumulative count workflow for approved, non-deleted SharePoint Online items. |
| [Calculate_MonthlyNumber_SPO_PermDeletions.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Calculate_MonthlyNumber_SPO_PermDeletions.ps1) | Aggregates monthly permanent deletion counts for SharePoint Online items or records. |
| [Calculate_MonthlyNumber_SPO_PermDeletions-Automated.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Calculate_MonthlyNumber_SPO_PermDeletions-Automated.ps1) | Automates monthly permanent deletion reporting for SharePoint Online content. |

## Power Platform Scripts

| Script | Description |
|---|---|
| [Get-PP-Info.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Get-PP-Info.ps1) | Inventories Power Platform environments, Power Automate flows, Power Apps, connectors, and endpoint URLs using delegated interactive authentication. |

## Disclaimer

The sample scripts are provided AS IS without warranty of any kind. Use them at your own risk. See individual script headers for details.

## License

MIT License.
