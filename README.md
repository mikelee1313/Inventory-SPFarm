---

# Inventory-SPFarm PowerShell Scripts

This repository contains PowerShell scripts for inventorying, auditing, and extracting information from an on-premises SharePoint Farm. The scripts help administrators generate detailed reports about lists, libraries, workflows, and more.

## Table of Contents

- [Scripts Overview](#scripts-overview)
- [Prerequisites](#prerequisites)
- [Disclaimer](#disclaimer)
- [License](#license)

## Scripts Overview

| Script Name                                                                                      | Description                                                                                                   |
|--------------------------------------------------------------------------------------------------|---------------------------------------------------------------------------------------------------------------|
| [Inventory-SPFarm.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Inventory-SPFarm.ps1) | Generates a report of all lists and libraries in a SharePoint On-prem Farm, including item counts, size, last modified date, and URLs. Handles large lists with batching and includes logging. |
| [SharePoint_Farm_Inventory_Report.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/SharePoint_Farm_Inventory_Report.ps1) | Scans the entire SharePoint farm to collect detailed info about sites, lists, and libraries, including owners. Exports the results to CSV and logs processing activities. |
| [SharePoint_Farm_Inventory_Report_WithMembers.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/SharePoint_Farm_Inventory_Report_WithMembers.ps1) | Similar to the above, but also includes the members of SharePoint groups for each site, providing a comprehensive inventory including user/group membership. |
| [EmailLibrariesReport.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/EmailLibrariesReport.ps1) | Finds all document libraries configured with incoming email and exports details (URL, library name, owner, email alias) to CSV. |
| [Scan-SharePoint2010Workflows.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Scan-SharePoint2010Workflows.ps1) | Scans a SharePoint Web Application to retrieve 2010 workflow details and last run status. Outputs CSV reports for workflows and blocked sites. |

> For a complete and up-to-date list of scripts, visit the [repository code browser](https://github.com/mikelee1313/Inventory-SPFarm).

## Prerequisites

- PowerShell 5.1 or later
- SharePoint Management Shell (for on-prem environments)
- Appropriate SharePoint farm administrative permissions
- Update variables (such as SharePoint URLs or output paths) inside each script as needed

## Disclaimer

The sample scripts are provided AS IS without warranty of any kind. Use them at your own risk. See individual script headers for details.

## License

MIT License.

---
