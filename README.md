Here’s a README template for all the PowerShell scripts found in your mikelee1313/Inventory-SPFarm repository. This covers each script with a summary of its purpose and usage. (Note: There may be more scripts in the repo, but only four were found—search results are limited. For the full list, see the GitHub repository directly.)

---

# Inventory-SPFarm PowerShell Scripts

This repository contains PowerShell scripts designed to help SharePoint administrators inventory, audit, and extract information from an on-premises SharePoint Farm. The scripts focus on lists, libraries, workflows, and libraries configured with incoming email.

## Table of Contents

- [Scripts Overview](#scripts-overview)
- [Prerequisites](#prerequisites)
- [Usage](#usage)
- [Disclaimer](#disclaimer)
- [License](#license)

## Scripts Overview

| Script Name                             | Description                                                                                                   |
|-----------------------------------------|---------------------------------------------------------------------------------------------------------------|
| [Inventory-SPFarm.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/Inventory-SPFarm.ps1) | Generates a report of all lists and libraries in a SharePoint On-prem Farm, including item counts, size, last modified date, and URLs. Handles large lists with batching and includes logging. |
| [SharePoint_Farm_Inventory_Report.ps1](https://github.com/mikelee1313/Inventory-SPFarm/blob/main/SharePoint_Farm_Inventory_Report.ps1) | Scans the entire SharePoint farm to collect detailed info about sites, lists, and libraries, including owners. Exports the results to CSV and logs processing activities. |
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

Let me know if you want to add script usage examples, more detailed prerequisites, or have other scripts you want described!
