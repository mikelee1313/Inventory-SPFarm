# Find-StaleSPOFiles.ps1

## Overview

**Find-StaleSPOFiles.ps1** is a PowerShell script designed to scan SharePoint Online sites and identify all files and folders that have **not been modified within a specified number of months** (“stale” items). It connects using app-only authentication and iterates through sites listed in an input file, scanning document libraries and lists to report on stale content. Results are exported to CSV or Excel, with detailed logging and summary files.

---

## Features

- **Stale File Detection:** Identifies files/folders not modified since a configurable cutoff date.
- **Multi-Site Scanning:** Processes multiple SharePoint Online sites from a list.
- **Document Libraries & Lists:** Scans both document libraries and generic lists.
- **Exclusions:** Ignores common system folders and lists to improve accuracy.
- **Batch Processing:** Efficiently writes results in batches for performance.
- **Custom Output:** Exports results to CSV or Excel (XLSX) formats.
- **Summary Generation:** Produces summary files for overall and per-site statistics.
- **Robust Logging:** Logs all operations, including errors and progress, to a log file.
- **Throttling Handling:** Automatically retries on SharePoint throttling errors.

---

## Prerequisites

Before running the script, ensure you have the following:

- **PowerShell 5.1+** (Windows)
- **Modules:**
  - [PnP.PowerShell](https://www.powershellgallery.com/packages/PnP.PowerShell)
  - [ImportExcel](https://www.powershellgallery.com/packages/ImportExcel) (if using Excel output)
- **App Registration:** An Azure AD app with `Sites.FullControl.All` permissions, using a certificate for app-only authentication.
- **Input File:** A text file containing the SharePoint site URLs to scan (one per line).

---

## Configuration

Edit the following variables at the top of the script to match your environment:

```powershell
$appID        # Azure AD App registration ID
$thumbprint   # Certificate thumbprint for authentication
$tenant       # Azure AD Tenant ID
$inputFilePath # Path to text file containing site URLs
$monthsBack   # Number of months for "stale" cutoff
$batchSize    # Batch size for writing results
$maxItemsPerSheet # Max items per worksheet in Excel
$outputFormat # "CSV" or "XLSX"
$debug        # $true for debug logging, $false for info only
```

**Example:**
```powershell
$appID        = "xxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
$thumbprint   = "ABCDEF1234567890ABCDEF1234567890ABCDEF12"
$tenant       = "yyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy"
$inputFilePath = "C:\temp\SPOSiteList.txt"
$monthsBack   = 6
$outputFormat = "XLSX"
$debug        = $false
```

---

## Usage

1. **Prepare the Input File:**  
   Create a plain text file (e.g., `C:\temp\SPOSiteList.txt`) listing each SharePoint Online site URL to scan on a separate line.

2. **Update Script Configuration:**  
   Edit the variables at the top of the script as described above.

3. **Install Required PowerShell Modules:**  
   Open PowerShell as your user (not as Administrator) and run:
   ```powershell
   Install-Module PnP.PowerShell -Scope CurrentUser
   Install-Module ImportExcel -Scope CurrentUser   # Only if using XLSX output
   ```

4. **Run the Script:**  
   Execute the script from PowerShell:
   ```powershell
   .\Find-StaleSPOFiles.ps1
   ```

   > **Note:** You may need to set the execution policy:
   > ```powershell
   > Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
   > ```

---

## Output

After execution, the script generates the following files in your `%TEMP%` directory:

- **Main Results File:**  
  - `Stale_Files_Report_[timestamp].csv` or `.xlsx`  
    Contains details of all stale files/folders found.
- **Summary Files:**  
  - `Stale_Files_Report_[timestamp]_Summary.csv`  
    Overall scan statistics.
  - `Stale_Files_Report_[timestamp]_SiteSummary.csv`  
    Site-by-site breakdown of stale files found.
- **Log File:**  
  - `Stale_Files_Report_[timestamp].txt`  
    Execution log with timestamps, progress, and errors.

**Excel output** includes multiple worksheets if results exceed the specified maximum per sheet.

---

## Details Included in Output

Each item includes:
- Site URL
- Item Type (`File` or `Folder`)
- Library/List Name
- Item Path
- Item Name
- Created By
- Created Date
- Modified Date

---

## Advanced Options & Customization

- **Batch and Sheet Size:**  
  Adjust `$batchSize` and `$maxItemsPerSheet` for performance with large environments.
- **Ignored Folders/Lists:**  
  The `$ignoreFolders` array can be customized to skip additional system or irrelevant folders/lists.

---

## Logging & Troubleshooting

- All actions and errors are logged to the log file in `%TEMP%`.
- Enable `$debug = $true` for detailed diagnostics and progress messages.
- Common errors (e.g., missing modules, authentication issues, site access problems) are reported in the log and console.

---

## Examples

**Scan for files stale for 12+ months and output to Excel:**
```powershell
$monthsBack = 12
$outputFormat = "XLSX"
```
Run:
```powershell
.\Find-StaleSPOFiles.ps1
```

---

## Disclaimer

The script is provided “AS IS” without warranty. Use at your own risk; always test in a non-production environment.

---

## Author

- **Mike Lee**  
  [GitHub Repo](https://github.com/mikelee1313/Inventory-SPFarm)

---

## License

See [LICENSE](../LICENSE) in the repository for details.

