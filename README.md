This script generates a report of all lists and libraries in a SharePoint On-premises Farm, including file counts, total size, last modified date, and full URL. The output is saved to a CSV file, and a log file is generated to track the progress and any errors.

The script traverses through all site collections in the SharePoint Farm, processes each site and its sub-sites, and gathers details about all lists and libraries. It handles large lists by retrieving items in batches to avoid the list view threshold issue. The script also includes error handling and logging mechanisms

The output is logged and exported to a CSV file.

Example:

![image](https://github.com/user-attachments/assets/c154f008-fd77-40b7-a5d5-dba73fa57944)

![image](https://github.com/user-attachments/assets/004067ea-fe6a-43d8-ba40-df05ce50f169)
