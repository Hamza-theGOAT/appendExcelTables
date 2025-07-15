
# Append Excel Tables

This is a simple python script to collect the data from multiple tables in an Excel workbook and append them vertically in a new excel workbook.




## 🔧 Script Inputs – variables.json
The script uses the following input variables, all of which are defined in the variables.json file:

### path
Specifies the directory path where the Excel workbook containing the tables is located.

Example:
This is typically set to the project's storage_dir, but it can be updated to any absolute or relative folder path.

### wbName
The name of the Excel workbook (including the .xlsx extension) that contains the tables to be processed.

Example:
"WorkbookName.xlsx"

### wsName
The specific worksheet within the workbook (wbName) from which the tables will be read.

Example:
"Sheet1"

### tableKey
A prefix or keyword used to identify which folders contain the tables to be merged.
The script looks for folders whose names start with this key (e.g., "sales_table1", "sales_summary", if tableKey = "sales_").

Usage Note:
All target folders must begin with this key to be included in the merge process.

### tableColz
A list of column headers to extract from each identified table.
Only these specified columns will be selected during the merge.

Example:
["Product", "Quantity", "Price"]
## Required Python Libraries

- Pandas
- Openpyxl
- os
- json
- Datetime


## Output

The script will generate an unformated excel file in the `path` folder with the name of the input workbook `wbName` along with the timestamp of when the script ran.
