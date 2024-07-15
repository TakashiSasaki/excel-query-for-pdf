# PowerShell and Power Query Excel Automation Example Scripts

This repository contains PowerShell scripts to automate Excel tasks using COM objects and Power Query.

## Scripts

### 1. CreateExcelFileWithPowerQuery.ps1

This script creates an Excel file, adds a Power Query to it, and saves the results into a worksheet.

#### Features:
- Deletes existing Excel file if present
- Adds Power Query to the workbook
- Loads Power Query results into the worksheet
- Saves the new Excel file with the query results

#### Usage:
```powershell
# Run the script
.\CreateExcelFileWithPowerQuery.ps1
```

#### Example:
```powershell
# Example usage:
.\CreateExcelFileWithPowerQuery.ps1
# This will create an Excel file with a Power Query that converts numbers 1 to 10 into a table.
```

### 2. ExtractTableFromPdfToExcel.ps1

This script downloads a PDF file from a specified URL, extracts table data from the PDF using Power Query, and saves the data into an Excel worksheet.

#### Features:
- Downloads `table.pdf` from a specified URL to the current directory
- Deletes existing `table.pdf` and `table_data.xlsx` if present
- Adds Power Query to extract table from the downloaded PDF file
- Loads extracted table data into an Excel worksheet
- Saves the new Excel file with the extracted table data

#### Usage:
```powershell
# Run the script
.\ExtractTableFromPdfToExcel.ps1
```

#### Example:
```powershell
# Example usage:
.\ExtractTableFromPdfToExcel.ps1
# This will download the table.pdf file, extract table data, and save it into table_data.xlsx.
```

## Requirements
- PowerShell 5.1 or later
- Excel installed on the system

## License
This repository is licensed under the MIT License.

## Contributions
Contributions are welcome! Please fork this repository and submit a pull request for any improvements.

## Authors
- Takashi Sasaki - [Homepage](https://x.com/TakashiSasaki)
