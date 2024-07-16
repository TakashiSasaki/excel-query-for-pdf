# ExtractTablesFromPdf.ps1

## Overview
`ExtractTablesFromPdf.ps1` is a PowerShell script designed to extract tables from a specified PDF file and save them into an Excel workbook. The script uses Power Query to load the tables from the PDF and outputs the data to separate sheets in the Excel file.

## Prerequisites
- Windows operating system
- Microsoft Excel installed
- PowerShell 5.1 or later

## Installation
1. Ensure that Microsoft Excel is installed on your machine.
2. Download or clone the script to your local machine.

## Usage
1. Open PowerShell.
2. Navigate to the directory where `ExtractTablesFromPdf.ps1` is located.
3. Run the script with the following command:

```powershell
.\ExtractTablesFromPdf.ps1 -pdfFileName <path_to_pdf_file>
```

Replace `<path_to_pdf_file>` with the path to your PDF file.

### Example
```powershell
.\ExtractTablesFromPdf.ps1 -pdfFileName "C:\Users\YourName\Documents\example.pdf"
```

## Parameters
- `-pdfFileName`: The path to the PDF file from which tables will be extracted. This parameter is required.

## Features
- Extracts tables from a specified PDF file.
- Saves extracted tables into an Excel workbook with each table in a separate sheet.
- Provides progress indication during the extraction and saving process.

## Script Details
- The script resolves the PDF file path to an absolute path.
- Checks if the specified PDF file exists.
- Creates a new Excel workbook.
- Uses Power Query to load tables from the PDF file.
- Filters table IDs to include only those that start with "Table" followed by digits.
- Saves each extracted table in a new worksheet within the Excel workbook.
- Provides progress messages during the refresh operations.

## Error Handling
- The script checks if the PDF file exists and throws an error if it does not.
- If an Excel file with the same name as the PDF file already exists, it will be removed before creating a new one.

## Notes
- Ensure that the PDF file path provided is correct and accessible.
- The script must be run with appropriate permissions to create and write files in the specified directory.

## License
This script is provided "as is" without warranty of any kind. Use it at your own risk.

## Contributing
If you find any issues or have suggestions for improvements, feel free to open an issue or submit a pull request.

## Authors
- Takashi Sasaki
