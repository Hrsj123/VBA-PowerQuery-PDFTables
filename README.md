# Accessing Data from PDF Tables into VBA using Power Query

This VBA project enables you to extract table data from PDF files using Power Query and access it in VBA code.

## Features

- Reads tables from PDFs using Power Query in Microsoft 365.
- Stores extracted data in a dictionary for easy access.
- Processes multiple PDF files automatically.
- Cleans up temporary queries and worksheets after extraction.

## How It Works

1. `FetchDataFromPDFs(pdfPaths)`: Loops through provided PDF paths, extracts tables, and returns a dictionary.
2. `LoadDataTables(filePath)`: Fetches table data from a single PDF and stores it in a dictionary.
3. `GetPDFTablesIdList(filePath)`: Retrieves table IDs from the PDF for further processing. (Required in LoadDataTables function to iterate over each table in a pdf.)

## Usage

1. Add the VBA code to your Excel macro-enabled workbook.
2. Ensure Power Query is enabled in Excel (Microsoft 365 required).
3. Pass an array of PDF file paths to `FetchDataFromPDFs` to extract and store tables.

### Example

```vba
Dim pdfPaths As Variant
Dim extractedData As Scripting.Dictionary

pdfPaths = Array("C:\path\to\file1.pdf", "C:\path\to\file2.pdf")
Set extractedData = FetchDataFromPDFs(pdfPaths)

' Access extracted data
Dim sampleName As Variant
For Each sampleName In extractedData.Keys
    Debug.Print "Data for: " & sampleName
    Dim tableData As Scripting.Dictionary
    Set tableData = extractedData(sampleName)
    Dim key As Variant
    For Each key In tableData.Keys
        Debug.Print key, tableData(key)
    Next key
Next sampleName
```

## Requirements

- Microsoft Excel (Microsoft 365 recommended for Power Query support).
- Power Query enabled.

## Notes

- The extracted data is stored in a dictionary, which can be further processed as needed. (I prefer using Dictionary. Any other data structure will do.)
- The script automatically removes temporary queries after data extraction.

## Disclaimer

This script is provided as-is without any warranties. Use it at your own discretion.\
Also reading pdf tables using Power Query is not very reliable if tables are not well structured.

