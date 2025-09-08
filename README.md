# ExportAnalytical_CLA

[cite_start]This VBA module is designed for Microsoft Access to automate the export of analytical data into Excel workbooks and manage related PDF documents for specific customers[cite: 46].

## Key Features

- [cite_start]**Automated Data Export**: Exports data from the `Consumables` table to a formatted Excel workbook for customers with `Customer_Entity` set to 'CLA'[cite: 48, 49].
- [cite_start]**Dynamic File Naming and Folder Creation**: Generates a safe file path based on the customer name and the maximum `Posting Month` from the data[cite: 9, 10, 11]. [cite_start]It also creates a dedicated monthly folder if it does not exist[cite: 10].
- **Excel Workbook Generation**: Creates a new Excel workbook with two sheets:
    - [cite_start]**"Data" sheet**: Contains the raw data from the `Consumables` table, including a formatted header with key reporting parameters like the ledger, posting dates, and a pipe-delimited list of unique `Segment5` values[cite: 66, 67, 72, 73].
    - [cite_start]**"Overview" sheet**: Features a dynamic PivotTable summarizing the data[cite: 78, 79]. [cite_start]The pivot table is configured with specific row and column fields and includes custom styling for subtotals[cite: 80, 81, 99, 106].
- [cite_start]**Custom Formatting and Branding**: Applies consistent formatting (font, size, bolding) to the Excel headers[cite: 68, 84]. [cite_start]It also inserts a company logo onto the sheets, retrieved from a `Company` table in the Access database[cite: 17, 18, 24].
- [cite_start]**Logging**: Includes a simple text-file logger to record any problems encountered during the export process, such as path errors or save failures[cite: 11, 12, 114].
- **Helper Functions**: The module contains several helper functions to ensure data integrity and proper formatting:
    - [cite_start]`CleanFileName`: Removes illegal characters from strings to create valid file names[cite: 3, 4, 5].
    - [cite_start]`NormalizeKey`: Normalizes strings by trimming and replacing non-breaking spaces and tabs[cite: 8].
    - [cite_start]`ParsePostingMonthToDate`: Converts various date formats into a standard date value[cite: 35, 36, 37].
    - [cite_start]`GetOrCreateSheet`: Safely retrieves an existing Excel worksheet or creates a new one[cite: 6, 7].
    - [cite_start]`SetProperty`: Applies font and size properties to a specified range in an Excel sheet[cite: 16, 17].
    - [cite_start]`InsertCompanyLogo`: Embeds a company logo from a database attachment field into an Excel sheet[cite: 18, 20, 24].
    - [cite_start]`FieldLabelColumn`: Finds the column index for a given PivotTable field[cite: 38].
    - [cite_start]`PivotHideIfExists`, `PivotEnsureRowField`: Utility functions to configure the PivotTable layout programmatically[cite: 43, 44, 92].
    - [cite_start]`GetPivotOrder`: Captures the order of the row fields in the PivotTable[cite: 32].

## Requirements

- Microsoft Access
- [cite_start]Microsoft Excel (VBA uses late-binding, so a specific version reference is not required)[cite: 1].
- A database with the following tables:
    - `Consumables`: Contains the raw data for the reports.
    - [cite_start]`Customer`: Contains customer details, including `Customer_Name`, `Customer_Entity`, and `Customer_Path`[cite: 48].
    - [cite_start]`Company`: Contains company-specific information, including the `CompanyLogo` stored as an attachment[cite: 18].
- [cite_start]The `Microsoft Scripting Runtime` reference is required for the `Scripting.Dictionary` object used for de-duplication[cite: 51].

## How to Use

1.  **Place the code** in a standard module within your Access database.
2.  **Ensure all required tables** (`Consumables`, `Customer`, `Company`) are present and populated with data.
3.  **Run the `ExportAnalytical_CLA` subroutine** to initiate the export process.
