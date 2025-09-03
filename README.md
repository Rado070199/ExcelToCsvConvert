# Excel to CSV Converter

This is a simple console application written in **C#** that converts
Excel (`.xlsx`) files to CSV format.

The application uses the
[EPPlus](https://github.com/EPPlusSoftware/EPPlus) library for reading
Excel files.

------------------------------------------------------------------------

## 📦 Features

-   Reads Excel files (`.xlsx`).
-   Exports content to a CSV file with a chosen separator.
-   Supports escaping of special characters (quotes, separators,
    newlines, spaces).
-   Handles empty cells gracefully.
-   Provides detailed error messages for invalid input, missing files,
    or permission issues.

------------------------------------------------------------------------

## 🚀 Usage

``` bash
ExcelToCsvConverter.exe [ExcelPath] [CSVPath] [Separator] [WorksheetIndex]
```

### Arguments:

1.  **ExcelPath** -- full path to the source Excel file (`.xlsx`).
2.  **CSVPath** -- full path where the CSV file will be created.
3.  **Separator** -- character used to separate columns (e.g. `;` or
    `,`).
4.  **WorksheetIndex** -- index of the worksheet (starting from `0`).

------------------------------------------------------------------------

## 📝 Example

``` bash
ExcelToCsvConverter.exe "C:\data\source.xlsx" "C:\data\output.csv" ";" 0
```

This will read the first worksheet (`0`) from `source.xlsx` and export
it as a `;`-separated CSV file.

------------------------------------------------------------------------

## ⚠️ Error Handling

The application will display descriptive error messages in cases such
as: - Missing or invalid arguments - Excel file not found - Invalid
worksheet index - Permission denied for reading/writing files - Empty
worksheet

------------------------------------------------------------------------

## ⚖️ License

This project uses the **EPPlus** library under the [Polyform
Noncommercial
License](https://polyformproject.org/licenses/noncommercial/1.0.0/).
