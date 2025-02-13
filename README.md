# ExcelToHTMLConverter

A VBA macro that converts data from an Excel worksheet into an HTML table format, replacing spaces with non-breaking spaces (`&nbsp;`), and saves the result as a text file. This script helps users quickly export Excel data into an HTML structure for web use or other purposes.

## Features

- Converts an entire worksheet's data into an HTML table.
- Replaces spaces in cell values with `&nbsp;` to preserve spacing in the HTML output.
- Automatically detects the last row and column with data.
- Saves the HTML content as a `.txt` file in the default file path with a timestamped filename.

## Getting Started

### Prerequisites

You will need:
- Microsoft Excel (with macro support enabled)

### How to Use

1. Open Excel and press `Alt + F11` to open the VBA editor.
2. Insert a new module (`Insert > Module`), then paste the VBA code provided in the [ConvertSheetToHTML.bas](ConvertSheetToHTML.bas) file.
3. Save the Excel file as a macro-enabled workbook (`.xlsm`).
4. To run the macro:
   - Press `Alt + F8` and select `ConvertSheetToHTMLFormatWithNBSPAndSaveAsTextFile` from the list, then click `Run`.

The macro will create an HTML file with the data from the active worksheet, and save it in your default file path.