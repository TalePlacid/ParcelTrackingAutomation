# Parcel Tracking Automation

[한국어](README.md)

A Windows-oriented Python automation tool that reads tracking numbers from an Excel delivery list, searches a parcel tracking website, and writes delivery completion dates back to a copied workbook.

## Features

- Opens a file picker for selecting the source Excel file.
- Creates a copy named `updated_<original file name>` in an `updated` folder next to the source file.
- Reads target rows and tracking numbers from the first worksheet.
- Uses Playwright to open Chrome and search up to 10 tracking numbers at a time.
- Writes delivery completion dates to the configured Excel column and creates a `log.txt` processing log.
- Can be packaged with the included PyInstaller spec file (`OA_version1.3.spec`).

## Project Structure

| File | Purpose |
| --- | --- |
| `BatchRunner.py` | Entry point and orchestration |
| `ExcelAdapter.py` | Excel loading, target row extraction, and date updates |
| `WebBrowserAdapter.py` | Playwright-based parcel tracking automation |
| `DeliveryDateParser.py` | Parses delivery dates from tracking result text |
| `ColorDetector.py` | Determines target rows from cell color and undelivered-reason text |
| `FileAdapter.py` | Copies the source file and writes the process log |
| `Parcel.py`, `ParcelSheet.py` | Parcel data model |
| `OA_version1.3.spec` | PyInstaller build configuration |

## Requirements

- Python 3.10 or later
- Google Chrome, or Chrome installed by Playwright
- Windows is recommended
- `.xlsx` Excel files are recommended

> The file picker displays `.xls` files, but the application uses `openpyxl`, so legacy `.xls` workbooks may not be supported.

## Installation

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install openpyxl python-dotenv playwright pyinstaller
python -m playwright install chrome
```

If Google Chrome is already installed, the last command may be unnecessary in some environments.

## Configuration

Create a `.env` file next to the executable or in the source root. This file contains deployment-specific settings such as the production site address, page selectors, and Excel column numbers, so it must not be included in the repository or public documentation.

Get the configuration file separately from the internal release owner.

## Excel Processing Rules

1. Only the first worksheet is processed.
2. The header row is found by scanning the configured report-date column for the value `보고날짜`.
3. Rows after the header are processed until the report-date column becomes empty.
4. A row is collected when either condition is true:
   - The report-date cell background is considered blank or white.
   - The undelivered-reason detail cell contains the configured client name.
5. If the resent tracking number contains a 12-digit number, it takes priority. Otherwise, the default tracking number is used.
6. When a date is parsed from the tracking result, the date, format, font, alignment, and border are applied to the configured delivery-date cell.

## Usage

```powershell
python BatchRunner.py
```

After you choose an Excel file, the application creates the following files next to the source workbook:

```text
updated/
  updated_<original file name>
  log.txt
```

## Build

```powershell
pyinstaller OA_version1.3.spec
```

The executable is typically created at `dist/OA_version1.3.exe`. The spec file bundles `.env` into the executable, so handle production configuration carefully when creating release builds.

## Troubleshooting

| Symptom | What to check |
| --- | --- |
| Chrome cannot be found | Verify Google Chrome is installed, or run `python -m playwright install chrome`. |
| Environment variable errors | Confirm `.env` exists at runtime and that the internal deployment settings are complete. |
| Delivery dates are not written | Check that the production configuration matches the actual tracking page and Excel template. |
| Excel file cannot be opened | Save the workbook as `.xlsx` instead of legacy `.xls`. |
