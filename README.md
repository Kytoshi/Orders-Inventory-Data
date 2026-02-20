# Orders Inventory Data

A Windows desktop application that automates daily data extraction from internal website and SAP systems, then processes the downloaded files into a consolidated Excel report.

## What It Does

1. **Web Downloads** — Logs into the interal portal via headless Chrome and downloads:
   - Material Shortage Report
   - Daily Order Fulfillment Reports (Completed, Incompletes, Billing)

2. **SAP Data Extraction** — Launches SAP GUI, runs three transactions in parallel:
   - MO Backorders (MB25)
   - MB51 material movements
   - Daily MO MB25

3. **Excel Report Engine** — Opens the excel report workbook, refreshes pivot tables, and appends summary rows across multiple sheets (MO YR SUMMARY, DN AO YR SUMMARY, SO YR COMP, SO YR INCMP, MO %).

All three phases can be run individually or as a full pipeline (downloads in parallel, then Excel report).

## Requirements

- Windows 10/11
- Python 3.10+
- SAP GUI with scripting enabled
- Google Chrome (for headless web downloads)

### Python Dependencies

```
PySide6
selenium
webdriver-manager
pywin32
pythoncom
psutil
send2trash
```

## Project Structure

```
├── App.spec                     # PyInstaller build spec
├── AMSO Logo v2.ico             # Application icon
├── AMSO Logo v2.png             # Application icon (high-res)
└── AMS_Orders/modules/
    ├── App.py                   # PySide6 GUI application
    ├── config.py                # Configuration loader
    ├── web_download.py          # PDBS web scraping (Selenium)
    ├── sap_download.py          # SAP GUI scripting automation
    ├── excel_report.py          # Excel pivot table processing
    ├── excel_manager.py         # Thread-safe Excel COM wrapper
    ├── helpers.py               # SAP connection, business day calc
    ├── file_utils.py            # File operations (copy, cleanup, download wait)
    └── logger.py                # Rotating file + console logger
```
