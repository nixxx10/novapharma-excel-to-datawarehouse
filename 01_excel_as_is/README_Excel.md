# NovaPharma_Commercial_Model — how to set it up

1. Open `NovaPharma_Commercial_Model.xlsx` in Excel. Press F9 once if the values look empty (the file is set to recalculate on open).
2. Import the macros: **Alt+F11 → File → Import File…** and import the six files in `vba/` (modCommon first). Then **File → Save As → Excel Macro-Enabled Workbook (.xlsm)**.
3. Optional: add a Developer-tab button or a Quick Access shortcut for each macro (ImportMonthlyFile, RefreshConsolidation, ValidateData, SnapshotForecastVersion, ExportToCSV).

## Try the monthly process

* Run **ImportMonthlyFile** and select `ERP_Sales_202609.csv` (the September 2026 extract, 1,574 rows, contains two duplicated lines on purpose).
* Change `Parameters!CurrentMonth` to `202609`, run **RefreshConsolidation** (takes 1–2 min: ~530k formulas).
* Run **ValidateData** — it will FAIL: the raw data ships with 45 duplicates, 18 unknown customer codes, 70 lower-case SKU codes and 25 SKU codes with trailing spaces. That is intentional: it is the "as-is" data quality the migration has to fix.
* Run **SnapshotForecastVersion** to create `FC 2026-Q3` frozen to `202609`.
* Run **ExportToCSV** to produce the flat files the Python ETL will read.

## What each sheet does

See the `README` sheet inside the workbook (sheet map, functions used, known limitations).
