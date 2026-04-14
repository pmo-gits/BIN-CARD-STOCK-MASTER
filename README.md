INVENTORY LEDGER SYSTEM
Google Apps Script — Project README
Version 1.0   |   April 2026

1. Project Overview
This project is a Google Apps Script automation layer for a Google Sheets–based inventory management system. It provides a custom menu-driven interface to push stock movements — opening stock, received goods, and issued goods — into their respective ledger sheets and a unified IN-OUT ENTRY log, reducing manual data entry errors and maintaining a clean audit trail.

Attribute	Detail
Platform	Google Sheets + Google Apps Script (V8 runtime)
Language	JavaScript (GAS)
Trigger	Custom menu: Update Script
Files	4 .gs files (Menu + 3 operations)
Target Users	Warehouse / Store team members

2. Sheet Structure
The system reads from entry sheets and writes to ledger sheets. All three operations also append a log entry to the shared IN-OUT ENTRY sheet.

Sheet Name	Role	Read / Write
PRICE LIST MASTER	Master item catalogue with opening stock quantities	Read
OPENING STOCK LEDGER	Permanent record of all opening stock entries	Write
RECEIVED ENTRY	Temporary staging area for incoming goods	Read + Clear
RECEIVED LEDGER	Permanent record of all received goods	Write
ISSUED ENTRY	Temporary staging area for issued goods	Read + Clear
ISSUED LEDGER	Permanent record of all issued goods	Write
IN-OUT ENTRY	Unified chronological log of all movements	Write




3. File Descriptions
01_Menu.gs
Registers the onOpen() trigger which builds the custom menu Update Script in the Google Sheets toolbar. Contains three menu items, each calling one of the three core functions.

02_OpeningStockLedgerPush.gs
Implements updateOpeningStock(). This function reads qualifying rows from PRICE LIST MASTER and pushes them to OPENING STOCK LEDGER and IN-OUT ENTRY. After writing, it marks the source column with the text OPENING STOCK UPDATED and applies red cell background with sheet protection to prevent re-processing.

Key behaviour
•	Must be run from the PRICE LIST MASTER tab — enforced by active sheet check.
•	Duplicate prevention: skips any ITEMCODE already present in OPENING STOCK LEDGER.
•	Formula column detection: uses osGetFormulaColumnFlags_() to skip formula cells in IN-OUT ENTRY when writing.
•	Marks processed rows with OPENING STOCK UPDATED and locks the cell with sheet protection.

Required columns — PRICE LIST MASTER
Column Header	Used For
Butler Item Code	Item identifier (ITEMCODE in ledger)
Description	Material name
Color	Item colour
BIN CARD NUMBER	Physical storage location reference
SUPPLIER NAME	Vendor name
OPENING STOCK ENTRY	Opening quantity — must be non-empty and not already marked

03_UpdateReceivedLedger.gs
Implements updateReceivedLedger(). Reads all fully-populated rows from RECEIVED ENTRY, appends them to RECEIVED LEDGER and IN-OUT ENTRY, then clears the key input columns on the source sheet so it is ready for the next batch of entries.

Key behaviour
•	Must be run from the RECEIVED ENTRY tab.
•	Validates all required columns are non-empty before processing any row.
•	After writing, clears ITEMCODE, VENDOR PO NO, RECEIVED QTY, RELATED TO CUSTOMER NAME, RELATED TO CUSTOMER PO NO columns from processed rows.
•	Does not deduplicate — all valid rows are pushed every time the function is run.

Required columns — RECEIVED ENTRY
Column Header	Notes
ITEMCODE	
MATERIAL NAME	
COLOR	
BIN CARD NUMBER	
SUPPLIER NAME	
VENDOR PO NO	Purchase order reference
RECEIVED QTY	Quantity received
RELATED TO CUSTOMER NAME	End customer
RELATED TO CUSTOMER PO NO	Customer PO reference
OPENING STOCK ENTRY	Must be non-empty (acts as a guard)
UOM	Unit of measure — not written to ledger, used for validation only

04_UpdateIssuedLedger.gs
Implements updateIssuedLedger(). Reads qualifying rows from ISSUED ENTRY and appends them to ISSUED LEDGER and IN-OUT ENTRY. Clears the key input columns after a successful write.

Key behaviour
•	Must be run from the ISSUED ENTRY tab.
•	Validates all required columns before processing — same all-or-nothing row logic as the received ledger.
•	After writing, clears ITEMCODE, ISSUED QTY, RELATED TO CUSTOMER NAME, RELATED TO CUSTOMER PO NO, RELATED TO STYLE NAME columns.
•	Uses batched contiguous range clearing for performance.

Required columns — ISSUED ENTRY
Column Header	Notes
ITEMCODE	
MATERIAL NAME	
COLOR	
BIN CARD NUMBER	
ISSUED QTY	Quantity issued
RELATED TO CUSTOMER NAME	
RELATED TO CUSTOMER PO NO	
OPENING STOCK ENTRY	Must be non-empty (guard condition — not written to ledger)
UOM	Unit of measure — not written to ledger, used for validation only

4. How to Use
4.1  Updating Opening Stock
1.	Open the Google Sheet and navigate to the PRICE LIST MASTER tab.
2.	Ensure item rows have all required columns filled, including a numeric value in OPENING STOCK ENTRY.
3.	Click the menu: Update Script → Update Opening Stock.
4.	Processed rows will turn red and be locked. A toast notification confirms the count.

4.2  Updating Received Ledger
5.	Navigate to the RECEIVED ENTRY tab.
6.	Fill all required columns for each received item row.
7.	Click: Update Script → Update Received Ledger.
8.	Key columns are cleared from the source sheet after writing.

4.3  Updating Issued Ledger
9.	Navigate to the ISSUED ENTRY tab.
10.	Fill all required columns for each issued item row.
11.	Click: Update Script → Update Issued Ledger.
12.	Key columns are cleared from the source sheet after writing.

5. Shared Helper Functions
Several utility functions are used across the scripts. Due to how Apps Script merges all .gs files into a single global scope, some functions exist in multiple files with identical or near-identical implementations. This is intentional in some cases (prefixed variants for Opening Stock) and a known technical debt in others.

Function	File(s)	Purpose
osCreateHeaderMap_()	02	Builds a header→column-index map for a row array
createHeaderMap_()	03, 04	Same as above — duplicate defined in both files
osGetExistingItemCodes_()	02	Returns a Set of all ITEMCODEs already in the ledger
osAppendRows_()	02	Appends rows to a sheet starting after last occupied row
osAppendRowsSkippingFormulaColumns_()	02	Appends rows but skips any column containing a formula
osGetFormulaColumnFlags_()	02	Returns boolean array — true for columns with formulas in row 2
osGetNextAppendRow_()	02	Finds the first truly empty row after all data
appendRows_()	03, 04	Same as osAppendRows_ — duplicate defined in both files
getNextRow_()	03	Variant of getNextAppendRow_ using getDataRange()
getNextAppendRow_()	04	Variant of getNextAppendRow_ using getLastRow()
clearCells_()	03	Clears specific columns row-by-row (unoptimised)
clearTransferredEntryCellsSafely_()	04	Clears specific columns using batched contiguous ranges
osBuildContiguousRanges_()	02	Groups sorted row numbers into contiguous range objects
buildContiguousRanges_()	04	Identical to above — duplicate
osMarkProcessedOpeningStock_()	02	Marks cells OPENING STOCK UPDATED, red fill, locked

6. Known Issues & Recommendations

ℹ️  Info  The items below are technical observations and suggested improvements. The script is functional as-is; these are recommendations for future maintenance.

6.1  Duplicate Helper Functions (Critical)
createHeaderMap_(), appendRows_(), buildContiguousRanges_() and their variants are defined independently in files 03 and 04. Since all Apps Script files share one global namespace, these duplicate definitions will cause a runtime error when both files are deployed together.

⚠️  Action Required  Recommended fix: Extract all common helpers into a new file 00_Helpers.gs and remove the duplicates from files 03 and 04. The os-prefixed versions in file 02 can be renamed to match.

6.2  clearCells_() in File 03 is Unoptimised
The clearCells_() function in 03_UpdateReceivedLedger clears cells one at a time in a nested loop. For large datasets this can be slow and may hit Apps Script execution time limits. File 04 already uses the correct batched approach (clearTransferredEntryCellsSafely_()).
📝 Suggestion  Recommended fix: Replace clearCells_() in file 03 with the same contiguous-range approach used in file 04.

6.3  No User Confirmation Before Write
All three functions begin writing immediately on click without asking the user to confirm the number of rows to be processed. An accidental double-click or wrong tab could cause duplicate entries in the ledger (particularly for Received/Issued, which do not deduplicate).
📝 Suggestion  Recommended fix: Add a ui.alert() confirmation dialog showing the row count before committing the write.

6.4  OPENING STOCK ENTRY Guard in Issued Entry
The requiredHeaders array in 04_UpdateIssuedLedger includes OPENING STOCK ENTRY as a mandatory field, but this value is never written to ISSUED LEDGER or IN-OUT ENTRY. Its purpose as a non-empty guard is unclear and may confuse future maintainers.
📝 Suggestion  Recommended fix: Either document the intent of this guard clearly in a code comment, or remove it if it is no longer needed.

6.5  getNextRow_() vs getNextAppendRow_() Inconsistency
File 03 uses getDataRange().getValues() to find the last row, while file 04 uses getLastRow() + a manual scan. Both approaches work, but they may behave differently on sheets with trailing empty rows. Consolidating to one approach in 00_Helpers.gs (recommended: the explicit scan in file 04) will ensure consistent behaviour.

7. Deployment Instructions
13.	Open the Google Sheet that contains (or will contain) the inventory tabs.
14.	Go to Extensions → Apps Script.
15.	Create four script files named 01_Menu.gs, 02_OpeningStockLedgerPush.gs, 03_UpdateReceivedLedger.gs, and 04_UpdateIssuedLedger.gs.
16.	Paste the respective script contents into each file.
17.	Save and close the Apps Script editor.
18.	Reload the spreadsheet. The Update Script menu will appear in the toolbar.
19.	On first run, Google will request permission to access the spreadsheet — click Review Permissions and grant access.

8. Required Permissions
Permission	Reason
spreadsheets (read/write)	Reading source sheets, writing to ledgers and IN-OUT ENTRY
spreadsheets (protect ranges)	Locking OPENING STOCK UPDATED cells in PRICE LIST MASTER
ui (alerts, toasts)	Showing error messages and success notifications

9. Change Log
Version	Date	Change
1.0	April 2026	Initial release — Opening Stock, Received Ledger, Issued Ledger automation

This document is auto-generated. For questions contact the script maintainer.
