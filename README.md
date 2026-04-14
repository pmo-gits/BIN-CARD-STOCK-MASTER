📦 Inventory Ledger System — Google Apps Script
A menu-driven Google Apps Script automation for a Google Sheets inventory management system. Pushes stock movements (opening stock, received goods, issued goods) into permanent ledger sheets and a unified IN-OUT ENTRY log — eliminating manual copy-paste and maintaining a clean audit trail.
---
📁 File Structure
File	Function	Description
`01_Menu.gs`	`onOpen()`	Registers the Update Script custom menu in the toolbar
`02_OpeningStockLedgerPush.gs`	`updateOpeningStock()`	Pushes opening stock from PRICE LIST MASTER → ledger + IN-OUT ENTRY
`03_UpdateReceivedLedger.gs`	`updateReceivedLedger()`	Pushes received goods from RECEIVED ENTRY → ledger + IN-OUT ENTRY
`04_UpdateIssuedLedger.gs`	`updateIssuedLedger()`	Pushes issued goods from ISSUED ENTRY → ledger + IN-OUT ENTRY
---
🗂️ Sheet Structure
Sheet Name	Role	Access
`PRICE LIST MASTER`	Master item catalogue with opening stock quantities	Read
`OPENING STOCK LEDGER`	Permanent record of all opening stock entries	Write
`RECEIVED ENTRY`	Staging area for incoming goods	Read + Clear
`RECEIVED LEDGER`	Permanent record of all received goods	Write
`ISSUED ENTRY`	Staging area for issued goods	Read + Clear
`ISSUED LEDGER`	Permanent record of all issued goods	Write
`IN-OUT ENTRY`	Unified chronological log of all movements	Write
---
⚙️ How Each Script Works
`02_OpeningStockLedgerPush.gs` — Update Opening Stock
Must be run from the `PRICE LIST MASTER` tab
Reads rows where all required fields are filled and `OPENING STOCK ENTRY` is not already marked
Skips any `ITEMCODE` already present in `OPENING STOCK LEDGER` (deduplication)
Detects formula columns in `IN-OUT ENTRY` and skips them when writing
After writing, marks the processed cells with `OPENING STOCK UPDATED`, applies red fill, and locks the cell with sheet protection
Required columns in PRICE LIST MASTER:
Column	Maps To
`Butler Item Code`	`ITEMCODE`
`Description`	`MATERIAL NAME`
`Color`	`COLOR`
`BIN CARD NUMBER`	`BIN CARD NUMBER`
`SUPPLIER NAME`	`SUPPLIER NAME`
`OPENING STOCK ENTRY`	Opening quantity (guard + value)
---
`03_UpdateReceivedLedger.gs` — Update Received Ledger
Must be run from the `RECEIVED ENTRY` tab
All required columns must be non-empty for a row to qualify
Appends qualifying rows to `RECEIVED LEDGER` and `IN-OUT ENTRY`
Clears the following columns from the source sheet after writing:
`ITEMCODE`, `VENDOR PO NO`, `RECEIVED QTY`, `RELATED TO CUSTOMER NAME`, `RELATED TO CUSTOMER PO NO`
Required columns in RECEIVED ENTRY:
`ITEMCODE`, `MATERIAL NAME`, `COLOR`, `BIN CARD NUMBER`, `SUPPLIER NAME`, `VENDOR PO NO`, `RECEIVED QTY`, `RELATED TO CUSTOMER NAME`, `RELATED TO CUSTOMER PO NO`, `OPENING STOCK ENTRY`, `UOM`
> `UOM` and `OPENING STOCK ENTRY` are used for validation only — they are not written to the ledger.
---
`04_UpdateIssuedLedger.gs` — Update Issued Ledger
Must be run from the `ISSUED ENTRY` tab
Same all-or-nothing row validation as the received ledger
Appends qualifying rows to `ISSUED LEDGER` and `IN-OUT ENTRY`
Clears the following columns from the source sheet after writing:
`ITEMCODE`, `ISSUED QTY`, `RELATED TO CUSTOMER NAME`, `RELATED TO CUSTOMER PO NO`, `RELATED TO STYLE NAME`
Uses batched contiguous range clearing for performance
Required columns in ISSUED ENTRY:
`ITEMCODE`, `MATERIAL NAME`, `COLOR`, `BIN CARD NUMBER`, `ISSUED QTY`, `RELATED TO CUSTOMER NAME`, `RELATED TO CUSTOMER PO NO`, `OPENING STOCK ENTRY`
> `OPENING STOCK ENTRY` is a guard condition only — it is not written to the ledger.
---
🚀 Deployment
Open the target Google Sheet
Go to Extensions → Apps Script
Create four script files and paste the contents of each `.gs` file
Save and close the editor
Reload the spreadsheet — the Update Script menu will appear
On first run, grant the requested permissions
---
🔐 Required Permissions
Permission	Reason
Spreadsheet read/write	Reading source sheets, writing to ledgers
Protect ranges	Locking `OPENING STOCK UPDATED` cells
UI (alerts, toasts)	Error messages and success notifications
---
🛠️ Helper Functions Reference
Function	File(s)	Purpose

`osCreateHeaderMap_()`	02	Builds header → column-index map
`createHeaderMap_()`	03, 04	Same as above — ⚠️ duplicate
`osGetExistingItemCodes_()`	02	Returns Set of ITEMCODEs already in ledger
`osAppendRows_()`	02	Appends rows after last occupied row
`osAppendRowsSkippingFormulaColumns_()`	02	Appends rows, skips formula columns
`osGetFormulaColumnFlags_()`	02	Boolean array — true if column has formula in row 2
`osGetNextAppendRow_()`	02	Finds first truly empty row after all data
`appendRows_()`	03, 04	Same as osAppendRows_ — ⚠️ duplicate
`getNextRow_()`	03	Next append row via `getDataRange()`
`getNextAppendRow_()`	04	Next append row via `getLastRow()` + scan
`clearCells_()`	03	Clears columns row-by-row (unoptimised)
`clearTransferredEntryCellsSafely_()`	04	Clears columns using batched contiguous ranges
`osBuildContiguousRanges_()`	02	Groups sorted row numbers into range objects
`buildContiguousRanges_()`	04	Identical to above — ⚠️ duplicate
`osMarkProcessedOpeningStock_()`	02	Marks cells updated, red fill, locked
---
⚠️ Known Issues & Recommendations
🔴 Critical — Duplicate Helper Functions
`createHeaderMap_()`, `appendRows_()`, `buildContiguousRanges_()` are defined independently in both files 03 and 04. Since all Apps Script `.gs` files share a single global namespace, these duplicate definitions will cause a runtime conflict when deployed together.
Fix: Extract all common helpers into a new `00_Helpers.gs` file and remove duplicates from files 03 and 04.
---
🟡 `clearCells_()` is Unoptimised (File 03)
Clears cells one at a time in a nested loop. For large datasets this is slow and may hit execution time limits. File 04 already uses the correct batched approach.
Fix: Replace `clearCells_()` in file 03 with the `clearTransferredEntryCellsSafely_()` pattern from file 04.
---
🟡 No Confirmation Dialog Before Write
All three functions write immediately on click. An accidental double-run can cause duplicate entries in the Received/Issued ledgers (which do not deduplicate).
Fix: Add a `ui.alert()` confirmation showing the row count before committing.
---
🟡 `OPENING STOCK ENTRY` Guard in Issued Entry
`OPENING STOCK ENTRY` is listed as a required field in `04_UpdateIssuedLedger` but is never written to any ledger. Its purpose as a guard is unclear.
Fix: Add a code comment explaining the intent, or remove it if no longer needed.
---
🟡 Inconsistent `getNextRow_()` vs `getNextAppendRow_()`
File 03 uses `getDataRange()` to find the last row; file 04 uses `getLastRow()` + a manual scan. Both work, but may behave differently on sheets with trailing empty rows.
Fix: Consolidate to one approach in `00_Helpers.gs` (prefer the explicit scan from file 04).
---
📋 Change Log
Version	Date	Notes
1.0	April 2026	Initial release — Opening Stock, Received Ledger, Issued Ledger automation
---
Internal use only. For questions contact the script maintainer.
