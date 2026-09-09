# Leave Manager workbook (Excel + VBA)

Firefighters plot leave directly on a Calendar sheet via dropdowns. Balances, deductions and slot bumping are handled by VBA. A Google Sheets + Apps Script version of the same engine is in `google_sheets/`.

Station parameters (name, year, anchor duty date, entitlement, expiry periods, slots per day) live in `config.py` and on the hidden Config sheet. The values in this repository are placeholders.

## Setup

1. Generate the workbook: `pip install openpyxl && python generate_workbook.py`
2. Open the `.xlsx` in Excel and save it as a macro-enabled workbook (`.xlsm`).
3. Import the six modules from `vba/` (VBA editor: File > Import File).
4. Paste the event handlers into the Calendar sheet module and ThisWorkbook:

```vba
' Calendar sheet
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    modCalendarEvents.CacheOldValue Target
End Sub
Private Sub Worksheet_Change(ByVal Target As Range)
    modCalendarEvents.HandleCalendarChange Target
End Sub

' ThisWorkbook
Private Sub Workbook_Open()
    modInit.OnWorkbookOpen
End Sub
```

5. Add personnel on the Personnel sheet, then run `RefreshAll` (Alt+F8 / Option+F8).

## Sheets

- **Calendar**: one row per duty date, slot columns with dropdowns, on-leave and available counts.
- **Personnel**: master list with calculated balances, and below it an append-only credits and debits ledger. Every booking writes DEBIT rows linked to the credit pool they drew from; removing a booking writes a reversing CREDIT rather than deleting anything.
- **Dashboard**: key metrics, a clearing tracker rating how hard it will be for each person to use their remaining leave before they finish service, monthly slot usage, and an expiry watch.
- **Config** and **NameLists** (hidden): parameters and shift-filtered dropdown sources.

## Engine

- Plotting a name deducts the configured cost per shift using soonest-expiry-first across that person's credit pools, mixing leave types if needed, and rejects the booking if the balance is short.
- Clearing a name reverses the deduction and shifts everyone in lower-priority slots up one.
- Swapping a name reverses the old person and deducts the new one, restoring the original if the new person cannot cover it.
- Balances are computed from the ledger by code, never stored as formulas.

## Files

```
generate_workbook.py   builds the .xlsx
config.py              station parameters and public holidays
sheets/                one module per sheet
vba/                   Excel engine (6 modules)
google_sheets/         Apps Script engine (6 scripts) and its own README
```
