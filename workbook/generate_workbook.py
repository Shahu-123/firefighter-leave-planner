#!/usr/bin/env python3
"""Generate the Station Leave Manager Excel workbook.

Usage:
    python generate_workbook.py [--output FILENAME]

Creates 5 sheets (3 visible + 2 hidden support):
  1. Calendar   - dropdown-based leave plotting (primary interface)
  2. Personnel  - personnel info + leave balances + credits log
  3. Dashboard  - officer analytics & clearing difficulty tracker
  4. Config     - station parameters & SG public holidays (hidden)
  5. NameLists  - shift-filtered name lists for dropdowns (hidden)
"""

import argparse
from openpyxl import Workbook

from sheets.sh_calendar import create_calendar_sheet
from sheets.sh_personnel import create_personnel_sheet
from sheets.sh_dashboard import create_dashboard_sheet
from sheets.sh_config import create_config_sheet
from sheets.sh_namelists import create_namelists_sheet


def generate(output_path: str):
    wb = Workbook()
    default = wb.active
    default.title = "_tmp"

    print("Creating Calendar sheet...")
    create_calendar_sheet(wb)

    print("Creating Personnel sheet...")
    create_personnel_sheet(wb)

    print("Creating Dashboard sheet...")
    create_dashboard_sheet(wb)

    print("Creating Config sheet...")
    create_config_sheet(wb)

    print("Creating NameLists sheet (hidden)...")
    create_namelists_sheet(wb)

    # Remove temp default sheet
    del wb["_tmp"]

    # Calendar is the active sheet
    wb.active = 0

    wb.save(output_path)
    print(f"\nWorkbook saved to: {output_path}")
    print(f"Sheets: {', '.join(wb.sheetnames)}")
    print()
    print("Setup steps:")
    print("  1. Open the .xlsx in Excel")
    print("  2. File > Save As > .xlsm (Macro-Enabled Workbook)")
    print("  3. Alt+F11 > File > Import each .bas file from vba/ folder")
    print("  4. Double-click 'Calendar' sheet in Project Explorer,")
    print("     paste the 3-line event handler (see README)")
    print("  5. Save. Start adding personnel on the Personnel sheet.")


def main():
    parser = argparse.ArgumentParser(description="Generate Station Leave Manager workbook")
    parser.add_argument(
        "--output", "-o",
        default="Station_Leave_Manager_2026.xlsx",
        help="Output filename (default: Station_Leave_Manager_2026.xlsx)",
    )
    args = parser.parse_args()
    generate(args.output)


if __name__ == "__main__":
    main()
