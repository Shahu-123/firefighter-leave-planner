"""Generate the Dashboard sheet for officer analytics.

Sections:
A. ORD Leave Clearing Tracker (difficulty indicators)
B. Monthly Slot Overview (rota management view)
C. Leave Expiry Watch
D. Key Metrics summary
"""

from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.workbook.defined_name import DefinedName


TITLE_FONT = Font(bold=True, size=16)
SECTION_FONT = Font(bold=True, size=13, color="FFFFFF")
SECTION_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
HEADER_FONT = Font(bold=True, size=10, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
METRIC_LABEL = Font(bold=True, size=11)
METRIC_VALUE = Font(bold=True, size=18, color="2F5496")
THIN_BORDER = Border(
    left=Side(style="thin", color="B4B4B4"),
    right=Side(style="thin", color="B4B4B4"),
    top=Side(style="thin", color="B4B4B4"),
    bottom=Side(style="thin", color="B4B4B4"),
)
CARD_FILL = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
GREEN_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
AMBER_FILL = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
RED_FILL = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")


def _section_header(ws, row, title, end_col):
    ws.cell(row=row, column=1, value=title).font = SECTION_FONT
    for c in range(1, end_col + 1):
        ws.cell(row=row, column=c).fill = SECTION_FILL
    ws.merge_cells(f"A{row}:{chr(64 + end_col)}{row}")
    ws.row_dimensions[row].height = 22


def _col_headers(ws, row, headers, fills=None, start_col=1):
    for i, h in enumerate(headers):
        c = start_col + i
        cell = ws.cell(row=row, column=c, value=h)
        cell.font = HEADER_FONT
        cell.fill = fills[i] if fills else HEADER_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = THIN_BORDER


def _data_area(ws, start_row, num_rows, num_cols, start_col=1):
    for r in range(start_row, start_row + num_rows):
        for c in range(start_col, start_col + num_cols):
            ws.cell(row=r, column=c).border = THIN_BORDER


def create_dashboard_sheet(wb):
    ws = wb.create_sheet("Dashboard")

    # Title
    ws.cell(row=1, column=1, value="STATION Leave Dashboard").font = TITLE_FONT
    ws.merge_cells("A1:K1")
    ws.cell(row=2, column=1,
            value='Data auto-refreshes when you plot/remove leave. Manual refresh: Alt+F8 > RefreshDashboard'
    ).font = Font(italic=True, color="666666", size=9)
    ws.merge_cells("A2:K2")

    # =============================================
    # SECTION A: Key Metrics Cards (row 4)
    # =============================================
    metrics_row = 4
    _section_header(ws, metrics_row, "Key Metrics", 11)
    metrics_row += 1

    # Metric cards: 5 cards across
    card_labels = [
        "Active Personnel", "Avg Leave Balance",
        "Expiring < 30 Days", "Slots Filled (Month)",
        "Upcoming ORDs",
    ]
    card_col_starts = [1, 3, 5, 7, 9]

    for i, (label, cs) in enumerate(zip(card_labels, card_col_starts)):
        # Label
        ws.cell(row=metrics_row, column=cs, value=label).font = Font(bold=True, size=9, color="666666")
        ws.merge_cells(start_row=metrics_row, start_column=cs, end_row=metrics_row, end_column=cs + 1)
        ws.cell(row=metrics_row, column=cs).alignment = Alignment(horizontal="center")
        # Value cell
        val_cell = ws.cell(row=metrics_row + 1, column=cs, value="-")
        val_cell.font = METRIC_VALUE
        val_cell.alignment = Alignment(horizontal="center")
        val_cell.fill = CARD_FILL
        ws.merge_cells(start_row=metrics_row + 1, start_column=cs, end_row=metrics_row + 1, end_column=cs + 1)
        for c in [cs, cs + 1]:
            ws.cell(row=metrics_row + 1, column=c).border = THIN_BORDER

    def _def(name, ref):
        wb.defined_names.add(DefinedName(name, attr_text=ref))

    _def("dash_MetricActivePersonnel", f"Dashboard!$A${metrics_row + 1}")
    _def("dash_MetricAvgBalance", f"Dashboard!$C${metrics_row + 1}")
    _def("dash_MetricExpiring30", f"Dashboard!$E${metrics_row + 1}")
    _def("dash_MetricSlotsFilled", f"Dashboard!$G${metrics_row + 1}")
    _def("dash_MetricUpcomingORDs", f"Dashboard!$I${metrics_row + 1}")

    # =============================================
    # SECTION B: ORD Leave Clearing Tracker (rows 8-45)
    # =============================================
    ord_start = metrics_row + 4
    _section_header(ws, ord_start, "ORD Leave Clearing Tracker", 11)

    ord_headers = [
        "Name", "Shift", "ORD Date", "Days to ORD",
        "Leave Bal", "Shifts to Clear",
        "Shifts Avail", "Open Slots Avail",
        "Difficulty", "Status",
    ]
    _col_headers(ws, ord_start + 1, ord_headers)
    _data_area(ws, ord_start + 2, 30, 10)
    _def("dash_ORDStart", f"Dashboard!$A${ord_start + 2}")

    # Difficulty explanation
    explain_row = ord_start + 33
    ws.cell(row=explain_row, column=1,
            value="Difficulty = Shifts to Clear / Open Slots Available. "
                  "Status: Easy (<30%) | Moderate (30-60%) | Hard (60-80%) | Critical (>80%)"
    ).font = Font(italic=True, color="888888", size=9)
    ws.merge_cells(f"A{explain_row}:J{explain_row}")

    # =============================================
    # SECTION C: Monthly Slot Overview (right side, rows 8-20)
    # =============================================
    monthly_start = ord_start
    monthly_col = 12  # Column L
    ws.cell(row=monthly_start, column=monthly_col, value="Monthly Rota Overview").font = SECTION_FONT
    for c in range(monthly_col, monthly_col + 7):
        ws.cell(row=monthly_start, column=c).fill = SECTION_FILL
    ws.merge_cells(start_row=monthly_start, start_column=monthly_col,
                   end_row=monthly_start, end_column=monthly_col + 6)

    monthly_headers = ["Month", "S1 Slots", "S2 Slots", "S3 Slots", "Total", "Avg Avail", "Fill %"]
    _col_headers(ws, monthly_start + 1, monthly_headers, start_col=monthly_col)
    _data_area(ws, monthly_start + 2, 9, 7, start_col=monthly_col)
    _def("dash_MonthlyStart", f"Dashboard!$L${monthly_start + 2}")

    # =============================================
    # SECTION D: Leave Expiry Watch (below monthly, right side)
    # =============================================
    expiry_start = monthly_start + 13
    ws.cell(row=expiry_start, column=monthly_col, value="Leave Expiry Watch").font = SECTION_FONT
    for c in range(monthly_col, monthly_col + 7):
        ws.cell(row=expiry_start, column=c).fill = SECTION_FILL
    ws.merge_cells(start_row=expiry_start, start_column=monthly_col,
                   end_row=expiry_start, end_column=monthly_col + 6)

    expiry_headers = ["Name", "Type", "Days", "Expiry", "Days Left", "Urgency"]
    _col_headers(ws, expiry_start + 1, expiry_headers, start_col=monthly_col)
    _data_area(ws, expiry_start + 2, 20, 6, start_col=monthly_col)
    _def("dash_ExpiryStart", f"Dashboard!$L${expiry_start + 2}")

    # =============================================
    # SECTION E: Per-Shift Utilization (below ORD tracker)
    # =============================================
    util_start = explain_row + 2
    _section_header(ws, util_start, "Leave Utilization by Shift", 6)

    util_headers = ["Metric", "Shift 1", "Shift 2", "Shift 3", "Overall"]
    _col_headers(ws, util_start + 1, util_headers)
    _data_area(ws, util_start + 2, 6, 5)
    _def("dash_UtilStart", f"Dashboard!$A${util_start + 2}")

    # Pre-fill metric labels
    util_metrics = [
        "VL Utilization %", "PHOL Utilization %",
        "OIL Utilization %", "Avg Balance / Person",
        "Active Personnel", "ORDs in Next 30 Days",
    ]
    for i, m in enumerate(util_metrics):
        ws.cell(row=util_start + 2 + i, column=1, value=m).font = Font(bold=True, size=10)

    # Column widths
    col_widths = {
        "A": 18, "B": 7, "C": 12, "D": 11, "E": 9, "F": 13,
        "G": 12, "H": 14, "I": 10, "J": 10, "K": 4,
        "L": 10, "M": 9, "N": 9, "O": 9, "P": 8, "Q": 10, "R": 8,
    }
    for letter, w in col_widths.items():
        ws.column_dimensions[letter].width = w

    # Freeze
    ws.freeze_panes = "A4"

    return ws
