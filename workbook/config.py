"""Station configuration and Singapore public holidays for 2026."""

from datetime import date

STATION_CONFIG = {
    "station_name": "STATION",
    "year": 2026,
    "anchor_date": date(2026, 1, 1),  # set to a date on which anchor_shift is known to be on duty
    "anchor_shift": 1,
    "shifts": [1, 2, 3],
    "slots_l1": 2,
    "slots_l2": 2,
    "slots_kiv": 2,
    "vl_days_full_year": 14,
    "oil_expiry_months": 3,
    "phol_expiry_months": 6,
    "plot_window_months": 1,  # Can plot up to end of next month
    "shift_leave_cost_days": 2,
    "start_month": 4,  # April 2026 onwards
}

# Singapore gazetted public holidays 2026
SG_PUBLIC_HOLIDAYS = [
    (date(2026, 1, 1), "New Year's Day"),
    (date(2026, 1, 29), "Chinese New Year"),
    (date(2026, 1, 30), "Chinese New Year Day 2"),
    (date(2026, 3, 20), "Hari Raya Puasa"),
    (date(2026, 4, 3), "Good Friday"),
    (date(2026, 5, 1), "Labour Day"),
    (date(2026, 5, 16), "Vesak Day"),
    (date(2026, 5, 27), "Hari Raya Haji"),
    (date(2026, 8, 9), "National Day"),
    (date(2026, 10, 25), "Deepavali"),
    (date(2026, 12, 25), "Christmas Day"),
]


def get_shift_for_date(target_date, anchor_date=None, anchor_shift=None):
    """Calculate which shift (1, 2, or 3) is on duty for a given date.

    Uses 24h on / 48h off rotation: shifts cycle 1 -> 2 -> 3 -> 1 ...
    """
    if anchor_date is None:
        anchor_date = STATION_CONFIG["anchor_date"]
    if anchor_shift is None:
        anchor_shift = STATION_CONFIG["anchor_shift"]

    days_diff = (target_date - anchor_date).days
    cycle_pos = days_diff % 3
    if cycle_pos < 0:
        cycle_pos += 3
    # anchor_shift is 1-based; shift order is [1, 2, 3]
    anchor_idx = anchor_shift - 1
    result_idx = (anchor_idx + cycle_pos) % 3
    return result_idx + 1  # back to 1-based


def get_month_names():
    """Return list of (month_number, month_name) for the active period."""
    import calendar
    start = STATION_CONFIG["start_month"]
    return [(m, calendar.month_abbr[m]) for m in range(start, 13)]
