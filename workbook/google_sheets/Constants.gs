// ============================================================
// Constants.gs - Sheet names, column indices, configuration
// Station Leave Manager (Google Sheets + Apps Script)
// One rota only
// ============================================================

// --- Sheet names ---
var CAL_SHEET = 'Calendar';
var P_SHEET   = 'Personnel';
var DASH_SHEET = 'Dashboard';
var CFG_SHEET  = 'Config';
var NL_SHEET   = 'NameLists';

// --- Leave types ---
var LEAVE_VL   = 'VL';
var LEAVE_PHOL = 'PHOL';
var LEAVE_OIL  = 'OIL';

// --- Transaction types ---
var TXN_CREDIT = 'CREDIT';
var TXN_DEBIT  = 'DEBIT';

// --- Calendar columns ---
var CAL_COL_DATE    = 1;   // A
var CAL_COL_DAY     = 2;   // B
var CAL_SLOT_FIRST  = 3;   // C  (L1 #1)
var CAL_SLOT_LAST   = 8;   // H  (KIV #2)
var CAL_COL_ONLEAVE = 9;   // I
var CAL_COL_AVAIL   = 10;  // J
var CAL_COL_REMARKS = 11;  // K
var CAL_HEADER_ROW  = 2;
var CAL_NUM_COLS    = 11;

// --- Personnel columns ---
var P_DATA_START  = 4;
var P_MAX_ROWS    = 50;
var P_COL_NAME    = 1;   // A
var P_COL_RANK    = 2;   // B
var P_COL_JOIN    = 3;   // C
var P_COL_ORD     = 4;   // D
var P_COL_STATUS  = 5;   // E
var P_COL_VL_ENT  = 6;   // F
var P_COL_VL_USED = 7;   // G
var P_COL_VL_BAL  = 8;   // H
var P_COL_PHOL_CR   = 9;  // I
var P_COL_PHOL_USED = 10;  // J
var P_COL_PHOL_BAL  = 11;  // K
var P_COL_OIL_CR    = 12;  // L
var P_COL_OIL_USED  = 13;  // M
var P_COL_OIL_BAL   = 14;  // N
var P_COL_TOTAL     = 15;  // O
var P_NUM_COLS      = 15;

// --- Leave Credits table (on Personnel sheet, below personnel) ---
var LC_SECTION_LABEL_ROW = 56;
var LC_HEADER_ROW = 57;
var LC_DATA_START = 58;
var LC_MAX_ROWS   = 500;
var LC_COL_ID      = 1;   // A
var LC_COL_NAME    = 2;   // B
var LC_COL_TXN     = 3;   // C
var LC_COL_LEAVE   = 4;   // D
var LC_COL_DAYS    = 5;   // E
var LC_COL_DATE    = 6;   // F
var LC_COL_EXPIRY  = 7;   // G
var LC_COL_SOURCE  = 8;   // H
var LC_COL_CALREF  = 9;   // I
var LC_COL_REMARKS = 10;  // J
var LC_NUM_COLS    = 10;

// --- Station defaults ---
var STATION_NAME      = 'STATION';
var YEAR              = 2026;
var ANCHOR_DATE       = new Date(2026, 0, 1); // set to a known duty date
var ANCHOR_SHIFT      = 1;
var ROTA              = 1;
var VL_DAYS_FULL_YEAR = 14;
var SHIFT_LEAVE_COST  = 2;
var PHOL_EXPIRY_MONTHS = 6;
var OIL_EXPIRY_MONTHS  = 3;
var SLOTS_PER_DAY     = 6;
var START_MONTH       = 1;  // January (full year)

// --- SG 2026 Public Holidays ---
var SG_HOLIDAYS = {
  '2026-01-01': "New Year's Day",
  '2026-01-29': 'Chinese New Year',
  '2026-01-30': 'Chinese New Year Day 2',
  '2026-03-20': 'Hari Raya Puasa',
  '2026-04-03': 'Good Friday',
  '2026-05-01': 'Labour Day',
  '2026-05-16': 'Vesak Day',
  '2026-05-27': 'Hari Raya Haji',
  '2026-08-09': 'National Day',
  '2026-10-25': 'Deepavali',
  '2026-12-25': 'Christmas Day'
};

// ============================================================
// Helpers
// ============================================================

/** Format a JS Date as 'yyyy-MM-dd' for holiday lookup */
function formatDateKey_(d) {
  var y = d.getFullYear();
  var m = ('0' + (d.getMonth() + 1)).slice(-2);
  var day = ('0' + d.getDate()).slice(-2);
  return y + '-' + m + '-' + day;
}

/** Check if two dates represent the same calendar day */
function datesEqual_(d1, d2) {
  if (!(d1 instanceof Date) || !(d2 instanceof Date)) return false;
  return d1.getFullYear() === d2.getFullYear() &&
         d1.getMonth() === d2.getMonth() &&
         d1.getDate() === d2.getDate();
}

/** Get all Shift 1 duty dates from Jan 1 to Dec 31 */
function getShift1Dates_() {
  var startDate = new Date(YEAR, 0, 1);  // Jan 1
  var endDate = new Date(YEAR, 11, 31);  // Dec 31

  // Find first Rota date on or after Jan 1
  var daysDiff = Math.round((startDate.getTime() - ANCHOR_DATE.getTime()) / 86400000);
  var mod = ((daysDiff % 3) + 3) % 3; // always positive
  var offset = (mod === 0) ? 0 : (3 - mod);
  var first = new Date(startDate.getTime() + offset * 86400000);

  var dates = [];
  var current = new Date(first.getTime());
  while (current <= endDate) {
    dates.push(new Date(current.getTime()));
    current.setDate(current.getDate() + 3);
  }
  return dates;
}

/** Count Shift 1 duty days between two dates (inclusive) */
function countShift1DutyDays_(fromDate, toDate) {
  var count = 0;
  var current = new Date(ANCHOR_DATE.getTime());
  // Wind forward/backward to first date >= fromDate
  if (current < fromDate) {
    var skip = Math.floor((fromDate - current) / (3 * 86400000));
    current = new Date(current.getTime() + skip * 3 * 86400000);
    if (current < fromDate) current.setDate(current.getDate() + 3);
  }
  while (current <= toDate) {
    if (current >= fromDate) count++;
    current.setDate(current.getDate() + 3);
  }
  return count;
}
