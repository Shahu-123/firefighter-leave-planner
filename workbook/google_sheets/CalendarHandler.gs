// ============================================================
// CalendarHandler.gs - Handles edits on the Calendar sheet
// Station Leave Manager (Google Sheets + Apps Script)
//
// Installed as an installable onEdit trigger by setupWorkbook().
// When a name is selected/removed/swapped in a slot cell,
// this handles leave deduction, reversal, and slot bumping.
// ============================================================

/**
 * Installable onEdit trigger (set up by setupWorkbook).
 * Has full authorization: can use getUi(), PropertiesService, etc.
 */
function onSheetEdit(e) {
  // Prevent recursion from programmatic cell changes
  var props = PropertiesService.getScriptProperties();
  var lockUntil = parseInt(props.getProperty('EDIT_LOCK') || '0');
  if (Date.now() < lockUntil) return;

  var sheet = e.range.getSheet();
  if (sheet.getName() !== CAL_SHEET) return;

  // Only process single-cell edits in slot columns
  if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;
  var row = e.range.getRow();
  var col = e.range.getColumn();
  if (col < CAL_SLOT_FIRST || col > CAL_SLOT_LAST) return;

  // Check this is a data row (column A has a Date)
  var dateVal = sheet.getRange(row, CAL_COL_DATE).getValue();
  if (!(dateVal instanceof Date)) return;

  var newValue = (e.value === undefined || e.value === null) ? '' : e.value.toString().trim();
  var oldValue = (e.oldValue === undefined || e.oldValue === null) ? '' : e.oldValue.toString().trim();
  if (newValue === oldValue) return;

  // Lock edits for 10s to prevent trigger recursion
  props.setProperty('EDIT_LOCK', (Date.now() + 10000).toString());

  try {
    var calDate = dateVal;
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // --- CASE 1: Name removed (cell cleared) ---
    if (newValue === '' && oldValue !== '') {
      reverseDeduction(oldValue, calDate);
      bumpSlotsUp_(sheet, row, col);
      refreshPersonBalance(oldValue);
      ss.toast(oldValue + ' removed from ' + fmtDate_(calDate), 'Leave Reversed', 4);
    }

    // --- CASE 2: Name added ---
    else if (newValue !== '' && oldValue === '') {
      if (isAlreadyOnDate_(sheet, row, col, newValue)) {
        ss.toast(newValue + ' is already on leave on this date.', 'Duplicate', 5);
        e.range.clearContent();
        return;
      }
      if (!deductLeaveForBooking(newValue, calDate)) {
        ss.toast('Insufficient balance for ' + newValue + '. Need ' + SHIFT_LEAVE_COST + ' days.',
                 'Cannot Book', 5);
        e.range.clearContent();
        return;
      }
      refreshPersonBalance(newValue);
      ss.toast(newValue + ' booked on ' + fmtDate_(calDate), 'Leave Booked', 4);
    }

    // --- CASE 3: Name swapped ---
    else if (newValue !== '' && oldValue !== '' && newValue !== oldValue) {
      // Reverse old person first
      reverseDeduction(oldValue, calDate);

      if (isAlreadyOnDate_(sheet, row, col, newValue)) {
        // Undo reversal - re-deduct for old person
        deductLeaveForBooking(oldValue, calDate);
        e.range.setValue(oldValue);
        refreshPersonBalance(oldValue);
        ss.toast(newValue + ' is already on leave on this date.', 'Duplicate', 5);
        return;
      }

      if (!deductLeaveForBooking(newValue, calDate)) {
        // Undo reversal
        deductLeaveForBooking(oldValue, calDate);
        e.range.setValue(oldValue);
        refreshPersonBalance(oldValue);
        ss.toast('Insufficient balance for ' + newValue + '.', 'Cannot Book', 5);
        return;
      }

      refreshPersonBalance(oldValue);
      refreshPersonBalance(newValue);
      ss.toast('Swapped: ' + oldValue + ' \u2192 ' + newValue, 'Slot Updated', 4);
    }

  } finally {
    // Release lock
    PropertiesService.getScriptProperties().deleteProperty('EDIT_LOCK');
  }
}

// ============================================================
// Helpers
// ============================================================

/** Check if a name already appears in another slot on the same row */
function isAlreadyOnDate_(sheet, row, excludeCol, displayName) {
  var values = sheet.getRange(row, CAL_SLOT_FIRST, 1, SLOTS_PER_DAY).getValues()[0];
  for (var i = 0; i < values.length; i++) {
    var c = CAL_SLOT_FIRST + i;
    if (c === excludeCol) continue;
    if (values[i] && values[i].toString().trim() === displayName) return true;
  }
  return false;
}

/** Bump slots left: shift everyone right of the cleared column one position left */
function bumpSlotsUp_(sheet, row, clearedCol) {
  var numCols = CAL_SLOT_LAST - clearedCol + 1;
  if (numCols <= 1) {
    // Last slot was cleared, nothing to bump
    return;
  }
  var values = sheet.getRange(row, clearedCol, 1, numCols).getValues()[0];
  // Shift left: remove first, append empty
  values.shift();
  values.push('');
  sheet.getRange(row, clearedCol, 1, numCols).setValues([values]);
}

/** Format date for display in toasts */
function fmtDate_(d) {
  var months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return d.getDate() + '-' + months[d.getMonth()];
}
