// ============================================================
// Menu.gs - Custom menu, refresh functions, VL auto-credit
// Station Leave Manager (Google Sheets + Apps Script)
// ============================================================

/**
 * Simple trigger: creates the custom menu on spreadsheet open.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Station Leave Manager')
    .addItem('Refresh All', 'refreshAll')
    .addSeparator()
    .addItem('Refresh Balances', 'refreshAllBalances')
    .addItem('Refresh Name Lists', 'refreshNameLists')
    .addItem('Refresh Dashboard', 'refreshDashboard')
    .addSeparator()
    .addItem('Auto-Credit VL', 'checkAndCreditVL')
    .addItem('Credit PHOL for Past PHs', 'autoCreditPHOL')
    .addSeparator()
    .addItem('Re-run Setup', 'setupWorkbook')
    .addToUi();
}

/**
 * Full refresh: name lists, VL credit, balances, dashboard.
 */
function refreshAll() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Refreshing...', 'Station Leave Manager', 2);

  refreshNameLists();
  checkAndCreditVL();
  autoCreditPHOL();
  refreshAllBalances();
  refreshDashboard();

  ss.toast('All data refreshed.', 'Station Leave Manager', 4);
}

// ============================================================
// Refresh NameLists (dropdown source for Calendar)
// ============================================================

function refreshNameLists() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pSheet = ss.getSheetByName(P_SHEET);
  var nlSheet = ss.getSheetByName(NL_SHEET);

  // Clear existing names (keep header in A1)
  nlSheet.getRange('A2:A' + (P_MAX_ROWS + 1)).clearContent();

  // Read personnel
  var pData = pSheet.getRange(P_DATA_START, 1, P_MAX_ROWS, P_COL_STATUS).getValues();

  var names = [];
  for (var i = 0; i < pData.length; i++) {
    var name = pData[i][P_COL_NAME - 1].toString().trim();
    var rank = pData[i][P_COL_RANK - 1].toString().trim();
    var status = pData[i][P_COL_STATUS - 1].toString().trim();
    if (name === '' || status !== 'Active') continue;

    var dn = rank !== '' ? rank + ' ' + name : name;
    names.push([dn]);
  }

  // Sort alphabetically
  names.sort(function(a, b) { return a[0].localeCompare(b[0]); });

  if (names.length > 0) {
    nlSheet.getRange(2, 1, names.length, 1).setValues(names);
  }
}

// ============================================================
// Auto-credit VL for personnel who don't have it yet
// ============================================================

function checkAndCreditVL() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pSheet = ss.getSheetByName(P_SHEET);
  var pData = pSheet.getRange(P_DATA_START, 1, P_MAX_ROWS, P_NUM_COLS).getValues();

  // Read credits table to check existing VL credits
  var lastRow = getLastCreditRow_(pSheet);
  var creditData = [];
  if (lastRow >= LC_DATA_START) {
    creditData = pSheet.getRange(LC_DATA_START, 1, lastRow - LC_DATA_START + 1, LC_NUM_COLS).getValues();
  }

  // Build set of people who already have VL credit
  var hasVL = {};
  for (var i = 0; i < creditData.length; i++) {
    var row = creditData[i];
    if (row[LC_COL_TXN - 1].toString().trim() === TXN_CREDIT &&
        row[LC_COL_LEAVE - 1].toString().trim() === LEAVE_VL) {
      hasVL[row[LC_COL_NAME - 1].toString().trim()] = true;
    }
  }

  for (var i = 0; i < pData.length; i++) {
    var name = pData[i][P_COL_NAME - 1].toString().trim();
    var rank = pData[i][P_COL_RANK - 1].toString().trim();
    var status = pData[i][P_COL_STATUS - 1].toString().trim();
    if (name === '' || status !== 'Active') continue;

    var dn = rank !== '' ? rank + ' ' + name : name;
    if (hasVL[dn]) continue;

    var vlEnt = parseFloat(pData[i][P_COL_VL_ENT - 1]);
    if (isNaN(vlEnt) || vlEnt <= 0) continue;

    // Get join date (or start of year)
    var joinDate = pData[i][P_COL_JOIN - 1];
    if (!(joinDate instanceof Date)) joinDate = new Date(YEAR, 0, 1);

    var expiryDate = new Date(YEAR, 11, 31);
    creditLeave(dn, LEAVE_VL, vlEnt, joinDate, expiryDate, 'Annual VL credit');
  }
}

// ============================================================
// Auto-credit PHOL for all past SG public holidays during
// each person's service period (join date to ORD/year-end).
// Not shift-dependent — every PH earns 1 PHOL.
// ============================================================

function autoCreditPHOL() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pSheet = ss.getSheetByName(P_SHEET);
  var today = new Date();

  // Read credits table
  var lastRow = getLastCreditRow_(pSheet);
  var creditData = [];
  if (lastRow >= LC_DATA_START) {
    creditData = pSheet.getRange(LC_DATA_START, 1, lastRow - LC_DATA_START + 1, LC_NUM_COLS).getValues();
  }

  // Gather all past PH dates (no shift check)
  var phDates = [];
  var keys = Object.keys(SG_HOLIDAYS);
  for (var i = 0; i < keys.length; i++) {
    var parts = keys[i].split('-');
    var phDate = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
    if (phDate > today) continue; // only past PHs
    phDates.push({date: phDate, name: SG_HOLIDAYS[keys[i]]});
  }

  if (phDates.length === 0) {
    ss.toast('No past public holidays found.', 'PHOL', 4);
    return;
  }

  // Get active personnel
  var pData = pSheet.getRange(P_DATA_START, 1, P_MAX_ROWS, P_NUM_COLS).getValues();
  var credited = 0;

  for (var pi = 0; pi < pData.length; pi++) {
    var name = pData[pi][P_COL_NAME - 1].toString().trim();
    var rank = pData[pi][P_COL_RANK - 1].toString().trim();
    var status = pData[pi][P_COL_STATUS - 1].toString().trim();
    if (name === '' || status !== 'Active') continue;
    var dn = rank !== '' ? rank + ' ' + name : name;

    // Service period: join date to ORD date (or Dec 31)
    var joinDate = pData[pi][P_COL_JOIN - 1];
    if (!(joinDate instanceof Date)) joinDate = new Date(YEAR, 0, 1);
    var ordDate = pData[pi][P_COL_ORD - 1];
    var endDate = (ordDate instanceof Date && ordDate.getTime() > 0) ? ordDate : new Date(YEAR, 11, 31);

    for (var phi = 0; phi < phDates.length; phi++) {
      var ph = phDates[phi];

      // PH must fall within person's service period
      if (ph.date < joinDate || ph.date > endDate) continue;

      // Check if already credited for this PH
      var alreadyCredited = false;
      for (var ci = 0; ci < creditData.length; ci++) {
        var cr = creditData[ci];
        if (cr[LC_COL_NAME - 1].toString().trim() === dn &&
            cr[LC_COL_TXN - 1].toString().trim() === TXN_CREDIT &&
            cr[LC_COL_LEAVE - 1].toString().trim() === LEAVE_PHOL &&
            cr[LC_COL_REMARKS - 1].toString().indexOf(ph.name) !== -1) {
          alreadyCredited = true;
          break;
        }
      }
      if (alreadyCredited) continue;

      // Credit 1 PHOL per PH, expires PHOL_EXPIRY_MONTHS after PH date
      var expiry = new Date(ph.date.getTime());
      expiry.setMonth(expiry.getMonth() + PHOL_EXPIRY_MONTHS);

      creditLeave(dn, LEAVE_PHOL, 1, ph.date, expiry, 'PHOL: ' + ph.name);
      credited++;

      // Re-read credit data to prevent duplicate writes
      lastRow = getLastCreditRow_(pSheet);
      if (lastRow >= LC_DATA_START) {
        creditData = pSheet.getRange(LC_DATA_START, 1, lastRow - LC_DATA_START + 1, LC_NUM_COLS).getValues();
      }
    }
  }

  ss.toast('Credited ' + credited + ' PHOL entries.', 'PHOL Auto-Credit', 5);
}
