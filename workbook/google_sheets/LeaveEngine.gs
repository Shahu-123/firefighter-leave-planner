// ============================================================
// LeaveEngine.gs - Credit pools, balance calculation, deductions
// Station Leave Manager (Google Sheets + Apps Script)
// ============================================================

// ============================================================
// Credit Pool Management
// ============================================================

/**
 * Get sorted credit pools (soonest-expiry-first) for a person.
 * Returns array of {id, leaveType, remaining, expiry} objects.
 * Only includes unexpired pools with remaining > 0.
 */
function getSortedCreditPools(displayName, asOfDate) {
  if (!asOfDate) asOfDate = new Date();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(P_SHEET);

  // Batch read the entire credits table
  var lastRow = getLastCreditRow_(sheet);
  if (lastRow < LC_DATA_START) return [];

  var numRows = lastRow - LC_DATA_START + 1;
  var data = sheet.getRange(LC_DATA_START, 1, numRows, LC_NUM_COLS).getValues();

  // Pass 1: gather CREDIT entries
  var credits = []; // {id, leaveType, days, expiry, row}
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (row[LC_COL_NAME - 1].toString().trim() !== displayName) continue;
    if (row[LC_COL_TXN - 1].toString().trim() !== TXN_CREDIT) continue;
    var expiry = row[LC_COL_EXPIRY - 1];
    if (!(expiry instanceof Date)) continue;
    if (expiry < asOfDate) continue; // expired

    var creditId = parseInt(row[LC_COL_ID - 1]);
    if (isNaN(creditId)) continue;

    credits.push({
      id: creditId,
      leaveType: row[LC_COL_LEAVE - 1].toString().trim(),
      days: parseFloat(row[LC_COL_DAYS - 1]) || 0,
      expiry: expiry
    });
  }

  // Pass 2: sum debits per credit source
  var debitsByCreditId = {};
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (row[LC_COL_NAME - 1].toString().trim() !== displayName) continue;
    if (row[LC_COL_TXN - 1].toString().trim() !== TXN_DEBIT) continue;
    var sourceId = parseInt(row[LC_COL_SOURCE - 1]);
    if (isNaN(sourceId)) continue;
    var days = parseFloat(row[LC_COL_DAYS - 1]) || 0;
    debitsByCreditId[sourceId] = (debitsByCreditId[sourceId] || 0) + days;
  }

  // Build pools with remaining balance
  var pools = [];
  for (var i = 0; i < credits.length; i++) {
    var c = credits[i];
    var debited = debitsByCreditId[c.id] || 0;
    var remaining = c.days - debited;
    if (remaining > 0.001) {
      pools.push({
        id: c.id,
        leaveType: c.leaveType,
        remaining: remaining,
        expiry: c.expiry
      });
    }
  }

  // Sort: soonest expiry first, then OIL < PHOL < VL for same date
  pools.sort(function(a, b) {
    if (a.expiry.getTime() !== b.expiry.getTime()) {
      return a.expiry.getTime() - b.expiry.getTime();
    }
    return leavePriority_(a.leaveType) - leavePriority_(b.leaveType);
  });

  return pools;
}

function leavePriority_(lt) {
  switch (lt) {
    case LEAVE_OIL:  return 1;
    case LEAVE_PHOL: return 2;
    case LEAVE_VL:   return 3;
    default:         return 99;
  }
}

// ============================================================
// Balance Queries
// ============================================================

function getBalance(displayName, leaveType) {
  var pools = getSortedCreditPools(displayName);
  var total = 0;
  for (var i = 0; i < pools.length; i++) {
    if (pools[i].leaveType === leaveType) total += pools[i].remaining;
  }
  return total;
}

function getTotalBalance(displayName) {
  var pools = getSortedCreditPools(displayName);
  var total = 0;
  for (var i = 0; i < pools.length; i++) {
    total += pools[i].remaining;
  }
  return total;
}

function getTotalCredits(displayName, leaveType) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(P_SHEET);
  var lastRow = getLastCreditRow_(sheet);
  if (lastRow < LC_DATA_START) return 0;
  var data = sheet.getRange(LC_DATA_START, 1, lastRow - LC_DATA_START + 1, LC_NUM_COLS).getValues();
  var total = 0;
  for (var i = 0; i < data.length; i++) {
    if (data[i][LC_COL_NAME - 1].toString().trim() === displayName &&
        data[i][LC_COL_TXN - 1].toString().trim() === TXN_CREDIT &&
        data[i][LC_COL_LEAVE - 1].toString().trim() === leaveType) {
      total += parseFloat(data[i][LC_COL_DAYS - 1]) || 0;
    }
  }
  return total;
}

function getTotalDebits(displayName, leaveType) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(P_SHEET);
  var lastRow = getLastCreditRow_(sheet);
  if (lastRow < LC_DATA_START) return 0;
  var data = sheet.getRange(LC_DATA_START, 1, lastRow - LC_DATA_START + 1, LC_NUM_COLS).getValues();
  var total = 0;
  for (var i = 0; i < data.length; i++) {
    if (data[i][LC_COL_NAME - 1].toString().trim() === displayName &&
        data[i][LC_COL_TXN - 1].toString().trim() === TXN_DEBIT &&
        data[i][LC_COL_LEAVE - 1].toString().trim() === leaveType) {
      total += parseFloat(data[i][LC_COL_DAYS - 1]) || 0;
    }
  }
  return total;
}

// ============================================================
// Leave Deduction (soonest-expiry-first)
// ============================================================

/**
 * Deduct leave for a booking. Creates DEBIT entries in the credits log.
 * Returns true if successful, false if insufficient balance.
 */
function deductLeaveForBooking(displayName, calDate) {
  var cost = SHIFT_LEAVE_COST;
  var pools = getSortedCreditPools(displayName);

  // Check total available
  var total = 0;
  for (var i = 0; i < pools.length; i++) total += pools[i].remaining;
  if (total < cost - 0.001) return false;

  // Plan deductions
  var plan = []; // {sourceId, leaveType, amount, expiry}
  var remaining = cost;
  for (var i = 0; i < pools.length; i++) {
    if (remaining <= 0.001) break;
    var amt = Math.min(pools[i].remaining, remaining);
    plan.push({
      sourceId: pools[i].id,
      leaveType: pools[i].leaveType,
      amount: amt,
      expiry: pools[i].expiry
    });
    remaining -= amt;
  }

  if (remaining > 0.001) return false;

  // Write DEBIT entries
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(P_SHEET);
  for (var i = 0; i < plan.length; i++) {
    var p = plan[i];
    var newRow = findNextEmptyCreditRow_(sheet);
    var newId = getNextCreditId_(sheet);

    var rowData = [];
    rowData[LC_COL_ID - 1]      = newId;
    rowData[LC_COL_NAME - 1]    = displayName;
    rowData[LC_COL_TXN - 1]     = TXN_DEBIT;
    rowData[LC_COL_LEAVE - 1]   = p.leaveType;
    rowData[LC_COL_DAYS - 1]    = p.amount;
    rowData[LC_COL_DATE - 1]    = new Date();
    rowData[LC_COL_EXPIRY - 1]  = p.expiry;
    rowData[LC_COL_SOURCE - 1]  = p.sourceId;
    rowData[LC_COL_CALREF - 1]  = calDate;
    rowData[LC_COL_REMARKS - 1] = 'Leave ' + Utilities.formatDate(calDate, 'Asia/Singapore', 'dd-MMM-yyyy');

    sheet.getRange(newRow, 1, 1, LC_NUM_COLS).setValues([rowData]);
  }

  return true;
}

// ============================================================
// Reversal: undo deductions for a person on a specific date
// ============================================================

function reverseDeduction(displayName, calDate) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(P_SHEET);
  var lastRow = getLastCreditRow_(sheet);
  if (lastRow < LC_DATA_START) return;

  var data = sheet.getRange(LC_DATA_START, 1, lastRow - LC_DATA_START + 1, LC_NUM_COLS).getValues();

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (row[LC_COL_NAME - 1].toString().trim() !== displayName) continue;
    if (row[LC_COL_TXN - 1].toString().trim() !== TXN_DEBIT) continue;

    var calRef = row[LC_COL_CALREF - 1];
    if (!(calRef instanceof Date)) continue;
    if (!datesEqual_(calRef, calDate)) continue;

    // Create reversal CREDIT
    var newRow = findNextEmptyCreditRow_(sheet);
    var newId = getNextCreditId_(sheet);

    var revData = [];
    revData[LC_COL_ID - 1]      = newId;
    revData[LC_COL_NAME - 1]    = displayName;
    revData[LC_COL_TXN - 1]     = TXN_CREDIT;
    revData[LC_COL_LEAVE - 1]   = row[LC_COL_LEAVE - 1];
    revData[LC_COL_DAYS - 1]    = row[LC_COL_DAYS - 1];
    revData[LC_COL_DATE - 1]    = new Date();
    revData[LC_COL_EXPIRY - 1]  = row[LC_COL_EXPIRY - 1];
    revData[LC_COL_SOURCE - 1]  = '';
    revData[LC_COL_CALREF - 1]  = calDate;
    revData[LC_COL_REMARKS - 1] = 'Reversal: ' + Utilities.formatDate(calDate, 'Asia/Singapore', 'dd-MMM-yyyy');

    sheet.getRange(newRow, 1, 1, LC_NUM_COLS).setValues([revData]);
  }
}

// ============================================================
// Credit Leave
// ============================================================

function creditLeave(displayName, leaveType, days, creditDate, expiryDate, remarks) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(P_SHEET);
  var newRow = findNextEmptyCreditRow_(sheet);
  var newId = getNextCreditId_(sheet);

  var rowData = [];
  rowData[LC_COL_ID - 1]      = newId;
  rowData[LC_COL_NAME - 1]    = displayName;
  rowData[LC_COL_TXN - 1]     = TXN_CREDIT;
  rowData[LC_COL_LEAVE - 1]   = leaveType;
  rowData[LC_COL_DAYS - 1]    = days;
  rowData[LC_COL_DATE - 1]    = creditDate;
  rowData[LC_COL_EXPIRY - 1]  = expiryDate;
  rowData[LC_COL_SOURCE - 1]  = '';
  rowData[LC_COL_CALREF - 1]  = '';
  rowData[LC_COL_REMARKS - 1] = remarks || '';

  sheet.getRange(newRow, 1, 1, LC_NUM_COLS).setValues([rowData]);
}

// ============================================================
// Refresh Balances on Personnel sheet
// ============================================================

function refreshPersonBalance(displayName) {
  var pRow = findPersonRow_(displayName);
  if (pRow === 0) return;

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(P_SHEET);

  var vlUsed   = getTotalDebits(displayName, LEAVE_VL);
  var vlBal    = getBalance(displayName, LEAVE_VL);
  var pholCr   = getTotalCredits(displayName, LEAVE_PHOL);
  var pholUsed = getTotalDebits(displayName, LEAVE_PHOL);
  var pholBal  = getBalance(displayName, LEAVE_PHOL);
  var oilCr    = getTotalCredits(displayName, LEAVE_OIL);
  var oilUsed  = getTotalDebits(displayName, LEAVE_OIL);
  var oilBal   = getBalance(displayName, LEAVE_OIL);
  var totalBal = getTotalBalance(displayName);

  var vals = [[vlUsed, vlBal]];
  sheet.getRange(pRow, P_COL_VL_USED, 1, 2).setValues(vals);

  sheet.getRange(pRow, P_COL_PHOL_CR, 1, 3).setValues([[pholCr, pholUsed, pholBal]]);
  sheet.getRange(pRow, P_COL_OIL_CR, 1, 3).setValues([[oilCr, oilUsed, oilBal]]);
  sheet.getRange(pRow, P_COL_TOTAL).setValue(totalBal);
}

function refreshAllBalances() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(P_SHEET);
  var data = sheet.getRange(P_DATA_START, 1, P_MAX_ROWS, 2).getValues();

  for (var i = 0; i < data.length; i++) {
    var name = data[i][P_COL_NAME - 1].toString().trim();
    var rank = data[i][P_COL_RANK - 1].toString().trim();
    if (name === '') continue;
    var dn = rank !== '' ? rank + ' ' + name : name;
    refreshPersonBalance(dn);
  }
}

// ============================================================
// Pro-rated VL
// ============================================================

function calcProRatedVL(joinDate, ordDate) {
  var startD = joinDate > new Date(YEAR, 0, 1) ? joinDate : new Date(YEAR, 0, 1);
  var endD;
  if (ordDate && ordDate instanceof Date && ordDate.getTime() > 0 &&
      ordDate < new Date(YEAR, 11, 31)) {
    endD = ordDate;
  } else {
    endD = new Date(YEAR, 11, 31);
  }

  var months = (endD.getFullYear() - startD.getFullYear()) * 12 +
               (endD.getMonth() - startD.getMonth()) + 1;
  if (months < 0) months = 0;
  if (months > 12) months = 12;

  return Math.ceil(VL_DAYS_FULL_YEAR * months / 12);
}

// ============================================================
// Personnel lookup helpers
// ============================================================

function findPersonRow_(displayName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(P_SHEET);
  var data = sheet.getRange(P_DATA_START, 1, P_MAX_ROWS, 2).getValues();
  for (var i = 0; i < data.length; i++) {
    var name = data[i][P_COL_NAME - 1].toString().trim();
    var rank = data[i][P_COL_RANK - 1].toString().trim();
    if (name === '') continue;
    var dn = rank !== '' ? rank + ' ' + name : name;
    if (dn === displayName) return P_DATA_START + i;
  }
  return 0;
}

function getDisplayName_(row) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(P_SHEET);
  var name = sheet.getRange(row, P_COL_NAME).getValue().toString().trim();
  var rank = sheet.getRange(row, P_COL_RANK).getValue().toString().trim();
  if (name === '') return '';
  return rank !== '' ? rank + ' ' + name : name;
}

// ============================================================
// Credits table helpers
// ============================================================

function findNextEmptyCreditRow_(sheet) {
  var data = sheet.getRange(LC_DATA_START, LC_COL_ID, LC_MAX_ROWS, 1).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === '' || data[i][0] === null || data[i][0] === undefined) {
      return LC_DATA_START + i;
    }
  }
  return LC_DATA_START + LC_MAX_ROWS; // overflow
}

function getNextCreditId_(sheet) {
  var data = sheet.getRange(LC_DATA_START, LC_COL_ID, LC_MAX_ROWS, 1).getValues();
  var maxId = 0;
  for (var i = 0; i < data.length; i++) {
    var v = parseInt(data[i][0]);
    if (!isNaN(v) && v > maxId) maxId = v;
  }
  return maxId + 1;
}

function getLastCreditRow_(sheet) {
  var data = sheet.getRange(LC_DATA_START, LC_COL_ID, LC_MAX_ROWS, 1).getValues();
  var last = LC_DATA_START - 1;
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] !== '' && data[i][0] !== null && data[i][0] !== undefined) {
      last = LC_DATA_START + i;
    }
  }
  return last;
}
