// ============================================================
// Setup.gs - Creates all sheets, formatting, validation
// Station Leave Manager (Google Sheets + Apps Script)
// Run setupWorkbook() once to initialise a blank spreadsheet.
// ============================================================

function setupWorkbook() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  // Guard against accidental re-run
  if (ss.getSheetByName(CAL_SHEET)) {
    var resp = ui.alert(
      'Setup',
      'This workbook already has a Calendar sheet.\nRe-running setup will DELETE all existing data.\n\nContinue?',
      ui.ButtonSet.YES_NO
    );
    if (resp !== ui.Button.YES) return;
    [CAL_SHEET, P_SHEET, DASH_SHEET, CFG_SHEET, NL_SHEET].forEach(function(name) {
      var s = ss.getSheetByName(name);
      if (s) { ss.deleteSheet(s); }
    });
  }

  // Create sheets in order (support sheets first)
  createNameListsSheet_(ss);
  createConfigSheet_(ss);
  createCalendarSheet_(ss);
  createPersonnelSheet_(ss);
  createDashboardSheet_(ss);

  // Hide support sheets
  ss.getSheetByName(CFG_SHEET).hideSheet();
  ss.getSheetByName(NL_SHEET).hideSheet();

  // Delete default Sheet1
  try {
    var s1 = ss.getSheetByName('Sheet1');
    if (s1) ss.deleteSheet(s1);
  } catch (e) { /* ignore */ }

  // Install onEdit trigger
  installTriggers_();

  // Activate Calendar
  ss.setActiveSheet(ss.getSheetByName(CAL_SHEET));
  SpreadsheetApp.flush();

  ui.alert('Setup Complete',
    'Workbook is ready.\n\n' +
    '1. Go to Personnel sheet and add your people (row 4+)\n' +
    '2. Use  Station Leave Manager > Refresh All  to populate dropdowns\n' +
    '3. Go to Calendar and start plotting leave via the dropdowns',
    ui.ButtonSet.OK);
}

// ============================================================
// Trigger installer
// ============================================================

function installTriggers_() {
  // Remove any existing project triggers for this spreadsheet
  ScriptApp.getProjectTriggers().forEach(function(t) {
    ScriptApp.deleteTrigger(t);
  });
  // Installable onEdit (has full auth, can use getUi, longer timeout)
  ScriptApp.newTrigger('onSheetEdit')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
}

// ============================================================
// NameLists sheet (hidden helper for dropdowns)
// ============================================================

function createNameListsSheet_(ss) {
  var sheet = ss.insertSheet(NL_SHEET);
  sheet.getRange('A1').setValue('Display Name').setFontWeight('bold');
  // Reserve 50 rows for names
  sheet.getRange('A2:A' + (P_MAX_ROWS + 1)).setNumberFormat('@'); // plain text
  sheet.setColumnWidth(1, 180);
}

// ============================================================
// Config sheet
// ============================================================

function createConfigSheet_(ss) {
  var sheet = ss.insertSheet(CFG_SHEET);

  // Header
  sheet.getRange('A1').setValue('Station Configuration').setFontSize(12).setFontWeight('bold');
  sheet.getRange('A1:B1').merge().setBackground('#1a237e').setFontColor('white');

  var params = [
    ['Parameter', 'Value'],
    ['StationName', STATION_NAME],
    ['Year', YEAR],
    ['AnchorDate', ANCHOR_DATE],
    ['AnchorShift', ANCHOR_SHIFT],
    ['Rota', ROTA],
    ['VLDaysFullYear', VL_DAYS_FULL_YEAR],
    ['ShiftLeaveCostDays', SHIFT_LEAVE_COST],
    ['PHOLExpiryMonths', PHOL_EXPIRY_MONTHS],
    ['OILExpiryMonths', OIL_EXPIRY_MONTHS],
    ['SlotsPerShiftDay', SLOTS_PER_DAY],
    ['StartMonth', START_MONTH]
  ];
  sheet.getRange(2, 1, params.length, 2).setValues(params);
  sheet.getRange(2, 1, 1, 2).setFontWeight('bold').setBackground('#e8eaf6');
  sheet.getRange(5, 2).setNumberFormat('dd-MMM-yyyy'); // AnchorDate

  // Public holidays table
  var phRow = params.length + 4;
  sheet.getRange(phRow, 1).setValue('Singapore Public Holidays ' + YEAR)
    .setFontSize(11).setFontWeight('bold');
  sheet.getRange(phRow, 1, 1, 2).merge().setBackground('#1a237e').setFontColor('white');

  sheet.getRange(phRow + 1, 1, 1, 2)
    .setValues([['Date', 'Holiday']])
    .setFontWeight('bold').setBackground('#e8eaf6');

  var keys = Object.keys(SG_HOLIDAYS).sort();
  var phData = keys.map(function(k) {
    var parts = k.split('-');
    return [new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2])), SG_HOLIDAYS[k]];
  });
  if (phData.length > 0) {
    sheet.getRange(phRow + 2, 1, phData.length, 2).setValues(phData);
    sheet.getRange(phRow + 2, 1, phData.length, 1).setNumberFormat('dd-MMM-yyyy');
  }

  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 200);
}

// ============================================================
// Calendar sheet
// ============================================================

function createCalendarSheet_(ss) {
  var sheet = ss.insertSheet(CAL_SHEET);
  var dates = getShift1Dates_();

  // Group by month
  var byMonth = {};
  dates.forEach(function(d) {
    var m = d.getMonth();
    if (!byMonth[m]) byMonth[m] = [];
    byMonth[m].push(d);
  });

  var monthNames = ['JANUARY','FEBRUARY','MARCH','APRIL','MAY','JUNE',
    'JULY','AUGUST','SEPTEMBER','OCTOBER','NOVEMBER','DECEMBER'];
  var dayAbbr = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];

  // Build all rows
  var allData = [];
  var rowMeta = []; // track row types for formatting

  // Row 1: title
  allData.push(['STATION Rota \u2014 Leave Calendar ' + YEAR, '', '', '', '', '', '', '', '', '', '']);
  rowMeta.push('title');

  // Row 2: headers
  allData.push(['Date', 'Day', 'L1', 'L1', 'L2', 'L2', 'KIV', 'KIV', 'On Lv', 'Avail', 'Remarks']);
  rowMeta.push('header');

  var sortedMonths = Object.keys(byMonth).map(Number).sort(function(a, b) { return a - b; });

  for (var mi = 0; mi < sortedMonths.length; mi++) {
    var m = sortedMonths[mi];
    // Month separator
    allData.push([monthNames[m] + ' ' + YEAR, '', '', '', '', '', '', '', '', '', '']);
    rowMeta.push('separator');

    var mDates = byMonth[m];
    for (var di = 0; di < mDates.length; di++) {
      var d = mDates[di];
      var dateKey = formatDateKey_(d);
      var isPH = !!SG_HOLIDAYS[dateKey];
      var row = allData.length + 1; // 1-indexed row this item will occupy

      allData.push([
        d,
        dayAbbr[d.getDay()],
        '', '', '', '', '', '',
        '=COUNTA(C' + row + ':H' + row + ')',
        '=COUNTIF(Personnel!E' + P_DATA_START + ':E' + (P_DATA_START + P_MAX_ROWS - 1) + ',"Active")-I' + row,
        isPH ? 'PH: ' + SG_HOLIDAYS[dateKey] : ''
      ]);
      rowMeta.push(isPH ? 'phdata' : 'data');
    }
  }

  // Batch write all data
  sheet.getRange(1, 1, allData.length, CAL_NUM_COLS).setValues(allData);

  // --- Formatting ---

  // Title
  sheet.getRange(1, 1, 1, CAL_NUM_COLS).merge()
    .setFontSize(14).setFontWeight('bold')
    .setBackground('#1a237e').setFontColor('white')
    .setHorizontalAlignment('center');

  // Headers
  sheet.getRange(2, 1, 1, CAL_NUM_COLS)
    .setFontWeight('bold').setBackground('#e8eaf6')
    .setHorizontalAlignment('center');
  // Slot header colors
  sheet.getRange(2, 3, 1, 2).setBackground('#c6efce'); // L1 green
  sheet.getRange(2, 5, 1, 2).setBackground('#ffe699'); // L2 amber
  sheet.getRange(2, 7, 1, 2).setBackground('#f4cccc'); // KIV red

  // Freeze top 2 rows (no column freeze — conflicts with merged title row)
  sheet.setFrozenRows(2);

  // Process each row for formatting
  for (var i = 0; i < rowMeta.length; i++) {
    var r = i + 1; // 1-indexed
    var type = rowMeta[i];

    if (type === 'separator') {
      sheet.getRange(r, 1, 1, CAL_NUM_COLS).merge()
        .setFontWeight('bold').setFontSize(11)
        .setBackground('#c5cae9').setHorizontalAlignment('center');
    }
    else if (type === 'data' || type === 'phdata') {
      // Date format
      sheet.getRange(r, CAL_COL_DATE).setNumberFormat('dd-MMM');
      // Slot cell backgrounds
      sheet.getRange(r, 3, 1, 2).setBackground('#e8f5e9'); // L1 light green
      sheet.getRange(r, 5, 1, 2).setBackground('#fff8e1'); // L2 light amber
      sheet.getRange(r, 7, 1, 2).setBackground('#fce4ec'); // KIV light red
      // Center day, on-lv, avail
      sheet.getRange(r, CAL_COL_DAY).setHorizontalAlignment('center');
      sheet.getRange(r, CAL_COL_ONLEAVE, 1, 2).setHorizontalAlignment('center');

      if (type === 'phdata') {
        sheet.getRange(r, 1, 1, 2).setBackground('#daeef3');
        sheet.getRange(r, CAL_COL_REMARKS).setFontColor('#0070c0').setFontWeight('bold');
      }
    }
  }

  // Data validation: dropdowns for all slot cells referencing NameLists
  var nlSheet = ss.getSheetByName(NL_SHEET);
  var nlRange = nlSheet.getRange('A2:A' + (P_MAX_ROWS + 1));
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(nlRange, true)
    .setAllowInvalid(true)  // show warning but allow (avoids issues with empty list)
    .build();

  for (var i = 0; i < rowMeta.length; i++) {
    if (rowMeta[i] === 'data' || rowMeta[i] === 'phdata') {
      var r = i + 1;
      sheet.getRange(r, CAL_SLOT_FIRST, 1, SLOTS_PER_DAY).setDataValidation(rule);
    }
  }

  // Column widths
  sheet.setColumnWidth(1, 75);   // Date
  sheet.setColumnWidth(2, 40);   // Day
  for (var c = 3; c <= 8; c++) sheet.setColumnWidth(c, 130); // Slots
  sheet.setColumnWidth(9, 50);   // On Lv
  sheet.setColumnWidth(10, 50);  // Avail
  sheet.setColumnWidth(11, 180); // Remarks

  // Light grid borders on data area
  var lastDataRow = allData.length;
  sheet.getRange(2, 1, lastDataRow - 1, CAL_NUM_COLS)
    .setBorder(true, true, true, true, true, true, '#bdbdbd', SpreadsheetApp.BorderStyle.SOLID);
}

// ============================================================
// Personnel sheet
// ============================================================

function createPersonnelSheet_(ss) {
  var sheet = ss.insertSheet(P_SHEET);

  // --- Title ---
  sheet.getRange('A1').setValue('Personnel \u2014 Rota');
  sheet.getRange('A1:O1').merge()
    .setFontSize(14).setFontWeight('bold')
    .setBackground('#1a237e').setFontColor('white')
    .setHorizontalAlignment('center');

  // --- Headers (row 3) ---
  var headers = [
    'Name', 'Rank', 'Join Date', 'ORD Date', 'Status',
    'VL Ent', 'VL Used', 'VL Bal',
    'PHOL Cr', 'PHOL Used', 'PHOL Bal',
    'OIL Cr', 'OIL Used', 'OIL Bal',
    'Total Bal'
  ];
  sheet.getRange(3, 1, 1, P_NUM_COLS).setValues([headers])
    .setFontWeight('bold').setBackground('#e8eaf6')
    .setHorizontalAlignment('center');

  // Sub-header coloring
  sheet.getRange(3, 6, 1, 3).setBackground('#c6efce');   // VL green
  sheet.getRange(3, 9, 1, 3).setBackground('#daeef3');    // PHOL blue
  sheet.getRange(3, 12, 1, 3).setBackground('#ffe699');   // OIL amber
  sheet.getRange(3, 15).setBackground('#d5a6bd');          // Total purple

  // Data area formatting (rows 4-53)
  for (var r = P_DATA_START; r < P_DATA_START + P_MAX_ROWS; r++) {
    // Date columns
    sheet.getRange(r, P_COL_JOIN).setNumberFormat('dd-MMM-yyyy');
    sheet.getRange(r, P_COL_ORD).setNumberFormat('dd-MMM-yyyy');
  }

  // Status data validation
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'ORD', 'Posted Out'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(P_DATA_START, P_COL_STATUS, P_MAX_ROWS, 1).setDataValidation(statusRule);

  // Balance columns: light background to show they're calculated
  sheet.getRange(P_DATA_START, P_COL_VL_USED, P_MAX_ROWS, 2).setBackground('#f0f0f0');  // VL Used + Bal
  sheet.getRange(P_DATA_START, P_COL_PHOL_USED, P_MAX_ROWS, 2).setBackground('#f0f0f0');
  sheet.getRange(P_DATA_START, P_COL_OIL_USED, P_MAX_ROWS, 2).setBackground('#f0f0f0');
  sheet.getRange(P_DATA_START, P_COL_TOTAL, P_MAX_ROWS, 1).setBackground('#f0f0f0');

  // Freeze header rows (no column freeze — conflicts with merged title row)
  sheet.setFrozenRows(3);

  // --- Leave Credits section ---
  sheet.getRange(LC_SECTION_LABEL_ROW, 1).setValue('Leave Credits & Debits Log');
  sheet.getRange(LC_SECTION_LABEL_ROW, 1, 1, LC_NUM_COLS).merge()
    .setFontSize(12).setFontWeight('bold')
    .setBackground('#1a237e').setFontColor('white')
    .setHorizontalAlignment('center');

  var lcHeaders = ['#', 'Name', 'Txn', 'Leave', 'Days', 'Date', 'Expiry', 'Source#', 'Cal Ref', 'Remarks'];
  sheet.getRange(LC_HEADER_ROW, 1, 1, LC_NUM_COLS).setValues([lcHeaders])
    .setFontWeight('bold').setBackground('#e8eaf6')
    .setHorizontalAlignment('center');

  // Date format for credits section
  sheet.getRange(LC_DATA_START, LC_COL_DATE, LC_MAX_ROWS, 1).setNumberFormat('dd-MMM-yyyy');
  sheet.getRange(LC_DATA_START, LC_COL_EXPIRY, LC_MAX_ROWS, 1).setNumberFormat('dd-MMM-yyyy');
  sheet.getRange(LC_DATA_START, LC_COL_CALREF, LC_MAX_ROWS, 1).setNumberFormat('dd-MMM-yyyy');

  // Column widths
  sheet.setColumnWidth(1, 140); // Name
  sheet.setColumnWidth(2, 60);  // Rank
  sheet.setColumnWidth(3, 95);  // Join
  sheet.setColumnWidth(4, 95);  // ORD
  sheet.setColumnWidth(5, 65);  // Status
  for (var c = 6; c <= 15; c++) sheet.setColumnWidth(c, 70);

  // Borders
  sheet.getRange(3, 1, P_MAX_ROWS + 1, P_NUM_COLS)
    .setBorder(true, true, true, true, true, true, '#bdbdbd', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(LC_HEADER_ROW, 1, 2, LC_NUM_COLS)
    .setBorder(true, true, true, true, true, true, '#bdbdbd', SpreadsheetApp.BorderStyle.SOLID);
}

// ============================================================
// Dashboard sheet
// ============================================================

function createDashboardSheet_(ss) {
  var sheet = ss.insertSheet(DASH_SHEET);

  // Title
  sheet.getRange('A1').setValue('STATION Rota \u2014 Dashboard');
  sheet.getRange('A1:F1').merge()
    .setFontSize(14).setFontWeight('bold')
    .setBackground('#1a237e').setFontColor('white')
    .setHorizontalAlignment('center');

  // --- Overall Leave Situation (rows 3-5) ---
  sheet.getRange(3, 1).setValue('OVERALL LEAVE SITUATION')
    .setFontWeight('bold').setFontSize(10).setFontColor('#666666');
  // Row 4: status badge (filled by refreshDashboard)
  sheet.getRange(4, 1, 1, 6).merge()
    .setFontSize(20).setFontWeight('bold').setHorizontalAlignment('center');
  // Row 5: summary text
  sheet.getRange(5, 1, 1, 6).merge()
    .setFontSize(11).setWrap(true).setHorizontalAlignment('center');

  // --- Quick Stats (row 7 label, row 8 values) ---
  sheet.getRange(7, 1).setValue('KEY NUMBERS')
    .setFontWeight('bold').setFontSize(10).setFontColor('#666666');
  // Row 8 filled by refreshDashboard

  // --- Personnel Status (rows 10-41) ---
  sheet.getRange(10, 1).setValue('PERSONNEL STATUS');
  sheet.getRange(10, 1, 1, 6).merge()
    .setFontSize(12).setFontWeight('bold')
    .setBackground('#4a148c').setFontColor('white');

  var pHeaders = ['Name', 'Balance', 'ORD Date', 'Days Left', 'Status', 'Forecast'];
  sheet.getRange(11, 1, 1, 6).setValues([pHeaders])
    .setFontWeight('bold').setBackground('#e8eaf6').setHorizontalAlignment('center');
  sheet.getRange(12, 1, 30, 6)
    .setBorder(true, true, true, true, true, true, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);

  // --- Monthly Outlook (rows 43-56) ---
  sheet.getRange(43, 1).setValue('MONTHLY OUTLOOK');
  sheet.getRange(43, 1, 1, 5).merge()
    .setFontSize(12).setFontWeight('bold')
    .setBackground('#1b5e20').setFontColor('white');

  var mHeaders = ['Month', 'Duty Days', 'Slots Used', 'Remaining', 'Status'];
  sheet.getRange(44, 1, 1, 5).setValues([mHeaders])
    .setFontWeight('bold').setBackground('#e8eaf6').setHorizontalAlignment('center');
  sheet.getRange(45, 1, 12, 5)
    .setBorder(true, true, true, true, true, true, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);

  // --- Action Items (rows 58-79) ---
  sheet.getRange(58, 1).setValue('ACTION ITEMS');
  sheet.getRange(58, 1, 1, 5).merge()
    .setFontSize(12).setFontWeight('bold')
    .setBackground('#b71c1c').setFontColor('white');

  var aHeaders = ['Priority', 'Person', 'Issue', 'Action Needed', ''];
  sheet.getRange(59, 1, 1, 5).setValues([aHeaders])
    .setFontWeight('bold').setBackground('#e8eaf6').setHorizontalAlignment('center');
  sheet.getRange(60, 1, 20, 5)
    .setBorder(true, true, true, true, true, true, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);

  // Column widths
  sheet.setColumnWidth(1, 150);  // Name / Month / Priority
  sheet.setColumnWidth(2, 80);   // Balance / Duty Days
  sheet.setColumnWidth(3, 95);   // ORD Date / Used
  sheet.setColumnWidth(4, 80);   // Days Left / Remaining
  sheet.setColumnWidth(5, 100);  // Status
  sheet.setColumnWidth(6, 350);  // Forecast (wide for text)
}
