// ============================================================
// Dashboard.gs - Commander-friendly dashboard with plain-language
// statuses and forecasts.
// Station Leave Manager (Google Sheets + Apps Script)
// ============================================================

function refreshDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dash = ss.getSheetByName(DASH_SHEET);
  var pSheet = ss.getSheetByName(P_SHEET);
  var calSheet = ss.getSheetByName(CAL_SHEET);
  var today = new Date();

  // --- Read personnel data ---
  var pData = pSheet.getRange(P_DATA_START, 1, P_MAX_ROWS, P_NUM_COLS).getValues();
  var active = [];
  for (var i = 0; i < pData.length; i++) {
    var name = pData[i][P_COL_NAME - 1].toString().trim();
    if (name === '') continue;
    var rank = pData[i][P_COL_RANK - 1].toString().trim();
    var dn = rank !== '' ? rank + ' ' + name : name;
    var status = pData[i][P_COL_STATUS - 1].toString().trim();
    if (status !== 'Active') continue;
    active.push({
      displayName: dn,
      ordDate: pData[i][P_COL_ORD - 1],
      vlEnt:    parseFloat(pData[i][P_COL_VL_ENT - 1]) || 0,
      vlUsed:   parseFloat(pData[i][P_COL_VL_USED - 1]) || 0,
      vlBal:    parseFloat(pData[i][P_COL_VL_BAL - 1]) || 0,
      pholCr:   parseFloat(pData[i][P_COL_PHOL_CR - 1]) || 0,
      pholUsed: parseFloat(pData[i][P_COL_PHOL_USED - 1]) || 0,
      pholBal:  parseFloat(pData[i][P_COL_PHOL_BAL - 1]) || 0,
      oilCr:    parseFloat(pData[i][P_COL_OIL_CR - 1]) || 0,
      oilUsed:  parseFloat(pData[i][P_COL_OIL_USED - 1]) || 0,
      oilBal:   parseFloat(pData[i][P_COL_OIL_BAL - 1]) || 0,
      totalBal: parseFloat(pData[i][P_COL_TOTAL - 1]) || 0
    });
  }

  // --- Compute each person's status & forecast ---
  var statuses = [];
  for (var i = 0; i < active.length; i++) {
    statuses.push(assessPerson_(active[i], today));
  }
  // Sort: most urgent first
  statuses.sort(function(a, b) { return a.priority - b.priority; });

  // --- Monthly outlook ---
  var monthly = computeMonthly_(calSheet);

  // --- Action items ---
  var actions = buildActionItems_(statuses, active, today);

  // --- Overall status ---
  var overall = computeOverall_(statuses, actions);

  // --- Write to Dashboard ---
  clearDashData_(dash);
  writeOverall_(dash, overall);
  writeQuickStats_(dash, active, calSheet, today);
  writePersonnel_(dash, statuses);
  writeMonthly_(dash, monthly);
  writeActions_(dash, actions);
}

// ============================================================
// Assess individual person
// ============================================================

function assessPerson_(p, today) {
  var bal = p.totalBal;
  var result = {
    name: p.displayName,
    balance: bal,
    ordDate: null,
    daysLeft: '',
    status: '',
    statusColor: '',
    forecast: '',
    priority: 99
  };

  // No balance
  if (bal <= 0) {
    result.status = 'All Clear';
    result.statusColor = '#c6efce';
    result.forecast = 'No leave to clear.';
    result.priority = 5;
    return result;
  }

  var hasOrd = (p.ordDate instanceof Date && p.ordDate >= today);

  if (hasOrd) {
    result.ordDate = p.ordDate;
    var daysToOrd = Math.ceil((p.ordDate - today) / 86400000);
    result.daysLeft = daysToOrd + 'd';
    var shiftsNeeded = Math.ceil(bal / SHIFT_LEAVE_COST);
    var dutyDays = countShift1DutyDays_(today, p.ordDate);

    if (dutyDays === 0) {
      result.status = 'Escalate';
      result.statusColor = '#e06666';
      result.forecast = 'No duty days before ORD. Cannot clear ' + bal + ' days of leave. Needs commander action.';
      result.priority = 1;
    } else {
      var ratio = shiftsNeeded / dutyDays;
      if (ratio >= 0.8) {
        result.status = 'Escalate';
        result.statusColor = '#e06666';
        result.forecast = 'Needs ' + shiftsNeeded + ' shifts but only ' + dutyDays + ' duty days left. Unlikely to clear. Recommend extra slots.';
        result.priority = 1;
      } else if (ratio >= 0.6) {
        result.status = 'Plan Now';
        result.statusColor = '#f6b26b';
        result.forecast = 'Tight. Must plot ' + shiftsNeeded + ' of ' + dutyDays + ' remaining duty days. Start immediately.';
        result.priority = 2;
      } else if (ratio >= 0.3) {
        result.status = 'Monitor';
        result.statusColor = '#ffe699';
        result.forecast = 'On pace but don\'t delay. ' + shiftsNeeded + ' shifts needed, ' + dutyDays + ' days available.';
        result.priority = 3;
      } else {
        result.status = 'On Track';
        result.statusColor = '#c6efce';
        result.forecast = 'Comfortable. ' + dutyDays + ' duty days to clear ' + shiftsNeeded + ' shifts. No rush.';
        result.priority = 4;
      }
    }
    return result;
  }

  // No ORD: assess by leave usage and expiry
  var monthsLeft = 12 - today.getMonth();
  var pools = getSortedCreditPools(p.displayName);
  var expiringDays = 0;
  var in30 = new Date(today.getTime() + 30 * 86400000);
  for (var j = 0; j < pools.length; j++) {
    if (pools[j].expiry <= in30) expiringDays += pools[j].remaining;
  }

  if (expiringDays > 0) {
    result.status = 'Plan Now';
    result.statusColor = '#f6b26b';
    result.forecast = Math.round(expiringDays * 10) / 10 + ' day(s) expiring within 30 days. Plot leave now.';
    result.priority = 2;
  } else if (bal > 10 && monthsLeft <= 6) {
    result.status = 'Monitor';
    result.statusColor = '#ffe699';
    result.forecast = 'High balance (' + bal + ' days) with ' + monthsLeft + ' months left in year.';
    result.priority = 3;
  } else {
    result.status = 'On Track';
    result.statusColor = '#c6efce';
    result.forecast = 'Leave balance manageable. No issues.';
    result.priority = 4;
  }
  return result;
}

// ============================================================
// Monthly outlook
// ============================================================

function computeMonthly_(calSheet) {
  var lastRow = calSheet.getLastRow();
  if (lastRow < 3) return [];
  var calData = calSheet.getRange(3, 1, lastRow - 2, CAL_SLOT_LAST).getValues();

  var monthNames = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  var stats = {};

  for (var i = 0; i < calData.length; i++) {
    var d = calData[i][0];
    if (!(d instanceof Date)) continue;
    var m = d.getMonth();
    if (!stats[m]) stats[m] = {dutyDays: 0, used: 0, total: 0};
    stats[m].dutyDays++;
    stats[m].total += SLOTS_PER_DAY;
    for (var c = CAL_SLOT_FIRST - 1; c < CAL_SLOT_LAST; c++) {
      if (calData[i][c] && calData[i][c].toString().trim() !== '') stats[m].used++;
    }
  }

  var result = [];
  var months = Object.keys(stats).map(Number).sort(function(a, b) { return a - b; });
  for (var i = 0; i < months.length; i++) {
    var m = months[i];
    var s = stats[m];
    var remaining = s.total - s.used;
    var pctFree = s.total > 0 ? Math.round(remaining / s.total * 100) : 0;
    var status, color;

    if (pctFree > 80)      { status = 'Wide Open';    color = '#c6efce'; }
    else if (pctFree > 50) { status = 'Available';     color = '#d9ead3'; }
    else if (pctFree > 20) { status = 'Filling Up';    color = '#ffe699'; }
    else if (pctFree > 0)  { status = 'Almost Full';   color = '#f6b26b'; }
    else                   { status = 'Full';           color = '#e06666'; }

    result.push({
      month: monthNames[m] + ' ' + YEAR,
      dutyDays: s.dutyDays,
      used: s.used,
      remaining: remaining,
      status: status,
      color: color
    });
  }
  return result;
}

// ============================================================
// Action items
// ============================================================

function buildActionItems_(statuses, active, today) {
  var items = [];
  var in30 = new Date(today.getTime() + 30 * 86400000);

  for (var i = 0; i < statuses.length; i++) {
    var s = statuses[i];
    if (s.status === 'Escalate') {
      items.push({
        priority: 'URGENT',
        priorityColor: '#e06666',
        person: s.name,
        issue: 'Cannot clear leave before ORD',
        action: 'Approve extra leave slots or discuss ORD extension.'
      });
    } else if (s.status === 'Plan Now') {
      items.push({
        priority: 'HIGH',
        priorityColor: '#f6b26b',
        person: s.name,
        issue: s.ordDate ? 'Limited time before ORD' : 'Leave expiring soon',
        action: s.ordDate ? 'Schedule leave plotting this week.' : 'Use expiring leave within 30 days.'
      });
    }
  }

  // Check for expiring credits across all active personnel
  for (var i = 0; i < active.length; i++) {
    var pools = getSortedCreditPools(active[i].displayName);
    for (var j = 0; j < pools.length; j++) {
      var p = pools[j];
      if (p.expiry <= in30 && p.remaining > 0) {
        var daysUntil = Math.ceil((p.expiry - today) / 86400000);
        // Don't duplicate if already covered by personnel status
        var alreadyCovered = false;
        for (var k = 0; k < items.length; k++) {
          if (items[k].person === active[i].displayName) { alreadyCovered = true; break; }
        }
        if (alreadyCovered) continue;

        var urgency = daysUntil <= 14 ? 'HIGH' : 'MEDIUM';
        var urgColor = daysUntil <= 14 ? '#f6b26b' : '#ffe699';
        items.push({
          priority: urgency,
          priorityColor: urgColor,
          person: active[i].displayName,
          issue: Math.round(p.remaining * 10) / 10 + ' ' + p.leaveType + ' expiring in ' + daysUntil + ' days',
          action: 'Plot leave before ' + fmtDate_(p.expiry) + '.'
        });
        break; // one item per person
      }
    }
  }

  // Sort: URGENT first, then HIGH, MEDIUM
  items.sort(function(a, b) {
    var order = {URGENT: 1, HIGH: 2, MEDIUM: 3, LOW: 4};
    return (order[a.priority] || 5) - (order[b.priority] || 5);
  });

  return items;
}

// ============================================================
// Overall status
// ============================================================

function computeOverall_(statuses, actions) {
  var counts = {Escalate: 0, 'Plan Now': 0, Monitor: 0, 'On Track': 0, 'All Clear': 0};
  for (var i = 0; i < statuses.length; i++) {
    counts[statuses[i].status] = (counts[statuses[i].status] || 0) + 1;
  }

  var urgentActions = 0;
  for (var i = 0; i < actions.length; i++) {
    if (actions[i].priority === 'URGENT') urgentActions++;
  }

  var badge, color, summary;

  if (counts.Escalate > 0) {
    badge = 'ACTION REQUIRED';
    color = '#e06666';
    summary = counts.Escalate + ' personnel cannot clear leave before ORD. Commander action needed. See action items below.';
  } else if (counts['Plan Now'] > 0 || urgentActions > 0) {
    badge = 'ATTENTION NEEDED';
    color = '#f6b26b';
    summary = counts['Plan Now'] + ' personnel need to plot leave urgently. Review action items.';
  } else if (counts.Monitor > 0) {
    badge = 'MONITORING';
    color = '#ffe699';
    summary = 'All manageable, but ' + counts.Monitor + ' personnel should plan leave soon. No immediate risk.';
  } else {
    badge = 'HEALTHY';
    color = '#c6efce';
    summary = 'All personnel on track. No urgent action needed.';
  }

  return {badge: badge, color: color, summary: summary};
}

// ============================================================
// Write functions
// ============================================================

function clearDashData_(dash) {
  // Clear data areas only, keep headers
  dash.getRange(4, 1, 2, 9).clearContent().setBackground(null);  // overall
  dash.getRange(8, 1, 1, 9).clearContent();                       // stats
  dash.getRange(12, 1, 30, 6).clearContent().setBackground(null); // personnel
  dash.getRange(45, 1, 12, 5).clearContent().setBackground(null); // monthly
  dash.getRange(60, 1, 20, 5).clearContent().setBackground(null); // actions
}

function writeOverall_(dash, overall) {
  // Row 4: status badge
  dash.getRange(4, 1, 1, 9).merge()
    .setValue(overall.badge)
    .setBackground(overall.color)
    .setFontSize(20).setFontWeight('bold')
    .setHorizontalAlignment('center');
  // Row 5: summary
  dash.getRange(5, 1, 1, 9).merge()
    .setValue(overall.summary)
    .setFontSize(11).setFontColor('#333333')
    .setWrap(true).setHorizontalAlignment('center');
}

function writeQuickStats_(dash, active, calSheet, today) {
  var count = active.length;
  var sumBal = 0;
  for (var i = 0; i < active.length; i++) sumBal += active[i].totalBal;
  var avgBal = count > 0 ? Math.round(sumBal / count * 10) / 10 : 0;

  // Expiring within 30 days count
  var in30 = new Date(today.getTime() + 30 * 86400000);
  var expiringCount = 0;
  for (var i = 0; i < active.length; i++) {
    var pools = getSortedCreditPools(active[i].displayName);
    for (var j = 0; j < pools.length; j++) {
      if (pools[j].expiry <= in30) { expiringCount++; break; }
    }
  }

  // Slot fill for future dates
  var lastRow = calSheet.getLastRow();
  var totalSlots = 0, filled = 0;
  if (lastRow >= 3) {
    var calData = calSheet.getRange(3, 1, lastRow - 2, CAL_SLOT_LAST).getValues();
    for (var i = 0; i < calData.length; i++) {
      if (!(calData[i][0] instanceof Date) || calData[i][0] < today) continue;
      totalSlots += SLOTS_PER_DAY;
      for (var c = CAL_SLOT_FIRST - 1; c < CAL_SLOT_LAST; c++) {
        if (calData[i][c] && calData[i][c].toString().trim() !== '') filled++;
      }
    }
  }
  var fillPct = totalSlots > 0 ? Math.round(filled / totalSlots * 100) : 0;

  // ORDs within 90 days
  var in90 = new Date(today.getTime() + 90 * 86400000);
  var ordCount = 0;
  for (var i = 0; i < active.length; i++) {
    var ord = active[i].ordDate;
    if (ord instanceof Date && ord >= today && ord <= in90) ordCount++;
  }

  dash.getRange(8, 1, 1, 9).setValues([[
    'Personnel: ' + count,
    '',
    'Avg Balance: ' + avgBal + 'd',
    '',
    'Expiring: ' + expiringCount,
    '',
    'Slots Filled: ' + fillPct + '%',
    '',
    'ORDs (90d): ' + ordCount
  ]]).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground('#f5f5f5');
}

function writePersonnel_(dash, statuses) {
  var max = Math.min(statuses.length, 30);
  for (var i = 0; i < max; i++) {
    var s = statuses[i];
    var row = 12 + i;
    dash.getRange(row, 1, 1, 6).setValues([[
      s.name,
      s.balance + 'd',
      s.ordDate instanceof Date ? s.ordDate : '',
      s.daysLeft,
      s.status,
      s.forecast
    ]]);
    if (s.ordDate instanceof Date) {
      dash.getRange(row, 3).setNumberFormat('dd-MMM-yyyy');
    }
    // Color the status cell
    dash.getRange(row, 5).setBackground(s.statusColor).setFontWeight('bold');
    // Light row tint
    dash.getRange(row, 1, 1, 4).setBackground(s.statusColor + '40'); // won't work as hex+alpha
    // Just tint the whole row very lightly
    var lightBg = s.priority <= 2 ? '#fff2f2' : (s.priority === 3 ? '#fffde7' : null);
    if (lightBg) dash.getRange(row, 1, 1, 4).setBackground(lightBg);
  }
}

function writeMonthly_(dash, monthly) {
  for (var i = 0; i < monthly.length && i < 12; i++) {
    var m = monthly[i];
    var row = 45 + i;
    dash.getRange(row, 1, 1, 5).setValues([[
      m.month, m.dutyDays, m.used, m.remaining, m.status
    ]]);
    dash.getRange(row, 5).setBackground(m.color).setFontWeight('bold');
  }
}

function writeActions_(dash, actions) {
  var max = Math.min(actions.length, 20);
  for (var i = 0; i < max; i++) {
    var a = actions[i];
    var row = 60 + i;
    dash.getRange(row, 1, 1, 5).setValues([[
      a.priority, a.person, a.issue, a.action, ''
    ]]);
    dash.getRange(row, 1).setBackground(a.priorityColor).setFontWeight('bold')
      .setFontColor('white');
  }

  if (max === 0) {
    dash.getRange(60, 1).setValue('No action items. All clear.');
    dash.getRange(60, 1).setFontColor('#666666').setFontStyle('italic');
  }
}
