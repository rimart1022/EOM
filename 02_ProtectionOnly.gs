/****************************************************
 V16 FAST PROTECTION FUNCTIONS ONLY
 - Protections do NOT sync staff permissions.
 - Use one sheet protection per sheet.
 - Reuses existing Carlisle protection instead of deleting/recreating.
 - Weekly can be run one sheet at a time to avoid timeout.
 - Does not write into merged cells.
****************************************************/

function protectionGroups_() {
  return {
    PURCHASES: ['PURCHASES'],
    DAILY_SALES: ['DAILY SALES'],
    DAILY_SALES_BREAKDOWN: ['DAILY SALES BREAKDOWN'],
    EXPENSES: ['EXPENSES'],
    STOCK_MOVEMENT: ['STOCK MOVEMENT APPROVAL LOG'],
    CS_SHEETS: ['CS MINI-MART','CS LAUNDRY','CS BAR','CS RESTAURANT','CS STORE','CS KITCHEN'],
    WEEKLY_MR: ['M.R MINI-MART','M.R BUSH BAR','M.R KITCHEN','2 M.R MINI-MART','2 M.R BUSH BAR','2 M.R KITCHEN','3 M.R MINI-MART','3 M.R BUSH BAR','3 M.R KITCHEN','4 M.R MINI-MART','4 M.R BUSH BAR','4 M.R KITCHEN','5 M.R MINI-MART','5 M.R BUSH BAR','5 M.R KITCHEN'],
    MR_KITCHEN_U: ['M.R KITCHEN U'],
    ADMIN: ['MASTER_PRICELIST','MASTER PRICE LIST','SYSTEM_ACCESS','SYSTEM_SETTINGS','SYSTEM_LOGS','EOM EDIT LOG','STOCK CHANGE LOG','STOCK AUDIT SUMMARY','AUDIT CHECK WEEK 1','AUDIT CHECK WEEK 2','AUDIT CHECK WEEK 3','AUDIT CHECK WEEK 4','AUDIT CHECK WEEK 5','ISSUED STOCK REPORT','DAMAGE REPORT','STAFF LIABILITY REPORT'],
    SYSTEM: ['SYSTEM_ACCESS','SYSTEM_SETTINGS','SYSTEM_LOGS'],
    LOGS: ['SYSTEM_LOGS','EOM EDIT LOG','STOCK CHANGE LOG'],
    REPORTS: ['STOCK AUDIT SUMMARY','AUDIT CHECK WEEK 1','AUDIT CHECK WEEK 2','AUDIT CHECK WEEK 3','AUDIT CHECK WEEK 4','AUDIT CHECK WEEK 5','ISSUED STOCK REPORT','DAMAGE REPORT','STAFF LIABILITY REPORT'],
    MASTER_PRICE: ['MASTER_PRICELIST','MASTER PRICE LIST']
  };
}

function protectionLock_(callback) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) throw new Error('Another Carlisle protection task is already running. Try again shortly.');
  try { return callback(); } finally { lock.releaseLock(); }
}

function safeRangeA1_(sheet, a1) {
  try {
    const r = sheet.getRange(a1);
    if (r.getLastRow() > sheet.getMaxRows() || r.getLastColumn() > sheet.getMaxColumns()) return null;
    return r;
  } catch (e) { return null; }
}

function safeRangeRC_(sheet, row, col, rows, cols) {
  try {
    if (row < 1 || col < 1 || rows < 1 || cols < 1) return null;
    const maxRows = sheet.getMaxRows();
    const maxCols = sheet.getMaxColumns();
    if (row > maxRows || col > maxCols) return null;
    rows = Math.min(rows, maxRows - row + 1);
    cols = Math.min(cols, maxCols - col + 1);
    if (rows < 1 || cols < 1) return null;
    return sheet.getRange(row, col, rows, cols);
  } catch (e) { return null; }
}

function addRange_(arr, range) { if (range) arr.push(range); }

function lastMeaningfulRow_(sheet, startRow, keyCol) {
  const max = sheet.getMaxRows();
  if (max < startRow) return startRow - 1;
  const numRows = max - startRow + 1;
  const vals = sheet.getRange(startRow, keyCol, numRows, 1).getDisplayValues();
  let last = startRow - 1;
  for (let i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][0] || '').trim() !== '') { last = startRow + i; break; }
  }
  return Math.max(last, startRow);
}

function isWeekOneSheet_(name) {
  return /^(M\.R MINI-MART|M\.R BUSH BAR|M\.R KITCHEN)$/i.test(name);
}
function isWeeklyMRSheet_(name) {
  return /^(\d\s+)?M\.R (MINI-MART|BUSH BAR|KITCHEN)$/i.test(name);
}
function isCSSheet_(name) {
  return /^CS (MINI-MART|LAUNDRY|BAR|RESTAURANT|STORE|KITCHEN)$/i.test(name);
}

function editableRangesForSheet_(sheet) {
  const name = sheet.getName();
  const maxRows = sheet.getMaxRows();
  const ranges = [];

  if (name === 'PURCHASES') {
    addRange_(ranges, safeRangeA1_(sheet, 'A:C'));
    addRange_(ranges, safeRangeA1_(sheet, 'E:I'));
    addRange_(ranges, safeRangeA1_(sheet, 'K:K'));
  } else if (name === 'DAILY SALES') {
    addRange_(ranges, safeRangeA1_(sheet, 'A3:C3'));
    addRange_(ranges, safeRangeA1_(sheet, 'O5:O35'));
  } else if (name === 'EXPENSES') {
    addRange_(ranges, safeRangeA1_(sheet, 'A3:C3'));
    addRange_(ranges, safeRangeA1_(sheet, 'A5:I502'));
    addRange_(ranges, safeRangeA1_(sheet, 'K5:K502'));
  } else if (name === 'DAILY SALES BREAKDOWN') {
    if (maxRows >= 5) {
      addRange_(ranges, safeRangeRC_(sheet, 5, 1, maxRows - 4, 8));   // A:H
      addRange_(ranges, safeRangeRC_(sheet, 5, 10, maxRows - 4, 3));  // J:L
      addRange_(ranges, safeRangeRC_(sheet, 5, 14, maxRows - 4, 1));  // N
      addRange_(ranges, safeRangeRC_(sheet, 5, 16, maxRows - 4, 4));  // P:S
    }
  } else if (name === 'STOCK MOVEMENT APPROVAL LOG') {
    addRange_(ranges, safeRangeA1_(sheet, 'A3:C3'));
    if (maxRows >= 7) {
      addRange_(ranges, safeRangeRC_(sheet, 7, 1, maxRows - 6, 3)); // A:C
      addRange_(ranges, safeRangeRC_(sheet, 7, 5, maxRows - 6, 5)); // E:I
    }
  } else if (isWeeklyMRSheet_(name)) {
    ['D2:H2','J2:N2','P2:T2','V2:Z2','AB2:AF2','AH2:AL2','AN2:AR2'].forEach(a1 => addRange_(ranges, safeRangeA1_(sheet, a1)));
    if (isWeekOneSheet_(name)) {
      const last = lastMeaningfulRow_(sheet, 8, 2);
      addRange_(ranges, safeRangeRC_(sheet, 8, 4, last - 7, 1));
    }
  } else if (name === 'M.R KITCHEN U') {
    ['A3:B3','D2:F2','J2:L2','P2:R2','V2:X2'].forEach(a1 => addRange_(ranges, safeRangeA1_(sheet, a1)));
    const last = lastMeaningfulRow_(sheet, 8, 2);
    addRange_(ranges, safeRangeRC_(sheet, 8, 4, last - 7, 2));
  } else if (isCSSheet_(name)) {
    addRange_(ranges, safeRangeA1_(sheet, 'E2:F3'));
    const last = lastMeaningfulRow_(sheet, 5, 2);
    const headerRows = Math.min(8, sheet.getMaxRows());
    const allHeaders = sheet.getRange(1, 1, headerRows, sheet.getMaxColumns()).getDisplayValues();
    const cols = new Set();
    allHeaders.forEach(row => row.forEach((h, i) => {
      const k = key_(h);
      if (k.includes('OPENING STOCK') || k === 'OPENING' || k.includes('PHYSICAL COUNT')) cols.add(i + 1);
    }));
    cols.forEach(col => addRange_(ranges, safeRangeRC_(sheet, 5, col, Math.max(last - 4, 1), 1)));
    addRange_(ranges, safeRangeRC_(sheet, 5, 31, Math.max(last - 4, 1), 1)); // AE5:AE
  }
  return ranges;
}

function existingSheets_(names) {
  const ss = SpreadsheetApp.getActive();
  return names.map(n => ss.getSheetByName(n)).filter(Boolean);
}

function getOrCreateCarlisleProtection_(sheet) {
  const desc = 'Carlisle EOM protection - ' + sheet.getName();
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)
    .filter(p => String(p.getDescription() || '') === desc);
  // Remove duplicate Carlisle protections on THIS sheet only.
  for (let i = 1; i < protections.length; i++) protections[i].remove();
  if (protections[0]) return protections[0];
  return sheet.protect().setDescription(desc);
}

function protectSheetFast_(sheet) {
  const p = getOrCreateCarlisleProtection_(sheet);
  p.setWarningOnly(false);
  p.setDescription('Carlisle EOM protection - ' + sheet.getName());
  p.setUnprotectedRanges(editableRangesForSheet_(sheet));
  // Do not sync staff editors here. That is handled separately by Permission Sync.
  return sheet.getName();
}

function protectSheetsFast_(sheetNames, label, maxSheetsPerRun) {
  return protectionLock_(() => {
    const sheets = existingSheets_(sheetNames).slice(0, maxSheetsPerRun || sheetNames.length);
    if (!sheets.length) throw new Error('No matching sheets found for group: ' + label);
    const done = [];
    sheets.forEach(s => done.push(protectSheetFast_(s)));
    log_('PROTECTION', 'Fast protection applied: ' + label + ' => ' + done.join(', '));
    uiAlert_('Done: ' + label + '\nSheets protected: ' + done.length + '\n' + done.join('\n'));
  });
}

function protect_Purchases(){ protectSheetsFast_(protectionGroups_().PURCHASES, 'Purchases'); }
function protect_DailySales(){ protectSheetsFast_(protectionGroups_().DAILY_SALES, 'Daily Sales'); }
function protect_DailySalesBreakdown(){ protectSheetsFast_(protectionGroups_().DAILY_SALES_BREAKDOWN, 'Daily Sales Breakdown'); }
function protect_Expenses(){ protectSheetsFast_(protectionGroups_().EXPENSES, 'Expenses'); }
function protect_StockMovement(){ protectSheetsFast_(protectionGroups_().STOCK_MOVEMENT, 'Stock Movement Approval Log'); }
function protect_MRKitchenU(){ protectSheetsFast_(protectionGroups_().MR_KITCHEN_U, 'M.R Kitchen U'); }
function protect_SystemSheets(){ protectSheetsFast_(protectionGroups_().SYSTEM, 'System Sheets', 1); }
function protect_LogSheets(){ protectSheetsFast_(protectionGroups_().LOGS, 'Log Sheets', 1); }
function protect_ReportSheets(){ protectSheetsFast_(protectionGroups_().REPORTS, 'Report Sheets', 1); }
function protect_MasterPriceList(){ protectSheetsFast_(protectionGroups_().MASTER_PRICE, 'Master Pricelist', 1); }

function resetAdminProtectionQueue() {
  PropertiesService.getDocumentProperties().deleteProperty('ADMIN_PROTECT_INDEX');
  uiAlert_('Owner/Admin protection queue reset. Run Protect Owner/Admin Sheets - Next Sheet.');
}
function protect_AdminSheets_Next() {
  return protectionLock_(() => {
    const names = protectionGroups_().ADMIN.slice();
    const props = PropertiesService.getDocumentProperties();
    let idx = Number(props.getProperty('ADMIN_PROTECT_INDEX') || 0);
    if (idx >= names.length) {
      props.deleteProperty('ADMIN_PROTECT_INDEX');
      uiAlert_('Owner/Admin protection queue already complete.');
      return;
    }
    const ss = SpreadsheetApp.getActive();
    let protectedName = null;
    while (idx < names.length) {
      const sh = ss.getSheetByName(names[idx]);
      idx++;
      if (sh) { protectedName = protectSheetFast_(sh); break; }
    }
    props.setProperty('ADMIN_PROTECT_INDEX', String(idx));
    if (idx >= names.length) props.deleteProperty('ADMIN_PROTECT_INDEX');
    log_('PROTECTION', 'Owner/Admin next sheet protected: ' + protectedName);
    uiAlert_(protectedName ? ('Protected: ' + protectedName + '\nProgress: ' + idx + '/' + names.length) : 'No more Owner/Admin sheets found.');
  });
}

// Backward-compatible old menu/function name. It protects ONE admin/system/log/report sheet per run, not the whole group.
function protect_AdminSheets(){ protect_AdminSheets_Next(); }
function protect_CSSheets(){ protectSheetsFast_(protectionGroups_().CS_SHEETS, 'CS Sheets', 2); }

function weeklyGroupsV13_() {
  return {
    WEEK1: ['M.R MINI-MART','M.R BUSH BAR','M.R KITCHEN'],
    WEEK2: ['2 M.R MINI-MART','2 M.R BUSH BAR','2 M.R KITCHEN'],
    WEEK3: ['3 M.R MINI-MART','3 M.R BUSH BAR','3 M.R KITCHEN'],
    WEEK4: ['4 M.R MINI-MART','4 M.R BUSH BAR','4 M.R KITCHEN'],
    WEEK5: ['5 M.R MINI-MART','5 M.R BUSH BAR','5 M.R KITCHEN'],
    MINIMART: ['M.R MINI-MART','2 M.R MINI-MART','3 M.R MINI-MART','4 M.R MINI-MART','5 M.R MINI-MART'],
    BUSHBAR: ['M.R BUSH BAR','2 M.R BUSH BAR','3 M.R BUSH BAR','4 M.R BUSH BAR','5 M.R BUSH BAR'],
    KITCHEN: ['M.R KITCHEN','2 M.R KITCHEN','3 M.R KITCHEN','4 M.R KITCHEN','5 M.R KITCHEN']
  };
}

function protect_Weekly_Week1(){ protectSheetsFast_(weeklyGroupsV13_().WEEK1, 'Weekly M.R - Week 1', 1); }
function protect_Weekly_Week2(){ protectSheetsFast_(weeklyGroupsV13_().WEEK2, 'Weekly M.R - Week 2', 1); }
function protect_Weekly_Week3(){ protectSheetsFast_(weeklyGroupsV13_().WEEK3, 'Weekly M.R - Week 3', 1); }
function protect_Weekly_Week4(){ protectSheetsFast_(weeklyGroupsV13_().WEEK4, 'Weekly M.R - Week 4', 1); }
function protect_Weekly_Week5(){ protectSheetsFast_(weeklyGroupsV13_().WEEK5, 'Weekly M.R - Week 5', 1); }
function protect_Weekly_MiniMart(){ protectSheetsFast_(weeklyGroupsV13_().MINIMART, 'Weekly M.R - Mini-Mart', 1); }
function protect_Weekly_BushBar(){ protectSheetsFast_(weeklyGroupsV13_().BUSHBAR, 'Weekly M.R - Bush Bar', 1); }
function protect_Weekly_Kitchen(){ protectSheetsFast_(weeklyGroupsV13_().KITCHEN, 'Weekly M.R - Kitchen', 1); }

function protect_ActiveSheetOnly() {
  return protectionLock_(() => {
    const sheet = SpreadsheetApp.getActiveSheet();
    if (!sheet) throw new Error('No active sheet found.');
    protectSheetFast_(sheet);
    log_('PROTECTION', 'Fast active sheet protection: ' + sheet.getName());
    uiAlert_('Done: protected active sheet only\n' + sheet.getName());
  });
}

function getWeeklyQueue_() { return protectionGroups_().WEEKLY_MR.slice(); }
function resetWeeklyProtectionQueue() {
  PropertiesService.getDocumentProperties().deleteProperty('WEEKLY_PROTECT_INDEX');
  uiAlert_('Weekly protection queue reset. Run Protect Weekly - Next Sheet.');
}
function protect_Weekly_NextSheet() {
  return protectionLock_(() => {
    const names = getWeeklyQueue_();
    const props = PropertiesService.getDocumentProperties();
    let idx = Number(props.getProperty('WEEKLY_PROTECT_INDEX') || 0);
    if (idx >= names.length) {
      props.deleteProperty('WEEKLY_PROTECT_INDEX');
      uiAlert_('Weekly protection queue already complete.');
      return;
    }
    const ss = SpreadsheetApp.getActive();
    let protectedName = null;
    while (idx < names.length) {
      const sh = ss.getSheetByName(names[idx]);
      idx++;
      if (sh) { protectedName = protectSheetFast_(sh); break; }
    }
    props.setProperty('WEEKLY_PROTECT_INDEX', String(idx));
    if (idx >= names.length) props.deleteProperty('WEEKLY_PROTECT_INDEX');
    log_('PROTECTION', 'Weekly next sheet protected: ' + protectedName);
    uiAlert_(protectedName ? ('Protected: ' + protectedName + '\nProgress: ' + idx + '/' + names.length) : 'No more weekly sheets found.');
  });
}

function clearCarlisleProtections() {
  const ss = SpreadsheetApp.getActive();
  ss.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => {
    if (String(p.getDescription() || '').startsWith('Carlisle EOM protection - ')) p.remove();
  });
  ss.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => {
    if (String(p.getDescription() || '').startsWith('Carlisle EOM protection - ')) p.remove();
  });
  log_('PROTECTION', 'Cleared Carlisle protections only.');
  uiAlert_('Carlisle protections cleared.');
}
