/****************************************************
 V16 FAST PROTECTION FUNCTIONS ONLY
 - Structural Protections: Defines WHAT is editable.
 - Access Control: SYSTEM_ACCESS defines WHO can open the file.
 - Use one sheet protection per sheet.
 - EXCLUDES formula cells from cashier editable ranges.
****************************************************/

function protectionGroups_() {
  return {
    PURCHASES: ['PURCHASES'],
    DAILY_SALES: ['DAILY SALES'],
    DAILY_SALES_BREAKDOWN: ['DAILY SALES BREAKDOWN'],
    EXPENSES: ['EXPENSES'],
    STOCK_MOVEMENT: ['STOCK MOVEMENT APPROVAL LOG'],
    CS_SHEETS: ['CS STORE','CS MINI-MART','CS RESTAURANT','CS LAUNDRY','CS BAR','CS KITCHEN'],
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

function existingSheets_(names) {
  const ss = SpreadsheetApp.getActive();
  const all = ss.getSheets();
  const out = [];
  names.forEach(n => {
    const k = key_(n);
    const sh = all.find(s => key_(s.getName()) === k);
    if (sh) out.push(sh);
  });
  return out;
}

function safeRangeA1_(sheet, a1) {
  try {
    const r = sheet.getRange(a1);
    const maxR = sheet.getMaxRows();
    const maxC = sheet.getMaxColumns();
    let startR = r.getRow();
    let startC = r.getColumn();
    let numR = r.getNumRows();
    let numC = r.getNumColumns();
    if (startR > maxR || startC > maxC) return null;
    numR = Math.min(numR, maxR - startR + 1);
    numC = Math.min(numC, maxC - startC + 1);
    return sheet.getRange(startR, startC, numR, numC);
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

function pruneFormulas_(range) {
  if (!range) return [];
  const sheet = range.getSheet();
  const formulas = range.getFormulas();
  const startRow = range.getRow();
  const startCol = range.getColumn();
  const numRows = formulas.length;
  const numCols = formulas[0].length;

  const unprotected = [];
  for (let c = 0; c < numCols; c++) {
    let currentStart = -1;
    for (let r = 0; r < numRows; r++) {
      const hasFormula = !!formulas[r][c];
      if (!hasFormula) {
        if (currentStart === -1) currentStart = r;
      } else {
        if (currentStart !== -1) {
          unprotected.push(sheet.getRange(startRow + currentStart, startCol + c, r - currentStart, 1));
          currentStart = -1;
        }
      }
    }
    if (currentStart !== -1) {
      unprotected.push(sheet.getRange(startRow + currentStart, startCol + c, numRows - currentStart, 1));
    }
  }
  return unprotected;
}

function addRange_(arr, range) {
  if (!range) return;
  const pruned = pruneFormulas_(range);
  pruned.forEach(r => arr.push(r));
}

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
  const n = key_(name);
  return n === 'M.R MINI-MART' || n === 'M.R BUSH BAR' || n === 'M.R KITCHEN';
}
function isWeeklyMRSheet_(name) {
  const n = key_(name);
  return n.includes('M.R') && (n.includes('MINI-MART') || n.includes('BUSH BAR') || n.includes('KITCHEN')) && !n.includes('KITCHEN U');
}

function editableRangesForSheet_(sheet) {
  const name = key_(sheet.getName());
  const maxRows = sheet.getMaxRows();
  const ranges = [];

  const row1004 = Math.min(maxRows, 1004);

  // Core Transaction Sheets
  if (name === 'PURCHASES') {
    addRange_(ranges, safeRangeA1_(sheet, 'A3:C3'));
    addRange_(ranges, safeRangeRC_(sheet, 5, 1, row1004 - 4, 3)); // A5:C1004
    addRange_(ranges, safeRangeRC_(sheet, 5, 5, row1004 - 4, 5)); // E5:I1004
    addRange_(ranges, safeRangeRC_(sheet, 5, 11, row1004 - 4, 1)); // K5:K1004
  } else if (name === 'EXPENSES') {
    addRange_(ranges, safeRangeA1_(sheet, 'A3:C3'));
    addRange_(ranges, safeRangeRC_(sheet, 5, 1, 502 - 4, 9)); // A5:I502
    addRange_(ranges, safeRangeRC_(sheet, 5, 11, 502 - 4, 1)); // K5:K502
  } else if (name === 'STOCK MOVEMENT APPROVAL LOG') {
    addRange_(ranges, safeRangeA1_(sheet, 'A3:C3'));
    addRange_(ranges, safeRangeRC_(sheet, 7, 1, row1004 - 6, 3)); // A7:C1004
    addRange_(ranges, safeRangeRC_(sheet, 7, 5, row1004 - 6, 5)); // E7:I1004
  } else if (name === 'DAILY SALES') {
    addRange_(ranges, safeRangeA1_(sheet, 'A3:C3'));
    addRange_(ranges, safeRangeA1_(sheet, 'O5:O35'));
  } else if (name === 'DAILY SALES BREAKDOWN') {
    addRange_(ranges, safeRangeA1_(sheet, 'A3:C3'));
    addRange_(ranges, safeRangeRC_(sheet, 5, 1, row1004 - 4, 8));   // A5:H (Editable) - I is Restricted
    addRange_(ranges, safeRangeRC_(sheet, 5, 10, row1004 - 4, 3));  // J5:L (Editable) - M is Restricted
    addRange_(ranges, safeRangeRC_(sheet, 5, 14, Math.min(maxRows, 10004) - 4, 1)); // N5:N (Editable) - O is Restricted
    addRange_(ranges, safeRangeRC_(sheet, 5, 16, row1004 - 4, 4));  // P5:S (Editable)

  // CS Sheets
  } else if (isCSSheet_(name)) {
    addRange_(ranges, safeRangeA1_(sheet, 'E2:F3'));
    const lastRow = lastMeaningfulRow_(sheet, 5, 2); // Use Column B (Name) to find last row
    const numRows = lastRow - 4;
    if (numRows > 0) {
      addRange_(ranges, safeRangeRC_(sheet, 5, 5, numRows, 2));  // E5:F (Opening Stock)
      addRange_(ranges, safeRangeRC_(sheet, 5, 31, numRows, 1)); // AE5:AE (Physical Count)
    }

  // Weekly M.R Sheets
  } else if (isWeeklyMRSheet_(name)) {
    addRange_(ranges, safeRangeA1_(sheet, 'A3:B3'));
    ['D2:H2','J2:N2','P2:T2','V2:Z2','AB2:AF2','AH2:AL2','AN2:AR2'].forEach(a1 => addRange_(ranges, safeRangeA1_(sheet, a1)));

    const lastRow = lastMeaningfulRow_(sheet, 8, 2); // Column B
    const numRows = lastRow - 7;
    if (numRows > 0) {
      // Sales columns: F, L, R, X, AD, AJ, AP (cols 6, 12, 18, 24, 30, 36, 42)
      [6, 12, 18, 24, 30, 36, 42].forEach(col => addRange_(ranges, safeRangeRC_(sheet, 8, col, numRows, 1)));
      // Opening Stock (Week 1 only): D8:D
      if (isWeekOneSheet_(name)) {
        addRange_(ranges, safeRangeRC_(sheet, 8, 4, numRows, 1));
      }
    }

  // M.R KITCHEN U
  } else if (name === 'M.R KITCHEN U') {
    addRange_(ranges, safeRangeA1_(sheet, 'A3:B3'));
    ['D2:F2','J2:L2','P2:R2','V2:X2'].forEach(a1 => addRange_(ranges, safeRangeA1_(sheet, a1)));
    const lastRow = lastMeaningfulRow_(sheet, 8, 1); // Column A
    const numRows = lastRow - 7;
    if (numRows > 0) addRange_(ranges, safeRangeRC_(sheet, 8, 4, numRows, 2)); // D8:E (Opening Stock)
  }

  return ranges;
}

function getOrCreateCarlisleProtection_(sheet) {
  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());
  const desc = 'Carlisle EOM protection - ' + sheet.getName();
  const p = sheet.protect().setDescription(desc);
  const owners = getOwnerEmails_();
  if (owners.length) {
    p.removeEditors(p.getEditors());
    p.addEditors(owners);
  }
  return p;
}

function protectSheetFast_(sheet) {
  const p = getOrCreateCarlisleProtection_(sheet);
  p.setWarningOnly(false);
  const unprotected = editableRangesForSheet_(sheet);
  if (unprotected.length) {
    p.setUnprotectedRanges(unprotected.slice(0, 50));
  }
  return sheet.getName();
}

function protectSheetsFast_(sheetNames, label) {
  return protectionLock_(() => {
    const sheets = existingSheets_(sheetNames);
    if (!sheets.length) throw new Error('No matching sheets found for group: ' + label);
    const done = [];
    sheets.forEach(s => done.push(protectSheetFast_(s)));
    log_('PROTECTION', 'Fast protection applied: ' + label + ' => ' + done.join(', '));
    uiAlert_('Done: ' + label + '\nSheets protected: ' + done.length + '\n' + done.join('\n'));
  });
}

function runQueueProtection_(queueName, groupName, label) {
  return protectionLock_(() => {
    const names = protectionGroups_()[groupName].slice();
    const props = PropertiesService.getDocumentProperties();
    let idx = Number(props.getProperty(queueName) || 0);
    if (idx >= names.length) {
      props.deleteProperty(queueName);
      uiAlert_(label + ' protection queue already complete.');
      return;
    }
    const allSheets = SpreadsheetApp.getActive().getSheets();
    let protectedName = null;
    while (idx < names.length) {
      const targetKey = key_(names[idx]);
      const sh = allSheets.find(s => key_(s.getName()) === targetKey);
      idx++;
      if (sh) { protectedName = protectSheetFast_(sh); break; }
    }
    props.setProperty(queueName, String(idx));
    if (idx >= names.length) props.deleteProperty(queueName);
    log_('PROTECTION', label + ' next sheet protected: ' + protectedName);
    uiAlert_(protectedName ? ('Protected: ' + protectedName + '\nProgress: ' + idx + '/' + names.length) : 'No more sheets found.');
  });
}

function protect_Purchases(){ protectSheetsFast_(protectionGroups_().PURCHASES, 'Purchases'); }
function protect_DailySales(){ protectSheetsFast_(protectionGroups_().DAILY_SALES, 'Daily Sales'); }
function protect_DailySalesBreakdown(){ protectSheetsFast_(protectionGroups_().DAILY_SALES_BREAKDOWN, 'Daily Sales Breakdown'); }
function protect_Expenses(){ protectSheetsFast_(protectionGroups_().EXPENSES, 'Expenses'); }
function protect_StockMovement(){ protectSheetsFast_(protectionGroups_().STOCK_MOVEMENT, 'Stock Movement Approval Log'); }
function protect_MRKitchenU(){ protectSheetsFast_(protectionGroups_().MR_KITCHEN_U, 'M.R Kitchen U'); }

function protect_AdminSheets_Next() { runQueueProtection_('ADMIN_PROTECT_INDEX', 'ADMIN', 'Owner/Admin'); }
function resetAdminProtectionQueue() { PropertiesService.getDocumentProperties().deleteProperty('ADMIN_PROTECT_INDEX'); uiAlert_('Admin queue reset.'); }

function protect_CSSheets_Next() { runQueueProtection_('CS_PROTECT_INDEX', 'CS_SHEETS', 'CS Sheets'); }
function resetCSSheetsProtectionQueue() { PropertiesService.getDocumentProperties().deleteProperty('CS_PROTECT_INDEX'); uiAlert_('CS Sheets queue reset.'); }

function protect_CSSheets(){ protect_CSSheets_Next(); }

function weeklyGroupsV13_() {
  return {
    WEEK1_2: ['M.R MINI-MART','M.R BUSH BAR','M.R KITCHEN','2 M.R MINI-MART','2 M.R BUSH BAR','2 M.R KITCHEN'],
    WEEK3_4: ['3 M.R MINI-MART','3 M.R BUSH BAR','3 M.R KITCHEN','4 M.R MINI-MART','4 M.R BUSH BAR','4 M.R KITCHEN'],
    WEEK5: ['5 M.R MINI-MART','5 M.R BUSH BAR','5 M.R KITCHEN']
  };
}

function protect_Weekly_Weeks1_2(){ protectSheetsFast_(weeklyGroupsV13_().WEEK1_2, 'Weekly M.R - Weeks 1-2'); }
function protect_Weekly_Weeks3_4(){ protectSheetsFast_(weeklyGroupsV13_().WEEK3_4, 'Weekly M.R - Weeks 3-4'); }
function protect_Weekly_Week5(){ protectSheetsFast_(weeklyGroupsV13_().WEEK5, 'Weekly M.R - Week 5'); }

function protect_ActiveSheetOnly() {
  return protectionLock_(() => {
    const sheet = SpreadsheetApp.getActiveSheet();
    if (!sheet) throw new Error('No active sheet found.');
    protectSheetFast_(sheet);
    log_('PROTECTION', 'Fast active sheet protection: ' + sheet.getName());
    uiAlert_('Done: protected active sheet only\n' + sheet.getName());
  });
}

function resetWeeklyProtectionQueue() { PropertiesService.getDocumentProperties().deleteProperty('WEEKLY_PROTECT_INDEX'); uiAlert_('Weekly queue reset.'); }
function protect_Weekly_NextSheet() { runQueueProtection_('WEEKLY_PROTECT_INDEX', 'WEEKLY_MR', 'Weekly M.R'); }

function clearCarlisleProtections() {
  const ss = SpreadsheetApp.getActive();
  ss.getSheets().forEach(sh => {
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
    sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());
  });
  log_('PROTECTION', 'Cleared ALL protections in the workbook.');
  uiAlert_('All protections cleared.');
}
