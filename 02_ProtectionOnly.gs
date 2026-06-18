/****************************************************
 V16 FAST PROTECTION FUNCTIONS ONLY
 - Protections do NOT sync staff permissions.
 - Use one sheet protection per sheet.
 - Reuses existing Carlisle protection instead of deleting/recreating.
 - Weekly can be run one sheet at a time to avoid timeout.
 - Formula cells are never editable.
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
    ADMIN: ['MASTER_PRICELIST','SYSTEM_ACCESS','SYSTEM_SETTINGS','SYSTEM_LOGS','EOM EDIT LOG','STOCK CHANGE LOG','STOCK AUDIT SUMMARY','AUDIT CHECK WEEK 1','AUDIT CHECK WEEK 2','AUDIT CHECK WEEK 3','AUDIT CHECK WEEK 4','AUDIT CHECK WEEK 5','ISSUED STOCK REPORT','DAMAGE REPORT','STAFF LIABILITY REPORT','UTILIZED REPORT'],
    MASTER_PRICE: ['MASTER_PRICELIST']
  };
}

function protectionLock_(callback) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) throw new Error('Another Carlisle protection task is already running. Try again shortly.');
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

function pruneFormulas_(sheet, ranges) {
  if (!ranges || !ranges.length) return [];
  const mask = {};
  ranges.forEach(range => {
    if (!range) return;
    const formulas = range.getFormulas();
    const startRow = range.getRow();
    const startCol = range.getColumn();
    for (let r = 0; r < formulas.length; r++) {
      for (let c = 0; c < formulas[r].length; c++) {
        if (formulas[r][c] === "") {
          const row = startRow + r;
          const col = startCol + c;
          if (!mask[col]) mask[col] = {};
          mask[col][row] = true;
        }
      }
    }
  });
  const consolidated = [];
  const columns = Object.keys(mask).map(Number).sort((a,b) => a - b);
  columns.forEach(col => {
    const rows = Object.keys(mask[col]).map(Number).sort((a,b) => a - b);
    if (!rows.length) return;
    let startRow = rows[0];
    for (let i = 1; i <= rows.length; i++) {
      if (i === rows.length || rows[i] !== rows[i-1] + 1) {
        consolidated.push(sheet.getRange(startRow, col, rows[i-1] - startRow + 1, 1));
        if (i < rows.length) startRow = rows[i];
      }
    }
  });
  return consolidated;
}

function editableRangesForSheet_(sheet) {
  const name = sheet.getName();
  const maxRows = sheet.getMaxRows();
  const ranges = [];
  const getCappedRow = (start, cap) => Math.min(lastMeaningfulRow_(sheet, start, 2), cap);

  if (name === 'PURCHASES') {
    addRange_(ranges, safeRangeA1_(sheet, 'A3:C3'));
    addRange_(ranges, safeRangeA1_(sheet, 'A5:C'));
    addRange_(ranges, safeRangeA1_(sheet, 'E5:I'));
    addRange_(ranges, safeRangeA1_(sheet, 'K:K'));
  } else if (name === 'DAILY SALES') {
    addRange_(ranges, safeRangeA1_(sheet, 'A3:C3'));
    addRange_(ranges, safeRangeA1_(sheet, 'O5:O35'));
  } else if (name === 'EXPENSES') {
    addRange_(ranges, safeRangeA1_(sheet, 'A3:C3'));
    addRange_(ranges, safeRangeA1_(sheet, 'A5:I502'));
    addRange_(ranges, safeRangeA1_(sheet, 'K5:K502'));
  } else if (name === 'DAILY SALES BREAKDOWN') {
    addRange_(ranges, safeRangeA1_(sheet, 'A3:C3'));
    if (maxRows >= 5) {
      addRange_(ranges, safeRangeRC_(sheet, 5, 1, maxRows - 4, 8));
      addRange_(ranges, safeRangeRC_(sheet, 5, 10, maxRows - 4, 3));
      addRange_(ranges, safeRangeRC_(sheet, 5, 14, maxRows - 4, 1));
      addRange_(ranges, safeRangeRC_(sheet, 5, 16, maxRows - 4, 4));
    }
  } else if (name === 'STOCK MOVEMENT APPROVAL LOG') {
    addRange_(ranges, safeRangeA1_(sheet, 'A3:C3'));
    if (maxRows >= 7) {
      addRange_(ranges, safeRangeRC_(sheet, 7, 1, maxRows - 6, 3));
      addRange_(ranges, safeRangeRC_(sheet, 7, 5, maxRows - 6, 5));
    }
  } else if (isWeeklyMRSheet_(name)) {
    addRange_(ranges, safeRangeA1_(sheet, 'A3:B3'));
    ['D2:H2','J2:N2','P2:T2','V2:Z2','AB2:AF2','AH2:AL2','AN2:AR2'].forEach(a1 => addRange_(ranges, safeRangeA1_(sheet, a1)));
    let cap = 0;
    if (name === 'M.R KITCHEN') cap = 106;
    else if (name === 'M.R BUSH BAR') cap = 102;
    else if (name === 'M.R MINI-MART') cap = 145;
    if (cap > 0 && isWeekOneSheet_(name)) {
      const last = getCappedRow(8, cap);
      if (last >= 8) addRange_(ranges, safeRangeRC_(sheet, 8, 4, last - 7, 1));
    }
  } else if (name === 'M.R KITCHEN U') {
    addRange_(ranges, safeRangeA1_(sheet, 'A3:B3'));
    ['D2:F2','J2:L2','P2:R2','V2:X2'].forEach(a1 => addRange_(ranges, safeRangeA1_(sheet, a1)));
    const last = getCappedRow(8, 70);
    if (last >= 8) addRange_(ranges, safeRangeRC_(sheet, 8, 4, last - 7, 1));
  } else if (isCSSheet_(name)) {
    addRange_(ranges, safeRangeA1_(sheet, 'E2:F3'));
    const caps = {'CS STORE':311,'CS MINI-MART':100,'CS RESTAURANT':94,'CS LAUNDRY':50,'CS BAR':74,'CS KITCHEN':50};
    const cap = caps[name] || 0;
    const last = getCappedRow(5, cap);
    const numRows = Math.max(last - 4, 0);
    if (numRows > 0) {
      addRange_(ranges, safeRangeRC_(sheet, 5, 5, numRows, 2)); // E:F
      if (name !== 'CS LAUNDRY') addRange_(ranges, safeRangeRC_(sheet, 5, 15, numRows, 1)); // O
      addRange_(ranges, safeRangeRC_(sheet, 5, 31, numRows, 1)); // AE
    }
  }
  return pruneFormulas_(sheet, ranges);
}

function existingSheets_(names) {
  const ss = SpreadsheetApp.getActive();
  return names.map(n => ss.getSheetByName(n)).filter(Boolean);
}

function protectSheetFast_(sheet) {
  const owners = getOwnerEmails_();
  // Clear all existing sheet and range protections
  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());

  const p = sheet.protect().setDescription('Carlisle EOM protection - ' + sheet.getName());
  p.setWarningOnly(false);
  p.removeEditors(p.getEditors());
  if (owners.length) p.addEditors(owners);

  const ranges = editableRangesForSheet_(sheet);
  if (ranges.length) p.setUnprotectedRanges(ranges);
  return sheet.getName();
}

function protectSheetsFast_(sheetNames, label, maxSheetsPerRun) {
  return protectionLock_(() => {
    const sheets = existingSheets_(sheetNames).slice(0, maxSheetsPerRun || sheetNames.length);
    if (!sheets.length) return;
    const done = [];
    sheets.forEach(s => done.push(protectSheetFast_(s)));
    log_('PROTECTION', 'V16 Protection applied: ' + label + ' => ' + done.join(', '));
    uiAlert_('Done: ' + label + '\nSheets protected: ' + done.length + '\n' + done.join('\n'));
  });
}

function protect_Purchases(){ protectSheetsFast_(protectionGroups_().PURCHASES, 'Purchases'); }
function protect_DailySales(){ protectSheetsFast_(protectionGroups_().DAILY_SALES, 'Daily Sales'); }
function protect_DailySalesBreakdown(){ protectSheetsFast_(protectionGroups_().DAILY_SALES_BREAKDOWN, 'Daily Sales Breakdown'); }
function protect_Expenses(){ protectSheetsFast_(protectionGroups_().EXPENSES, 'Expenses'); }
function protect_StockMovement(){ protectSheetsFast_(protectionGroups_().STOCK_MOVEMENT, 'Stock Movement Approval Log'); }
function protect_MRKitchenU(){ protectSheetsFast_(protectionGroups_().MR_KITCHEN_U, 'M.R Kitchen U'); }
function protect_MasterPriceList(){ protectSheetsFast_(protectionGroups_().MASTER_PRICE, 'Master Pricelist', 1); }
function protect_CSSheets(){ protectSheetsFast_(protectionGroups_().CS_SHEETS, 'CS Sheets', 2); }

function resetAdminProtectionQueue() {
  PropertiesService.getDocumentProperties().deleteProperty('ADMIN_PROTECT_INDEX');
  uiAlert_('Owner/Admin protection queue reset.');
}
function protect_AdminSheets_Next() {
  return protectionLock_(() => {
    const names = protectionGroups_().ADMIN.slice();
    const props = PropertiesService.getDocumentProperties();
    let idx = Number(props.getProperty('ADMIN_PROTECT_INDEX') || 0);
    if (idx >= names.length) { props.deleteProperty('ADMIN_PROTECT_INDEX'); uiAlert_('Queue complete.'); return; }
    const ss = SpreadsheetApp.getActive();
    let protectedName = null;
    while (idx < names.length) {
      const sh = ss.getSheetByName(names[idx++]);
      if (sh) { protectedName = protectSheetFast_(sh); break; }
    }
    props.setProperty('ADMIN_PROTECT_INDEX', String(idx));
    if (idx >= names.length) props.deleteProperty('ADMIN_PROTECT_INDEX');
    uiAlert_(protectedName ? ('Protected: ' + protectedName) : 'No more sheets found.');
  });
}

function protect_ActiveSheetOnly() {
  return protectionLock_(() => {
    const sheet = SpreadsheetApp.getActiveSheet();
    protectSheetFast_(sheet);
    log_('PROTECTION', 'Active sheet protection: ' + sheet.getName());
    uiAlert_('Done: ' + sheet.getName());
  });
}

function resetWeeklyProtectionQueue() {
  PropertiesService.getDocumentProperties().deleteProperty('WEEKLY_PROTECT_INDEX');
  uiAlert_('Weekly queue reset.');
}
function protect_Weekly_NextSheet() {
  return protectionLock_(() => {
    const names = protectionGroups_().WEEKLY_MR.slice();
    const props = PropertiesService.getDocumentProperties();
    let idx = Number(props.getProperty('WEEKLY_PROTECT_INDEX') || 0);
    if (idx >= names.length) { props.deleteProperty('WEEKLY_PROTECT_INDEX'); uiAlert_('Queue complete.'); return; }
    const ss = SpreadsheetApp.getActive();
    let protectedName = null;
    while (idx < names.length) {
      const sh = ss.getSheetByName(names[idx++]);
      if (sh) { protectedName = protectSheetFast_(sh); break; }
    }
    props.setProperty('WEEKLY_PROTECT_INDEX', String(idx));
    if (idx >= names.length) props.deleteProperty('WEEKLY_PROTECT_INDEX');
    uiAlert_(protectedName ? ('Protected: ' + protectedName) : 'No more sheets found.');
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
  uiAlert_('Carlisle protections cleared.');
}
