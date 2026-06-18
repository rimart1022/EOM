/****************************************************
 CARLISLE E.O.M STOCK SYSTEM - V16 DECOUPLED ARCHITECTURE
 - SYSTEM_ACCESS is the single source of truth for users/permissions.
 - Shared utilities are centralized here for zero redundancy.
****************************************************/

const CONFIG = {
  SYSTEM_ACCESS_SHEET: 'SYSTEM_ACCESS',
  LOG_SHEET: 'SYSTEM_LOGS',
  EDIT_LOG_SHEET: 'EOM EDIT LOG',
  STOCK_CHANGE_LOG: 'STOCK CHANGE LOG',
  MASTER_PRICE_LIST: 'MASTER_PRICELIST',

  SYSTEM_ACCESS_HEADERS: ['Email', 'Full Name', 'Role', 'Active', 'Notes', 'Sheets Controlled'],

  // Rule 2: MD and MANAGING DIRECTOR allowed to give approval.
  OWNER_ROLES: ['OWNER', 'MANAGING DIRECTOR', 'MD'],

  ROLE_OPTIONS: [
    'OWNER','MANAGING DIRECTOR','MD','MANAGER','DUTY SUPERVISOR',
    'CASHIER','WAITSTAFF','STORE CLERK','INVENTORY LEAD','INVENTORY LEAD I','INVENTORY LEAD II',
    'KITCHEN ASSISTANT','COOK','LAUNDRY','HOUSEKEEPER','HOUSEKEEPING LEAD',
    'PORTER','SECURITY','MOPOL','GARDENER','MARKETER','SOCIAL MEDIA','INTERN'
  ],

  SHEET_CONTROL_OPTIONS: [
    'ALL','OWNER SHEETS','ADMIN SHEETS','SYSTEM SHEETS','LOG SHEETS','REPORT SHEETS','REPORTS','CS SHEETS','ALL CS SHEETS','WEEKLY M.R SHEETS',
    'PURCHASES','DAILY SALES','DAILY SALES BREAKDOWN','EXPENSES','STOCK MOVEMENT APPROVAL LOG',
    'CS BAR','CS KITCHEN','CS MINI-MART','CS LAUNDRY','CS BAR','CS RESTAURANT','CS STORE',
    'MASTER_PRICELIST','SYSTEM_ACCESS','SYSTEM_LOGS','SYSTEM_SETTINGS','EOM EDIT LOG','STOCK CHANGE LOG',
    'DAMAGE REPORT','ISSUED STOCK REPORT','STAFF LIABILITY REPORT','STOCK AUDIT SUMMARY','FORMULA ERROR CHECK'
  ],

  DEPARTMENT_OPTIONS: [
    'CS MINI-MART','CS BAR','CS RESTAURANT','CS LAUNDRY','CS STORE','CS KITCHEN',
    'M.R MINI-MART','M.R BUSH BAR','M.R KITCHEN','M.R KITCHEN U',
    'Store: Housekeeping','Store: Reception','Store: Maintenance','Store: Kitchen','Store: Functionaries','BOOKING'
  ],

  MOVEMENT_TYPES: ['SOLD','ISSUED','DAMAGED','UTILIZED','ADDED','OPENING'],

  GROUP_ALIASES: {
    'ALL': ['*'],
    'OWNER SHEETS': ['MASTER_PRICELIST','SYSTEM_ACCESS','SYSTEM_SETTINGS','SYSTEM_LOGS','EOM EDIT LOG','STOCK CHANGE LOG','STOCK AUDIT SUMMARY','ISSUED STOCK REPORT','DAMAGE REPORT','STAFF LIABILITY REPORT'],
    'ADMIN SHEETS': ['MASTER_PRICELIST','SYSTEM_ACCESS','SYSTEM_SETTINGS','SYSTEM_LOGS','EOM EDIT LOG','STOCK CHANGE LOG','STOCK AUDIT SUMMARY','ISSUED STOCK REPORT','DAMAGE REPORT','STAFF LIABILITY REPORT'],
    'CS SHEETS': ['CS STORE','CS MINI-MART','CS RESTAURANT','CS LAUNDRY','CS BAR','CS KITCHEN'],
    'ALL CS SHEETS': ['CS STORE','CS MINI-MART','CS RESTAURANT','CS LAUNDRY','CS BAR','CS KITCHEN'],
    'WEEKLY M.R SHEETS': ['M.R MINI-MART','M.R BUSH BAR','M.R KITCHEN','2 M.R MINI-MART','3 M.R MINI-MART','4 M.R MINI-MART','5 M.R MINI-MART','2 M.R BUSH BAR','3 M.R BUSH BAR','4 M.R BUSH BAR','5 M.R BUSH BAR','2 M.R KITCHEN','3 M.R KITCHEN','4 M.R KITCHEN','5 M.R KITCHEN']
  }
};

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Carlisle EOM')
    .addSubMenu(ui.createMenu('Menu 1: Protections Carlisle EOM')
      .addItem('Protect Purchases', 'protect_Purchases')
      .addItem('Protect Daily Sales', 'protect_DailySales')
      .addItem('Protect Daily Sales Breakdown', 'protect_DailySalesBreakdown')
      .addItem('Protect Expenses', 'protect_Expenses')
      .addItem('Protect Stock Movement Approval Log', 'protect_StockMovement')
      .addSeparator()
      .addItem('Protect Weekly M.R Weeks 1-2', 'protect_Weekly_Weeks1_2')
      .addItem('Protect Weekly M.R Weeks 3-4', 'protect_Weekly_Weeks3_4')
      .addItem('Protect Weekly M.R Week 5', 'protect_Weekly_Week5')
      .addItem('Protect M.R Kitchen U', 'protect_MRKitchenU')
      .addSeparator()
      .addItem('Protect CS Sheets - Next Sheet', 'protect_CSSheets_Next')
      .addItem('Protect Admin Sheets - Next Sheet', 'protect_AdminSheets_Next')
      .addSeparator()
      .addItem('Clear All Protections', 'clearCarlisleProtections'))
    .addSubMenu(ui.createMenu('Menu 2: Access Control Carlisle EOM')
      .addItem('Rebuild SYSTEM_ACCESS', 'rebuildSystemAccess')
      .addItem('Setup SYSTEM_ACCESS Dropdowns', 'setupSystemAccessDropdowns')
      .addItem('Setup Log Dropdowns', 'setupStockMovementDropdowns')
      .addItem('Validate SYSTEM_ACCESS', 'validateSystemAccess')
      .addItem('Sync User Permissions', 'syncUserPermissions_All'))
    .addSubMenu(ui.createMenu('Carlisle Reports')
      .addItem('Generate Damage Report', 'generateDamageReport')
      .addItem('Generate Issued Stock Report', 'generateIssuedStockReport')
      .addItem('Generate Utilized Report', 'generateUtilizedReport')
      .addItem('Generate Staff Liability Report', 'generateStaffLiabilityReport')
      .addItem('Generate Stock Audit Summary', 'generateStockAuditSummary')
      .addItem('Check Formula Errors', 'checkFormulaErrors'))
    .addSubMenu(ui.createMenu('Utilities')
      .addItem('Refresh Approval Timestamps', 'refreshApprovalTimestamps')
      .addItem('Sync Master Price Items', 'syncMasterPriceItemsFromDepartments'))
    .addToUi();
}

/** SHARED UTILITIES **/

function key_(v) { return String(v || '').trim().toUpperCase().replace(/\s+/g, ' '); }

function splitList_(v) { return String(v || '').split(/[,\n;]+/).map(v => key_(v)).filter(Boolean); }

function log_(type, message) {
  try {
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName(CONFIG.LOG_SHEET);
    if (!sh) sh = ss.insertSheet(CONFIG.LOG_SHEET);
    sh.appendRow([new Date(), type, Session.getActiveUser().getEmail(), message]);
  } catch (e) {}
}

function uiAlert_(message) {
  try { SpreadsheetApp.getUi().alert(message); } catch (e) { console.warn(message); }
}

function isCSSheet_(name) {
  const n = key_(name);
  return n.startsWith('CS ');
}

function isWeeklyMRSheet_(name) {
  const n = key_(name);
  return n.includes('M.R') && (n.includes('MINI-MART') || n.includes('BUSH BAR') || n.includes('KITCHEN')) && !n.includes('KITCHEN U');
}

function isWeekOneSheet_(name) {
  const n = key_(name);
  return n === 'M.R MINI-MART' || n === 'M.R BUSH BAR' || n === 'M.R KITCHEN';
}

function lastMeaningfulRow_(sheet, startRow, keyCol) {
  const max = sheet.getMaxRows();
  if (max < startRow) return startRow - 1;
  const vals = sheet.getRange(startRow, keyCol, max - startRow + 1, 1).getDisplayValues();
  let last = startRow - 1;
  for (let i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][0] || '').trim() !== '') { last = startRow + i; break; }
  }
  return Math.max(last, startRow);
}

function findHeaderCol_(sheet, names, headerRows) {
  headerRows = headerRows || 10;
  const maxRows = Math.min(headerRows, sheet.getMaxRows());
  if (maxRows === 0) return null;
  const vals = sheet.getRange(1, 1, maxRows, sheet.getMaxColumns()).getValues();
  const wanted = names.map(key_);
  for (let r = 0; r < vals.length; r++) {
    for (let c = 0; c < vals[r].length; c++) {
      if (wanted.includes(key_(vals[r][c]))) return { row: r + 1, col: c + 1 };
    }
  }
  return null;
}
