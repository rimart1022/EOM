/****************************************************
 CARLISLE E.O.M STOCK SYSTEM - V16 DECOUPLED ARCHITECTURE
 - No hard-coded staff emails.
 - SYSTEM_ACCESS is the source of users and sheet permissions.
 - Protections and permission-sync are separated to avoid timeouts.
****************************************************/

const CONFIG = {
  SYSTEM_ACCESS_SHEET: 'SYSTEM_ACCESS',
  LOG_SHEET: 'SYSTEM_LOGS',
  EDIT_LOG_SHEET: 'EOM EDIT LOG',
  STOCK_CHANGE_LOG: 'STOCK CHANGE LOG',
  MASTER_PRICE_LIST: 'MASTER_PRICELIST',

  SYSTEM_ACCESS_HEADERS: ['Email', 'Full Name', 'Role', 'Active', 'Notes', 'Sheets Controlled'],
  OWNER_ROLES: ['OWNER', 'MANAGING DIRECTOR', 'MD'],

  ROLE_OPTIONS: [
    'OWNER','MANAGING DIRECTOR','MD','MANAGER','DUTY SUPERVISOR',
    'CASHIER','WAITSTAFF','STORE CLERK','INVENTORY LEAD','INVENTORY LEAD I','INVENTORY LEAD II',
    'KITCHEN ASSISTANT','COOK','LAUNDRY','HOUSEKEEPER','HOUSEKEEPING LEAD',
    'PORTER','SECURITY','MOPOL','GARDENER','MARKETER','SOCIAL MEDIA','INTERN'
  ],

  // These are options that can be typed or selected in SYSTEM_ACCESS > Sheets Controlled.
  SHEET_CONTROL_OPTIONS: [
    'ALL','OWNER SHEETS','ADMIN SHEETS','SYSTEM SHEETS','LOG SHEETS','REPORT SHEETS','REPORTS','CS SHEETS','ALL CS SHEETS','WEEKLY M.R SHEETS',
    'WEEKLY M.R WEEK 1-2','WEEKLY M.R WEEK 3-4','WEEKLY M.R WEEK 5',
    'PURCHASES','DAILY SALES','DAILY SALES BREAKDOWN','EXPENSES','STOCK MOVEMENT APPROVAL LOG',
    'CS MINI-MART','CS BAR','CS RESTAURANT','CS LAUNDRY','CS STORE','CS KITCHEN',
    'M.R MINI-MART 1-5','M.R BUSH BAR 1-5','M.R KITCHEN 1-5',
    'M.R MINI-MART','M.R BUSH BAR','M.R KITCHEN','M.R KITCHEN U',
    'MASTER_PRICELIST','SYSTEM_ACCESS','SYSTEM_LOGS','SYSTEM_SETTINGS','EOM EDIT LOG','STOCK CHANGE LOG',
    'DAMAGE REPORT','ISSUED STOCK REPORT','STAFF LIABILITY REPORT','STOCK AUDIT SUMMARY','FORMULA ERROR CHECK'
  ],

  DEPARTMENT_OPTIONS: [
    'CS MINI-MART','CS BAR','CS RESTAURANT','CS LAUNDRY','CS STORE','CS KITCHEN',
    'M.R MINI-MART','M.R BUSH BAR','M.R KITCHEN','M.R KITCHEN U',
    'Store: Housekeeping','Store: Reception','Store: Maintenance','Store: Kitchen','Store: Functionaries','BOOKING'
  ],

  MOVEMENT_TYPES: [
    'SOLD','ISSUED','DAMAGED','UTILIZED','ADDED','OPENING'
  ],

  GROUP_ALIASES: {
    'ALL': ['*'],
    'OWNER SHEETS': ['MASTER_PRICELIST','MASTER PRICE LIST','SYSTEM_ACCESS','SYSTEM_SETTINGS','SYSTEM_LOGS','EOM EDIT LOG','STOCK CHANGE LOG','STOCK AUDIT SUMMARY','AUDIT CHECK WEEK 1','AUDIT CHECK WEEK 2','AUDIT CHECK WEEK 3','AUDIT CHECK WEEK 4','AUDIT CHECK WEEK 5','ISSUED STOCK REPORT','DAMAGE REPORT','STAFF LIABILITY REPORT'],
    'ADMIN SHEETS': ['MASTER_PRICELIST','MASTER PRICE LIST','SYSTEM_ACCESS','SYSTEM_SETTINGS','SYSTEM_LOGS','EOM EDIT LOG','STOCK CHANGE LOG','STOCK AUDIT SUMMARY','AUDIT CHECK WEEK 1','AUDIT CHECK WEEK 2','AUDIT CHECK WEEK 3','AUDIT CHECK WEEK 4','AUDIT CHECK WEEK 5','ISSUED STOCK REPORT','DAMAGE REPORT','STAFF LIABILITY REPORT'],
    'SYSTEM SHEETS': ['SYSTEM_ACCESS','SYSTEM_SETTINGS','SYSTEM_LOGS'],
    'LOG SHEETS': ['SYSTEM_LOGS','EOM EDIT LOG','STOCK CHANGE LOG'],
    'REPORT SHEETS': ['STOCK AUDIT SUMMARY','AUDIT CHECK WEEK 1','AUDIT CHECK WEEK 2','AUDIT CHECK WEEK 3','AUDIT CHECK WEEK 4','AUDIT CHECK WEEK 5','ISSUED STOCK REPORT','DAMAGE REPORT','STAFF LIABILITY REPORT'],
    'REPORTS': ['DAMAGE REPORT','ISSUED STOCK REPORT','STAFF LIABILITY REPORT','STOCK AUDIT SUMMARY','FORMULA ERROR CHECK','AUDIT CHECK WEEK 1','AUDIT CHECK WEEK 2','AUDIT CHECK WEEK 3','AUDIT CHECK WEEK 4','AUDIT CHECK WEEK 5'],
    'PURCHASES': ['PURCHASES'],
    'DAILY SALES': ['DAILY SALES'],
    'DAILY SALES BREAKDOWN': ['DAILY SALES BREAKDOWN'],
    'EXPENSES': ['EXPENSES'],
    'STOCK MOVEMENT APPROVAL LOG': ['STOCK MOVEMENT APPROVAL LOG'],
    'MASTER_PRICELIST': ['MASTER_PRICELIST'],
    'CS SHEETS': ['CS MINI-MART','CS LAUNDRY','CS BAR','CS RESTAURANT','CS STORE','CS KITCHEN'],
    'ALL CS SHEETS': ['CS MINI-MART','CS LAUNDRY','CS BAR','CS RESTAURANT','CS STORE','CS KITCHEN'],
    'CS MINI-MART': ['CS MINI-MART'],
    'CS LAUNDRY': ['CS LAUNDRY'],
    'CS BAR': ['CS BAR'],
    'CS RESTAURANT': ['CS RESTAURANT'],
    'CS STORE': ['CS STORE'],
    'CS KITCHEN': ['CS KITCHEN'],
    'WEEKLY M.R SHEETS': ['M.R MINI-MART','M.R BUSH BAR','M.R KITCHEN','2 M.R MINI-MART','3 M.R MINI-MART','4 M.R MINI-MART','5 M.R MINI-MART','2 M.R BUSH BAR','3 M.R BUSH BAR','4 M.R BUSH BAR','5 M.R BUSH BAR','2 M.R KITCHEN','3 M.R KITCHEN','4 M.R KITCHEN','5 M.R KITCHEN'],
    'WEEKLY M.R WEEK 1-2': ['M.R MINI-MART','M.R BUSH BAR','M.R KITCHEN','2 M.R MINI-MART','2 M.R BUSH BAR','2 M.R KITCHEN'],
    'WEEKLY M.R WEEK 3-4': ['3 M.R MINI-MART','3 M.R BUSH BAR','3 M.R KITCHEN','4 M.R MINI-MART','4 M.R BUSH BAR','4 M.R KITCHEN'],
    'WEEKLY M.R WEEK 5': ['5 M.R MINI-MART','5 M.R BUSH BAR','5 M.R KITCHEN'],
    'M.R MINI-MART 1-5': ['M.R MINI-MART','2 M.R MINI-MART','3 M.R MINI-MART','4 M.R MINI-MART','5 M.R MINI-MART'],
    'M.R BUSH BAR 1-5': ['M.R BUSH BAR','2 M.R BUSH BAR','3 M.R BUSH BAR','4 M.R BUSH BAR','5 M.R BUSH BAR'],
    'M.R KITCHEN 1-5': ['M.R KITCHEN','2 M.R KITCHEN','3 M.R KITCHEN','4 M.R KITCHEN','5 M.R KITCHEN'],
    'M.R MINI-MART': ['M.R MINI-MART','2 M.R MINI-MART','3 M.R MINI-MART','4 M.R MINI-MART','5 M.R MINI-MART'],
    'M.R BUSH BAR': ['M.R BUSH BAR','2 M.R BUSH BAR','3 M.R BUSH BAR','4 M.R BUSH BAR','5 M.R BUSH BAR'],
    'M.R KITCHEN': ['M.R KITCHEN','2 M.R KITCHEN','3 M.R KITCHEN','4 M.R KITCHEN','5 M.R KITCHEN']
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
      .addItem('Clear All Protections', 'clearCarlisleProtections')
      .addSeparator()
      .addItem('Reset Admin Protection Queue', 'resetAdminProtectionQueue')
      .addItem('Reset CS Sheets Protection Queue', 'resetCSSheetsProtectionQueue')
      .addItem('Reset Weekly Protection Queue', 'resetWeeklyProtectionQueue'))
    .addSubMenu(ui.createMenu('Menu 2: Access Control Carlisle EOM')
      .addItem('Rebuild SYSTEM_ACCESS', 'rebuildSystemAccess')
      .addItem('Setup SYSTEM_ACCESS Dropdowns', 'setupSystemAccessDropdowns')
      .addItem('Setup Log Dropdowns', 'setupStockMovementDropdowns')
      .addItem('Setup Sales Breakdown Dropdowns', 'setupDailySalesBreakdownDropdowns')
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
      .addItem('Run Master Price Cost Sync', 'syncMasterPriceCostsFromApprovedMovements')
      .addItem('Run Quantity Movement Sync', 'syncQuantitiesFromApprovedMovements')
      .addItem('Sync Master Price Items from Departments', 'syncMasterPriceItemsFromDepartments'))
    .addToUi();
}

function key_(value) {
  return String(value || '').trim().toUpperCase().replace(/\s+/g, ' ');
}

function splitList_(value) {
  return String(value || '')
    .split(/[,\n;]+/)
    .map(v => key_(v))
    .filter(Boolean);
}

function log_(type, message) {
  try {
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName(CONFIG.LOG_SHEET);
    if (!sh) sh = ss.insertSheet(CONFIG.LOG_SHEET);
    if (sh.getLastRow() === 0) sh.appendRow(['Timestamp','Type','User','Message']);
    sh.appendRow([new Date(), type, Session.getActiveUser().getEmail(), message]);
  } catch (e) {}
}

function uiAlert_(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
    console.warn("UI Alert suppressed: " + message);
  }
}
