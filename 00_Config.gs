/****************************************************
 CARLISLE E.O.M STOCK SYSTEM - V16 OWNER + SYSTEM PROTECTIONS FIXED
 - No hard-coded staff emails.
 - SYSTEM_ACCESS is the source of users and sheet permissions.
 - Protections and permission-sync are separated to avoid timeouts.
****************************************************/

const CONFIG = {
  SYSTEM_ACCESS_SHEET: 'SYSTEM_ACCESS',
  LOG_SHEET: 'SYSTEM_LOGS',
  EDIT_LOG_SHEET: 'EOM EDIT LOG',
  MASTER_PRICE_LIST: 'MASTER_PRICELIST',

  SYSTEM_ACCESS_HEADERS: ['Email', 'Full Name', 'Role', 'Active', 'Notes', 'Sheets Controlled'],
  OWNER_ROLES: ['OWNER', 'MANAGING DIRECTOR', 'MD'],

  ROLE_OPTIONS: [
    'OWNER','MANAGING DIRECTOR','MD','MANAGER','DUTY SUPERVISOR',
    'CASHIER','WAITSTAFF','STORE CLERK','INVENTORY LEAD','INVENTORY LEAD I','INVENTORY LEAD II',
    'KITCHEN ASSISTANT','COOK','LAUNDRY','HOUSEKEEPER','HOUSEKEEPING LEAD',
    'PORTER','SECURITY','MOPOL','GARDENER','MARKETER','SOCIAL MEDIA','INTERN'
  ],

  SHEET_CONTROL_OPTIONS: [
    'ALL','OWNER SHEETS','ADMIN SHEETS','SYSTEM SHEETS','LOG SHEETS','REPORT SHEETS','CS SHEETS','WEEKLY M.R SHEETS',
    'PURCHASES','DAILY SALES','DAILY SALES BREAKDOWN','EXPENSES','STOCK MOVEMENT APPROVAL LOG',
    'CS MINI-MART','CS BAR','CS RESTAURANT','CS LAUNDRY','CS STORE','CS KITCHEN',
    'M.R MINI-MART','M.R MINI-MART SHEETS','M.R BUSH BAR','M.R BUSH BAR SHEETS','M.R KITCHEN','M.R KITCHEN SHEETS','M.R KITCHEN U',
    'MASTER_PRICELIST','MASTER PRICE LIST','SYSTEM_ACCESS','SYSTEM_LOGS','SYSTEM_SETTINGS','EOM EDIT LOG','STOCK CHANGE LOG',
    'DAMAGE REPORT','ISSUED STOCK REPORT','STAFF LIABILITY REPORT','STOCK AUDIT SUMMARY','UTILIZED REPORT'
  ],

  GROUP_ALIASES: {
    'ALL': ['*'],
    'OWNER SHEETS': ['MASTER_PRICELIST','SYSTEM_ACCESS','SYSTEM_SETTINGS','SYSTEM_LOGS','EOM EDIT LOG','STOCK CHANGE LOG','STOCK AUDIT SUMMARY','AUDIT CHECK WEEK 1','AUDIT CHECK WEEK 2','AUDIT CHECK WEEK 3','AUDIT CHECK WEEK 4','AUDIT CHECK WEEK 5','ISSUED STOCK REPORT','DAMAGE REPORT','STAFF LIABILITY REPORT','UTILIZED REPORT'],
    'ADMIN SHEETS': ['MASTER_PRICELIST','SYSTEM_ACCESS','SYSTEM_SETTINGS','SYSTEM_LOGS','EOM EDIT LOG','STOCK CHANGE LOG','STOCK AUDIT SUMMARY','AUDIT CHECK WEEK 1','AUDIT CHECK WEEK 2','AUDIT CHECK WEEK 3','AUDIT CHECK WEEK 4','AUDIT CHECK WEEK 5','ISSUED STOCK REPORT','DAMAGE REPORT','STAFF LIABILITY REPORT','UTILIZED REPORT'],
    'SYSTEM SHEETS': ['SYSTEM_ACCESS','SYSTEM_SETTINGS','SYSTEM_LOGS'],
    'LOG SHEETS': ['SYSTEM_LOGS','EOM EDIT LOG','STOCK CHANGE LOG'],
    'REPORT SHEETS': ['STOCK AUDIT SUMMARY','AUDIT CHECK WEEK 1','AUDIT CHECK WEEK 2','AUDIT CHECK WEEK 3','AUDIT CHECK WEEK 4','AUDIT CHECK WEEK 5','ISSUED STOCK REPORT','DAMAGE REPORT','STAFF LIABILITY REPORT','UTILIZED REPORT'],
    'PURCHASES': ['PURCHASES'],
    'DAILY SALES': ['DAILY SALES'],
    'DAILY SALES BREAKDOWN': ['DAILY SALES BREAKDOWN'],
    'EXPENSES': ['EXPENSES'],
    'STOCK MOVEMENT APPROVAL LOG': ['STOCK MOVEMENT APPROVAL LOG'],
    'MASTER_PRICELIST': ['MASTER_PRICELIST'],
    'MASTER PRICE LIST': ['MASTER_PRICELIST'],
    'CS SHEETS': ['CS MINI-MART','CS LAUNDRY','CS BAR','CS RESTAURANT','CS STORE','CS KITCHEN'],
    'WEEKLY M.R SHEETS': ['M.R MINI-MART','M.R BUSH BAR','M.R KITCHEN','2 M.R MINI-MART','2 M.R BUSH BAR','2 M.R KITCHEN','3 M.R MINI-MART','3 M.R BUSH BAR','3 M.R KITCHEN','4 M.R MINI-MART','4 M.R BUSH BAR','4 M.R KITCHEN','5 M.R MINI-MART','5 M.R BUSH BAR','5 M.R KITCHEN'],
    'M.R MINI-MART SHEETS': ['M.R MINI-MART','2 M.R MINI-MART','3 M.R MINI-MART','4 M.R MINI-MART','5 M.R MINI-MART'],
    'M.R BUSH BAR SHEETS': ['M.R BUSH BAR','2 M.R BUSH BAR','3 M.R BUSH BAR','4 M.R BUSH BAR','5 M.R BUSH BAR'],
    'M.R KITCHEN SHEETS': ['M.R KITCHEN','2 M.R KITCHEN','3 M.R KITCHEN','4 M.R KITCHEN','5 M.R KITCHEN'],
    'M.R KITCHEN U': ['M.R KITCHEN U']
  }
};

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Carlisle EOM')
    .addSubMenu(ui.createMenu('Access Control')
      .addItem('Rebuild SYSTEM_ACCESS', 'rebuildSystemAccess')
      .addItem('Setup SYSTEM_ACCESS Dropdowns', 'setupSystemAccessDropdowns')
      .addItem('Validate SYSTEM_ACCESS', 'validateSystemAccess')
      .addSeparator()
      .addItem('Sync User Permissions - ALL', 'syncUserPermissions_All')
      .addItem('Sync User Permissions - Purchases', 'syncUserPermissions_Purchases')
      .addItem('Sync User Permissions - CS Sheets', 'syncUserPermissions_CSSheets')
      .addItem('Sync User Permissions - Weekly M.R Sheets', 'syncUserPermissions_WeeklyMR')
      .addSeparator()
      .addItem('Sync Permissions - ACTIVE Sheet Only', 'syncUserPermissions_ActiveSheetOnly'))
    .addSubMenu(ui.createMenu('Protections Only')
      .addItem('Protect Purchases', 'protect_Purchases')
      .addItem('Protect Daily Sales', 'protect_DailySales')
      .addItem('Protect Daily Sales Breakdown', 'protect_DailySalesBreakdown')
      .addItem('Protect Expenses', 'protect_Expenses')
      .addItem('Protect Stock Movement Approval Log', 'protect_StockMovement')
      .addItem('Protect CS Sheets', 'protect_CSSheets')
      .addSeparator()
      .addItem('Protect Weekly - Next Sheet', 'protect_Weekly_NextSheet')
      .addItem('Reset Weekly Protection Queue', 'resetWeeklyProtectionQueue')
      .addSeparator()
      .addItem('Protect M.R Kitchen U', 'protect_MRKitchenU')
      .addItem('Protect ACTIVE Sheet Only', 'protect_ActiveSheetOnly')
      .addItem('Protect Master Pricelist', 'protect_MasterPriceList')
      .addItem('Protect Owner/Admin Sheets - Next Sheet', 'protect_AdminSheets_Next')
      .addItem('Reset Owner/Admin Protection Queue', 'resetAdminProtectionQueue')
      .addSeparator()
      .addItem('Clear Protections - ALL', 'clearCarlisleProtections'))
    .addSubMenu(ui.createMenu('Reports')
      .addItem('Generate Damage Report', 'generateDamageReport')
      .addItem('Generate Issued Stock Report', 'generateIssuedStockReport')
      .addItem('Generate Staff Liability Report', 'generateStaffLiabilityReport')
      .addItem('Generate Utilized Report', 'generateUtilizedReport')
      .addItem('Generate Stock Audit Summary', 'generateStockAuditSummary')
      .addItem('Check Formula Errors', 'checkFormulaErrors'))
    .addSubMenu(ui.createMenu('Utilities')
      .addItem('Refresh Approval Timestamps', 'refreshApprovalTimestamps')
      .addItem('Run Master Price Cost Sync', 'syncMasterPriceCostsFromApprovedMovements'))
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
  SpreadsheetApp.getUi().alert(message);
}
