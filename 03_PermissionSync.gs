/****************************************************
 PERMISSION-SYNC FUNCTIONS ONLY.
 - Syncs Workbook (Spreadsheet) editors for all active staff.
 - Syncs Protection editors for Owners/MDs only.
 - Cashiers edit unprotected ranges via Workbook editor access.
****************************************************/

function syncWorkbookEditors_() {
  const ss = SpreadsheetApp.getActive();
  const records = getAccessRecords_(); // Active users only
  const owners = getOwnerEmails_();
  const targetEmails = records.map(r => r.email.toLowerCase());
  const currentUser = Session.getActiveUser().getEmail().toLowerCase();

  // Add all active staff as spreadsheet editors
  targetEmails.forEach(email => {
    try { ss.addEditor(email); } catch(e) {}
  });

  // Remove users who are no longer in SYSTEM_ACCESS or are inactive
  // But never remove Owners, MDs, or the current user.
  const currentEditors = ss.getEditors();
  currentEditors.forEach(e => {
    const email = e.getEmail().toLowerCase();
    if (!targetEmails.includes(email) && !owners.includes(email) && email !== currentUser) {
      try { ss.removeEditor(email); } catch(e) {}
    }
  });
}

function syncPermissionsForSheetNames_(sheetNames, label) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) throw new Error('Could not acquire lock for permission sync.');
  try {
    const ss = SpreadsheetApp.getActive();
    const owners = getOwnerEmails_();
    const targetNames = new Set(sheetNames.map(key_));
    let count = 0;

    ss.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => {
      const range = p.getRange();
      if (!range) return;
      const shName = range.getSheet().getName();
      if (!targetNames.has(key_(shName))) return;

      // Protection editors are ONLY Owners/MDs.
      // This prevents normal staff (Cashiers) from editing protected areas/formulas.
      try { p.removeEditors(p.getEditors()); } catch (e) {}
      if (owners.length) p.addEditors(owners);
      count++;
    });
    log_('PERMISSION_SYNC', 'Synced ' + label + ' protections for ' + owners.length + ' owners. Total sheets: ' + count);
    uiAlert_('Done syncing permissions for: ' + label + '\nWorkbook editors synced.\nProtection editors limited to Owners/MDs.');
  } finally {
    lock.releaseLock();
  }
}

function syncUserPermissions_All() {
  syncWorkbookEditors_();
  const groups = protectionGroups_();
  const all = [].concat(groups.PURCHASES, groups.DAILY_SALES, groups.DAILY_SALES_BREAKDOWN, groups.EXPENSES, groups.STOCK_MOVEMENT, groups.CS_SHEETS, groups.WEEKLY_MR, groups.MR_KITCHEN_U, groups.ADMIN);
  syncPermissionsForSheetNames_(Array.from(new Set(all)), 'ALL');
}

function syncUserPermissions_Purchases(){ syncPermissionsForSheetNames_(protectionGroups_().PURCHASES, 'Purchases'); }
function syncUserPermissions_CSSheets(){ syncPermissionsForSheetNames_(protectionGroups_().CS_SHEETS, 'CS Sheets'); }
function syncUserPermissions_WeeklyMR(){ syncPermissionsForSheetNames_(protectionGroups_().WEEKLY_MR.concat(protectionGroups_().MR_KITCHEN_U), 'Weekly M.R'); }

function syncUserPermissions_ActiveSheetOnly() {
  const s = SpreadsheetApp.getActiveSheet();
  if (!s) return;
  syncPermissionsForSheetNames_([s.getName()], 'Active: ' + s.getName());
}
