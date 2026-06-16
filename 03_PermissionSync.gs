/****************************************************
 PERMISSION-SYNC FUNCTIONS ONLY.
 These read SYSTEM_ACCESS and update protection editors.
 Run separately from protection creation to avoid timeouts.
****************************************************/

function usersForSheetFromAccess_(sheetName) {
  const owners = getOwnerEmails_();
  const directUsers = usersForSheet_(sheetName);
  return Array.from(new Set(owners.concat(directUsers).filter(Boolean)));
}

function syncPermissionsForSheetNames_(sheetNames, label) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) throw new Error('Another Carlisle permission task is already running. Try again shortly.');
  try {
    const ss = SpreadsheetApp.getActive();
    const targetNames = new Set(sheetNames.map(key_));
    let count = 0;
    ss.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => {
      const range = p.getRange();
      if (!range) return;
      const sheetName = range.getSheet().getName();
      if (!targetNames.has(key_(sheetName))) return;
      const editors = usersForSheetFromAccess_(sheetName);
      try { p.removeEditors(p.getEditors()); } catch (e) {}
      if (editors.length) p.addEditors(editors);
      count++;
    });
    log_('PERMISSION_SYNC', 'Synced editors for ' + label + ': ' + count + ' protections.');
    uiAlert_('Done syncing permissions: ' + label + '\nProtections updated: ' + count);
  } finally {
    lock.releaseLock();
  }
}

function syncUserPermissions_All() {
  const ss = SpreadsheetApp.getActive();
  const allNames = ss.getSheets().map(sh => sh.getName());
  syncPermissionsForSheetNames_(allNames, 'ALL');
}
function syncUserPermissions_Purchases(){ syncPermissionsForSheetNames_(protectionGroups_().PURCHASES, 'Purchases'); }
function syncUserPermissions_CSSheets(){ syncPermissionsForSheetNames_(protectionGroups_().CS_SHEETS, 'CS Sheets'); }
function syncUserPermissions_WeeklyMR(){ syncPermissionsForSheetNames_(protectionGroups_().WEEKLY_MR.concat(protectionGroups_().MR_KITCHEN_U), 'Weekly M.R Sheets'); }

/****************************************************
 V13 SMALLER PERMISSION SYNC GROUPS
 Run these after the corresponding protections if needed.
****************************************************/
function syncUserPermissions_Weekly_Week1(){ syncPermissionsForSheetNames_(weeklyGroupsV13_().WEEK1, 'Weekly M.R - Week 1'); }
function syncUserPermissions_Weekly_Week2(){ syncPermissionsForSheetNames_(weeklyGroupsV13_().WEEK2, 'Weekly M.R - Week 2'); }
function syncUserPermissions_Weekly_Week3(){ syncPermissionsForSheetNames_(weeklyGroupsV13_().WEEK3, 'Weekly M.R - Week 3'); }
function syncUserPermissions_Weekly_Week4(){ syncPermissionsForSheetNames_(weeklyGroupsV13_().WEEK4, 'Weekly M.R - Week 4'); }
function syncUserPermissions_Weekly_Week5(){ syncPermissionsForSheetNames_(weeklyGroupsV13_().WEEK5, 'Weekly M.R - Week 5'); }

function syncUserPermissions_ActiveSheetOnly() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (!sheet) throw new Error('No active sheet found.');
  syncPermissionsForSheetNames_([sheet.getName()], 'Active sheet only: ' + sheet.getName());
}
