/****************************************************
 PERMISSION-SYNC FUNCTIONS ONLY.
 These read SYSTEM_ACCESS and update Spreadsheet-level editors.
 Sheet-level structural protections (locked ranges) are maintained
 by Menu 1 and remain fixed once applied.
****************************************************/

/**
 * Returns an array of user emails that should have edit access to the spreadsheet
 * based on SYSTEM_ACCESS sheet.
 */
function getActiveStaffEmails_(records) {
  return Array.from(new Set(records.map(r => r.email).filter(Boolean)));
}

function syncUserPermissions_All() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) throw new Error('Another Carlisle permission task is already running. Try again shortly.');
  try {
    const ss = SpreadsheetApp.getActive();
    const records = getAccessRecords_(); // Read once
    if (!records.length) throw new Error('No active staff found in SYSTEM_ACCESS.');

    const staffEmails = getActiveStaffEmails_(records);
    const ownerEmails = getOwnerEmails_();

    const currentEditors = ss.getEditors().map(e => e.getEmail().toLowerCase());
    const staffLower = staffEmails.map(e => e.toLowerCase());

    // Add missing staff as spreadsheet editors
    const toAdd = staffEmails.filter(e => !currentEditors.includes(e.toLowerCase()));
    if (toAdd.length) ss.addEditors(toAdd);

    // Logic: We don't remove editors here because they might have been added manually.
    // However, if we want strict control, we would remove anyone not in staffLower or ownerEmails.

    log_('PERMISSION_SYNC', 'Synced ' + staffEmails.length + ' staff members to spreadsheet editors.');
    uiAlert_('Spreadsheet editors synced with SYSTEM_ACCESS.\nTotal staff: ' + staffEmails.length);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Note: syncPermissionsForSheetNames_ is deprecated in the new architecture
 * because structural sheet protection should only have Owners as editors.
 * Cashiers edit through unblocked ranges.
 */
function syncUserPermissions_ActiveSheetOnly() {
  syncUserPermissions_All();
}
