/****************************************************
 PERMISSION-SYNC FUNCTIONS ONLY.
 These read SYSTEM_ACCESS and update Spreadsheet-level editors.
****************************************************/

function syncUserPermissions_All() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) throw new Error('Lock timeout.');
  try {
    const ss = SpreadsheetApp.getActive();
    const records = getAccessRecords_();
    if (!records.length) throw new Error('No active staff found.');

    const staffEmails = Array.from(new Set(records.map(r => r.email)));
    const ownerEmails = getOwnerEmails_();

    const currentEditors = ss.getEditors().map(e => e.getEmail().toLowerCase());
    const staffLower = staffEmails.map(e => e.toLowerCase());
    const ownersLower = ownerEmails.map(o => o.toLowerCase());

    // Add missing staff
    const toAdd = staffEmails.filter(e => !currentEditors.includes(e.toLowerCase()));
    if (toAdd.length) ss.addEditors(toAdd);

    // Remove inactive/deleted staff (except owners)
    const toRemove = currentEditors.filter(e => !staffLower.includes(e) && !ownersLower.includes(e));
    if (toRemove.length) {
      toRemove.forEach(e => { try { ss.removeEditor(e); } catch (err) {} });
    }

    log_('PERMISSION_SYNC', 'Synced ' + staffEmails.length + ' staff members. Added: ' + toAdd.length + ', Removed: ' + toRemove.length);
    uiAlert_('Editors synced with SYSTEM_ACCESS.\nActive Staff: ' + staffEmails.length + '\nRemoved: ' + toRemove.length);
  } finally {
    lock.releaseLock();
  }
}
