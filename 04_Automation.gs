/****************************************************
 AUTOMATION: triggers, approval workflow, sync.
 Kept separate from protection/permission logic.
****************************************************/

function onEdit(e) {
  try {
    if (!e || !e.range) return;

    // Check permission via SYSTEM_ACCESS
    if (checkSheetPermission_(e)) return;

    // Rule 1: Approval lock - prevent non-authorized users from editing APPROVED rows.
    if (preventApprovedEdit_(e)) return;

    auditEdit_(e);
    handleApprovalEdit_(e);
    handleMasterPriceCostUpdate_(e);
  } catch (err) {
    log_('ON_EDIT_ERROR', err.message);
  }
}

/**
 * Enforces 'Sheets Controlled' from SYSTEM_ACCESS.
 */
function checkSheetPermission_(e) {
  const userEmail = Session.getActiveUser().getEmail();
  const shName = e.range.getSheet().getName();

  const sys = ['SYSTEM_ACCESS','SYSTEM_LOGS','EOM EDIT LOG','STOCK CHANGE LOG'];
  if (sys.includes(shName.toUpperCase())) return false;

  const authEmails = usersForSheet_(shName);
  if (!authEmails.map(u => u.toLowerCase()).includes(userEmail.toLowerCase())) {
    if (!getOwnerEmails_().map(o => o.toLowerCase()).includes(userEmail.toLowerCase())) {
      e.range.setValue(e.oldValue || '');
      uiAlert_('Access Denied: You are not authorized for sheet "' + shName + '".');
      return true;
    }
  }
  return false;
}

/**
 * Rule 1: Prevent editing approved rows unless MD/Owner.
 */
function preventApprovedEdit_(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== 'STOCK MOVEMENT APPROVAL LOG') return false;

  const statusCell = findHeaderCol_(sh, ['STATUS'], 10);
  if (!statusCell || e.range.getRow() <= statusCell.row) return false;

  const status = key_(sh.getRange(e.range.getRow(), statusCell.col).getValue());
  if (status !== 'APPROVED') return false;

  const userEmail = Session.getActiveUser().getEmail();
  if (!getOwnerEmails_().map(o => o.toLowerCase()).includes(userEmail.toLowerCase())) {
    e.range.setValue(e.oldValue || '');
    uiAlert_('Action Denied: This entry is APPROVED and locked.');
    return true;
  }
  return false;
}

function auditEdit_(e) {
  try {
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName(CONFIG.EDIT_LOG_SHEET);
    if (!sh) return;
    sh.appendRow([new Date(), Session.getActiveUser().getEmail(), e.range.getSheet().getName(), e.range.getA1Notation(), e.oldValue || '', e.value || '']);
  } catch (e) {}
}

function handleApprovalEdit_(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== 'STOCK MOVEMENT APPROVAL LOG') return;
  const statusCell = findHeaderCol_(sh, ['STATUS'], 10);
  if (!statusCell || e.range.getColumn() !== statusCell.col || e.range.getRow() <= statusCell.row) return;

  const userEmail = Session.getActiveUser().getEmail();
  if (!getOwnerEmails_().map(o => o.toLowerCase()).includes(userEmail.toLowerCase())) {
    e.range.setValue(e.oldValue || '');
    uiAlert_('Access Denied: Only MD or Owners can approve movements.');
    return;
  }

  const stat = key_(e.value);
  if (stat === 'APPROVED') {
    const approvedBy = findHeaderCol_(sh, ['APPROVED BY','APPROVER'], 10);
    const approvedDate = findHeaderCol_(sh, ['APPROVAL DATE'], 10);
    if (approvedBy) sh.getRange(e.range.getRow(), approvedBy.col).setValue(userEmail);
    if (approvedDate) sh.getRange(e.range.getRow(), approvedDate.col).setValue(new Date());

    syncMovementToTarget_(sh, e.range.getRow());
  }
}

function handleMasterPriceCostUpdate_(e) {
  if (e.range.getSheet().getName() !== 'PURCHASES' || e.range.getColumn() !== 11 || e.range.getRow() < 5) return;

  const row = e.range.getRow();
  const codeHeader = findHeaderCol_(e.range.getSheet(), ['ITEM CODE'], 10);
  if (!codeHeader) return;

  const code = String(e.range.getSheet().getRange(row, codeHeader.col).getValue()).trim();
  const cost = Number(e.value);
  if (!code || isNaN(cost)) return;

  const master = SpreadsheetApp.getActive().getSheetByName(CONFIG.MASTER_PRICE_LIST);
  if (!master) return;

  const mCodeH = findHeaderCol_(master, ['ITEM CODE'], 10);
  const mCostH = findHeaderCol_(master, ['COST PRICE'], 10);
  if (!mCodeH || !mCostH) return;

  const mData = master.getRange(mCodeH.row + 1, mCodeH.col, master.getLastRow() - mCodeH.row, 1).getValues();
  for (let i = 0; i < mData.length; i++) {
    if (String(mData[i][0]).trim() === code) {
      master.getRange(mCodeH.row + 1 + i, mCostH.col).setValue(cost);
      break;
    }
  }
}

function syncMovementToTarget_(sh, row) {
  const deptCol = findHeaderCol_(sh, ['DEPARTMENT'], 10);
  const codeCol = findHeaderCol_(sh, ['ITEM CODE'], 10);
  const qtyCol = findHeaderCol_(sh, ['QTY'], 10);
  const typeCol = findHeaderCol_(sh, ['MOVEMENT TYPE'], 10);
  if (!deptCol || !codeCol || !qtyCol || !typeCol) return;

  const deptName = String(sh.getRange(row, deptCol.col).getValue()).trim();
  const code = String(sh.getRange(row, codeCol.col).getValue()).trim();
  const qty = Number(sh.getRange(row, qtyCol.col).getValue());
  const type = key_(sh.getRange(row, typeCol.col).getValue());

  const targetSh = SpreadsheetApp.getActive().getSheetByName(deptName);
  if (!targetSh) return;

  const tCodeH = findHeaderCol_(targetSh, ['ITEM CODE'], 10);
  if (!tCodeH) return;

  let targetColName = null;
  if (type === 'SOLD' || type === 'UTILIZED') targetColName = ['SOLD','SALES'];
  else if (type === 'DAMAGED') targetColName = ['DAMAGED','DAMAGE'];
  else if (type === 'ISSUED') targetColName = ['ISSUED'];
  else if (type === 'ADDED') targetColName = ['ADDED STOCK'];

  if (!targetColName) return;
  const tMetricH = findHeaderCol_(targetSh, targetColName, 10);
  if (!tMetricH) return;

  const tData = targetSh.getRange(tCodeH.row + 1, tCodeH.col, targetSh.getLastRow() - tCodeH.row, 1).getValues();
  for (let i = 0; i < tData.length; i++) {
    if (String(tData[i][0]).trim() === code) {
      const cur = Number(targetSh.getRange(tCodeH.row + 1 + i, tMetricH.col).getValue() || 0);
      targetSh.getRange(tCodeH.row + 1 + i, tMetricH.col).setValue(cur + qty);
      break;
    }
  }
}

function syncMasterPriceItemsFromDepartments() {
  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName(CONFIG.MASTER_PRICE_LIST);
  if (!master) return;

  const mCodeH = findHeaderCol_(master, ['ITEM CODE'], 10) || {col: 2};
  const mItemH = findHeaderCol_(master, ['ITEM'], 10) || {col: 1};

  const codes = new Set(master.getRange(2, mCodeH.col, Math.max(master.getLastRow() - 1, 1), 1).getValues().map(v => String(v[0]).trim()));
  const newItems = [];

  ss.getSheets().forEach(sh => {
    if (!isCSSheet_(sh.getName()) && !isWeeklyMRSheet_(sh.getName())) return;
    const sCodeH = findHeaderCol_(sh, ['ITEM CODE'], 10);
    const sItemH = findHeaderCol_(sh, ['ITEM'], 10);
    if (!sCodeH || !sItemH) return;

    const data = sh.getRange(sCodeH.row + 1, 1, Math.max(sh.getLastRow() - sCodeH.row, 1), sh.getLastColumn()).getValues();
    data.forEach(r => {
      const c = String(r[sCodeH.col - 1]).trim();
      const itm = String(r[sItemH.col - 1]).trim();
      if (c && !codes.has(c)) {
        const row = [];
        row[mCodeH.col - 1] = c;
        row[mItemH.col - 1] = itm;
        newItems.push(row);
        codes.add(c);
      }
    });
  });

  if (newItems.length) {
    master.getRange(master.getLastRow() + 1, 1, newItems.length, newItems[0].length).setValues(newItems);
    uiAlert_('Added ' + newItems.length + ' new items to Master Price List.');
  }
}

function refreshApprovalTimestamps() {
  uiAlert_('Function not needed in V16 Real-time approval.');
}
