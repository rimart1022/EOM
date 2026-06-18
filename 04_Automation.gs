/****************************************************
 AUTOMATION: approval, audit, and security.
****************************************************/

function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    const name = sh.getName();

    // Sheet-level access check
    const user = Session.getActiveUser().getEmail();
    const owners = getOwnerEmails_();
    if (!owners.includes(user)) {
      const allowed = usersForSheet_(name);
      if (!allowed.includes(user)) {
        e.range.setValue(e.oldValue);
        uiAlert_('Access Denied: You are not assigned to ' + name);
        return;
      }
    }

    if (preventApprovedEdit_(e)) return;
    auditEdit_(e);
    handleApprovalEdit_(e);
    handleMasterPriceCostUpdate_(e);
  } catch (err) {
    log_('ON_EDIT_ERROR', err.message || err);
  }
}

function preventApprovedEdit_(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== 'STOCK MOVEMENT APPROVAL LOG') return false;
  const statusCell = findHeaderCol_(sh, ['STATUS'], 10);
  if (!statusCell) return false;

  const row = e.range.getRow();
  if (row <= statusCell.row) return false;

  const currentStatus = key_(sh.getRange(row, statusCell.col).getValue());
  if (currentStatus === 'APPROVED') {
    const owners = getOwnerEmails_();
    if (!owners.includes(Session.getActiveUser().getEmail())) {
      e.range.setValue(e.oldValue);
      uiAlert_('This row is APPROVED and locked. Contact an Owner.');
      return true;
    }
  }
  return false;
}

function handleApprovalEdit_(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== 'STOCK MOVEMENT APPROVAL LOG') return;
  const statusCell = findHeaderCol_(sh, ['STATUS'], 10);
  if (!statusCell || e.range.getColumn() !== statusCell.col || e.range.getRow() <= statusCell.row) return;

  const owners = getOwnerEmails_();
  if (!owners.includes(Session.getActiveUser().getEmail())) {
    e.range.setValue(e.oldValue);
    uiAlert_('Only Owners/MDs can change status.');
    return;
  }

  const approvedBy = findHeaderCol_(sh, ['APPROVED BY','APPROVER'], 10);
  const approvedDate = findHeaderCol_(sh, ['APPROVAL DATE','APPROVED DATE'], 10);
  if (approvedBy) sh.getRange(e.range.getRow(), approvedBy.col).setValue(Session.getActiveUser().getEmail());
  if (approvedDate) sh.getRange(e.range.getRow(), approvedDate.col).setValue(new Date());
}

function auditEdit_(e) {
  try {
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName(CONFIG.EDIT_LOG_SHEET);
    if (!sh) return;
    if (sh.getLastRow() === 0) sh.appendRow(['Timestamp','User','Sheet','Cell','Old Value','New Value']);
    sh.appendRow([new Date(), Session.getActiveUser().getEmail(), e.range.getSheet().getName(), e.range.getA1Notation(), e.oldValue || '', e.value || '']);
  } catch (err) {}
}

function findHeaderCol_(sheet, names, headerRows) {
  headerRows = headerRows || 10;
  const maxRows = Math.min(headerRows, sheet.getMaxRows());
  const vals = sheet.getRange(1, 1, maxRows, sheet.getMaxColumns()).getValues();
  const wanted = names.map(key_);
  for (let r = 0; r < vals.length; r++) {
    for (let c = 0; c < vals[r].length; c++) {
      const k = key_(vals[r][c]);
      if (wanted.includes(k)) return { row: r + 1, col: c + 1 };
    }
  }
  return null;
}

function refreshApprovalTimestamps() {
  const sh = SpreadsheetApp.getActive().getSheetByName('STOCK MOVEMENT APPROVAL LOG');
  if (!sh) return;
  const statusCell = findHeaderCol_(sh, ['STATUS'], 10);
  const approvedBy = findHeaderCol_(sh, ['APPROVED BY','APPROVER'], 10);
  const approvedDate = findHeaderCol_(sh, ['APPROVAL DATE','APPROVED DATE'], 10);
  if (!statusCell || !approvedBy || !approvedDate) return;
  const last = sh.getLastRow();
  for (let r = statusCell.row + 1; r <= last; r++) {
    const status = key_(sh.getRange(r, statusCell.col).getValue());
    if (['APPROVED','REJECTED','UNDER REVIEW'].includes(status)) {
      if (!sh.getRange(r, approvedBy.col).getValue()) sh.getRange(r, approvedBy.col).setValue('SYSTEM CHECK');
      if (!sh.getRange(r, approvedDate.col).getValue()) sh.getRange(r, approvedDate.col).setValue(new Date());
    }
  }
  uiAlert_('Timestamps refreshed.');
}

function handleMasterPriceCostUpdate_(e) {
  const name = e.range.getSheet().getName();
  if (!['PURCHASES','STOCK MOVEMENT APPROVAL LOG'].includes(name)) return;
  const header = findHeaderCol_(e.range.getSheet(), ['UNIT COST','COST PRICE','UNIT VALUE'], 10);
  if (header && e.range.getColumn() === header.col) syncMasterPriceCostsFromApprovedMovements();
}

function syncMasterPriceCostsFromApprovedMovements() {
  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName(CONFIG.MASTER_PRICE_LIST);
  if (!master) return;
  const mCode = findHeaderCol_(master, ['ITEM CODE','CODE'], 10) || {row: 1, col: 1};
  const mCost = findHeaderCol_(master, ['COST PRICE','UNIT COST'], 10);
  if (!mCost) return;

  const codeToRow = {};
  const lastM = master.getLastRow();
  if (lastM > mCode.row) {
    const codes = master.getRange(mCode.row + 1, mCode.col, lastM - mCode.row, 1).getValues();
    codes.forEach((v, i) => { if (v[0]) codeToRow[String(v[0]).trim()] = mCode.row + 1 + i; });
  }

  ['PURCHASES','STOCK MOVEMENT APPROVAL LOG'].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const codeCol = findHeaderCol_(sh, ['ITEM CODE','CODE'], 10);
    const costCol = findHeaderCol_(sh, ['UNIT COST','COST PRICE','UNIT VALUE'], 10);
    const statusCol = findHeaderCol_(sh, ['STATUS'], 10);
    if (!codeCol || !costCol) return;
    const start = Math.max(codeCol.row, costCol.row) + 1;
    for (let r = start; r <= sh.getLastRow(); r++) {
      const status = statusCol ? key_(sh.getRange(r, statusCol.col).getValue()) : 'APPROVED';
      if (status === 'REJECTED') continue;
      const code = String(sh.getRange(r, codeCol.col).getValue() || '').trim();
      const cost = sh.getRange(r, costCol.col).getValue();
      if (!code || isNaN(Number(cost)) || cost === '') continue;
      const mRow = codeToRow[code];
      if (mRow) master.getRange(mRow, mCost.col).setValue(Number(cost));
    }
  });
}
