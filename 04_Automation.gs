/****************************************************
 AUTOMATION: approval timestamps, edit log, master price cost update.
 Kept separate from protection/permission logic.
****************************************************/

function onEdit(e) {
  try {
    if (!e || !e.range) return;
    auditEdit_(e);
    handleApprovalEdit_(e);
    handleMasterPriceCostUpdate_(e);
    logStockChange_(e);
  } catch (err) {
    log_('ON_EDIT_ERROR', err.message || err);
  }
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

function logStockChange_(e) {
  try {
    const sh = e.range.getSheet();
    const name = sh.getName();
    // Only log significant stock-related changes
    if (!isCSSheet_(name) && !isWeeklyMRSheet_(name) && name !== 'PURCHASES' && name !== 'STOCK MOVEMENT APPROVAL LOG') return;

    const ss = SpreadsheetApp.getActive();
    let logSh = ss.getSheetByName(CONFIG.STOCK_CHANGE_LOG);
    if (!logSh) return;
    if (logSh.getLastRow() === 0) logSh.appendRow(['Timestamp','User','Sheet','Cell','Change']);
    logSh.appendRow([new Date(), Session.getActiveUser().getEmail(), name, e.range.getA1Notation(), (e.oldValue || 'BLANK') + ' -> ' + (e.value || 'BLANK')]);
  } catch (err) {}
}

function findHeaderCol_(sheet, names, headerRows) {
  headerRows = headerRows || 10;
  const maxRows = Math.min(headerRows, sheet.getMaxRows());
  if (maxRows === 0) return null;
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

function isOwner_(email) {
  const records = getAccessRecords_();
  const user = records.find(r => r.email.toLowerCase() === email.toLowerCase());
  if (!user) return false;
  return CONFIG.OWNER_ROLES.includes(user.role) || user.sheets.includes('ALL') || user.sheets.includes('OWNER SHEETS');
}

function handleApprovalEdit_(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== 'STOCK MOVEMENT APPROVAL LOG') return;
  const statusCell = findHeaderCol_(sh, ['STATUS'], 10);
  if (!statusCell || e.range.getColumn() !== statusCell.col || e.range.getRow() <= statusCell.row) return;

  const userEmail = Session.getActiveUser().getEmail();
  const status = key_(e.value);
  if (!['APPROVED','REJECTED','UNDER REVIEW','PENDING'].includes(status)) return;

  // Owner-only approval enforcement
  if (!isOwner_(userEmail)) {
    e.range.setValue(e.oldValue || '');
    uiAlert_('Access Denied: Only Owners or MDs can approve or reject stock movements.');
    return;
  }

  const approvedBy = findHeaderCol_(sh, ['APPROVED BY','APPROVER'], 10);
  const approvedDate = findHeaderCol_(sh, ['APPROVAL DATE','APPROVED DATE'], 10);
  if (approvedBy) sh.getRange(e.range.getRow(), approvedBy.col).setValue(userEmail);
  if (approvedDate) sh.getRange(e.range.getRow(), approvedDate.col).setValue(new Date());
}

function handleMasterPriceCostUpdate_(e) {
  const sheetName = e.range.getSheet().getName();
  if (!['PURCHASES','STOCK MOVEMENT APPROVAL LOG'].includes(sheetName)) return;

  let isCostCol = false;
  if (sheetName === 'PURCHASES' && e.range.getColumn() === 11) { // Column K
    isCostCol = true;
  } else {
    const header = findHeaderCol_(e.range.getSheet(), ['UNIT COST','COST PRICE','UNIT VALUE','PURCHASE UNIT PRICE'], 10);
    if (header && e.range.getColumn() === header.col) isCostCol = true;
  }

  if (isCostCol) syncMasterPriceCostsFromApprovedMovements();
}

function syncMasterPriceCostsFromApprovedMovements() {
  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName(CONFIG.MASTER_PRICE_LIST) || ss.getSheetByName('MASTER PRICE LIST');
  if (!master) throw new Error('MASTER_PRICELIST not found.');

  const mCode = findHeaderCol_(master, ['ITEM CODE','CODE'], 10) || {row: 1, col: 1};
  // Explicitly target Column C (3) or search for Cost header
  const mCost = findHeaderCol_(master, ['COST PRICE','UNIT COST','COST'], 10) || {row: 1, col: 3};

  const codeToRow = {};
  const lastM = master.getLastRow();
  if (lastM > mCode.row) {
    const codes = master.getRange(mCode.row + 1, mCode.col, lastM - mCode.row, 1).getValues();
    codes.forEach((v, i) => { if (v[0]) codeToRow[String(v[0]).trim()] = mCode.row + 1 + i; });
  }

  let updates = 0;
  ['PURCHASES','STOCK MOVEMENT APPROVAL LOG'].forEach(sheetName => {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return;
    const codeCol = findHeaderCol_(sh, ['ITEM CODE','CODE'], 10) || {row: 1, col: (sheetName === 'PURCHASES' ? 2 : 1)};
    // Explicitly target Purchases K (11) if header not found
    const costCol = findHeaderCol_(sh, ['UNIT COST','COST PRICE','UNIT VALUE','PURCHASE UNIT PRICE'], 10) || {row: 1, col: (sheetName === 'PURCHASES' ? 11 : 1)};
    const statusCol = findHeaderCol_(sh, ['STATUS'], 10);

    const start = Math.max(codeCol.row, costCol.row, statusCol ? statusCol.row : 1) + 1;
    for (let r = start; r <= sh.getLastRow(); r++) {
      const status = statusCol ? key_(sh.getRange(r, statusCol.col).getValue()) : 'APPROVED';
      if (status === 'REJECTED') continue;

      const code = String(sh.getRange(r, codeCol.col).getValue() || '').trim();
      const cost = sh.getRange(r, costCol.col).getValue();
      if (!code || cost === '' || isNaN(Number(cost))) continue;

      const mRow = codeToRow[code];
      if (!mRow) continue;

      const current = master.getRange(mRow, mCost.col).getValue();
      if (Number(current) !== Number(cost)) {
        master.getRange(mRow, mCost.col).setValue(Number(cost));
        updates++;
      }
    }
  });
  log_('MASTER_PRICE_SYNC', 'Updated cost prices: ' + updates);
  uiAlert_('Master Price List cost sync complete. Updates made: ' + updates);
}

function syncMasterPriceItemsFromDepartments() {
  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName(CONFIG.MASTER_PRICE_LIST) || ss.getSheetByName('MASTER PRICE LIST');
  if (!master) throw new Error('MASTER_PRICELIST not found.');

  const mCode = findHeaderCol_(master, ['ITEM CODE','CODE'], 10) || {row: 1, col: 2};
  const mItem = findHeaderCol_(master, ['ITEM','DESCRIPTION','PRODUCT'], 10) || {row: 1, col: 1};

  const existingCodes = new Set();
  const lastM = master.getLastRow();
  if (lastM > mCode.row) {
    master.getRange(mCode.row + 1, mCode.col, lastM - mCode.row, 1).getValues().forEach(v => {
      if (v[0]) existingCodes.add(String(v[0]).trim());
    });
  }

  let newItems = 0;
  const sheetsToSync = ss.getSheets().filter(sh => isCSSheet_(sh.getName()) || isWeeklyMRSheet_(sh.getName()));

  sheetsToSync.forEach(sh => {
    const codeCol = findHeaderCol_(sh, ['ITEM CODE','CODE'], 10);
    const itemCol = findHeaderCol_(sh, ['ITEM','DESCRIPTION'], 10);
    if (!codeCol || !itemCol) return;

    const lastRow = sh.getLastRow();
    if (lastRow <= codeCol.row) return;

    const data = sh.getRange(codeCol.row + 1, 1, lastRow - codeCol.row, sh.getLastColumn()).getValues();
    data.forEach(row => {
      const code = String(row[codeCol.col - 1] || '').trim();
      const item = String(row[itemCol.col - 1] || '').trim();
      if (code && !existingCodes.has(code)) {
        master.appendRow([item, code, '']); // Minimal entry: Item, Code, [Cost]
        existingCodes.add(code);
        newItems++;
      }
    });
  });

  log_('MASTER_PRICE_ITEM_SYNC', 'New items added to Master Price List: ' + newItems);
  uiAlert_('Master Price List item sync complete. New items added: ' + newItems);
}

function refreshApprovalTimestamps() {
  const sh = SpreadsheetApp.getActive().getSheetByName('STOCK MOVEMENT APPROVAL LOG');
  if (!sh) throw new Error('STOCK MOVEMENT APPROVAL LOG not found.');
  const statusCell = findHeaderCol_(sh, ['STATUS'], 10);
  const approvedBy = findHeaderCol_(sh, ['APPROVED BY','APPROVER'], 10);
  const approvedDate = findHeaderCol_(sh, ['APPROVAL DATE','APPROVED DATE'], 10);
  if (!statusCell || !approvedBy || !approvedDate) throw new Error('Required approval columns not found.');
  const last = sh.getLastRow();
  let count = 0;
  for (let r = statusCell.row + 1; r <= last; r++) {
    const status = key_(sh.getRange(r, statusCell.col).getValue());
    if (['APPROVED','REJECTED','UNDER REVIEW'].includes(status)) {
      if (!sh.getRange(r, approvedBy.col).getValue()) sh.getRange(r, approvedBy.col).setValue('SYSTEM CHECK');
      if (!sh.getRange(r, approvedDate.col).getValue()) sh.getRange(r, approvedDate.col).setValue(new Date());
      count++;
    }
  }
  uiAlert_('Approval timestamp refresh complete. Rows checked: ' + count);
}
