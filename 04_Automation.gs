/****************************************************
 AUTOMATION: approval timestamps, edit log, master price cost update, quantity sync.
 Kept separate from protection/permission logic.
****************************************************/

function onEdit(e) {
  try {
    if (!e || !e.range) return;

    // Check if the user has specific permission to edit this sheet from SYSTEM_ACCESS
    if (checkSheetPermission_(e)) return;

    // Prevent editing approved rows in STOCK MOVEMENT APPROVAL LOG
    if (preventApprovedEdit_(e)) return;

    auditEdit_(e);
    handleApprovalEdit_(e);
    handleMasterPriceCostUpdate_(e);
    logStockChange_(e);
  } catch (err) {
    log_('ON_EDIT_ERROR', err.message || err);
  }
}

/**
 * Enforces 'Sheets Controlled' from SYSTEM_ACCESS at the edit level.
 */
function checkSheetPermission_(e) {
  const userEmail = Session.getActiveUser().getEmail();
  const sheetName = e.range.getSheet().getName();

  const systemSheets = ['SYSTEM_ACCESS','SYSTEM_LOGS','SYSTEM_SETTINGS','EOM EDIT LOG','STOCK CHANGE LOG'];
  if (systemSheets.includes(sheetName.toUpperCase())) return false;

  if (isAuthorizedApprover_(userEmail)) return false;

  const authorizedUsers = usersForSheet_(sheetName);
  if (!authorizedUsers.map(u => u.toLowerCase()).includes(userEmail.toLowerCase())) {
    e.range.setValue(e.oldValue || '');
    uiAlert_('Access Denied: You are not assigned to control the sheet "' + sheetName + '" in SYSTEM_ACCESS.');
    return true;
  }
  return false;
}

/**
 * Prevents non-approvers from editing rows that are already marked as APPROVED.
 */
function preventApprovedEdit_(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== 'STOCK MOVEMENT APPROVAL LOG') return false;

  const statusCell = findHeaderCol_(sh, ['STATUS'], 10);
  if (!statusCell) return false;

  const row = e.range.getRow();
  if (row <= statusCell.row) return false;

  const status = key_(sh.getRange(row, statusCell.col).getValue());
  if (status !== 'APPROVED') return false;

  const userEmail = Session.getActiveUser().getEmail();
  if (!isAuthorizedApprover_(userEmail)) {
    e.range.setValue(e.oldValue || '');
    uiAlert_('Action Denied: This movement has already been APPROVED and is now locked. Only MDs or Owners can modify it.');
    return true;
  }
  return false;
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

function isAuthorizedApprover_(email) {
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

  if (!isAuthorizedApprover_(userEmail)) {
    e.range.setValue(e.oldValue || '');
    uiAlert_('Access Denied: Only MDs or Owners can approve or reject stock movements.');
    return;
  }

  const approvedBy = findHeaderCol_(sh, ['APPROVED BY','APPROVER'], 10);
  const approvedDate = findHeaderCol_(sh, ['APPROVAL DATE','APPROVED DATE'], 10);
  if (approvedBy) sh.getRange(e.range.getRow(), approvedBy.col).setValue(userEmail);
  if (approvedDate) sh.getRange(e.range.getRow(), approvedDate.col).setValue(new Date());

  if (status === 'APPROVED') {
    syncSingleMovementToMasterPriceList_(sh, e.range.getRow());
    syncSingleMovementToDepartment_(sh, e.range.getRow());
  }
}

function handleMasterPriceCostUpdate_(e) {
  const sheetName = e.range.getSheet().getName();
  if (sheetName !== 'PURCHASES') return;
  if (e.range.getColumn() === 11 && e.range.getRow() >= 5) {
    syncMasterPriceCostsFromApprovedMovements();
  }
}

function syncSingleMovementToMasterPriceList_(sh, row) {
  const codeHeader = findHeaderCol_(sh, ['ITEM CODE','CODE'], 10);
  const costHeader = findHeaderCol_(sh, ['UNIT COST','COST PRICE','UNIT VALUE','PURCHASE UNIT PRICE'], 10);
  if (!codeHeader || !costHeader) return;

  const code = String(sh.getRange(row, codeHeader.col).getValue() || '').trim();
  const cost = sh.getRange(row, costHeader.col).getValue();
  if (!code || isNaN(Number(cost))) return;

  const master = SpreadsheetApp.getActive().getSheetByName(CONFIG.MASTER_PRICE_LIST) || SpreadsheetApp.getActive().getSheetByName('MASTER PRICE LIST');
  if (!master) return;

  const mCodeHeader = findHeaderCol_(master, ['ITEM CODE','CODE'], 10) || {row: 1, col: 2};
  const mCostHeader = findHeaderCol_(master, ['COST PRICE','UNIT COST','COST'], 10) || {row: 1, col: 3};

  const mData = master.getRange(mCodeHeader.row + 1, mCodeHeader.col, Math.max(master.getLastRow() - mCodeHeader.row, 1), 1).getValues();
  for (let i = 0; i < mData.length; i++) {
    if (String(mData[i][0]).trim() === code) {
      master.getRange(mCodeHeader.row + 1 + i, mCostHeader.col).setValue(Number(cost));
      break;
    }
  }
}

function syncSingleMovementToDepartment_(sh, row) {
  const deptCol = findHeaderCol_(sh, ['DEPARTMENT'], 10);
  const codeCol = findHeaderCol_(sh, ['ITEM CODE','CODE'], 10);
  const qtyCol = findHeaderCol_(sh, ['QTY','QUANTITY'], 10);
  const typeCol = findHeaderCol_(sh, ['MOVEMENT TYPE'], 10);
  if (!deptCol || !codeCol || !qtyCol || !typeCol) return;

  const deptName = String(sh.getRange(row, deptCol.col).getValue() || '').trim();
  const code = String(sh.getRange(row, codeCol.col).getValue() || '').trim();
  const qty = Number(sh.getRange(row, qtyCol.col).getValue() || 0);
  const type = key_(sh.getRange(row, typeCol.col).getValue());

  const deptSh = SpreadsheetApp.getActive().getSheetByName(deptName);
  if (!deptSh) return;

  const dCode = findHeaderCol_(deptSh, ['ITEM CODE','CODE'], 10);
  if (!dCode) return;

  let targetCol = null;
  if (type === 'SOLD' || type === 'UTILIZED') targetCol = findHeaderCol_(deptSh, ['SOLD'], 10);
  else if (type === 'DAMAGED' || type === 'DAMAGE') targetCol = findHeaderCol_(deptSh, ['DAMAGED','DAMAGE'], 10);
  else if (type === 'ISSUED') targetCol = findHeaderCol_(deptSh, ['ISSUED'], 10);
  else if (type === 'ADDED') targetCol = findHeaderCol_(deptSh, ['ADDED STOCK','ADDED'], 10);

  if (!targetCol) return;

  const lastD = deptSh.getLastRow();
  const dData = deptSh.getRange(dCode.row + 1, dCode.col, Math.max(lastD - dCode.row, 1), 1).getValues();
  for (let i = 0; i < dData.length; i++) {
    if (String(dData[i][0]).trim() === code) {
      const currentQty = Number(deptSh.getRange(dCode.row + 1 + i, targetCol.col).getValue() || 0);
      deptSh.getRange(dCode.row + 1 + i, targetCol.col).setValue(currentQty + qty);
      break;
    }
  }
}

function syncMasterPriceCostsFromApprovedMovements() {
  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName(CONFIG.MASTER_PRICE_LIST) || ss.getSheetByName('MASTER PRICE LIST');
  if (!master) throw new Error('MASTER_PRICELIST not found.');

  const mCodeHeader = findHeaderCol_(master, ['ITEM CODE','CODE'], 10) || {row: 1, col: 2};
  const mCostHeader = findHeaderCol_(master, ['COST PRICE','UNIT COST','COST'], 10) || {row: 1, col: 3};

  const lastM = master.getLastRow();
  if (lastM <= mCodeHeader.row) return;

  const mData = master.getRange(mCodeHeader.row + 1, 1, lastM - mCodeHeader.row, master.getLastColumn()).getValues();
  const codeToIdx = {};
  mData.forEach((row, i) => {
    const code = String(row[mCodeHeader.col - 1] || '').trim();
    if (code) codeToIdx[code] = i;
  });

  let updates = 0;
  const pSh = ss.getSheetByName('PURCHASES');
  if (pSh) {
    const pCodeHeader = findHeaderCol_(pSh, ['ITEM CODE','CODE'], 10) || {row: 1, col: 2};
    const lastP = pSh.getLastRow();
    if (lastP >= 5) {
      const pData = pSh.getRange(5, 1, lastP - 4, Math.max(pCodeHeader.col, 11)).getValues();
      pData.forEach(row => {
        const code = String(row[pCodeHeader.col - 1] || '').trim();
        const cost = row[10]; // Column K
        if (code && codeToIdx[code] !== undefined && cost !== '' && !isNaN(Number(cost))) {
          const idx = codeToIdx[code];
          if (Number(mData[idx][mCostHeader.col - 1]) !== Number(cost)) {
            mData[idx][mCostHeader.col - 1] = Number(cost);
            updates++;
          }
        }
      });
    }
  }

  const logSh = ss.getSheetByName('STOCK MOVEMENT APPROVAL LOG');
  if (logSh) {
    const lCode = findHeaderCol_(logSh, ['ITEM CODE','CODE'], 10);
    const lCost = findHeaderCol_(logSh, ['UNIT COST','COST PRICE','UNIT VALUE','PURCHASE UNIT PRICE'], 10);
    const lStatus = findHeaderCol_(logSh, ['STATUS'], 10);
    if (lCode && lCost && lStatus) {
      const lastL = logSh.getLastRow();
      if (lastL > lStatus.row) {
        const lData = logSh.getRange(lStatus.row + 1, 1, lastL - lStatus.row, logSh.getLastColumn()).getValues();
        lData.forEach(row => {
          if (key_(row[lStatus.col - 1]) === 'APPROVED') {
            const code = String(row[lCode.col - 1] || '').trim();
            const cost = row[lCost.col - 1];
            if (code && codeToIdx[code] !== undefined && cost !== '' && !isNaN(Number(cost))) {
              const idx = codeToIdx[code];
              if (Number(mData[idx][mCostHeader.col - 1]) !== Number(cost)) {
                mData[idx][mCostHeader.col - 1] = Number(cost);
                updates++;
              }
            }
          }
        });
      }
    }
  }

  if (updates > 0) {
    const outRange = master.getRange(mCodeHeader.row + 1, mCostHeader.col, mData.length, 1);
    const outValues = mData.map(r => [r[mCostHeader.col - 1]]);
    outRange.setValues(outValues);
    log_('MASTER_PRICE_SYNC', 'Updated cost prices: ' + updates);
    uiAlert_('Master Price List cost sync complete. Updates made: ' + updates);
  }
}

function syncQuantitiesFromApprovedMovements() {
  const ss = SpreadsheetApp.getActive();
  const logSh = ss.getSheetByName('STOCK MOVEMENT APPROVAL LOG');
  if (!logSh) return;

  const lDept = findHeaderCol_(logSh, ['DEPARTMENT'], 10);
  const lCode = findHeaderCol_(logSh, ['ITEM CODE','CODE'], 10);
  const lQty = findHeaderCol_(logSh, ['QTY','QUANTITY'], 10);
  const lType = findHeaderCol_(logSh, ['MOVEMENT TYPE'], 10);
  const lStatus = findHeaderCol_(logSh, ['STATUS'], 10);

  if (!lDept || !lCode || !lQty || !lType || !lStatus) return;

  const lastL = logSh.getLastRow();
  if (lastL <= lStatus.row) return;

  const lData = logSh.getRange(lStatus.row + 1, 1, lastL - lStatus.row, logSh.getLastColumn()).getValues();
  const aggregates = {};

  lData.forEach(row => {
    if (key_(row[lStatus.col - 1]) !== 'APPROVED') return;
    const dept = String(row[lDept.col - 1] || '').trim();
    const code = String(row[lCode.col - 1] || '').trim();
    const qty = Number(row[lQty.col - 1] || 0);
    const type = key_(row[lType.col - 1]);

    if (!dept || !code || isNaN(qty)) return;
    if (!aggregates[dept]) aggregates[dept] = {};
    if (!aggregates[dept][code]) aggregates[dept][code] = { sold: 0, damaged: 0, issued: 0, added: 0 };

    if (type === 'SOLD' || type === 'UTILIZED') aggregates[dept][code].sold += qty;
    else if (type === 'DAMAGED' || type === 'DAMAGE') aggregates[dept][code].damaged += qty;
    else if (type === 'ISSUED') aggregates[dept][code].issued += qty;
    else if (type === 'ADDED') aggregates[dept][code].added += qty;
  });

  let updates = 0;
  Object.keys(aggregates).forEach(deptName => {
    const deptSh = ss.getSheetByName(deptName);
    if (!deptSh) return;

    const dCode = findHeaderCol_(deptSh, ['ITEM CODE','CODE'], 10);
    if (!dCode) return;

    const dSold = findHeaderCol_(deptSh, ['SOLD'], 10);
    const dDamaged = findHeaderCol_(deptSh, ['DAMAGED','DAMAGE'], 10);
    const dIssued = findHeaderCol_(deptSh, ['ISSUED'], 10);
    const dAdded = findHeaderCol_(deptSh, ['ADDED STOCK','ADDED'], 10);

    const lastD = deptSh.getLastRow();
    if (lastD <= dCode.row) return;

    const dRange = deptSh.getRange(dCode.row + 1, 1, lastD - dCode.row, deptSh.getLastColumn());
    const dData = dRange.getValues();
    let deptUpdated = false;

    dData.forEach((row, i) => {
      const code = String(row[dCode.col - 1] || '').trim();
      const move = aggregates[deptName][code];
      if (!move) return;

      if (dSold && move.sold > 0) { row[dSold.col - 1] = move.sold; deptUpdated = true; }
      if (dDamaged && move.damaged > 0) { row[dDamaged.col - 1] = move.damaged; deptUpdated = true; }
      if (dIssued && move.issued > 0) { row[dIssued.col - 1] = move.issued; deptUpdated = true; }
      if (dAdded && move.added > 0) { row[dAdded.col - 1] = move.added; deptUpdated = true; }
    });

    if (deptUpdated) {
      dRange.setValues(dData);
      updates++;
    }
  });

  if (updates > 0) log_('QUANTITY_SYNC', 'Updated quantities in ' + updates + ' departments.');
}

function syncMasterPriceItemsFromDepartments() {
  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName(CONFIG.MASTER_PRICE_LIST) || ss.getSheetByName('MASTER PRICE LIST');
  if (!master) throw new Error('MASTER_PRICELIST not found.');

  const mCodeCol = (findHeaderCol_(master, ['ITEM CODE','CODE'], 10) || {row: 1, col: 2}).col;
  const mItemCol = (findHeaderCol_(master, ['ITEM','DESCRIPTION','PRODUCT'], 10) || {row: 1, col: 1}).col;

  const existingCodes = new Set();
  const lastM = master.getLastRow();
  if (lastM > 1) {
    master.getRange(2, mCodeCol, lastM - 1, 1).getValues().forEach(v => {
      if (v[0]) existingCodes.add(String(v[0]).trim());
    });
  }

  const newItems = [];
  const targetSheets = ss.getSheets().filter(sh => isCSSheet_(sh.getName()) || isWeeklyMRSheet_(sh.getName()));

  targetSheets.forEach(sh => {
    const codeHeader = findHeaderCol_(sh, ['ITEM CODE','CODE'], 10);
    const itemHeader = findHeaderCol_(sh, ['ITEM','DESCRIPTION'], 10);
    if (!codeHeader || !itemHeader) return;

    const lastRow = sh.getLastRow();
    if (lastRow <= codeHeader.row) return;

    const data = sh.getRange(codeHeader.row + 1, 1, lastRow - codeHeader.row, sh.getLastColumn()).getValues();
    data.forEach(row => {
      const code = String(row[codeHeader.col - 1] || '').trim();
      const item = String(row[itemHeader.col - 1] || '').trim();
      if (code && !existingCodes.has(code)) {
        const newRow = [];
        newRow[mItemCol - 1] = item;
        newRow[mCodeCol - 1] = code;
        newItems.push(newRow);
        existingCodes.add(code);
      }
    });
  });

  if (newItems.length > 0) {
    master.getRange(master.getLastRow() + 1, 1, newItems.length, newItems[0].length).setValues(newItems);
    log_('MASTER_PRICE_ITEM_SYNC', 'New items added: ' + newItems.length);
    uiAlert_('Sync complete. New items added to Master Price List: ' + newItems.length);
  }
}

function refreshApprovalTimestamps() {
  const sh = SpreadsheetApp.getActive().getSheetByName('STOCK MOVEMENT APPROVAL LOG');
  if (!sh) throw new Error('STOCK MOVEMENT APPROVAL LOG not found.');
  const statusCell = findHeaderCol_(sh, ['STATUS'], 10);
  const approvedBy = findHeaderCol_(sh, ['APPROVED BY','APPROVER'], 10);
  const approvedDate = findHeaderCol_(sh, ['APPROVAL DATE','APPROVED DATE'], 10);
  if (!statusCell || !approvedBy || !approvedDate) throw new Error('Required approval columns not found.');
  const last = sh.getLastRow();
  if (last <= statusCell.row) return;

  const data = sh.getRange(statusCell.row + 1, 1, last - statusCell.row, sh.getLastColumn()).getValues();
  let count = 0;
  data.forEach((row, i) => {
    const status = key_(row[statusCell.col - 1]);
    if (['APPROVED','REJECTED','UNDER REVIEW'].includes(status)) {
      if (!row[approvedBy.col - 1]) {
        sh.getRange(statusCell.row + 1 + i, approvedBy.col).setValue('SYSTEM CHECK');
        sh.getRange(statusCell.row + 1 + i, approvedDate.col).setValue(new Date());
        count++;
      }
    }
  });
  uiAlert_('Approval timestamp refresh complete. Rows updated: ' + count);
}
