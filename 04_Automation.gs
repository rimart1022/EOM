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

  if (!isOwner_(userEmail)) {
    e.range.setValue(e.oldValue || '');
    uiAlert_('Access Denied: Only Owners or MDs can approve or reject stock movements.');
    return;
  }

  const approvedBy = findHeaderCol_(sh, ['APPROVED BY','APPROVER'], 10);
  const approvedDate = findHeaderCol_(sh, ['APPROVAL DATE','APPROVED DATE'], 10);
  if (approvedBy) sh.getRange(e.range.getRow(), approvedBy.col).setValue(userEmail);
  if (approvedDate) sh.getRange(e.range.getRow(), approvedDate.col).setValue(new Date());

  if (status === 'APPROVED') syncMasterPriceCostsFromApprovedMovements();
}

function handleMasterPriceCostUpdate_(e) {
  const sheetName = e.range.getSheet().getName();
  if (sheetName !== 'PURCHASES') return;
  if (e.range.getColumn() === 11 && e.range.getRow() >= 5) {
    syncMasterPriceCostsFromApprovedMovements();
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

  // Batch read Master Price List
  const mData = master.getRange(mCodeHeader.row + 1, 1, lastM - mCodeHeader.row, master.getLastColumn()).getValues();
  const codeToIdx = {};
  mData.forEach((row, i) => {
    const code = String(row[mCodeHeader.col - 1] || '').trim();
    if (code) codeToIdx[code] = i;
  });

  let updates = 0;

  // 1. Sync from PURCHASES (K -> C)
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

  // 2. Sync from APPROVED Stock Movement Log
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

  // Batch commit updates
  if (updates > 0) {
    const outRange = master.getRange(mCodeHeader.row + 1, mCostHeader.col, mData.length, 1);
    const outValues = mData.map(r => [r[mCostHeader.col - 1]]);
    outRange.setValues(outValues);
    log_('MASTER_PRICE_SYNC', 'Updated cost prices: ' + updates);
    uiAlert_('Master Price List cost sync complete. Updates made: ' + updates);
  }
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
