/****************************************************
 REPORTS restored.
 These are simple, safe report generators from existing workbook data.
****************************************************/

function ensureReportSheet_(name, headers) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clear();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#d9ead3');
  return sh;
}

function approvalLogRows_() {
  const sh = SpreadsheetApp.getActive().getSheetByName('STOCK MOVEMENT APPROVAL LOG');
  if (!sh) return {headers: [], rows: []};
  const headerRow = 5;
  const headers = sh.getRange(headerRow, 1, 1, sh.getLastColumn()).getValues()[0];
  const rows = sh.getLastRow() > headerRow ? sh.getRange(headerRow + 1, 1, sh.getLastRow() - headerRow, sh.getLastColumn()).getValues() : [];
  return {headers, rows};
}

function colIndex_(headers, names) {
  const keys = headers.map(key_);
  for (const n of names) {
    const idx = keys.indexOf(key_(n));
    if (idx >= 0) return idx;
  }
  return -1;
}

function movementReport_(movementTypes, targetSheet) {
  const data = approvalLogRows_();
  const headers = data.headers;
  const rows = data.rows;
  const cType = colIndex_(headers, ['MOVEMENT TYPE']);
  const cStatus = colIndex_(headers, ['STATUS']);
  const outHeaders = ['Report Date','Source Row'].concat(headers);
  const out = [];
  rows.forEach((r, i) => {
    const type = cType >= 0 ? key_(String(r[cType])) : '';
    const status = cStatus >= 0 ? key_(String(r[cStatus])) : '';
    if (movementTypes.includes(type) && status !== 'REJECTED') out.push([new Date(), i + 6].concat(r));
  });
  const sh = ensureReportSheet_(targetSheet, outHeaders);
  if (out.length) sh.getRange(2, 1, out.length, outHeaders.length).setValues(out);
  sh.autoResizeColumns(1, Math.min(outHeaders.length, 12));
  return out.length;
}

function generateDamageReport() {
  const count = movementReport_(['DAMAGE','DAMAGED'], 'DAMAGE REPORT');
  uiAlert_('Damage Report generated. Rows: ' + count);
}
function generateIssuedStockReport() {
  const count = movementReport_(['ISSUED','ISSUED STOCK','ISSUE TO DEPARTMENT'], 'ISSUED STOCK REPORT');
  uiAlert_('Issued Stock Report generated. Rows: ' + count);
}
function generateUtilizedReport() {
  const count = movementReport_(['UTILIZED'], 'UTILIZED REPORT');
  uiAlert_('Utilized Report generated. Rows: ' + count);
}
function generateStaffLiabilityReport() {
  const count = movementReport_(['DAMAGE','DAMAGED'], 'STAFF LIABILITY REPORT');
  uiAlert_('Staff Liability Report generated from approved/non-rejected damages. Rows: ' + count);
}

function clearAllReports() {
  const ss = SpreadsheetApp.getActive();
  const reports = ['DAMAGE REPORT','ISSUED STOCK REPORT','UTILIZED REPORT','STAFF LIABILITY REPORT','STOCK AUDIT SUMMARY','FORMULA ERROR CHECK'];
  reports.forEach(r => {
    const sh = ss.getSheetByName(r);
    if (sh) sh.clear();
  });
  uiAlert_('All report sheets cleared.');
}

function generateStockAuditSummary() {
  const ss = SpreadsheetApp.getActive();
  const csNames = ['CS STORE','CS MINI-MART','CS LAUNDRY','CS BAR','CS RESTAURANT','CS KITCHEN'];
  const shHeaders = ['Timestamp','Sheet','Item Code','Item','Opening','Added','Issued','Sold','Damaged','Physical Count','Closing Stock','Variance'];
  const sh = ensureReportSheet_('STOCK AUDIT SUMMARY', shHeaders);
  const out = [];

  csNames.forEach(name => {
    const s = ss.getSheetByName(name);
    if (!s) return;
    const last = s.getLastRow();
    if (last < 5) return;
    const data = s.getRange(1, 1, last, Math.min(s.getLastColumn(), 50)).getValues();
    const h = data[3]; // Row 4 headers
    const cIdx = {
      code: colIndex_(h, ['ITEM CODE','CODE']),
      name: colIndex_(h, ['ITEM NAME','ITEM','PARTICULARS']),
      open: colIndex_(h, ['OPENING STOCK','OPENING']),
      add: colIndex_(h, ['ADDED STOCK','ADDED']),
      iss: colIndex_(h, ['ISSUED STOCK','ISSUED']),
      sold: colIndex_(h, ['SOLD','SALES']),
      dmg: colIndex_(h, ['DAMAGED','DAMAGE']),
      phys: colIndex_(h, ['PHYSICAL COUNT','PHYSICAL']),
      close: colIndex_(h, ['CLOSING STOCK','CLOSING']),
      var: colIndex_(h, ['VARIANCE','VAR'])
    };

    for (let r = 4; r < data.length; r++) {
      const row = data[r];
      const itemCode = cIdx.code >= 0 ? row[cIdx.code] : '';
      const itemName = cIdx.name >= 0 ? row[cIdx.name] : '';
      if (!itemCode && !itemName) continue;

      const getVal = (idx) => (idx >= 0 && typeof row[idx] === 'number') ? row[idx] : 0;

      out.push([
        new Date(),
        name,
        itemCode,
        itemName,
        getVal(cIdx.open),
        getVal(cIdx.add),
        getVal(cIdx.iss),
        getVal(cIdx.sold),
        getVal(cIdx.dmg),
        getVal(cIdx.phys),
        getVal(cIdx.close),
        getVal(cIdx.var)
      ]);
    }
  });

  if (out.length) {
    sh.getRange(2, 1, out.length, shHeaders.length).setValues(out);
  }
  uiAlert_('Stock Audit Summary generated. Rows: ' + out.length);
}

function checkFormulaErrors() {
  const ss = SpreadsheetApp.getActive();
  const sh = ensureReportSheet_('FORMULA ERROR CHECK', ['Timestamp','Sheet','Cell','Formula','Displayed Value']);
  const errors = ['#REF!','#VALUE!','#DIV/0!','#NAME?','#N/A','#ERROR!'];
  const out = [];
  ss.getSheets().forEach(s => {
    const range = s.getDataRange();
    if (range.getNumRows() === 0) return;
    const formulas = range.getFormulas();
    const displays = range.getDisplayValues();
    formulas.forEach((row, r) => row.forEach((f, c) => {
      if (!f) return;
      const d = displays[r][c];
      const dStr = String(d);
      const fStr = String(f);
      if (errors.some(err => dStr.includes(err)) || errors.some(err => fStr.includes(err))) {
        out.push([new Date(), s.getName(), s.getRange(r + 1, c + 1).getA1Notation(), f, d]);
      }
    }));
  });
  if (out.length) sh.getRange(2, 1, out.length, 5).setValues(out);
  uiAlert_(out.length ? 'Formula errors found: ' + out.length : 'No formula error strings found.');
}
