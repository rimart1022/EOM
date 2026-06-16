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
  const headerRow = 6; // Carlisle approval log entries usually start around row 7.
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
    const type = cType >= 0 ? key_(r[cType]) : '';
    const status = cStatus >= 0 ? key_(r[cStatus]) : '';
    if (movementTypes.includes(type) && status !== 'REJECTED') out.push([new Date(), i + 7].concat(r));
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
function generateStaffLiabilityReport() {
  const count = movementReport_(['DAMAGE','DAMAGED'], 'STAFF LIABILITY REPORT');
  uiAlert_('Staff Liability Report generated from approved/non-rejected damages. Rows: ' + count);
}

function generateStockAuditSummary() {
  const ss = SpreadsheetApp.getActive();
  const csNames = ['CS MINI-MART','CS LAUNDRY','CS BAR','CS RESTAURANT','CS STORE','CS KITCHEN'];
  const sh = ensureReportSheet_('STOCK AUDIT SUMMARY', ['Timestamp','Sheet','Item Code','Item','Opening','Added','Issued','Sold','Damaged','Physical Count','Closing Stock','Variance']);
  const out = [];
  csNames.forEach(name => {
    const s = ss.getSheetByName(name);
    if (!s) return;
    const last = s.getLastRow();
    if (last < 5) return;
    const values = s.getRange(1, 1, last, s.getLastColumn()).getValues();
    // Generic scan: include visible rows with item/code in first columns and any non-zero variance-like last columns.
    for (let r = 4; r < values.length; r++) {
      const row = values[r];
      if (!row[0] && !row[1]) continue;
      const itemCode = row[0] || row[1] || '';
      const item = row[1] || row[2] || '';
      const nums = row.filter(v => typeof v === 'number');
      if (nums.length) out.push([new Date(), name, itemCode, item].concat(nums.slice(0, 8)));
    }
  });
  if (out.length) sh.getRange(2, 1, out.length, Math.min(out[0].length, 12)).setValues(out.map(r => r.slice(0, 12)));
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
      if (errors.some(err => String(d).includes(err)) || errors.some(err => String(f).includes(err))) {
        out.push([new Date(), s.getName(), s.getRange(r + 1, c + 1).getA1Notation(), f, d]);
      }
    }));
  });
  if (out.length) sh.getRange(2, 1, out.length, 5).setValues(out);
  uiAlert_(out.length ? 'Formula errors found: ' + out.length : 'No formula error strings found.');
}
