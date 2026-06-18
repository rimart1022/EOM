/****************************************************
 REPORTS
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
  const headerRow = 6;
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
  sh.autoResizeColumns(1, 12);
  return out.length;
}

function generateDamageReport() {
  movementReport_(['DAMAGE','DAMAGED'], 'DAMAGE REPORT');
}
function generateIssuedStockReport() {
  movementReport_(['ISSUED','ISSUED STOCK'], 'ISSUED STOCK REPORT');
}
function generateStaffLiabilityReport() {
  movementReport_(['DAMAGE','DAMAGED'], 'STAFF LIABILITY REPORT');
}
function generateUtilizedReport() {
  movementReport_(['UTILIZED','UTILIZED STOCK','BOOKING'], 'UTILIZED REPORT');
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
    for (let r = 4; r < values.length; r++) {
      const row = values[r];
      if (!row[0] && !row[1]) continue;
      const nums = row.filter(v => typeof v === 'number');
      if (nums.length) out.push([new Date(), name, row[0], row[1]].concat(nums.slice(0, 8)));
    }
  });
  if (out.length) sh.getRange(2, 1, out.length, 12).setValues(out.map(r => r.slice(0, 12)));
}

function checkFormulaErrors() {
  const ss = SpreadsheetApp.getActive();
  const sh = ensureReportSheet_('FORMULA ERROR CHECK', ['Timestamp','Sheet','Cell','Formula','Displayed Value']);
  const errors = ['#REF!','#VALUE!','#DIV/0!','#NAME?','#N/A','#ERROR!'];
  const out = [];
  ss.getSheets().forEach(s => {
    const range = s.getDataRange();
    const formulas = range.getFormulas();
    const displays = range.getDisplayValues();
    formulas.forEach((row, r) => row.forEach((f, c) => {
      if (!f) return;
      if (errors.some(err => String(displays[r][c]).includes(err))) {
        out.push([new Date(), s.getName(), s.getRange(r + 1, c + 1).getA1Notation(), f, displays[r][c]]);
      }
    }));
  });
  if (out.length) sh.getRange(2, 1, out.length, 5).setValues(out);
}
