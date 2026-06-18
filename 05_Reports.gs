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
  const hCol = findHeaderCol_(sh, ['MOVEMENT TYPE'], 10);
  const hRow = hCol ? hCol.row : 6;
  const headers = sh.getRange(hRow, 1, 1, sh.getLastColumn()).getValues()[0];
  const rows = sh.getLastRow() > hRow ? sh.getRange(hRow + 1, 1, sh.getLastRow() - hRow, sh.getLastColumn()).getValues() : [];
  return {headers, rows, hRow};
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
  const h = data.headers;
  const rows = data.rows;
  const cType = colIndex_(h, ['MOVEMENT TYPE']);
  const cStatus = colIndex_(h, ['STATUS']);
  const outH = ['Report Date','Source Row'].concat(h);
  const out = [];
  rows.forEach((r, i) => {
    const type = cType >= 0 ? key_(String(r[cType])) : '';
    const status = cStatus >= 0 ? key_(String(r[cStatus])) : '';
    if (movementTypes.includes(type) && status !== 'REJECTED') out.push([new Date(), i + data.hRow + 1].concat(r));
  });
  const sh = ensureReportSheet_(targetSheet, outH);
  if (out.length) sh.getRange(2, 1, out.length, outH.length).setValues(out);
  sh.autoResizeColumns(1, 12);
  return out.length;
}

function generateDamageReport() { movementReport_(['DAMAGE','DAMAGED'], 'DAMAGE REPORT'); }
function generateIssuedStockReport() { movementReport_(['ISSUED'], 'ISSUED STOCK REPORT'); }
function generateUtilizedReport() { movementReport_(['UTILIZED'], 'UTILIZED REPORT'); }
function generateStaffLiabilityReport() { movementReport_(['DAMAGE','DAMAGED'], 'STAFF LIABILITY REPORT'); }

function clearAllReports() {
  const reports = ['DAMAGE REPORT','ISSUED STOCK REPORT','UTILIZED REPORT','STAFF LIABILITY REPORT','STOCK AUDIT SUMMARY','FORMULA ERROR CHECK'];
  reports.forEach(r => {
    const sh = SpreadsheetApp.getActive().getSheetByName(r);
    if (sh) sh.clear();
  });
  uiAlert_('Reports cleared.');
}

function generateStockAuditSummary() {
  const ss = SpreadsheetApp.getActive();
  const cs = ['CS STORE','CS MINI-MART','CS LAUNDRY','CS BAR','CS RESTAURANT','CS KITCHEN'];
  const h = ['Timestamp','Sheet','Item Code','Item','Opening','Added','Issued','Sold','Damaged','Physical Count','Closing Stock','Variance'];
  const sh = ensureReportSheet_('STOCK AUDIT SUMMARY', h);
  const out = [];

  cs.forEach(name => {
    const s = ss.getSheetByName(name);
    if (!s || s.getLastRow() < 5) return;
    const data = s.getRange(1, 1, s.getLastRow(), Math.min(s.getLastColumn(), 50)).getValues();
    const hRow = data[3] || [];
    const ci = {
      code: colIndex_(hRow, ['ITEM CODE','CODE']),
      name: colIndex_(hRow, ['ITEM NAME','ITEM']),
      open: colIndex_(hRow, ['OPENING STOCK','OPENING']),
      add: colIndex_(hRow, ['ADDED STOCK','ADDED']),
      iss: colIndex_(hRow, ['ISSUED STOCK','ISSUED']),
      sold: colIndex_(hRow, ['SOLD','SALES']),
      dmg: colIndex_(hRow, ['DAMAGED','DAMAGE']),
      phys: colIndex_(hRow, ['PHYSICAL COUNT','PHYSICAL']),
      close: colIndex_(hRow, ['CLOSING STOCK','CLOSING']),
      var: colIndex_(hRow, ['VARIANCE','VAR'])
    };

    for (let r = 4; r < data.length; r++) {
      const row = data[r];
      if (!row[ci.code] && !row[ci.name]) continue;
      const val = (i) => (i >= 0 && typeof row[i] === 'number') ? row[i] : 0;
      out.push([new Date(), name, row[ci.code], row[ci.name], val(ci.open), val(ci.add), val(ci.iss), val(ci.sold), val(ci.dmg), val(ci.phys), val(ci.close), val(ci.var)]);
    }
  });
  if (out.length) sh.getRange(2, 1, out.length, h.length).setValues(out);
  uiAlert_('Audit Summary generated.');
}

function checkFormulaErrors() {
  const ss = SpreadsheetApp.getActive();
  const sh = ensureReportSheet_('FORMULA ERROR CHECK', ['Timestamp','Sheet','Cell','Formula','Displayed Value']);
  const errs = ['#REF!','#VALUE!','#DIV/0!','#NAME?','#N/A','#ERROR!'];
  const out = [];
  ss.getSheets().forEach(s => {
    const formulas = s.getDataRange().getFormulas();
    const displays = s.getDataRange().getDisplayValues();
    formulas.forEach((row, r) => row.forEach((f, c) => {
      if (f && errs.some(e => String(displays[r][c]).includes(e))) {
        out.push([new Date(), s.getName(), s.getRange(r+1, c+1).getA1Notation(), f, displays[r][c]]);
      }
    }));
  });
  if (out.length) sh.getRange(2, 1, out.length, 5).setValues(out);
  uiAlert_(out.length ? 'Errors found: ' + out.length : 'No errors found.');
}
