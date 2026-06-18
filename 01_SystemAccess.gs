/****************************************************
 SYSTEM_ACCESS tools.
 Rule 3: Preservation of existing dropdowns.
****************************************************/

function getSystemAccessSheet_(createIfMissing) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CONFIG.SYSTEM_ACCESS_SHEET);
  if (!sh && createIfMissing) sh = ss.insertSheet(CONFIG.SYSTEM_ACCESS_SHEET);
  return sh;
}

/**
 * Merges defaults with existing validation options in the sheet.
 */
function getExistingOptions_(sheet, column, defaults) {
  if (sheet.getLastRow() < 2) return defaults;
  const vals = sheet.getRange(2, column, sheet.getLastRow() - 1, 1).getValues()
    .map(v => String(v[0] || '').trim())
    .filter(Boolean);
  const combined = new Set(defaults);
  vals.forEach(v => combined.add(v));
  return Array.from(combined).sort();
}

function rebuildSystemAccess() {
  const sh = getSystemAccessSheet_(true);
  const oldRows = [];
  let roleOptions = CONFIG.ROLE_OPTIONS;
  let sheetOptions = CONFIG.SHEET_CONTROL_OPTIONS;

  if (sh.getLastRow() >= 1) {
    const data = sh.getRange(1, 1, sh.getLastRow(), Math.max(sh.getLastColumn(), 6)).getValues();
    const headers = data[0].map(key_);
    const m = {
      email: headers.indexOf('EMAIL') + 1,
      name: (headers.indexOf('FULL NAME') >= 0 ? headers.indexOf('FULL NAME') : headers.indexOf('NAME')) + 1,
      role: headers.indexOf('ROLE') + 1,
      active: headers.indexOf('ACTIVE') + 1,
      notes: headers.indexOf('NOTES') + 1,
      sheets: (headers.indexOf('SHEETS CONTROLLED') >= 0 ? headers.indexOf('SHEETS CONTROLLED') : headers.indexOf('SHEET CONTROLLED')) + 1
    };

    for (let r = 1; r < data.length; r++) {
      const email = String(data[r][m.email - 1] || '').trim();
      if (!email) continue;
      oldRows.push([
        email,
        data[r][m.name - 1] || '',
        data[r][m.role - 1] || '',
        data[r][m.active - 1] || 'Yes',
        data[r][m.notes - 1] || '',
        data[r][m.sheets - 1] || ''
      ]);
    }

    roleOptions = getExistingOptions_(sh, 8, CONFIG.ROLE_OPTIONS);
    sheetOptions = getExistingOptions_(sh, 10, CONFIG.SHEET_CONTROL_OPTIONS);
  }

  sh.clear();
  sh.getRange(1, 1, 1, CONFIG.SYSTEM_ACCESS_HEADERS.length).setValues([CONFIG.SYSTEM_ACCESS_HEADERS]).setFontWeight('bold').setBackground('#d9ead3');
  if (oldRows.length) sh.getRange(2, 1, oldRows.length, 6).setValues(oldRows);

  sh.getRange('H1').setValue('ROLE OPTIONS').setFontWeight('bold').setBackground('#cfe2f3');
  sh.getRange(2, 8, roleOptions.length, 1).setValues(roleOptions.map(x => [x]));
  sh.getRange('J1').setValue('SHEETS CONTROLLED OPTIONS').setFontWeight('bold').setBackground('#fce5cd');
  sh.getRange(2, 10, sheetOptions.length, 1).setValues(sheetOptions.map(x => [x]));

  sh.autoResizeColumns(1, 10);
  setupSystemAccessDropdowns();
  log_('SYSTEM_ACCESS', 'Rebuilt SYSTEM_ACCESS with preservation.');
  uiAlert_('SYSTEM_ACCESS rebuilt. Existing users and dropdown items preserved.');
}

function setupSystemAccessDropdowns() {
  const sh = getSystemAccessSheet_(true);
  const maxRows = Math.max(sh.getMaxRows() - 1, 50);
  const roleOptions = getExistingOptions_(sh, 8, CONFIG.ROLE_OPTIONS);
  const sheetOptions = getExistingOptions_(sh, 10, CONFIG.SHEET_CONTROL_OPTIONS);

  const roleRule = SpreadsheetApp.newDataValidation().requireValueInList(roleOptions, true).build();
  const activeRule = SpreadsheetApp.newDataValidation().requireValueInList(['Yes','No'], true).build();
  const sheetsRule = SpreadsheetApp.newDataValidation().requireValueInList(sheetOptions, true).setAllowInvalid(true).build();

  sh.getRange(2, 3, maxRows, 1).setDataValidation(roleRule);
  sh.getRange(2, 4, maxRows, 1).setDataValidation(activeRule);
  sh.getRange(2, 6, maxRows, 1).setDataValidation(sheetsRule);
}

function setupStockMovementDropdowns() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('STOCK MOVEMENT APPROVAL LOG');
  if (!sh) return;
  const deptCol = findHeaderCol_(sh, ['DEPARTMENT'], 10);
  const typeCol = findHeaderCol_(sh, ['MOVEMENT TYPE'], 10);
  const statusCol = findHeaderCol_(sh, ['STATUS'], 10);
  if (!deptCol || !typeCol || !statusCol) return;

  const maxRows = Math.max(sh.getMaxRows() - statusCol.row, 500);
  const deptRule = SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.DEPARTMENT_OPTIONS, true).build();
  const typeRule = SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.MOVEMENT_TYPES, true).build();
  const statusRule = SpreadsheetApp.newDataValidation().requireValueInList(['APPROVED','REJECTED','UNDER REVIEW','PENDING'], true).build();

  sh.getRange(statusCol.row + 1, deptCol.col, maxRows, 1).setDataValidation(deptRule);
  sh.getRange(statusCol.row + 1, typeCol.col, maxRows, 1).setDataValidation(typeRule);
  sh.getRange(statusCol.row + 1, statusCol.col, maxRows, 1).setDataValidation(statusRule);
  uiAlert_('Dropdowns updated for Log.');
}

function validateSystemAccess() {
  const sh = getSystemAccessSheet_(false);
  if (!sh) throw new Error('SYSTEM_ACCESS sheet missing.');
  const data = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 6).getValues();
  const errors = [];
  data.forEach((r, i) => {
    if (r[0] && !r[2]) errors.push('Row ' + (i + 2) + ': Missing Role');
  });
  if (errors.length) uiAlert_('Validation Errors:\n' + errors.join('\n'));
  else uiAlert_('Validation Passed.');
}

function getAccessRecords_() {
  const sh = getSystemAccessSheet_(false);
  if (!sh || sh.getLastRow() < 2) return [];
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();
  return data.map(r => ({
    email: String(r[0] || '').trim().toLowerCase(),
    role: key_(r[2]),
    active: key_(r[3]) === 'YES',
    sheets: splitList_(r[5])
  })).filter(r => r.email && r.active);
}

function getOwnerEmails_() {
  return getAccessRecords_()
    .filter(r => CONFIG.OWNER_ROLES.includes(r.role) || r.sheets.includes('ALL'))
    .map(r => r.email);
}

function expandControlsToSheets_(controls) {
  const out = new Set();
  controls.forEach(c => {
    const k = key_(c);
    const expanded = CONFIG.GROUP_ALIASES[k];
    if (expanded) expanded.forEach(s => out.add(key_(s)));
    else out.add(k);
  });
  return Array.from(out);
}

function usersForSheet_(sheetName) {
  const records = getAccessRecords_();
  const users = [];
  const sKey = key_(sheetName);
  records.forEach(r => {
    const expanded = expandControlsToSheets_(r.sheets);
    if (expanded.includes('*') || expanded.includes(sKey)) users.push(r.email);
  });
  return Array.from(new Set(users));
}
