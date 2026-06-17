/****************************************************
 SYSTEM_ACCESS tools.
 Rebuild preserves existing staff rows and merges dropdown options.
****************************************************/

function getSystemAccessSheet_(createIfMissing) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CONFIG.SYSTEM_ACCESS_SHEET);
  if (!sh && createIfMissing) sh = ss.insertSheet(CONFIG.SYSTEM_ACCESS_SHEET);
  return sh;
}

function getHeaderMap_(sheet, row) {
  row = row || 1;
  const headers = sheet.getRange(row, 1, 1, Math.max(sheet.getLastColumn(), CONFIG.SYSTEM_ACCESS_HEADERS.length)).getValues()[0];
  const map = {};
  headers.forEach((h, i) => {
    const k = key_(h);
    if (k) map[k] = i + 1;
  });
  return map;
}

/**
 * Reads existing options from the sheet to avoid overriding user-added items.
 */
function getExistingOptions_(sheet, column, defaults) {
  if (sheet.getLastRow() < 2) return defaults;
  const vals = sheet.getRange(2, column, sheet.getLastRow() - 1, 1).getValues()
    .map(v => String(v[0] || '').trim())
    .filter(Boolean);
  const combined = new Set(defaults);
  vals.forEach(v => combined.add(v));
  return Array.from(combined);
}

function rebuildSystemAccess() {
  const sh = getSystemAccessSheet_(true);
  const oldRows = [];
  let roleOptions = CONFIG.ROLE_OPTIONS;
  let sheetOptions = CONFIG.SHEET_CONTROL_OPTIONS;

  if (sh.getLastRow() >= 1) {
    const map = getHeaderMap_(sh, 1);
    const maxCol = sh.getLastColumn();
    const data = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), maxCol).getValues();
    data.forEach(r => {
      const email = r[(map.EMAIL || 1) - 1];
      if (!email) return;
      oldRows.push({
        email: email,
        fullName: r[(map['FULL NAME'] || map.NAME || 2) - 1] || '',
        role: r[(map.ROLE || 3) - 1] || '',
        active: r[(map.ACTIVE || map.STATUS || 4) - 1] || 'Yes',
        notes: r[(map.NOTES || 5) - 1] || '',
        sheets: r[(map['SHEETS CONTROLLED'] || map['SHEET CONTROLLED'] || map.SHEETS || 6) - 1] || ''
      });
    });

    // Read and merge existing dropdown options from columns H and J
    roleOptions = getExistingOptions_(sh, 8, CONFIG.ROLE_OPTIONS);
    sheetOptions = getExistingOptions_(sh, 10, CONFIG.SHEET_CONTROL_OPTIONS);
  }

  sh.clear();
  sh.getRange(1, 1, 1, CONFIG.SYSTEM_ACCESS_HEADERS.length).setValues([CONFIG.SYSTEM_ACCESS_HEADERS]).setFontWeight('bold').setBackground('#d9ead3');
  if (oldRows.length) {
    sh.getRange(2, 1, oldRows.length, 6).setValues(oldRows.map(r => [r.email, r.fullName, r.role, r.active, r.notes, r.sheets]));
  }

  sh.getRange('H1').setValue('ROLE OPTIONS').setFontWeight('bold').setBackground('#cfe2f3');
  sh.getRange(2, 8, roleOptions.length, 1).setValues(roleOptions.map(x => [x]));
  sh.getRange('J1').setValue('SHEETS CONTROLLED OPTIONS').setFontWeight('bold').setBackground('#fce5cd');
  sh.getRange(2, 10, sheetOptions.length, 1).setValues(sheetOptions.map(x => [x]));
  sh.autoResizeColumns(1, 10);

  setupSystemAccessDropdowns();
  log_('SYSTEM_ACCESS', 'Rebuilt SYSTEM_ACCESS, merged ' + roleOptions.length + ' roles and ' + sheetOptions.length + ' sheet options.');
  uiAlert_('SYSTEM_ACCESS rebuilt. Existing staff and user-added dropdown items were preserved.');
}

function setupSystemAccessDropdowns() {
  const sh = getSystemAccessSheet_(true);
  const maxRows = Math.max(sh.getMaxRows() - 1, 50);

  // Read merged options from columns H and J
  const roleOptions = getExistingOptions_(sh, 8, CONFIG.ROLE_OPTIONS);
  const sheetOptions = getExistingOptions_(sh, 10, CONFIG.SHEET_CONTROL_OPTIONS);

  const roleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(roleOptions, true)
    .setAllowInvalid(false)
    .build();
  const activeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes','No'], true)
    .setAllowInvalid(false)
    .build();
  const sheetsRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(sheetOptions, true)
    .setAllowInvalid(true) // allows comma-separated multiple entries.
    .build();

  sh.getRange(2, 3, maxRows, 1).setDataValidation(roleRule);
  sh.getRange(2, 4, maxRows, 1).setDataValidation(activeRule);
  sh.getRange(2, 6, maxRows, 1).setDataValidation(sheetsRule);
  log_('SYSTEM_ACCESS', 'Dropdowns updated.');
}

function validateSystemAccess() {
  const sh = getSystemAccessSheet_(false);
  if (!sh) throw new Error('SYSTEM_ACCESS sheet is missing. Run Rebuild SYSTEM_ACCESS.');
  const map = getHeaderMap_(sh, 1);
  const errors = [];
  ['EMAIL','ROLE'].forEach(h => { if (!map[h]) errors.push('Missing required column: ' + h); });
  if (errors.length) throw new Error(errors.join('\n'));

  const data = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 0), Math.max(sh.getLastColumn(), 6)).getValues();
  const seen = {};

  const roleOptions = getExistingOptions_(sh, 8, CONFIG.ROLE_OPTIONS).map(key_);
  const validRoles = new Set(roleOptions);
  const sheetOptions = getExistingOptions_(sh, 10, CONFIG.SHEET_CONTROL_OPTIONS).map(key_);
  const validControls = new Set(sheetOptions);

  data.forEach((r, idx) => {
    const row = idx + 2;
    const email = String(r[(map.EMAIL || 1) - 1] || '').trim();
    const role = key_(r[(map.ROLE || 3) - 1]);
    const active = key_(r[(map.ACTIVE || map.STATUS || 4) - 1] || 'YES');
    const sheets = splitList_(r[(map['SHEETS CONTROLLED'] || 6) - 1]);
    if (!email && !role && !sheets.length) return;
    if (!email) errors.push('Row ' + row + ': Email is blank.');
    if (email && seen[email.toLowerCase()]) errors.push('Row ' + row + ': Duplicate email ' + email);
    if (email) seen[email.toLowerCase()] = true;
    if (!role) errors.push('Row ' + row + ': Role is blank.');
    if (role && !validRoles.has(role)) errors.push('Row ' + row + ': Invalid role: ' + role);
    if (active && !['YES','NO','ACTIVE','INACTIVE','TRUE','FALSE'].includes(active)) errors.push('Row ' + row + ': Active must be Yes/No.');
    sheets.forEach(s => { if (!validControls.has(s)) errors.push('Row ' + row + ': Unknown Sheets Controlled option: ' + s); });
  });

  if (errors.length) {
    uiAlert_('SYSTEM_ACCESS validation found issues:\n\n' + errors.slice(0, 20).join('\n') + (errors.length > 20 ? '\n...more errors not shown' : ''));
    return false;
  }
  uiAlert_('SYSTEM_ACCESS validation passed.');
  return true;
}

function getAccessRecords_() {
  const sh = getSystemAccessSheet_(false);
  if (!sh) return [];
  const map = getHeaderMap_(sh, 1);
  if (!map.EMAIL || !map.ROLE) return [];
  const data = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 0), Math.max(sh.getLastColumn(), 6)).getValues();
  return data.map(r => {
    const email = String(r[(map.EMAIL || 1) - 1] || '').trim();
    const role = key_(r[(map.ROLE || 3) - 1]);
    const activeVal = key_(r[(map.ACTIVE || map.STATUS || 4) - 1] || 'YES');
    const sheets = splitList_(r[(map['SHEETS CONTROLLED'] || 6) - 1]);
    return { email, role, active: !['NO','INACTIVE','FALSE'].includes(activeVal), sheets };
  }).filter(r => r.email && r.role && r.active);
}

function getOwnerEmails_() {
  return getAccessRecords_()
    .filter(r => CONFIG.OWNER_ROLES.includes(r.role) || r.sheets.includes('ALL') || r.sheets.includes('OWNER SHEETS'))
    .map(r => r.email);
}

function expandControlsToSheets_(controls) {
  const out = new Set();
  controls.forEach(c => {
    const k = key_(c);
    const expanded = CONFIG.GROUP_ALIASES[k];
    if (expanded) {
      expanded.forEach(s => out.add(key_(s)));
    } else {
      out.add(k);
    }
  });
  return Array.from(out);
}
