/****************************************************
 V16 FAST PROTECTION FUNCTIONS ONLY
 Rule 5: Protect entire sheet except approved editable ranges.
 Rule 6: No redundant logic.
****************************************************/

function protectionLock_(callback) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) throw new Error('Lock timeout.');
  try { return callback(); } finally { lock.releaseLock(); }
}

function existingSheets_(names) {
  const ss = SpreadsheetApp.getActive();
  return names.map(n => ss.getSheetByName(n)).filter(Boolean);
}

function safeRangeA1_(sheet, a1) {
  try { return sheet.getRange(a1); } catch (e) { return null; }
}

function safeRangeRC_(sheet, r, c, rs, cs) {
  try { return sheet.getRange(r, c, rs, cs); } catch (e) { return null; }
}

function pruneFormulas_(range) {
  if (!range) return [];
  const sh = range.getSheet();
  const formulas = range.getFormulas();
  const startR = range.getRow();
  const startC = range.getColumn();
  const unprotected = [];

  for (let c = 0; c < formulas[0].length; c++) {
    let currentStart = -1;
    for (let r = 0; r < formulas.length; r++) {
      if (!formulas[r][c]) {
        if (currentStart === -1) currentStart = r;
      } else {
        if (currentStart !== -1) {
          unprotected.push(sh.getRange(startR + currentStart, startC + c, r - currentStart, 1));
          currentStart = -1;
        }
      }
    }
    if (currentStart !== -1) unprotected.push(sh.getRange(startR + currentStart, startC + c, formulas.length - currentStart, 1));
  }
  return unprotected;
}

function addRange_(arr, range) {
  if (!range) return;
  pruneFormulas_(range).forEach(r => arr.push(r));
}

function editableRangesForSheet_(sheet) {
  const name = key_(sheet.getName());
  const ranges = [];

  // Core Transaction Sheets
  if (name === 'PURCHASES') {
    addRange_(ranges, safeRangeA1_(sheet, 'A3:C3'));
    const lr = Math.min(lastMeaningfulRow_(sheet, 5, 1), 2000);
    if (lr >= 5) {
      addRange_(ranges, safeRangeRC_(sheet, 5, 1, lr - 4, 3)); // A5:C
      addRange_(ranges, safeRangeRC_(sheet, 5, 5, lr - 4, 5)); // E5:I
      addRange_(ranges, safeRangeRC_(sheet, 5, 11, lr - 4, 1)); // K5:K
    }
  } else if (name === 'EXPENSES') {
    addRange_(ranges, safeRangeA1_(sheet, 'A3:C3'));
    const lr = Math.min(lastMeaningfulRow_(sheet, 5, 1), 502);
    if (lr >= 5) {
      addRange_(ranges, safeRangeRC_(sheet, 5, 1, lr - 4, 9)); // A5:I
      addRange_(ranges, safeRangeRC_(sheet, 5, 11, lr - 4, 1)); // K5:K
    }
  } else if (name === 'STOCK MOVEMENT APPROVAL LOG') {
    addRange_(ranges, safeRangeA1_(sheet, 'A3:C3'));
    const lr = Math.min(lastMeaningfulRow_(sheet, 7, 1), 2000);
    if (lr >= 7) {
      addRange_(ranges, safeRangeRC_(sheet, 7, 1, lr - 6, 3)); // A7:C
      addRange_(ranges, safeRangeRC_(sheet, 7, 5, lr - 6, 5)); // E7:I
    }
  } else if (name === 'DAILY SALES') {
    addRange_(ranges, safeRangeA1_(sheet, 'A3:C3'));
    addRange_(ranges, safeRangeA1_(sheet, 'O5:O35'));
  } else if (name === 'DAILY SALES BREAKDOWN') {
    addRange_(ranges, safeRangeA1_(sheet, 'A3:C3'));
    const lr = Math.min(lastMeaningfulRow_(sheet, 5, 1), 5000);
    if (lr >= 5) {
      addRange_(ranges, safeRangeRC_(sheet, 5, 1, lr - 4, 8));   // A:H (I locked)
      addRange_(ranges, safeRangeRC_(sheet, 5, 10, lr - 4, 3));  // J:L (M locked)
      addRange_(ranges, safeRangeRC_(sheet, 5, 14, lr - 4, 1));  // N (O locked)
      addRange_(ranges, safeRangeRC_(sheet, 5, 16, lr - 4, 4));  // P:S
    }

  // CS Sheets (Rule 4: CS BAR and CS KITCHEN included via isCSSheet_)
  } else if (isCSSheet_(name)) {
    addRange_(ranges, safeRangeA1_(sheet, 'E2:F3'));
    const lr = lastMeaningfulRow_(sheet, 5, 2);
    if (lr >= 5) {
      addRange_(ranges, safeRangeRC_(sheet, 5, 5, lr - 4, 2));   // E5:F
      addRange_(ranges, safeRangeRC_(sheet, 5, 15, lr - 4, 1));  // O5:O
      addRange_(ranges, safeRangeRC_(sheet, 5, 31, lr - 4, 1));  // AE5:AE
    }

  // Weekly M.R Sheets
  } else if (isWeeklyMRSheet_(name)) {
    addRange_(ranges, safeRangeA1_(sheet, 'A3:B3'));
    ['D2:H2','J2:N2','P2:T2','V2:Z2','AB2:AF2','AH2:AL2','AN2:AR2'].forEach(a => addRange_(ranges, safeRangeA1_(sheet, a)));
    if (isWeekOneSheet_(name)) {
      const lr = lastMeaningfulRow_(sheet, 8, 2);
      if (lr >= 8) addRange_(ranges, safeRangeRC_(sheet, 8, 4, lr - 7, 1)); // D8:D
    }

  // M.R KITCHEN U
  } else if (name === 'M.R KITCHEN U') {
    addRange_(ranges, safeRangeA1_(sheet, 'A3:B3'));
    ['D2:F2','J2:L2','P2:R2','V2:X2'].forEach(a => addRange_(ranges, safeRangeA1_(sheet, a)));
    const lr = lastMeaningfulRow_(sheet, 8, 1);
    if (lr >= 8) addRange_(ranges, safeRangeRC_(sheet, 8, 4, lr - 7, 1)); // D8:D
  }

  return ranges;
}

function protectSheetFast_(sheet) {
  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());

  const p = sheet.protect().setDescription('Carlisle V16 Protection');
  const owners = getOwnerEmails_();
  if (owners.length) {
    p.removeEditors(p.getEditors());
    p.addEditors(owners);
  }

  const unprotected = editableRangesForSheet_(sheet);
  if (unprotected.length) p.setUnprotectedRanges(unprotected.slice(0, 50));
  return sheet.getName();
}

function protectSheetsFast_(sheetNames, label) {
  return protectionLock_(() => {
    const sheets = existingSheets_(sheetNames);
    if (!sheets.length) throw new Error('No sheets for: ' + label);
    const done = sheets.map(s => protectSheetFast_(s));
    uiAlert_('Done: ' + label + '\nProtected: ' + done.join(', '));
  });
}

function runQueueProtection_(queueName, groupNames, label) {
  return protectionLock_(() => {
    const names = [];
    groupNames.forEach(gn => {
      const g = CONFIG.GROUP_ALIASES[gn] || [gn];
      g.forEach(sn => names.push(sn));
    });

    const props = PropertiesService.getDocumentProperties();
    let idx = Number(props.getProperty(queueName) || 0);
    if (idx >= names.length) {
      props.deleteProperty(queueName);
      uiAlert_(label + ' queue complete.');
      return;
    }

    const sh = SpreadsheetApp.getActive().getSheetByName(names[idx]);
    let result = 'Not Found';
    if (sh) result = protectSheetFast_(sh);

    idx++;
    props.setProperty(queueName, String(idx));
    if (idx >= names.length) props.deleteProperty(queueName);
    uiAlert_('Protected: ' + result + '\nProgress: ' + idx + '/' + names.length);
  });
}

// Menu mappings
function protect_Purchases(){ protectSheetsFast_(['PURCHASES'], 'Purchases'); }
function protect_DailySales(){ protectSheetsFast_(['DAILY SALES'], 'Daily Sales'); }
function protect_DailySalesBreakdown(){ protectSheetsFast_(['DAILY SALES BREAKDOWN'], 'Sales Breakdown'); }
function protect_Expenses(){ protectSheetsFast_(['EXPENSES'], 'Expenses'); }
function protect_StockMovement(){ protectSheetsFast_(['STOCK MOVEMENT APPROVAL LOG'], 'Log'); }
function protect_MRKitchenU(){ protectSheetsFast_(['M.R KITCHEN U'], 'Kitchen U'); }

function protect_Weekly_Weeks1_2(){
  const g = CONFIG.GROUP_ALIASES['WEEKLY M.R SHEETS'];
  protectSheetsFast_(g.slice(0, 6), 'Weeks 1-2');
}
function protect_Weekly_Weeks3_4(){
  const g = CONFIG.GROUP_ALIASES['WEEKLY M.R SHEETS'];
  protectSheetsFast_(g.slice(6, 12), 'Weeks 3-4');
}
function protect_Weekly_Week5(){
  const g = CONFIG.GROUP_ALIASES['WEEKLY M.R SHEETS'];
  protectSheetsFast_(g.slice(12), 'Week 5');
}

function protect_CSSheets_Next() { runQueueProtection_('CS_Q', ['CS SHEETS'], 'CS Sheets'); }
function protect_AdminSheets_Next() { runQueueProtection_('ADM_Q', ['OWNER SHEETS'], 'Admin Sheets'); }

function clearCarlisleProtections() {
  SpreadsheetApp.getActive().getSheets().forEach(sh => {
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
    sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());
  });
  uiAlert_('All protections cleared.');
}
