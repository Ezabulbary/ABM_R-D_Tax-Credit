// ============================================================
// step2_emailPermutator.gs — Build "Email Permutator" sheet
//
// Reads from:  "ABM_R&D_Tax Credit" → Main_Crunchbase tab
// Writes to:   "ABM R&D Tax Credit - Automation" (this SS)
//              → Email Permutator tab
// ============================================================

// autoMode = true  -> called from the background auto-chain. No UI is
//                     available, so the append prompt is skipped and we
//                     always append new unique founder rows.
// autoMode = false -> called from the menu. Shows the append prompt.
function buildEmailPermutatorSheet(autoMode) {
  var automationSS = SpreadsheetApp.getActiveSpreadsheet(); // this spreadsheet
  var crunchbaseSS = SpreadsheetApp.openById(CONFIG.CRUNCHBASE_SS_ID);
  var mainSheet    = crunchbaseSS.getSheetByName(CONFIG.MAIN_SHEET_NAME);
  if (!mainSheet) throw new Error('"' + CONFIG.MAIN_SHEET_NAME + '" tab not found.');

  var HEADERS = [
    'Organization Name',    // A  0
    'First Name',           // B  1
    'Last Name',            // C  2
    'Domain',               // D  3
    'firstname@',           // E  4
    'Verification Status',  // F  5
    'firstname.lastname@',  // G  6
    'Verification Status',  // H  7
    'firstinitiallast@',    // I  8
    'Verification Status',  // J  9
    'Email',                // K 10
    'Verification Status',  // L 11
    'Verification Date'     // M 12  (filled by Step 3 when verified)
  ];

  // Get or create Email Permutator in this (Automation) SS
  var permSheet = automationSS.getSheetByName(CONFIG.PERMUTATOR_SHEET_NAME);
  var isNew     = false;
  if (!permSheet) {
    permSheet = automationSS.insertSheet(CONFIG.PERMUTATOR_SHEET_NAME);
    isNew     = true;
  }

  var lastRow       = permSheet.getLastRow();
  var existingKeys  = Object.create(null);

  if (!isNew && lastRow > 0) {
    // Only ask in manual mode; the background auto-chain always appends.
    if (!autoMode) {
      var resp = SpreadsheetApp.getUi().alert(
        '❓ Append Data',
        '"Email Permutator" already has data.\nAppend new founder rows to it?',
        SpreadsheetApp.getUi().ButtonSet.YES_NO
      );
      if (resp !== SpreadsheetApp.getUi().Button.YES) return 0;
    }

    // Load existing keys to prevent duplicates
    var existingVals = permSheet.getRange(1, 1, lastRow, 4).getValues();
    for (var j = 1; j < existingVals.length; j++) {
      var eRow = existingVals[j];
      var key  = [0, 1, 2, 3]
        .map(function(c) { return String(eRow[c]).trim().toLowerCase(); })
        .join('|||');
      existingKeys[key] = true;
    }
  } else {
    var hRow = permSheet.getRange(1, 1, 1, HEADERS.length);
    hRow.setValues([HEADERS]);
    hRow.setBackground('#1a1a2e').setFontColor('#e0e0e0').setFontWeight('bold').setFontSize(10);
    permSheet.setFrozenRows(1);
    lastRow = 1;
  }

  // Older Email Permutator sheets pre-date the Verification Date column.
  // Ensure the header exists at column 13 (it sits at the end, so no
  // existing data needs to shift).
  if (String(permSheet.getRange(1, 13).getValue()).trim() !== 'Verification Date') {
    permSheet.getRange(1, 13).setValue('Verification Date')
      .setBackground('#14532d').setFontColor('#86efac').setFontWeight('bold').setFontSize(10);
  }

  // Read Main_Crunchbase
  var mainData    = mainSheet.getDataRange().getValues();
  var mainHeaders = mainData[0].map(String);

  var orgCol    = mainHeaders.indexOf('Organization Name');
  var domainCol = mainHeaders.indexOf('Domain');

  // Collect Founder 1…N column indices
  var founderColIdxs = [];
  for (var h = 0; h < mainHeaders.length; h++) {
    if (/^Founder \d+$/.test(mainHeaders[h])) founderColIdxs.push(h);
  }

  var rows = [];

  for (var i = 1; i < mainData.length; i++) {
    var row     = mainData[i];
    var orgName = (orgCol >= 0 && row[orgCol])    ? String(row[orgCol]).trim()    : '';
    var domain  = (domainCol >= 0 && row[domainCol]) ? String(row[domainCol]).trim() : '';

    if (!orgName && !domain) continue;

    // Gather founders from Founder 1…N columns
    var founders = [];
    founderColIdxs.forEach(function(ci) {
      var v = row[ci] ? String(row[ci]).trim() : '';
      if (v) founders.push(v);
    });

    // Fallback to raw Founders column
    var foundersIdx = mainHeaders.indexOf('Founders');
    if (founders.length === 0 && foundersIdx >= 0 && row[foundersIdx]) {
      founders = splitFounders(String(row[foundersIdx]));
    }

    if (founders.length === 0 || !domain) continue;

    founders.forEach(function(fullName) {
      var parsed = parseFounderName(fullName);
      var fn     = parsed.firstName;
      var ln     = parsed.lastName;
      if (!fn) return;

      var fnL = fn.toLowerCase();
      var lnL = ln.toLowerCase();
      var fi  = fnL.charAt(0);

      var pat1 = fnL + '@' + domain;
      var pat2 = lnL ? fnL + '.' + lnL + '@' + domain : '';
      var pat3 = lnL ? fi + lnL + '@' + domain : '';

      var permKey = orgName.toLowerCase() + '|||' + fnL + '|||' + lnL + '|||' + domain.toLowerCase();
      if (permKey in existingKeys) return;
      existingKeys[permKey] = true;

      rows.push([orgName, fn, ln, domain, pat1, '', pat2, '', pat3, '', '', '', '']);
    });
  }

  if (rows.length > 0) {
    permSheet.getRange(lastRow + 1, 1, rows.length, HEADERS.length).setValues(rows);
  }

  permSheet.autoResizeColumns(1, HEADERS.length);
  [6, 8, 10, 12].forEach(function(col) {
    permSheet.getRange(1, col).setBackground('#2d2d44').setFontColor('#aaaaff');
  });
  permSheet.getRange(1, 11, 1, 3).setBackground('#14532d').setFontColor('#86efac');

  SpreadsheetApp.flush();
  return rows.length;
}

// ── Auto-chain entry point (runs from a background time trigger) ───
// Step 1.2 schedules this. It builds the permutator without any
// prompt, then schedules Step 3 (Reoon verification).
function autoStep2Start() {
  _deleteTriggersForFunction('autoStep2Start');
  try {
    buildEmailPermutatorSheet(true); // true = auto mode, no UI prompt
  } catch (e) {
    Logger.log('[autoStep2Start] ' + e.message + '\n' + e.stack);
  }
  _scheduleTrigger('autoStep3Start'); // continue the chain -> Step 3
}
