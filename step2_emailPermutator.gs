// ============================================================
// step2_emailPermutator.gs — Build "Email Permutator" sheet
// ============================================================

/**
 * Creates (or appends to) the "Email Permutator" sheet.
 * One row per founder, three email pattern columns + verification slots.
 * Returns total founder-rows written.
 */
function buildEmailPermutatorSheet() {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName(CONFIG.MAIN_SHEET_NAME);
  if (!mainSheet) throw new Error('"' + CONFIG.MAIN_SHEET_NAME + '" sheet not found.');

  // ── Headers ───────────────────────────────────────────────
  var HEADERS = [
    'Organization Name',   // A  0
    'First Name',          // B  1
    'Last Name',           // C  2
    'Domain',              // D  3
    'firstname@',          // E  4  — pattern 1 (displayed first)
    'Verification Status', // F  5
    'firstname.lastname@', // G  6  — pattern 2 (displayed second)
    'Verification Status', // H  7
    'firstinitiallast@',   // I  8  — pattern 3
    'Verification Status', // J  9
    'Email',              // K 10
    'Verification Status'  // L 11
  ];

  // ── Get or create Email Permutator sheet ──────────────────
  var permSheet = ss.getSheetByName(CONFIG.PERMUTATOR_SHEET_NAME);
  var isNew = false;
  if (!permSheet) {
    permSheet = ss.insertSheet(CONFIG.PERMUTATOR_SHEET_NAME);
    isNew = true;
  }

  var lastRow = permSheet.getLastRow();
  var existingKeys = Object.create(null);

  if (!isNew && lastRow > 0) {
    var resp = SpreadsheetApp.getUi().alert(
      '❓ Append Data',
      '"' + CONFIG.PERMUTATOR_SHEET_NAME + '" sheet already exists.\n' +
      'Do you want to append new founder email permutations to it?',
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    if (resp !== SpreadsheetApp.getUi().Button.YES) return 0;

    // Load existing keys (columns A to D) to prevent duplicates
    var existingValues = permSheet.getRange(1, 1, lastRow, 4).getValues();
    for (var j = 1; j < existingValues.length; j++) {
      var eOrg = String(existingValues[j][0]).trim().toLowerCase();
      var eFirst = String(existingValues[j][1]).trim().toLowerCase();
      var eLast = String(existingValues[j][2]).trim().toLowerCase();
      var eDomain = String(existingValues[j][3]).trim().toLowerCase();
      if (eOrg || eFirst || eLast || eDomain) {
        var key = eOrg + '|||' + eFirst + '|||' + eLast + '|||' + eDomain;
        existingKeys[key] = true;
      }
    }
  } else {
    // Write headers if new or completely empty
    var hRow = permSheet.getRange(1, 1, 1, HEADERS.length);
    hRow.setValues([HEADERS]);
    hRow.setBackground('#1a1a2e')
        .setFontColor('#e0e0e0')
        .setFontWeight('bold')
        .setFontSize(10);
    permSheet.setFrozenRows(1);
    lastRow = 1;
  }

  // ── Read main sheet ───────────────────────────────────────
  var mainData    = mainSheet.getDataRange().getValues();
  var mainHeaders = mainData[0].map(String);

  var orgCol     = mainHeaders.indexOf('Organization Name');
  var domainCol  = mainHeaders.indexOf('Domain');
  var foundersCol = mainHeaders.indexOf('Founders');

  // Collect Founder 1…N column indices
  var founderColIdxs = [];
  for (var h = 0; h < mainHeaders.length; h++) {
    if (/^Founder \d+$/.test(mainHeaders[h])) founderColIdxs.push(h);
  }

  // ── Build permutator rows ─────────────────────────────────
  var rows = [];

  for (var i = 1; i < mainData.length; i++) {
    var row = mainData[i];
    var orgName = orgCol >= 0 && row[orgCol] ? String(row[orgCol]).trim() : '';
    var domain  = domainCol >= 0 && row[domainCol] ? String(row[domainCol]).trim() : '';

    if (!orgName && !domain) continue;

    // Gather founders
    var founders = [];
    founderColIdxs.forEach(function(ci) {
      var v = row[ci] ? String(row[ci]).trim() : '';
      if (v) founders.push(v);
    });
    // Fallback: original Founders column
    if (founders.length === 0 && foundersCol >= 0 && row[foundersCol]) {
      founders = splitFounders(String(row[foundersCol]));
    }

    if (founders.length === 0 || !domain) continue;

    founders.forEach(function(fullName) {
      var parsed    = parseFounderName(fullName);
      var firstName = parsed.firstName;
      var lastName  = parsed.lastName;
      if (!firstName) return;

      var fn  = firstName.toLowerCase();
      var ln  = lastName.toLowerCase();
      var fi  = fn.charAt(0);              // first initial

      var pat1 = fn + '@' + domain;
      var pat2 = ln ? fn + '.' + ln + '@' + domain : '';
      var pat3 = ln ? fi + ln + '@' + domain : '';

      var permKey = orgName.toLowerCase() + '|||' + fn + '|||' + ln + '|||' + domain.toLowerCase();
      if (permKey in existingKeys) return; // skip duplicate
      existingKeys[permKey] = true;

      rows.push([
        orgName,   // A 0  Organization Name
        firstName, // B 1  First Name
        lastName,  // C 2  Last Name
        domain,    // D 3  Domain
        pat1,      // E 4  firstname@  (displayed 1st)
        '',        // F 5  Verification Status
        pat2,      // G 6  firstname.lastname@  (displayed 2nd)
        '',        // H 7  Verification Status
        pat3,      // I 8  firstinitiallast@
        '',        // J 9  Verification Status
        '',        // K 10 Verified Email Only
        ''         // L 11 Verification Status (final)
      ]);
    });
  }

  if (rows.length > 0) {
    permSheet.getRange(lastRow + 1, 1, rows.length, HEADERS.length).setValues(rows);
  }

  // ── Formatting ────────────────────────────────────────────
  permSheet.autoResizeColumns(1, HEADERS.length);

  // Highlight status header columns
  [6, 8, 10, 12].forEach(function(col) {
    permSheet.getRange(1, col)
      .setBackground('#2d2d44')
      .setFontColor('#aaaaff');
  });

  // Highlight Verified Email + Status columns
  permSheet.getRange(1, 11, 1, 2)
    .setBackground('#14532d')
    .setFontColor('#86efac');

  SpreadsheetApp.flush();
  return rows.length;
}

// ── Helpers ──────────────────────────────────────────────────

/**
 * Parses "John Smith Jr" → { firstName: "John", lastName: "Smith Jr" }
 * Single word → { firstName: "John", lastName: "" }
 */
function parseFounderName(fullName) {
  if (!fullName) return { firstName: '', lastName: '' };
  var parts = fullName.trim().split(/\s+/);
  if (parts.length === 1) return { firstName: parts[0], lastName: '' };
  return { firstName: parts[0], lastName: parts.slice(1).join(' ') };
}
