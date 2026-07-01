// ============================================================
// step1_processMain.gs — Step 1.1: Main_Crunchbase domain + founders
//
// Reads from: "ABM_R&D_Tax Credit" spreadsheet → Main_Crunchbase tab
// Writes to:  same tab (Domain column + Founder 1…N columns)
// ============================================================

var MAIN_TARGET_COLUMNS = [
  'Campaign Status',
  'How Many Leads',
  'Verification Status',
  'Organization Name',
  'Website',
  'Founded Date',
  'Number of Employees',
  'Full Description',
  'Industry',
  'Patents Granted',
  'Total Funding Amount',
  'Last Funding Type',
  'IT Spend (Aberdeen)',
  'Number of Articles',
  'Last Funding Date',
  'Most Recent Valuation Range',
  'Estimated Revenue Range',
  'Actively Hiring',
  'Founders',
  'Contact Email',
  'Downloads Last 30 Days',
  'Industry Groups',
  'Average Visits (6 months)',
  'Headquarters Location',
  'CB Rank (Organization)'
];

/**
 * Aligns headers, adds Domain column, splits Founders into Founder 1…N columns.
 * Returns number of processed data rows.
 */
function processMainSheet() {
  var ss    = SpreadsheetApp.openById(CONFIG.CRUNCHBASE_SS_ID);
  var sheet = ss.getSheetByName(CONFIG.MAIN_SHEET_NAME);
  if (!sheet) throw new Error(
    '"' + CONFIG.MAIN_SHEET_NAME + '" tab not found in ABM_R&D_Tax Credit spreadsheet.'
  );

  var headers = alignHeaders(sheet, MAIN_TARGET_COLUMNS);

  var allValues       = sheet.getDataRange().getValues();
  var websiteIdx      = headers.indexOf('Website');
  var foundersIdx     = headers.indexOf('Founders');
  var contactEmailIdx = headers.indexOf('Contact Email');

  if (websiteIdx  === -1) throw new Error('"Website" column not found.');
  if (foundersIdx === -1) throw new Error('"Founders" column not found.');

  // ── Determine max founders across all rows ─────────────────
  var maxFounders = 0;
  for (var i = 1; i < allValues.length; i++) {
    var raw = allValues[i][foundersIdx] ? String(allValues[i][foundersIdx]).trim() : '';
    if (raw) maxFounders = Math.max(maxFounders, splitFounders(raw).length);
  }
  maxFounders = Math.max(maxFounders, 1);

  // ── Ensure Domain column at the END ───────────────────────
  var domainIdx = headers.indexOf('Domain');
  if (domainIdx === -1) {
    domainIdx = headers.length;
    sheet.getRange(1, domainIdx + 1).setValue('Domain');
    headers.push('Domain');
  }

  // ── Ensure Founder 1…N columns at the END ─────────────────
  var founderColIdxs = [];
  for (var h = 0; h < headers.length; h++) {
    if (/^Founder \d+$/.test(headers[h])) founderColIdxs.push(h);
  }

  var existingF = founderColIdxs.length;
  var need      = maxFounders - existingF;

  if (need > 0) {
    var nextCol = headers.length;
    for (var f = existingF; f < maxFounders; f++) {
      sheet.getRange(1, nextCol + 1).setValue('Founder ' + (f + 1));
      founderColIdxs.push(nextCol);
      headers.push('Founder ' + (f + 1));
      nextCol++;
    }
  }

  // Re-read after column additions
  if (need > 0) {
    allValues = sheet.getDataRange().getValues();
    headers   = allValues[0].map(String);
    domainIdx = headers.indexOf('Domain');
    founderColIdxs = [];
    for (var hh = 0; hh < headers.length; hh++) {
      if (/^Founder \d+$/.test(headers[hh])) founderColIdxs.push(hh);
    }
  }

  // ── Build batch output arrays ──────────────────────────────
  var numDataRows    = allValues.length - 1;
  var domainOutput   = [];
  var founderOutput  = [];
  var processedCount = 0;

  for (var r = 1; r < allValues.length; r++) {
    var row     = allValues[r];
    var isEmpty = !row.some(function(v) { return v !== '' && v !== null && v !== undefined; });

    if (isEmpty) {
      domainOutput.push(['']);
      founderOutput.push(founderColIdxs.map(function() { return ''; }));
      continue;
    }

    // Domain — prefer Contact Email, fall back to Website
    var contactEmail = (contactEmailIdx >= 0 && row[contactEmailIdx])
                       ? String(row[contactEmailIdx]).trim() : '';
    var domain = domainFromEmail(contactEmail);
    if (!domain) {
      var website = row[websiteIdx] ? String(row[websiteIdx]).trim() : '';
      domain = extractDomain(website);
    }
    domainOutput.push([domain]);

    // Split founders
    var foundersRaw = row[foundersIdx] ? String(row[foundersIdx]).trim() : '';
    var parts       = splitFounders(foundersRaw);
    founderOutput.push(founderColIdxs.map(function(colIdx, pos) {
      return parts[pos] || '';
    }));

    processedCount++;
  }

  // ── Batch write ───────────────────────────────────────────
  if (numDataRows > 0) {
    sheet.getRange(2, domainIdx + 1, numDataRows, 1).setValues(domainOutput);
    if (founderColIdxs.length > 0) {
      sheet.getRange(2, founderColIdxs[0] + 1, numDataRows, founderColIdxs.length)
           .setValues(founderOutput);
    }
  }

  SpreadsheetApp.flush();
  return processedCount;
}
