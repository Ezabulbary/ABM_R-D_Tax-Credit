// ============================================================
// step1_processMain.gs — Domain extraction + Founder splitting
// ============================================================

var MAIN_TARGET_COLUMNS = [
  'How Many Leads',
  'Verification Status',
  'Campaign Status',
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

var APOLLO_TARGET_COLUMNS = [
  'First Name',
  'Last Name',
  'Title',
  'Company Name',
  'Company Name for Emails',
  'Email',
  'Verification Status',
  'Email Status',
  'Seniority',
  'Departments',
  'Sub Departments',
  'Mobile Phone',
  'Corporate Phone',
  'Other Phone',
  'Stage',
  '# Employees',
  'Industry',
  'Keywords',
  'Person Linkedin Url',
  'Website',
  'Company Linkedin Url',
  'Facebook Url',
  'Twitter Url',
  'City',
  'State',
  'Country',
  'Company Address',
  'Company City',
  'Company State',
  'Company Country',
  'Company Phone',
  'Company Founded Year',
  'headline',
  'Technologies',
  'Annual Revenue',
  'Total Funding',
  'Latest Funding 6 Month Growth',
  'Latest Funding 12 Month Growth',
  'Latest Funding 24 Month Growth',
  'Email Sent',
  'Email Open',
  'Email Bounced',
  'Replied',
  'Demoed',
  'Number of Retail Locations',
  'Apollo Contact Id',
  'Apollo Account Id'
];

/**
 * Main entry: aligns headers, adds Domain column & Founder 1…N columns at the END of "Main_Crunchbase".
 * Returns number of processed data rows.
 */
function processMainSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.MAIN_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.MAIN_SHEET_NAME);
  }

  // ── 1. Align headers to target list ───────────────────────
  var headers = alignHeaders(sheet, MAIN_TARGET_COLUMNS);

  var allValues = sheet.getDataRange().getValues();
  var websiteIdx      = headers.indexOf('Website');
  var foundersIdx     = headers.indexOf('Founders');
  var contactEmailIdx = headers.indexOf('Contact Email');

  if (websiteIdx  === -1) throw new Error('"Website" column not found.');
  if (foundersIdx === -1) throw new Error('"Founders" column not found.');

  // ── 2. Determine max founders across all rows ─────────────
  var maxFounders = 0;
  for (var i = 1; i < allValues.length; i++) {
    var raw = allValues[i][foundersIdx] ? String(allValues[i][foundersIdx]).trim() : '';
    if (raw) maxFounders = Math.max(maxFounders, splitFounders(raw).length);
  }
  maxFounders = Math.max(maxFounders, 1);

  // ── 3. Ensure Domain column exists at the END ─────────────
  var domainIdx = headers.indexOf('Domain');
  if (domainIdx === -1) {
    domainIdx = headers.length;
    sheet.getRange(1, domainIdx + 1).setValue('Domain');
    headers.push('Domain');
  }

  // ── 4. Ensure Founder 1…N columns exist at the END ────────
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

  // ── 5. Re-read sheet data to get correct column count ─────
  if (need > 0 || headers.indexOf('Domain') !== domainIdx) {
    allValues = sheet.getDataRange().getValues();
    headers   = allValues[0].map(String);
    domainIdx = headers.indexOf('Domain');
    founderColIdxs = [];
    for (var hh = 0; hh < headers.length; hh++) {
      if (/^Founder \d+$/.test(headers[hh])) founderColIdxs.push(hh);
    }
  }

  // ── 6. Build batch output arrays and write ────────────────
  var numDataRows   = allValues.length - 1;
  var domainOutput  = [];
  var founderOutput = [];
  var processedCount = 0;

  for (var r = 1; r < allValues.length; r++) {
    var row = allValues[r];
    var isEmptyRow = !row.some(function(v) { return v !== '' && v !== null && v !== undefined; });

    if (isEmptyRow) {
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

  // ── Batch write ──────────────────────────────────────────
  if (numDataRows > 0) {
    sheet.getRange(2, domainIdx + 1, numDataRows, 1).setValues(domainOutput);

    if (founderColIdxs.length > 0) {
      sheet.getRange(2, founderColIdxs[0] + 1, numDataRows, founderColIdxs.length)
           .setValues(founderOutput);
    }
  }

  SpreadsheetApp.flush();

  // ── Also extract domains for Apollo Leads tab ────────────
  _processApolloLeadsDomains();

  return processedCount;
}

// ── Helpers ──────────────────────────────────────────────────

/**
 * Extracts domain from an email address (part after @).
 */
function domainFromEmail(email) {
  if (!email) return '';
  var atIdx = email.lastIndexOf('@');
  if (atIdx === -1) return '';
  var domain = email.slice(atIdx + 1).trim().toLowerCase();
  return domain.indexOf('.') !== -1 ? domain : '';
}

/**
 * Strips protocol, www, path, port, query from a URL → bare domain.
 */
function extractDomain(url) {
  if (!url) return '';
  var d = url
    .replace(/^https?:\/\//i, '')  // remove protocol
    .replace(/^www\./i, '')         // remove www.
    .split('/')[0]                  // remove path
    .split('?')[0]                  // remove query
    .split('#')[0]                  // remove fragment
    .split(':')[0]                  // remove port
    .toLowerCase()
    .trim();
  return d;
}

/**
 * Splits a comma-separated founders string into trimmed name array.
 */
function splitFounders(str) {
  if (!str) return [];
  var sep = str.indexOf(',') !== -1 ? ',' :
            str.indexOf('|') !== -1 ? '|' :
            str.indexOf(';') !== -1 ? ';' : ',';
  return str.split(sep)
    .map(function(s) { return s.trim(); })
    .filter(function(s) { return s.length > 0; });
}

/**
 * Align headers of a sheet to a target list of columns.
 * Adds missing columns, keeps order of target columns, and appends extra columns at the end.
 */
function alignHeaders(sheet, targetColumns) {
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  var numRows = values.length;
  var numCols = values[0].length;

  // If sheet is completely empty or has just one empty cell
  if (numRows === 1 && numCols === 1 && values[0][0] === '') {
    sheet.getRange(1, 1, 1, targetColumns.length).setValues([targetColumns]);
    SpreadsheetApp.flush();
    return targetColumns;
  }

  var existingHeaders = values[0].map(String).map(function(h) { return h.trim(); });

  // Identify extra columns (e.g. Domain, Founder 1, etc.)
  var extraColumns = [];
  existingHeaders.forEach(function(h) {
    if (h && targetColumns.indexOf(h) === -1) {
      extraColumns.push(h);
    }
  });

  var finalHeaders = targetColumns.concat(extraColumns);

  // Re-map all rows to the final headers structure
  var alignedValues = [];
  for (var r = 0; r < numRows; r++) {
    var oldRow = values[r];
    var newRow = [];
    finalHeaders.forEach(function(header) {
      var oldIdx = existingHeaders.indexOf(header);
      newRow.push(oldIdx !== -1 ? oldRow[oldIdx] : '');
    });
    alignedValues.push(newRow);
  }

  // Clear existing sheet and write new aligned values
  sheet.clearContents();
  sheet.getRange(1, 1, alignedValues.length, finalHeaders.length).setValues(alignedValues);
  SpreadsheetApp.flush();

  return finalHeaders;
}

/**
 * Extracts domain from the "Email" / "Website" columns of the "Apollo Leads" sheet
 * and writes it into a "Domain" column at the END of that sheet.
 */
function _processApolloLeadsDomains() {
  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var apolloSheet = ss.getSheetByName(CONFIG.APOLLO_SHEET_NAME);
  if (!apolloSheet) {
    apolloSheet = ss.insertSheet(CONFIG.APOLLO_SHEET_NAME);
  }

  // Align headers of Apollo Leads to ensure "Email", "Website", "Verification Status" etc. are properly set up
  var headers = alignHeaders(apolloSheet, APOLLO_TARGET_COLUMNS);

  var data = apolloSheet.getDataRange().getValues();
  if (data.length < 2) return;  // header only — nothing to process

  var emailIdx   = headers.indexOf('Email');
  var websiteIdx = headers.indexOf('Website');
  if (emailIdx === -1 && websiteIdx === -1) return;

  // Ensure Domain column exists at the END
  var domainIdx = headers.indexOf('Domain');
  if (domainIdx === -1) {
    domainIdx = headers.length;
    apolloSheet.getRange(1, domainIdx + 1).setValue('Domain');
  }

  // Build domain output in one batch
  var numDataRows  = data.length - 1;
  var domainOutput = [];
  for (var r = 1; r < data.length; r++) {
    var email = (emailIdx >= 0 && data[r][emailIdx]) ? String(data[r][emailIdx]).trim() : '';
    var domain = domainFromEmail(email);
    if (!domain) {
      var website = (websiteIdx >= 0 && data[r][websiteIdx]) ? String(data[r][websiteIdx]).trim() : '';
      domain = extractDomain(website);
    }
    domainOutput.push([domain]);
  }

  apolloSheet.getRange(2, domainIdx + 1, numDataRows, 1).setValues(domainOutput);
  SpreadsheetApp.flush();
}
