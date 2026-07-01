// ============================================================
// step1b_apolloProcess.gs — Step 1.2: Apollo Leads split + domain
//
// Reads from: "Apollo Leads - Cleaner" spreadsheet
//             → any raw source tab (user selects)
// Writes to:  → "Apollo Leads" tab  (valid, unique rows)
//             → "Remaining"    tab  (duplicates / no email)
//             → adds Domain column at end of "Apollo Leads" tab
// ============================================================

var APOLLO_TARGET_COLUMNS = [
  'First Name', 'Last Name', 'Title', 'Company Name', 'Company Name for Emails',
  'Email', 'Verification Status', 'Email Status', 'Seniority', 'Departments',
  'Sub Departments', 'Mobile Phone', 'Corporate Phone', 'Other Phone', 'Stage',
  '# Employees', 'Industry', 'Keywords', 'Person Linkedin Url', 'Website',
  'Company Linkedin Url', 'Facebook Url', 'Twitter Url', 'City', 'State',
  'Country', 'Company Address', 'Company City', 'Company State', 'Company Country',
  'Company Phone', 'Company Founded Year', 'headline', 'Technologies',
  'Annual Revenue', 'Total Funding', 'Latest Funding 6 Month Growth',
  'Latest Funding 12 Month Growth', 'Latest Funding 24 Month Growth',
  'Email Sent', 'Email Open', 'Email Bounced', 'Replied', 'Demoed',
  'Number of Retail Locations', 'Apollo Contact Id', 'Apollo Account Id'
];

/**
 * Main entry for Step 1.2.
 * Reads the fixed "Insert File" tab, splits rows into Apollo Leads +
 * Remaining, then adds the Domain column. No tab picker prompt.
 */
function processApolloSheet() {
  var ui       = SpreadsheetApp.getUi();
  var apolloSS = SpreadsheetApp.openById(CONFIG.APOLLO_SS_ID);

  // Get or create destination tabs
  var apolloSheet = apolloSS.getSheetByName(CONFIG.APOLLO_SHEET_NAME);
  if (!apolloSheet) apolloSheet = apolloSS.insertSheet(CONFIG.APOLLO_SHEET_NAME);

  var remainingSheet = apolloSS.getSheetByName(CONFIG.APOLLO_REMAINING_NAME);
  if (!remainingSheet) remainingSheet = apolloSS.insertSheet(CONFIG.APOLLO_REMAINING_NAME);

  // The raw leads always live in the fixed "Insert File" tab.
  var sourceSheet = apolloSS.getSheetByName(CONFIG.APOLLO_INSERT_NAME);

  // No source tab, or source tab is empty -> just refresh the Domain column.
  if (!sourceSheet || sourceSheet.getLastRow() < 2) {
    _addApolloDomainsColumn(apolloSheet);
    ui.alert(
      '✅ Step 1.2 Complete!\n\n' +
      'Domain column updated in "Apollo Leads" tab.\n' +
      '("' + CONFIG.APOLLO_INSERT_NAME + '" tab is empty — nothing to split.)' +
      (_isAutoMode() ? '\n\n➡️ Step 2 (Email Permutator) will auto-start in 30 seconds…' : '')
    );
    _continueChain('autoStep2Start'); // only chains in one-click mode
    return;
  }

  var result = _splitApolloData(sourceSheet, apolloSheet, remainingSheet);
  _addApolloDomainsColumn(apolloSheet);

  ui.alert(
    '✅ Step 1.2 Complete!\n\n' +
    '✔ Added to "Apollo Leads":  ' + result.valid     + ' rows\n' +
    '✔ Moved to "Remaining":     ' + result.remaining + ' rows\n\n' +
    'Domain column updated in "Apollo Leads" tab.' +
    (_isAutoMode() ? '\n\n➡️ Step 2 (Email Permutator) will auto-start in 30 seconds…' : '')
  );
  _continueChain('autoStep2Start'); // only chains in one-click mode
}

// ── Core: split source tab → Apollo Leads + Remaining ────────────

function _splitApolloData(sourceSheet, apolloSheet, remainingSheet) {
  var apolloLastRow = apolloSheet.getLastRow();
  var apolloHeaders;

  if (apolloLastRow === 0) {
    // Fresh sheet — write standard headers
    apolloSheet.getRange(1, 1, 1, APOLLO_TARGET_COLUMNS.length)
               .setValues([APOLLO_TARGET_COLUMNS]);
    apolloHeaders = APOLLO_TARGET_COLUMNS.slice();
    apolloLastRow = 1;
  } else {
    apolloHeaders = apolloSheet
      .getRange(1, 1, 1, apolloSheet.getLastColumn())
      .getValues()[0]
      .map(String);
  }

  // Ensure Remaining has matching headers
  if (remainingSheet.getLastRow() === 0) {
    remainingSheet.getRange(1, 1, 1, apolloHeaders.length).setValues([apolloHeaders]);
  }

  // Build existing email index from Apollo Leads
  var emailColIdx    = apolloHeaders.indexOf('Email');
  var existingEmails = Object.create(null);

  if (apolloLastRow > 1 && emailColIdx !== -1) {
    var existingEmailVals = apolloSheet
      .getRange(2, emailColIdx + 1, apolloLastRow - 1, 1)
      .getValues();
    for (var e = 0; e < existingEmailVals.length; e++) {
      var em = String(existingEmailVals[e][0]).trim().toLowerCase();
      if (em) existingEmails[em] = true;
    }
  }

  // Read source data
  var sourceData    = sourceSheet.getDataRange().getValues();
  if (sourceData.length < 2) return { valid: 0, remaining: 0 };

  var sourceHeaders = sourceData[0].map(String);
  var headerMap     = Object.create(null);
  sourceHeaders.forEach(function(h, i) { headerMap[h] = i; });

  var validRows     = [];
  var remainingRows = [];

  for (var i = 1; i < sourceData.length; i++) {
    var row    = sourceData[i];
    var newRow = apolloHeaders.map(function(h) {
      return (h in headerMap) ? row[headerMap[h]] : '';
    });

    if (emailColIdx === -1) {
      remainingRows.push(newRow);
      continue;
    }

    var email = String(newRow[emailColIdx] || '').trim();
    if (!email) {
      remainingRows.push(newRow);
      continue;
    }

    var emailLower = email.toLowerCase();
    if (emailLower in existingEmails) {
      remainingRows.push(newRow);
    } else {
      existingEmails[emailLower] = true;
      validRows.push(newRow);
    }
  }

  // Write valid rows to Apollo Leads
  if (validRows.length > 0) {
    apolloSheet.getRange(apolloLastRow + 1, 1, validRows.length, apolloHeaders.length)
               .setValues(validRows);
  }

  // Write duplicates/no-email rows to Remaining
  if (remainingRows.length > 0) {
    var remLastRow = Math.max(remainingSheet.getLastRow(), 1);
    remainingSheet.getRange(remLastRow + 1, 1, remainingRows.length, apolloHeaders.length)
                  .setValues(remainingRows);
  }

  SpreadsheetApp.flush();
  return { valid: validRows.length, remaining: remainingRows.length };
}

// ── Domain column: extract domain for every row in Apollo Leads ──

function _addApolloDomainsColumn(apolloSheet) {
  var data = apolloSheet.getDataRange().getValues();
  if (data.length < 2) return;

  var headers    = data[0].map(String);
  var emailIdx   = headers.indexOf('Email');
  var websiteIdx = headers.indexOf('Website');

  // Ensure Domain column exists at the END
  var domainIdx = headers.indexOf('Domain');
  if (domainIdx === -1) {
    domainIdx = headers.length;
    apolloSheet.getRange(1, domainIdx + 1).setValue('Domain');
  }

  var numDataRows  = data.length - 1;
  var domainOutput = [];

  for (var r = 1; r < data.length; r++) {
    var row     = data[r];
    var email   = (emailIdx >= 0 && row[emailIdx])   ? String(row[emailIdx]).trim()   : '';
    var domain  = domainFromEmail(email);
    if (!domain) {
      var website = (websiteIdx >= 0 && row[websiteIdx]) ? String(row[websiteIdx]).trim() : '';
      domain = extractDomain(website);
    }
    domainOutput.push([domain]);
  }

  apolloSheet.getRange(2, domainIdx + 1, numDataRows, 1).setValues(domainOutput);
  SpreadsheetApp.flush();
}
