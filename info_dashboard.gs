// ============================================================
// info_dashboard.gs — the shared "Info" tab
//
// The "Info" tab exists in BOTH spreadsheets:
//   - "ABM_R&D_Tax Credit"
//   - "ABM R&D Tax Credit - Automation"
// It is updated together with the Final Format tab (at the end of
// Step 4) and shows a small stats dashboard.
//
// PROTECTION: the tab is sheet-protected so that only the spreadsheet
// OWNER can edit or delete it. Other editors can view it but cannot
// change or remove it. (The owner running the script bypasses the
// protection, so the dashboard still updates automatically.)
// ============================================================

// Builds a stats object from the primary Final Format sheet values
// (the 2D array including the header row).
function _computeFinalStats(finalValues) {
  var statusIdx = FINAL_COLUMNS.indexOf('Verification Status');
  var domainIdx = FINAL_COLUMNS.indexOf('Domain');
  var emailIdx  = FINAL_COLUMNS.indexOf('Email');

  var stats  = { total: 0, safe: 0, catchAll: 0, invalid: 0, other: 0, domains: 0, addedThisRun: 0 };
  var domSet = Object.create(null);

  for (var i = 1; i < finalValues.length; i++) {
    var row   = finalValues[i];
    var email = emailIdx  >= 0 ? row[emailIdx]  : '';
    var dom   = domainIdx >= 0 ? String(row[domainIdx] || '').trim().toLowerCase() : '';
    if (!email && !dom) continue; // skip blank rows

    stats.total++;

    var st = statusIdx >= 0 ? String(row[statusIdx] || '').trim().toLowerCase() : '';
    if      (st === 'safe')      stats.safe++;
    else if (st === 'catch_all') stats.catchAll++;
    else if (st === 'invalid')   stats.invalid++;
    else                         stats.other++;

    if (dom) domSet[dom] = true;
  }

  stats.domains = Object.keys(domSet).length;
  return stats;
}

// Writes + protects the Info tab in BOTH shared spreadsheets.
function _updateInfoTabs(stats) {
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  var rows = [
    ['ABM Lead Gen — Info Dashboard', ''],
    ['Last updated', now],
    ['', ''],
    ['Total leads (Final Format)', stats.total],
    ['Added in last run',          stats.addedThisRun || 0],
    ['Unique domains',             stats.domains],
    ['', ''],
    ['Safe',              stats.safe],
    ['Catch-all',         stats.catchAll],
    ['Invalid',           stats.invalid],
    ['Other / unknown',   stats.other],
    ['', ''],
    ['Do not delete this tab', 'Only the sheet owner can edit or delete it.']
  ];

  _getSharedSpreadsheets().forEach(function(t) {
    try {
      _writeInfoSheet(t.ss, rows);
    } catch (e) {
      Logger.log('[Info] Failed to update Info tab in ' + t.label + ': ' + e.message);
    }
  });
}

// Writes the dashboard rows into one spreadsheet's Info tab and locks it.
function _writeInfoSheet(ss, rows) {
  var sheet = ss.getSheetByName(CONFIG.INFO_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG.INFO_SHEET_NAME, 0); // create as first tab

  var body = sheet.getRange(1, 1, rows.length, 2);
  body.breakApart();     // undo any previous merge so setValues is safe
  sheet.clearContents(); // owner bypasses protection

  body.setValues(rows);

  // Title styling (row 1)
  sheet.getRange(1, 1, 1, 2)
    .setBackground('#0d1b2a').setFontColor('#93c5fd')
    .setFontWeight('bold').setFontSize(13);
  sheet.setColumnWidth(1, 280);
  sheet.setColumnWidth(2, 280);
  sheet.setFrozenRows(1);

  _protectInfoSheet(sheet);
  SpreadsheetApp.flush();
}

// Sheet-protects the Info tab so only the owner can edit/delete it.
function _protectInfoSheet(sheet) {
  try {
    var existing = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    var protection = existing.length > 0 ? existing[0] : sheet.protect();
    protection.setDescription('Info tab — only the owner can edit/delete');

    // Remove every non-owner editor. The owner is never removed and
    // always keeps edit rights, so the script can still update it.
    var editorEmails = protection.getEditors().map(function(u) { return u.getEmail(); });
    if (editorEmails.length > 0) protection.removeEditors(editorEmails);
    if (protection.canDomainEdit()) protection.setDomainEdit(false);
  } catch (e) {
    // If the runner is not the owner, protection can't be changed — log only.
    Logger.log('[Info] Could not set protection: ' + e.message);
  }
}
