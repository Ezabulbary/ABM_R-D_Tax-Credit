// ============================================================
// info_dashboard.gs — the shared, LIVE "Info" tab
//
// The "Info" tab exists in BOTH spreadsheets:
//   - "ABM_R&D_Tax Credit"
//   - "ABM R&D Tax Credit - Automation"
//
// It is a LIVE dashboard: the numbers are Google Sheets FORMULAS that
// point at the local "Final Format" tab, so they auto-recalculate the
// moment Final Format changes — even without running the workflow. The
// workflow only refreshes the layout, the "last run" time, and re-locks
// the tab.
//
// Final Format column reference (used by the formulas):
//   B = Domain,  E = Email,  F = Verification Status
//
// PROTECTION: sheet-protected so only the OWNER can edit or delete it.
// ============================================================

// Writes + protects the Info tab in BOTH shared spreadsheets.
// stats.addedThisRun = number of new leads added in the run that called this.
function _updateInfoTabs(stats) {
  stats = stats || {};
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  // Quoted local sheet name for use inside formulas: 'Final Format'
  var ff = "'" + CONFIG.FINAL_FORMAT_SHEET_NAME + "'";

  // Value column mixes static fields (run time / count) and LIVE formulas.
  var rows = [
    ['ABM Lead Gen — Info Dashboard', ''],
    ['Last workflow run',  now],
    ['Added in last run',  (stats.addedThisRun || 0)],
    ['', ''],
    ['Total leads (live)',    '=COUNTA(' + ff + '!E2:E)'],
    ['Unique domains (live)', '=COUNTUNIQUE(' + ff + '!B2:B)'],
    ['', ''],
    ['Safe (live)',      '=COUNTIF(' + ff + '!F2:F,"safe")'],
    ['Catch-all (live)', '=COUNTIF(' + ff + '!F2:F,"catch_all")'],
    ['Invalid (live)',   '=COUNTIF(' + ff + '!F2:F,"invalid")'],
    ['Other / unknown',  '=MAX(0,COUNTA(' + ff + '!E2:E)-COUNTIF(' + ff + '!F2:F,"safe")-COUNTIF(' + ff + '!F2:F,"catch_all")-COUNTIF(' + ff + '!F2:F,"invalid"))'],
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
// setValues stores "=..." strings as live formulas automatically.
function _writeInfoSheet(ss, rows) {
  var sheet = ss.getSheetByName(CONFIG.INFO_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG.INFO_SHEET_NAME, 0); // create as first tab

  var body = sheet.getRange(1, 1, rows.length, 2);
  body.breakApart();     // undo any previous merge so setValues is safe
  sheet.clearContents(); // owner bypasses protection

  body.setValues(rows);  // "=..." entries become live formulas

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
