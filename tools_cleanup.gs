// ============================================================
// tools_cleanup.gs — Clean up empty cells across all spreadsheets
// ============================================================

/**
 * Removes trailing empty rows and columns from every sheet
 * in all three spreadsheets to prevent the 10M cell limit error.
 */
function cleanUpAllSheets() {
  var targets = [
    { ss: SpreadsheetApp.getActiveSpreadsheet(), label: 'Automation (this spreadsheet)' },
    { ss: SpreadsheetApp.openById(CONFIG.CRUNCHBASE_SS_ID), label: 'ABM_R&D_Tax Credit' },
    { ss: SpreadsheetApp.openById(CONFIG.APOLLO_SS_ID),     label: 'Apollo Leads - Cleaner' },
  ];

  var totalCellsDeleted = 0;
  var report            = [];

  targets.forEach(function(target) {
    var ss      = target.ss;
    var sheets  = ss.getSheets();
    var ssCells = 0;

    sheets.forEach(function(sheet) {
      var maxRows = sheet.getMaxRows();
      var lastRow = Math.max(sheet.getLastRow(), 1);
      var maxCols = sheet.getMaxColumns();
      var lastCol = Math.max(sheet.getLastColumn(), 1);

      var rowsToDelete = maxRows - lastRow;
      var colsToDelete = maxCols - lastCol;

      if (rowsToDelete > 0) {
        sheet.deleteRows(lastRow + 1, rowsToDelete);
        ssCells += rowsToDelete * maxCols;
      }
      if (colsToDelete > 0) {
        sheet.deleteColumns(lastCol + 1, colsToDelete);
        ssCells += colsToDelete * lastRow;
      }
    });

    totalCellsDeleted += ssCells;
    report.push('• ' + target.label + ': ' +
      (ssCells > 0 ? ssCells.toLocaleString() + ' empty cells removed' : 'already clean'));
  });

  SpreadsheetApp.getUi().alert(
    '🧹 Cleanup Complete!\n\n' +
    report.join('\n') + '\n\n' +
    'Total cells removed: ' + totalCellsDeleted.toLocaleString() + '\n\n' +
    (totalCellsDeleted > 0
      ? 'The 10M cell limit should no longer block you.'
      : 'No extra empty cells found. If you still hit the limit, the actual data is too large — consider archiving old rows.')
  );
}

// ============================================================
// Fix "#ERROR! Formula parse error" cells (usually phone numbers
// that start with "+", "=" or "-" and got read as formulas).
//
// Recovers the original text from the broken formula and rewrites it
// as PLAIN TEXT in the Apollo Leads / Remaining / Insert File tabs.
// ============================================================
function fixFormulaErrorCells() {
  var apolloSS = SpreadsheetApp.openById(CONFIG.APOLLO_SS_ID);
  var names = [CONFIG.APOLLO_SHEET_NAME, CONFIG.APOLLO_REMAINING_NAME, CONFIG.APOLLO_INSERT_NAME];
  var totalFixed = 0;

  names.forEach(function(name) {
    var sh = apolloSS.getSheetByName(name);
    if (!sh || sh.getLastRow() < 1 || sh.getLastColumn() < 1) return;

    var rng      = sh.getDataRange();
    var formulas = rng.getFormulas();
    var values   = rng.getValues();
    var changed  = false;

    for (var r = 0; r < formulas.length; r++) {
      for (var c = 0; c < formulas[r].length; c++) {
        var f = formulas[r][c];
        if (f && f.charAt(0) === '=') {
          // Drop the leading "=" to recover the literal text,
          // e.g. "=+1 (555) 123-4567" -> "+1 (555) 123-4567"
          values[r][c] = f.substring(1);
          changed = true;
          totalFixed++;
        }
      }
    }

    if (changed) {
      rng.setNumberFormat('@'); // whole block to text so it never re-parses
      rng.setValues(values);
    }
  });

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    '🧹 Fixed ' + totalFixed + ' formula-error cell(s) into plain text.\n\n' +
    'Tip: before pasting new leads, set the "' + CONFIG.APOLLO_INSERT_NAME +
    '" tab to Plain Text (Format → Number → Plain text) so phone numbers ' +
    'never turn into #ERROR again.'
  );
}
