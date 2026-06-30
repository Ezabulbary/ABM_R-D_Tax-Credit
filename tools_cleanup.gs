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
