// ============================================================
// tools_cleanup.gs — Utility to clean up empty cells
// ============================================================

/**
 * Removes all empty rows at the bottom and empty columns at the right 
 * of every sheet in the active workbook. This is essential to prevent 
 * the 10,000,000 cell limit error in Google Sheets.
 */
function cleanUpWorkbook() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var totalCellsDeleted = 0;
  
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var maxRows = sheet.getMaxRows();
    var lastRow = sheet.getLastRow();
    var maxCols = sheet.getMaxColumns();
    var lastCol = sheet.getLastColumn();
    
    // Ensure we don't delete everything if the sheet is completely empty
    if (lastRow === 0) lastRow = 1;
    if (lastCol === 0) lastCol = 1;
    
    var rowsToDelete = maxRows - lastRow;
    var colsToDelete = maxCols - lastCol;
    
    if (rowsToDelete > 0) {
      sheet.deleteRows(lastRow + 1, rowsToDelete);
      totalCellsDeleted += (rowsToDelete * maxCols);
    }
    
    // After rows are deleted, the maxCols remains, but total cells calculation changes slightly
    // We already counted the deleted row cells based on maxCols. 
    // Now delete columns for the remaining rows (lastRow).
    if (colsToDelete > 0) {
      sheet.deleteColumns(lastCol + 1, colsToDelete);
      totalCellsDeleted += (colsToDelete * lastRow);
    }
  }
  
  if (totalCellsDeleted > 0) {
    SpreadsheetApp.getUi().alert(
      '🧹 Cleanup Complete!\n\n' +
      'Successfully removed ' + totalCellsDeleted.toLocaleString() + ' empty cells from the workbook.\n\n' +
      'You should no longer see the 10,000,000 cell limit error. You can now run Step 4 again.'
    );
  } else {
    SpreadsheetApp.getUi().alert(
      '🧹 Cleanup Complete!\n\n' +
      'No empty extra cells found at the edges of your sheets. If you still get the error, you may have too much actual data in this workbook.'
    );
  }
}
