// ============================================================
// helpers.gs — Shared utility functions used across all steps
// ============================================================

/**
 * Strips protocol, www, path, port, query from a URL → bare domain.
 */
function extractDomain(url) {
  if (!url) return '';
  return url
    .replace(/^https?:\/\//i, '')
    .replace(/^www\./i, '')
    .split('/')[0]
    .split('?')[0]
    .split('#')[0]
    .split(':')[0]
    .toLowerCase()
    .trim();
}

/**
 * Extracts domain from an email address (part after @).
 */
function domainFromEmail(email) {
  if (!email) return '';
  var atIdx = email.lastIndexOf('@');
  if (atIdx === -1) return '';
  var d = email.slice(atIdx + 1).trim().toLowerCase();
  return d.indexOf('.') !== -1 ? d : '';
}

/**
 * Splits a comma/pipe/semicolon-separated founders string.
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
 * Parses "John Smith Jr" → { firstName: "John", lastName: "Smith Jr" }
 */
function parseFounderName(fullName) {
  if (!fullName) return { firstName: '', lastName: '' };
  var parts = fullName.trim().split(/\s+/);
  if (parts.length === 1) return { firstName: parts[0], lastName: '' };
  return { firstName: parts[0], lastName: parts.slice(1).join(' ') };
}

/**
 * Aligns sheet headers to a target list.
 * Adds missing columns, reorders to match target, appends extra columns at end.
 */
function alignHeaders(sheet, targetColumns) {
  var values  = sheet.getDataRange().getValues();
  var numRows = values.length;

  if (numRows === 1 && values[0].length === 1 && values[0][0] === '') {
    sheet.getRange(1, 1, 1, targetColumns.length).setValues([targetColumns]);
    SpreadsheetApp.flush();
    return targetColumns.slice();
  }

  var existingHeaders = values[0].map(function(h) { return String(h).trim(); });
  var extraColumns    = [];
  existingHeaders.forEach(function(h) {
    if (h && targetColumns.indexOf(h) === -1) extraColumns.push(h);
  });

  var finalHeaders  = targetColumns.concat(extraColumns);
  var alignedValues = [];

  for (var r = 0; r < numRows; r++) {
    var oldRow = values[r];
    var newRow = finalHeaders.map(function(header) {
      var oldIdx = existingHeaders.indexOf(header);
      return oldIdx !== -1 ? oldRow[oldIdx] : '';
    });
    alignedValues.push(newRow);
  }

  sheet.clearContents();
  sheet.getRange(1, 1, alignedValues.length, finalHeaders.length).setValues(alignedValues);
  SpreadsheetApp.flush();
  return finalHeaders;
}

/**
 * Deletes all project triggers for a given handler function name.
 */
function _deleteTriggersForFunction(fnName) {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === fnName) ScriptApp.deleteTrigger(t);
  });
}

/**
 * Schedules a time-based trigger for fnName after CONFIG.TRIGGER_DELAY_MS.
 */
function _scheduleTrigger(fnName) {
  _deleteTriggersForFunction(fnName);
  ScriptApp.newTrigger(fnName)
    .timeBased()
    .after(CONFIG.TRIGGER_DELAY_MS)
    .create();
}

/**
 * Schedules a time-based trigger for fnName after a custom delay (ms).
 */
function _scheduleTriggerDelayed(fnName, delayMs) {
  _deleteTriggersForFunction(fnName);
  ScriptApp.newTrigger(fnName)
    .timeBased()
    .after(delayMs)
    .create();
}

// ── Workflow mode (one-click vs single-step) ──────────────────────
// 'auto'   -> the one-click full workflow; each step schedules the next.
// 'manual' -> a single menu step; nothing chains to the next step.
var PROP_WORKFLOW_MODE = 'WORKFLOW_MODE';

function _setWorkflowMode(mode) {
  PropertiesService.getScriptProperties().setProperty(PROP_WORKFLOW_MODE, mode);
}

function _isAutoMode() {
  return PropertiesService.getScriptProperties().getProperty(PROP_WORKFLOW_MODE) === 'auto';
}

// Schedules the next step ONLY when running the one-click workflow.
// In manual (single-step) mode this does nothing.
function _continueChain(fnName) {
  if (_isAutoMode()) _scheduleTrigger(fnName);
}

// ── Spreadsheet handles ───────────────────────────────────────────
// The Automation spreadsheet is the one the script is bound to.
function _getAutomationSS() { return SpreadsheetApp.getActiveSpreadsheet(); }
function _getCrunchbaseSS() { return SpreadsheetApp.openById(CONFIG.CRUNCHBASE_SS_ID); }
function _getApolloSS()     { return SpreadsheetApp.openById(CONFIG.APOLLO_SS_ID); }

// Returns the two spreadsheets that share the Info + Final Format tabs.
function _getSharedSpreadsheets() {
  return [
    { ss: _getAutomationSS(), label: 'ABM R&D Tax Credit - Automation' },
    { ss: _getCrunchbaseSS(), label: 'ABM_R&D_Tax Credit' }
  ];
}
