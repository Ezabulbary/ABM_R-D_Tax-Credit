// ============================================================
// config.gs — Global Configuration & API Key Management
// ============================================================

var CONFIG = {
  MAIN_SHEET_NAME         : 'Main_Crunchbase',
  PERMUTATOR_SHEET_NAME   : 'Email Permutator',
  FINAL_FORMAT_SHEET_NAME : 'Final Format',
  APOLLO_SHEET_NAME       : 'Apollo Leads',
  SLEEP_MS                : 700,            // ms between each Reoon API call
  MAX_RUN_MS              : 5 * 60 * 1000, // 5-min safety limit per run
  VALID_STATUSES          : ['safe', 'catch_all'],  // statuses included in Final Format
  EXTERNAL_SPREADSHEET_ID : '1_8HJOyl9pIKTmTSrqPZ3oD0tMwo4-7fdifQm7kzem0c',             // Enter an existing spreadsheet ID to export there
};

// ── API Key Helpers ───────────────────────────────────────────

/** Returns array of all stored Reoon API keys (in order). */
function getApiKeys() {
  var props = PropertiesService.getScriptProperties();
  var keys = [];
  for (var i = 1; i <= 20; i++) {
    var k = props.getProperty('REOON_API_KEY_' + i);
    if (!k) break;
    keys.push(k.trim());
  }
  return keys;
}

/** UI dialog to add / replace Reoon API keys. */
function manageApiKeys() {
  var ui    = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var curr  = getApiKeys();

  var currentInfo = curr.length > 0
    ? curr.map(function(k, i) {
        return 'Key ' + (i + 1) + ': ' + k.substring(0, 10) + '…';
      }).join('\n')
    : '(No keys saved)';

  var res = ui.prompt(
    '⚙️ Reoon API Keys',
    'Current keys:\n' + currentInfo + '\n\n' +
    'Enter new key(s) — separate multiple keys with a comma:\n' +
    'Example:  key_abc123, key_def456\n\n' +
    '(Leave blank to keep existing keys unchanged)',
    ui.ButtonSet.OK_CANCEL
  );

  if (res.getSelectedButton() !== ui.Button.OK) return;

  var input = res.getResponseText().trim();
  if (!input) { ui.alert('ℹ️ No changes made.'); return; }

  // Clear old keys
  for (var j = 1; j <= 20; j++) props.deleteProperty('REOON_API_KEY_' + j);

  // Save new keys
  var newKeys = input.split(',')
    .map(function(k) { return k.trim(); })
    .filter(function(k) { return k.length > 0; });

  newKeys.forEach(function(key, idx) {
    props.setProperty('REOON_API_KEY_' + (idx + 1), key);
  });

  ui.alert('✅ ' + newKeys.length + ' API key(s) saved!\n\n' +
    newKeys.map(function(k, i) {
      return 'Key ' + (i + 1) + ': ' + k.substring(0, 10) + '…';
    }).join('\n'));
}
