// ============================================================
// config.gs — Global Configuration & API Key Management
// ============================================================

var CONFIG = {
  // ── Spreadsheet IDs ───────────────────────────────────────
  // "ABM_R&D_Tax Credit" — Main_Crunchbase + Final Format tabs live here
  CRUNCHBASE_SS_ID        : '1_8HJOyl9pIKTmTSrqPZ3oD0tMwo4-7fdifQm7kzem0c',
  // "Apollo Leads - Cleaner" — Apollo Leads + Remaining tabs live here
  APOLLO_SS_ID            : '14eta9RZ_S4chtLoSlYqn55BCisyrPjDTjX6Am1Zys1M',

  // ── Tab Names ─────────────────────────────────────────────
  MAIN_SHEET_NAME         : 'Main_Crunchbase',   // in CRUNCHBASE_SS
  APOLLO_SHEET_NAME       : 'Apollo Leads',       // in APOLLO_SS
  APOLLO_REMAINING_NAME   : 'Remaining',          // in APOLLO_SS
  PERMUTATOR_SHEET_NAME   : 'Email Permutator',  // in Automation SS (this spreadsheet)
  FINAL_FORMAT_SHEET_NAME : 'Final Format',       // in CRUNCHBASE_SS

  // ── Timing ────────────────────────────────────────────────
  SLEEP_MS                : 700,              // ms between Reoon API calls
  MAX_RUN_MS              : 4.5 * 60 * 1000, // 4.5 min safety limit per batch
  TRIGGER_DELAY_MS        : 30 * 1000,       // 30 sec between auto-triggers

  // ── Verification ─────────────────────────────────────────
  VALID_STATUSES          : ['safe', 'catch_all'],

  // ── BigQuery (Step 4 also pushes leads here) ──────────────
  // Fill these in AFTER following BigQuery_Setup_Guide.txt, then
  // set BQ_ENABLED to true. While BQ_ENABLED is false the whole
  // BigQuery part is skipped silently — the rest of the workflow
  // still runs and writes the "Final Format" tab as usual.
  BQ_ENABLED              : false,   // turn true once setup is done
  BQ_PROJECT_ID           : '',      // e.g. 'abm-lead-gen-471203'
  BQ_DATASET_ID           : 'abm_leads',     // BigQuery dataset name
  BQ_TABLE_ID             : 'final_format',  // BigQuery table name
};

// ── API Key Helpers ───────────────────────────────────────────

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

function manageApiKeys() {
  var ui    = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var curr  = getApiKeys();

  var currentInfo = curr.length > 0
    ? curr.map(function(k, i) { return 'Key ' + (i + 1) + ': ' + k.substring(0, 10) + '…'; }).join('\n')
    : '(No keys saved)';

  var res = ui.prompt(
    '⚙️ Reoon API Keys',
    'Current keys:\n' + currentInfo + '\n\n' +
    'Enter new key(s) — separate multiple keys with a comma:\n' +
    '(Leave blank to keep existing keys unchanged)',
    ui.ButtonSet.OK_CANCEL
  );

  if (res.getSelectedButton() !== ui.Button.OK) return;
  var input = res.getResponseText().trim();
  if (!input) { ui.alert('ℹ️ No changes made.'); return; }

  for (var j = 1; j <= 20; j++) props.deleteProperty('REOON_API_KEY_' + j);

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
