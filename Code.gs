// ============================================================
// Code.gs — onOpen menu + step runners (entry points)
//
// Spreadsheet layout (3 tabs each):
//   "ABM_R&D_Tax Credit"               -> Info, Main_Crunchbase, Final Format
//   "ABM R&D Tax Credit - Automation"  -> Info, Email Permutator, Final Format
//   "Apollo Leads - Cleaner"           -> Apollo Leads, Insert File, Remaining
//
// Info + Final Format are SHARED tabs: they exist in BOTH the
// Crunchbase and Automation spreadsheets and are kept in sync.
//
// TWO WAYS TO RUN:
//   1. ONE-CLICK  -> "▶️ RUN FULL WORKFLOW": Step 1.1 -> 1.2 -> 2 -> 3
//                    -> 4 -> BigQuery, each step auto-starts the next.
//   2. STEP BY STEP -> run any single item; it runs ONLY that step and
//                    does NOT auto-continue to the next.
//
// The difference is the workflow mode flag (_setWorkflowMode):
//   'auto'   = one-click, each step schedules the next.
//   'manual' = single step, nothing chains.
// ============================================================

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 Lead Gen Tools')
    // ── Option 1: one click does everything ──
    .addItem('▶️  RUN FULL WORKFLOW (one click, Step 1 → 4 → BigQuery)', 'runFullWorkflow')
    .addSeparator()
    // ── Option 2: run each step on its own ──
    .addItem('Step 1.1 — Crunchbase: Domain & Founders Split', 'runStep1aOnly')
    .addItem('Step 1.2 — Apollo: Process, Split & Domain',     'runStep1b')
    .addItem('Step 2 — Build Email Permutator',                'runStep2')
    .addSeparator()
    .addSubMenu(ui.createMenu('Step 3 — Verify Emails (Reoon)')
      .addItem('▶️  Start Verification',    'startVerification')
      .addItem('⏩  Resume Verification',   'resumeVerification')
      .addItem('⛔  Stop Verification',     'stopVerification')
      .addItem('📊  Check Progress',        'checkProgress')
    )
    .addSubMenu(ui.createMenu('Step 4 — Build Final Format')
      .addItem('▶️  Start Building',        'startFinalFormat')
      .addItem('⏩  Resume Building',        'resumeFinalFormat')
      .addItem('⛔  Stop Building',          'stopFinalFormat')
    )
    .addSeparator()
    .addItem('⚙️  Manage Reoon API Keys',  'manageApiKeys')
    .addItem('🧹  Clean Up All Sheets',    'cleanUpAllSheets')
    .addToUi();
}

// ════════════════════════════════════════════════════════════
//  OPTION 1 — ONE-CLICK FULL WORKFLOW
// ════════════════════════════════════════════════════════════
// Sets auto mode, then runs Step 1.1 and Step 1.2 in the foreground
// (Step 1.2 reads the fixed "Insert File" tab — no prompt). Step 1.2
// then schedules Step 2, and every step after that chains itself.
function runFullWorkflow() {
  _setWorkflowMode('auto');

  // Step 1.1 — Main_Crunchbase
  try {
    var count = processMainSheet();
    SpreadsheetApp.getUi().alert(
      '✅ Step 1.1 Complete!\n\n' +
      '• Domain column updated\n' +
      '• Founders split: ' + count + ' row(s) processed\n\n' +
      '➡️ Continuing to Step 1.2 (Apollo)…'
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Step 1.1 Error:\n\n' + e.message);
    Logger.log(e.stack);
    return;
  }

  // Step 1.2 — Apollo (schedules Step 2 automatically in auto mode)
  try {
    processApolloSheet();
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Step 1.2 Error:\n\n' + e.message);
    Logger.log(e.stack);
  }
}

// ════════════════════════════════════════════════════════════
//  OPTION 2 — SINGLE STEP RUNNERS (no auto-continue)
// ════════════════════════════════════════════════════════════

// Step 1.1 only — Main_Crunchbase domain extraction + founders split.
function runStep1aOnly() {
  _setWorkflowMode('manual');
  try {
    var count = processMainSheet();
    SpreadsheetApp.getUi().alert(
      '✅ Step 1.1 Complete!\n\n' +
      '• Domain column updated\n' +
      '• Founders split: ' + count + ' row(s) processed\n\n' +
      'Sheet: Main_Crunchbase (ABM_R&D_Tax Credit)'
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Step 1.1 Error:\n\n' + e.message);
    Logger.log(e.stack);
  }
}

// Step 1.2 only — Apollo Leads split + domain (reads "Insert File" tab).
function runStep1b() {
  _setWorkflowMode('manual');
  try {
    processApolloSheet();
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Step 1.2 Error:\n\n' + e.message);
    Logger.log(e.stack);
  }
}

// Step 2 only — Build Email Permutator (with append prompt, no chaining).
function runStep2() {
  _setWorkflowMode('manual');
  try {
    var count = buildEmailPermutatorSheet(false); // false = show append prompt
    if (count === 0) return; // user cancelled the append
    SpreadsheetApp.getUi().alert(
      '✅ Step 2 Complete!\n\n' +
      '"Email Permutator" sheet has been built in this spreadsheet.\n' +
      'Total ' + count + ' founder row(s) added.'
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Step 2 Error:\n\n' + e.message);
    Logger.log(e.stack);
  }
}
