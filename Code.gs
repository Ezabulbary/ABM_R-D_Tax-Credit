// ============================================================
// Code.gs — onOpen menu + step runners (entry points)
//
// Spreadsheet layout:
//   "ABM R&D Tax Credit - Automation"  -> all code lives here
//                                         + Email Permutator tab
//   "ABM_R&D_Tax Credit"               -> Main_Crunchbase tab
//                                         + Final Format tab (output)
//   "Apollo Leads - Cleaner"           -> Apollo Leads tab
//                                         + Remaining tab
//
// AUTO-CHAIN (one click runs everything):
//   Step 1.1 -> (auto) Step 1.2 -> (auto) Step 2 -> (auto) Step 3
//            -> (auto) Step 4 -> (auto) push all leads to BigQuery
//
//   Each finished step schedules a 30-second background trigger
//   that launches the next step, so the user only has to start once.
// ============================================================

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 Lead Gen Tools')
    // One-click full pipeline (recommended)
    .addItem('▶️  RUN FULL WORKFLOW (Step 1 → 4 → BigQuery)', 'runStep1a')
    .addSeparator()
    // Individual steps (manual control / re-runs)
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

// ── Step 1.1 (FULL WORKFLOW entry) ────────────────────────────────
// Runs Step 1.1 on Main_Crunchbase, then immediately continues to
// Step 1.2 (Apollo) in the same foreground execution because the
// Apollo tab picker needs the UI. Step 1.2 then schedules the rest
// of the chain (2 -> 3 -> 4 -> BigQuery) in the background.
function runStep1a() {
  try {
    var count = processMainSheet();
    SpreadsheetApp.getUi().alert(
      '✅ Step 1.1 Complete!\n\n' +
      '• Domain column updated\n' +
      '• Founders split: ' + count + ' row(s) processed\n\n' +
      'Sheet: Main_Crunchbase (ABM_R&D_Tax Credit)\n\n' +
      '➡️ Continuing to Step 1.2 (Apollo)…'
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Step 1.1 Error:\n\n' + e.message);
    Logger.log(e.stack);
    return;
  }

  // Continue into Step 1.2 (foreground — needs the tab picker UI).
  // processApolloSheet() schedules Step 2 automatically when it finishes.
  runStep1b();
}

// ── Step 1.1 ONLY (no auto-continue) ──────────────────────────────
// Use this when you want to re-run Step 1.1 by itself.
function runStep1aOnly() {
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

// ── Step 1.2 — Apollo Leads split + domain extraction ─────────────
// processApolloSheet() schedules the background auto-chain (Step 2)
// at the end, so running this alone also completes the full workflow.
function runStep1b() {
  try {
    processApolloSheet();
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Step 1.2 Error:\n\n' + e.message);
    Logger.log(e.stack);
  }
}

// ── Step 2 — Build Email Permutator (manual run) ──────────────────
// Builds the permutator (with the append prompt) and then schedules
// Step 3 so the chain continues even from a manual Step 2 run.
function runStep2() {
  try {
    var count = buildEmailPermutatorSheet(false); // false = show append prompt
    if (count === 0) return; // user cancelled the append — stop the chain
    SpreadsheetApp.getUi().alert(
      '✅ Step 2 Complete!\n\n' +
      '"Email Permutator" sheet has been built in this spreadsheet.\n' +
      'Total ' + count + ' founder row(s) added.\n\n' +
      '➡️ Step 3 (Verify Emails) will auto-start in 30 seconds…'
    );
    _scheduleTrigger('autoStep3Start'); // continue the chain
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Step 2 Error:\n\n' + e.message);
    Logger.log(e.stack);
  }
}
