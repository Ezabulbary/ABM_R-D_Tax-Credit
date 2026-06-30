// ============================================================
// Code.gs — onOpen menu + step runners
//
// Spreadsheet layout:
//   "ABM R&D Tax Credit - Automation"  → all code lives here
//                                         + Email Permutator tab
//   "ABM_R&D_Tax Credit"               → Main_Crunchbase tab
//                                         + Final Format tab (output)
//   "Apollo Leads - Cleaner"           → Apollo Leads tab
//                                         + Remaining tab
// ============================================================

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 Lead Gen Tools')
    .addItem('Step 1.1 — Crunchbase: Domain & Founders Split',  'runStep1a')
    .addItem('Step 1.2 — Apollo: Process, Split & Domain',      'runStep1b')
    .addSeparator()
    .addItem('Step 2 — Build Email Permutator',                  'runStep2')
    .addSeparator()
    .addSubMenu(ui.createMenu('Step 3 — Verify Emails (Reoon)')
      .addItem('▶️  Start Verification',    'startVerification')
      .addItem('⏩  Resume Verification',   'resumeVerification')
      .addItem('⛔  Stop Verification',     'stopVerification')
      .addItem('📊  Check Progress',        'checkProgress')
    )
    .addSeparator()
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

// ── Step 1.1 — Main_Crunchbase domain extraction + founders split ──

function runStep1a() {
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

function runStep1b() {
  try {
    processApolloSheet();
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Step 1.2 Error:\n\n' + e.message);
    Logger.log(e.stack);
  }
}

// ── Step 2 — Build Email Permutator ───────────────────────────────

function runStep2() {
  try {
    var count = buildEmailPermutatorSheet();
    if (count === 0) return;
    SpreadsheetApp.getUi().alert(
      '✅ Step 2 Complete!\n\n' +
      '"Email Permutator" sheet has been built in this spreadsheet.\n' +
      'Total ' + count + ' founder row(s) added.\n\n' +
      'Now run Step 3 → Start Verification.'
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Step 2 Error:\n\n' + e.message);
    Logger.log(e.stack);
  }
}
