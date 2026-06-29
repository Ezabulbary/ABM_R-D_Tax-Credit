// ============================================================
// Code.gs — onOpen menu + step runners
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 Lead Gen Tools')
    .addItem('📖 Guidelines — Full Guide',            'showGuidelines')
    .addSeparator()
    .addItem('Step 1 — Domain Extract & Founders Split',  'runStep1')
    .addItem('Step 2 — Build Email Permutator Sheet',     'runStep2')
    .addSeparator()
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('Step 3 — Verify Emails (Reoon)')
        .addItem('▶️  Start Verification (fresh)',        'startVerification')
        .addItem('⏩  Resume Verification (continue)',    'resumeVerification')
        .addItem('⛔  Stop Verification',                 'stopVerification')
        .addItem('📊  Check Verification Progress',       'checkProgress')
    )
    .addSeparator()
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('Step 4 — Build Final Format Sheet')
        .addItem('▶️  Start Building (fresh)',            'startFinalFormat')
        .addItem('⏩  Resume Building (continue)',        'resumeFinalFormat')
        .addItem('⛔  Stop Building',                     'stopFinalFormat')
    )
    .addSeparator()
    .addItem('⚙️  Manage Reoon API Keys',                'manageApiKeys')
    .addItem('🧹  Clean Up Empty Cells (Fix 10M Limit)', 'cleanUpWorkbook')
    .addToUi();
}

/** Guidelines dialog — detailed description of all steps */
function showGuidelines() {
  var html = HtmlService.createHtmlOutputFromFile('guidelines')
    .setTitle('📖 Lead Gen Guidelines')
    .setWidth(680)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '📖 Lead Gen — Full Guidelines');
}

// ── Step Runners ─────────────────────────────────────────────

function runStep1() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var count = processMainSheet();
    var apolloNote = ss.getSheetByName(CONFIG.APOLLO_SHEET_NAME)
      ? '\n• Apollo Leads tab — Domain column updated'
      : '';
    SpreadsheetApp.getUi().alert(
      '✅ Step 1 Complete!\n\n' +
      '• Domain column has been added\n' +
      '• Founders split applied to ' + count + ' row(s)' +
      apolloNote
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Step 1 Error:\n\n' + e.message);
    Logger.log(e.stack);
  }
}

function runStep2() {
  try {
    var count = buildEmailPermutatorSheet();
    if (count === 0) return; // user cancelled
    SpreadsheetApp.getUi().alert(
      '✅ Step 2 Complete!\n\n' +
      '"Email Permutator" sheet has been created.\n' +
      'Total ' + count + ' Founder row(s) added.\n\n' +
      'Now run Step 3 → Start Verification.'
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Step 2 Error:\n\n' + e.message);
    Logger.log(e.stack);
  }
}
