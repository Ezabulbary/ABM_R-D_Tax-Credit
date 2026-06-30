// ============================================================
// step3_reoonVerifier.gs — Reoon API email verification
//
// Reads/Writes: "ABM R&D Tax Credit - Automation" → Email Permutator tab
//
// Auto-trigger flow:
//   1. Start → resets progress, runs first batch
//   2. Each batch runs ≤ 4.5 min, then saves progress
//   3. After 30 sec an auto-trigger fires the next batch
//   4. When all rows done → auto-starts Step 4 (Final Format)
//   5. Stop can interrupt at any time
// ============================================================

var PROP_ROW     = 'VERIFY_ROW';
var PROP_KEYIDX  = 'VERIFY_KEY_IDX';
var PROP_RUNNING = 'VERIFY_RUNNING';
var PROP_VTOTAL  = 'VERIFY_TOTAL';
var PROP_SSKIP   = 'VERIFY_SKIP';
var TRIGGER_FN_VERIFY = 'autoVerifyContinue';

// Column indices in Email Permutator (0-based)
var PCOL = {
  ORG       : 0,
  FIRST     : 1,
  LAST      : 2,
  DOMAIN    : 3,
  PAT_FN    : 4,   // firstname@
  STAT_FN   : 5,
  PAT_FNLN  : 6,   // firstname.lastname@
  STAT_FNLN : 7,
  PAT_FILN  : 8,   // firstinitiallast@
  STAT_FILN : 9,
  VER_EMAIL : 10,
  VER_STATUS: 11
};

// ════════════════════════════════════════════════════════════
//  PUBLIC — Menu actions
// ════════════════════════════════════════════════════════════

function startVerification() {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var permSheet = ss.getSheetByName(CONFIG.PERMUTATOR_SHEET_NAME);
  if (!permSheet) {
    SpreadsheetApp.getUi().alert('❌ "Email Permutator" not found. Run Step 2 first.');
    return;
  }

  var apiKeys = getApiKeys();
  if (apiKeys.length === 0) {
    SpreadsheetApp.getUi().alert('❌ No Reoon API keys found!\nUse "⚙️ Manage Reoon API Keys".');
    return;
  }

  // Find first unverified row
  var data      = permSheet.getDataRange().getValues();
  var startFrom = 1;
  for (var r = 1; r < data.length; r++) {
    var email  = data[r][10] ? String(data[r][10]).trim() : '';
    var status = data[r][11] ? String(data[r][11]).trim() : '';
    if (!email || !status) { startFrom = r; break; }
    startFrom = data.length; // all done
  }

  if (startFrom >= data.length) {
    SpreadsheetApp.getUi().alert('🎉 All rows already verified!');
    return;
  }

  var props = PropertiesService.getScriptProperties();
  props.setProperty(PROP_ROW,     String(startFrom));
  props.setProperty(PROP_KEYIDX,  '0');
  props.setProperty(PROP_RUNNING, 'true');
  props.setProperty(PROP_VTOTAL,  '0');
  props.setProperty(PROP_SSKIP,   '0');

  _deleteTriggersForFunction(TRIGGER_FN_VERIFY);

  SpreadsheetApp.getUi().alert(
    '▶️ Step 3 Starting!\n\n' +
    'Verification runs in background, auto-resumes every 30 sec.\n' +
    'When done, Step 4 (Final Format) starts automatically.\n\n' +
    'To check progress: "📊 Check Progress"\n' +
    'To stop: "⛔ Stop Verification"'
  );

  _runVerifyBatch();
}

function autoVerifyContinue() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty(PROP_RUNNING) !== 'true') {
    _deleteTriggersForFunction(TRIGGER_FN_VERIFY);
    return;
  }
  _runVerifyBatch();
}

function stopVerification() {
  var props    = PropertiesService.getScriptProperties();
  props.setProperty(PROP_RUNNING, 'false');
  _deleteTriggersForFunction(TRIGGER_FN_VERIFY);

  var verified = parseInt(props.getProperty(PROP_VTOTAL) || '0');
  var skipped  = parseInt(props.getProperty(PROP_SSKIP)  || '0');
  SpreadsheetApp.getUi().alert(
    '⛔ Verification stopped.\n\n' +
    '✅ Verified this session: ' + verified + '\n' +
    '⏭️ Already done (skipped): ' + skipped + '\n\n' +
    'Use "⏩ Resume" to continue later.'
  );
}

function resumeVerification() {
  var props = PropertiesService.getScriptProperties();
  var row   = parseInt(props.getProperty(PROP_ROW) || '1');

  if (row <= 1 && props.getProperty(PROP_RUNNING) !== 'true') {
    SpreadsheetApp.getUi().alert('ℹ️ No saved progress. Use "▶️ Start" to begin.');
    return;
  }

  props.setProperty(PROP_RUNNING, 'true');
  _deleteTriggersForFunction(TRIGGER_FN_VERIFY);

  SpreadsheetApp.getUi().alert('▶️ Resuming from row ' + (row + 1) + '…');
  _runVerifyBatch();
}

// ════════════════════════════════════════════════════════════
//  CORE — Single batch processor
// ════════════════════════════════════════════════════════════

function _runVerifyBatch() {
  var props     = PropertiesService.getScriptProperties();
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var permSheet = ss.getSheetByName(CONFIG.PERMUTATOR_SHEET_NAME);
  if (!permSheet) { props.setProperty(PROP_RUNNING, 'false'); return; }

  var apiKeys   = getApiKeys();
  var startRow  = parseInt(props.getProperty(PROP_ROW)    || '1');
  var keyIdx    = parseInt(props.getProperty(PROP_KEYIDX) || '0');
  var sesVerify = parseInt(props.getProperty(PROP_VTOTAL) || '0');
  var sesSkip   = parseInt(props.getProperty(PROP_SSKIP)  || '0');

  var data      = permSheet.getDataRange().getValues();
  var totalRows = data.length - 1;
  var startTime = Date.now();

  var i                = startRow;
  var dailyCreditsLeft = -1;

  function ensureActiveKey() {
    while (keyIdx < apiKeys.length) {
      if (dailyCreditsLeft === -1) {
        var credits = _getDailyCredits(apiKeys[keyIdx]);
        if (credits > 0) { dailyCreditsLeft = credits; return true; }
        keyIdx++;
      } else if (dailyCreditsLeft > 0) {
        return true;
      } else {
        keyIdx++;
        dailyCreditsLeft = -1;
      }
    }
    return false;
  }

  function verifyPattern(emailPattern, colIndex, currentStatus) {
    if (!emailPattern) return { status: '', stopped: false };
    if (currentStatus)  return { status: currentStatus, stopped: false };

    if (!ensureActiveKey()) {
      props.setProperty(PROP_ROW,    String(i));
      props.setProperty(PROP_KEYIDX, String(keyIdx));
      props.setProperty(PROP_VTOTAL, String(sesVerify));
      props.setProperty(PROP_SSKIP,  String(sesSkip));
      SpreadsheetApp.flush();
      _scheduleTriggerDelayed(TRIGGER_FN_VERIFY, 60 * 60 * 1000); // retry in 1 hour
      ss.toast('All daily credits exhausted. Auto-resuming in 1 hour…', '⏳', 30);
      return { status: '', stopped: true };
    }

    var r = _callReoon(emailPattern, apiKeys[keyIdx]);
    if (dailyCreditsLeft > 0) dailyCreditsLeft--;

    if (r.keyExhausted) {
      dailyCreditsLeft = 0;
      return verifyPattern(emailPattern, colIndex, currentStatus);
    }

    permSheet.getRange(i + 1, colIndex + 1).setValue(r.status);
    data[i][colIndex] = r.status;
    sesVerify++;
    Utilities.sleep(CONFIG.SLEEP_MS);
    return { status: r.status, stopped: false };
  }

  for (; i < data.length; i++) {

    // Time check every row
    if (Date.now() - startTime > CONFIG.MAX_RUN_MS) {
      props.setProperty(PROP_ROW,    String(i));
      props.setProperty(PROP_KEYIDX, String(keyIdx));
      props.setProperty(PROP_VTOTAL, String(sesVerify));
      props.setProperty(PROP_SSKIP,  String(sesSkip));
      SpreadsheetApp.flush();
      _scheduleTrigger(TRIGGER_FN_VERIFY);
      ss.toast(
        'Verified: ' + sesVerify + '  |  Remaining: ' + (totalRows - i + 1),
        '🔄 Auto-resuming in 30 sec… Row ' + (i + 1) + '/' + (totalRows + 1),
        15
      );
      return;
    }

    if (keyIdx >= apiKeys.length) {
      props.setProperty(PROP_ROW, String(i));
      SpreadsheetApp.flush();
      _scheduleTriggerDelayed(TRIGGER_FN_VERIFY, 60 * 60 * 1000);
      ss.toast('All daily credits exhausted. Auto-resuming in 1 hour…', '⏳', 30);
      return;
    }

    var row       = data[i];
    var firstName = row[PCOL.FIRST]  ? String(row[PCOL.FIRST]).trim()  : '';
    var domain    = row[PCOL.DOMAIN] ? String(row[PCOL.DOMAIN]).trim() : '';
    if (!firstName || !domain) continue;

    // Already fully verified
    var verEmail  = row[PCOL.VER_EMAIL]  ? String(row[PCOL.VER_EMAIL]).trim()  : '';
    var verStatus = row[PCOL.VER_STATUS] ? String(row[PCOL.VER_STATUS]).trim() : '';
    if (verEmail && verStatus) { sesSkip++; continue; }

    var foundGood = false;

    // 1st: firstname@
    var patFN  = row[PCOL.PAT_FN]  ? String(row[PCOL.PAT_FN]).trim()  : '';
    var statFN = row[PCOL.STAT_FN] ? String(row[PCOL.STAT_FN]).trim() : '';
    var resFN  = verifyPattern(patFN, PCOL.STAT_FN, statFN);
    if (resFN.stopped) return;
    statFN = resFN.status;
    if (statFN && CONFIG.VALID_STATUSES.indexOf(statFN) !== -1) {
      permSheet.getRange(i + 1, PCOL.VER_EMAIL  + 1).setValue(patFN);
      permSheet.getRange(i + 1, PCOL.VER_STATUS + 1).setValue(statFN);
      foundGood = true;
    }

    // 2nd: firstname.lastname@
    if (!foundGood) {
      var patFNLN  = row[PCOL.PAT_FNLN]  ? String(row[PCOL.PAT_FNLN]).trim()  : '';
      var statFNLN = row[PCOL.STAT_FNLN] ? String(row[PCOL.STAT_FNLN]).trim() : '';
      var resFNLN  = verifyPattern(patFNLN, PCOL.STAT_FNLN, statFNLN);
      if (resFNLN.stopped) return;
      statFNLN = resFNLN.status;
      if (statFNLN && CONFIG.VALID_STATUSES.indexOf(statFNLN) !== -1) {
        permSheet.getRange(i + 1, PCOL.VER_EMAIL  + 1).setValue(patFNLN);
        permSheet.getRange(i + 1, PCOL.VER_STATUS + 1).setValue(statFNLN);
        foundGood = true;
      }
    }

    // 3rd: firstinitiallast@
    if (!foundGood) {
      var patFILN  = row[PCOL.PAT_FILN]  ? String(row[PCOL.PAT_FILN]).trim()  : '';
      var statFILN = row[PCOL.STAT_FILN] ? String(row[PCOL.STAT_FILN]).trim() : '';
      var resFILN  = verifyPattern(patFILN, PCOL.STAT_FILN, statFILN);
      if (resFILN.stopped) return;
      statFILN = resFILN.status;
      if (statFILN && CONFIG.VALID_STATUSES.indexOf(statFILN) !== -1) {
        permSheet.getRange(i + 1, PCOL.VER_EMAIL  + 1).setValue(patFILN);
        permSheet.getRange(i + 1, PCOL.VER_STATUS + 1).setValue(statFILN);
        foundGood = true;
      }
    }

    // Fallback: record firstname@ even if unverified
    if (!foundGood && patFN) {
      permSheet.getRange(i + 1, PCOL.VER_EMAIL  + 1).setValue(patFN);
      permSheet.getRange(i + 1, PCOL.VER_STATUS + 1).setValue(statFN || 'unknown');
    }

    if ((i - startRow) % 10 === 0) SpreadsheetApp.flush();
  }

  // ── All rows done ─────────────────────────────────────────
  props.setProperty(PROP_RUNNING, 'false');
  props.setProperty(PROP_ROW,    String(data.length));
  props.setProperty(PROP_VTOTAL, String(sesVerify));
  props.setProperty(PROP_SSKIP,  String(sesSkip));
  SpreadsheetApp.flush();
  _deleteTriggersForFunction(TRIGGER_FN_VERIFY);

  ss.toast(
    '✅ Verified: ' + sesVerify + '  |  Skipped: ' + sesSkip + '\nStarting Step 4…',
    '🎉 Verification Complete! Auto-starting Final Format…',
    20
  );

  // Auto-chain to Step 4
  _scheduleTrigger('autoFinalFormatStart');
}

// Auto-trigger entry point that starts Step 4 after Step 3 completes
function autoFinalFormatStart() {
  _deleteTriggersForFunction('autoFinalFormatStart');
  startFinalFormat();
}

// ════════════════════════════════════════════════════════════
//  Reoon API
// ════════════════════════════════════════════════════════════

function _callReoon(email, apiKey) {
  var url = 'https://emailverifier.reoon.com/api/v1/verify' +
            '?email=' + encodeURIComponent(email) +
            '&key='   + encodeURIComponent(apiKey) +
            '&mode=power';
  for (var attempt = 0; attempt < 2; attempt++) {
    try {
      var resp = UrlFetchApp.fetch(url, {
        method: 'GET', muteHttpExceptions: true,
        headers: { 'Accept': 'application/json' }
      });
      var code = resp.getResponseCode();
      var body = resp.getContentText();
      Logger.log('[Reoon] %s → HTTP %s (attempt %s)', email, code, attempt + 1);

      if (code === 200) {
        var json = JSON.parse(body);
        return { status: json.status || 'unknown', keyExhausted: false };
      }
      if (code === 401 || code === 403) return { status: 'key_error', keyExhausted: true };
      if (code === 429) { Utilities.sleep(3000); continue; }

      try {
        var err = JSON.parse(body);
        var msg = (err.message || '').toLowerCase();
        if (msg.indexOf('credit') !== -1 || msg.indexOf('quota') !== -1 ||
            msg.indexOf('limit')  !== -1 || msg.indexOf('exceed') !== -1) {
          return { status: 'key_exhausted', keyExhausted: true };
        }
      } catch (ignore) {}
      return { status: 'error', keyExhausted: false };
    } catch (e) {
      Logger.log('[Reoon] Exception: ' + e.message);
      return { status: 'error', keyExhausted: false };
    }
  }
  return { status: 'error', keyExhausted: false };
}

function _getDailyCredits(apiKey) {
  var url = 'https://emailverifier.reoon.com/api/v1/check-account-balance/' +
            '?key=' + encodeURIComponent(apiKey);
  try {
    var resp = UrlFetchApp.fetch(url, {
      method: 'GET', muteHttpExceptions: true,
      headers: { 'Accept': 'application/json' }
    });
    if (resp.getResponseCode() === 200) {
      var json = JSON.parse(resp.getContentText());
      if (json.status === 'success' && json.api_status === 'active') {
        return parseInt(json.remaining_daily_credits || '0');
      }
    }
  } catch (e) {
    Logger.log('[Reoon Balance] ' + e.message);
  }
  return 0;
}

// ════════════════════════════════════════════════════════════
//  Progress checker
// ════════════════════════════════════════════════════════════

function checkProgress() {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var permSheet = ss.getSheetByName(CONFIG.PERMUTATOR_SHEET_NAME);
  if (!permSheet) {
    SpreadsheetApp.getUi().alert('"Email Permutator" not found.');
    return;
  }

  var props    = PropertiesService.getScriptProperties();
  var data     = permSheet.getDataRange().getValues();
  var total    = data.length - 1;
  var done = 0, safe = 0, catchAll = 0, inv = 0;

  for (var i = 1; i < data.length; i++) {
    var vs = data[i][PCOL.VER_STATUS] ? String(data[i][PCOL.VER_STATUS]).trim() : '';
    if (vs) done++;
    if (vs === 'safe')      safe++;
    if (vs === 'catch_all') catchAll++;
    if (vs === 'invalid')   inv++;
  }

  var pct      = total > 0 ? Math.round((done / total) * 100) : 0;
  var running  = props.getProperty(PROP_RUNNING) === 'true';
  var savedRow = parseInt(props.getProperty(PROP_ROW) || '1');

  SpreadsheetApp.getUi().alert(
    '📊 Verification Progress\n\n' +
    '🔄 Status         : ' + (running ? 'Running…' : 'Stopped') + '\n' +
    '📍 Next row       : ' + (savedRow + 1) + ' / ' + (total + 1) + '\n\n' +
    '📋 Total founders : ' + total     + '\n' +
    '✅ Verified done  : ' + done      + '  (' + pct + '%)\n' +
    '⏳ Remaining      : ' + (total - done) + '\n\n' +
    '🟢 Safe           : ' + safe      + '\n' +
    '🟡 Catch-all      : ' + catchAll  + '\n' +
    '🔴 Invalid        : ' + inv
  );
}
