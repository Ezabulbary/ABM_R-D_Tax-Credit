// ============================================================
// step3_reoonVerifier.gs — Reoon API email verification
//
// ✅ Execution timeout solution: Time-based Trigger + Script Properties
//
// How it works:
//   1. "Step 3 — Start Verification" → resets progress, starts first batch
//   2. Each batch runs for at most 4.5 minutes, then saves progress and stops
//   3. After 1 minute, an auto-trigger fires and starts the next batch
//   4. When all rows are done, the trigger is deleted and a toast is shown
//   5. "Step 3 — Stop Verification" can be used to stop at any time
// ============================================================

// ── Script Property keys ──────────────────────────────────────
var PROP_ROW     = 'VERIFY_ROW';      // Next row index (1-based, data index)
var PROP_KEYIDX  = 'VERIFY_KEY_IDX'; // Current API key index
var PROP_RUNNING = 'VERIFY_RUNNING'; // 'true' / 'false'
var PROP_VTOTAL  = 'VERIFY_TOTAL';   // Total verified this session
var PROP_SSKIP   = 'VERIFY_SKIP';    // Total skipped this session
var TRIGGER_FN   = 'autoVerifyContinue'; // Trigger handler name

// Column indices in Email Permutator (0-based)
// Display order : firstname.lastname@ (E4) → firstname@ (G6) → firstinitiallast@ (I8)
// Verify order  : firstname@ (G6) first → firstname.lastname@ (E4) → firstinitiallast@ (I8)
var PCOL = {
  ORG        : 0,
  FIRST      : 1,
  LAST       : 2,
  DOMAIN     : 3,
  PAT_FN     : 4,   // firstname@           (col E — verified 2nd)
  STAT_FN    : 5,
  PAT_FNLN   : 6,   // firstname.lastname@  (col G — verified 1st)
  STAT_FNLN  : 7,
  PAT_FILN   : 8,   // firstinitiallast@    (col I — verified 3rd)
  STAT_FILN  : 9,
  VER_EMAIL  : 10,
  VER_STATUS : 11
};

// ════════════════════════════════════════════════════════════
//  PUBLIC — Menu actions
// ════════════════════════════════════════════════════════════

/** Step 3 START — resets progress and starts fresh from the first unverified row */
function startVerification() {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var permSheet = ss.getSheetByName(CONFIG.PERMUTATOR_SHEET_NAME);
  if (!permSheet) {
    SpreadsheetApp.getUi().alert('❌ "Email Permutator" sheet not found. Please run Step 2 first.');
    return;
  }
  var apiKeys = getApiKeys();
  if (apiKeys.length === 0) {
    SpreadsheetApp.getUi().alert('❌ No Reoon API keys found!\nPlease add keys via "⚙️ Manage Reoon API Keys".');
    return;
  }

  // Find the first unverified row (where Email or Verification Status is empty)
  var data = permSheet.getDataRange().getValues();
  var startFrom = 1;
  for (var r = 1; r < data.length; r++) {
    var email = data[r][10] ? String(data[r][10]).trim() : '';
    var status = data[r][11] ? String(data[r][11]).trim() : '';
    if (email && status) {
      // Already has email and status, skip it
    } else {
      startFrom = r;
      break;
    }
  }

  if (startFrom >= data.length) {
    SpreadsheetApp.getUi().alert('🎉 All rows in "Email Permutator" are already verified!');
    return;
  }

  // Reset progress
  var props = PropertiesService.getScriptProperties();
  props.setProperty(PROP_ROW,     String(startFrom));
  props.setProperty(PROP_KEYIDX,  '0');
  props.setProperty(PROP_RUNNING, 'true');
  props.setProperty(PROP_VTOTAL,  '0');
  props.setProperty(PROP_SSKIP,   '0');

  // Remove any leftover triggers
  _deleteTriggers();

  SpreadsheetApp.getUi().alert(
    '▶️ Step 3 is starting!\n\n' +
    'Verification will run in the background starting from row ' + (startFrom + 1) + '.\n' +
    'It will auto-restart every 5–6 minutes.\n\n' +
    'To check progress:\nUse "📊 Check Verification Progress".\n\n' +
    'To stop:\nUse "⛔ Stop Verification".'
  );

  // Run first batch immediately
  _runBatch();
}

/** Auto-trigger handler — called by Google Trigger */
function autoVerifyContinue() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty(PROP_RUNNING) !== 'true') {
    _deleteTriggers();
    return;
  }
  _runBatch();
}

/** Step 3 STOP — deletes the trigger */
function stopVerification() {
  var props = PropertiesService.getScriptProperties();
  props.setProperty(PROP_RUNNING, 'false');
  _deleteTriggers();

  var verified = parseInt(props.getProperty(PROP_VTOTAL) || '0');
  var skipped  = parseInt(props.getProperty(PROP_SSKIP)  || '0');
  SpreadsheetApp.getUi().alert(
    '⛔ Verification has been stopped.\n\n' +
    '✅ Verified this session: ' + verified + '\n' +
    '⏭️ Already done (skipped): ' + skipped + '\n\n' +
    'To continue later, run Step 3 → Resume.'
  );
}

/** Step 3 RESUME — continues from where it left off */
function resumeVerification() {
  var props = PropertiesService.getScriptProperties();
  var row   = parseInt(props.getProperty(PROP_ROW) || '1');

  if (row <= 1 && props.getProperty(PROP_RUNNING) !== 'true') {
    SpreadsheetApp.getUi().alert('ℹ️ No saved progress found.\nUse "Start" to begin from the beginning.');
    return;
  }

  props.setProperty(PROP_RUNNING, 'true');
  _deleteTriggers();

  SpreadsheetApp.getUi().alert(
    '▶️ Resuming verification!\n\nWill start from row ' + (row + 1) + '.'
  );
  _runBatch();
}

// ════════════════════════════════════════════════════════════
//  CORE — Single batch processor (max 4.5 min per call)
// ════════════════════════════════════════════════════════════

function _runBatch() {
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
  var MAX_MS    = 2 * 60 * 1000; // 2 min — extra safety against execution timeout

  var i = startRow; // 1-based data index (0 = header)
  var dailyCreditsLeft = -1; // -1 means not yet fetched for current keyIdx

  // Helper to ensure the active API key has daily credits left.
  // Returns true if a key with credits is ready, false if all daily credits are exhausted.
  function ensureActiveKey() {
    while (keyIdx < apiKeys.length) {
      if (dailyCreditsLeft === -1) {
        var credits = _getDailyCredits(apiKeys[keyIdx]);
        if (credits > 0) {
          dailyCreditsLeft = credits;
          return true;
        } else {
          // No daily credits left on this key, rotate to next key
          keyIdx++;
        }
      } else if (dailyCreditsLeft > 0) {
        return true;
      } else {
        // dailyCreditsLeft is 0, rotate to next key
        keyIdx++;
        dailyCreditsLeft = -1;
      }
    }
    return false;
  }

  // Helper to verify a single pattern if it doesn't already have a status.
  // Returns { status: string, stopped: boolean }
  function verifyPattern(emailPattern, colIndex, currentStatus) {
    if (!emailPattern) {
      return { status: '', stopped: false };
    }
    if (currentStatus) {
      return { status: currentStatus, stopped: false };
    }

    if (!ensureActiveKey()) {
      props.setProperty(PROP_ROW,    String(i));
      props.setProperty(PROP_KEYIDX, String(keyIdx));
      props.setProperty(PROP_VTOTAL, String(sesVerify));
      props.setProperty(PROP_SSKIP,  String(sesSkip));
      SpreadsheetApp.flush();
      _scheduleNextDelayed(60); // Check again in 60 minutes
      ss.toast('Waiting for Daily Credits to reset.', '⏳ All Daily Credits exhausted! Auto-resuming in 1 hour...', 30);
      return { status: '', stopped: true };
    }

    var r = _callReoon(emailPattern, apiKeys[keyIdx]);

    if (dailyCreditsLeft > 0) {
      dailyCreditsLeft--;
    }

    if (r.keyExhausted) {
      // Rotate key and try again recursively
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

    // ── Time check — every single row ─────────────────────
    if (Date.now() - startTime > MAX_MS) {
      // Save progress and schedule next run
      props.setProperty(PROP_ROW,    String(i));
      props.setProperty(PROP_KEYIDX, String(keyIdx));
      props.setProperty(PROP_VTOTAL, String(sesVerify));
      props.setProperty(PROP_SSKIP,  String(sesSkip));
      SpreadsheetApp.flush();
      _scheduleNext();

      ss.toast(
        'Verified: ' + sesVerify + '  |  Remaining: ' + (totalRows - i + 1),
        '🔄 Auto-resuming in 1 min… Row ' + (i + 1) + '/' + (totalRows + 1),
        15
      );
      return;
    }

    // ── API keys exhausted check ───────────────────────────
    if (keyIdx >= apiKeys.length) {
      props.setProperty(PROP_ROW,    String(i));
      SpreadsheetApp.flush();
      _scheduleNextDelayed(60); // Check again in 60 minutes
      ss.toast(
        '✅ Verified: ' + sesVerify + '  |  Waiting for Daily Credits to reset.',
        '⏳ All Daily Credits exhausted! Auto-resuming in 1 hour...',
        30
      );
      return;
    }

    var row       = data[i];
    var firstName = row[PCOL.FIRST]  ? String(row[PCOL.FIRST]).trim()  : '';
    var domain    = row[PCOL.DOMAIN] ? String(row[PCOL.DOMAIN]).trim() : '';
    if (!firstName || !domain) continue;

    // Already fully verified → skip
    var verEmail  = row[PCOL.VER_EMAIL]  ? String(row[PCOL.VER_EMAIL]).trim()  : '';
    var verStatus = row[PCOL.VER_STATUS] ? String(row[PCOL.VER_STATUS]).trim() : '';
    if (verEmail && verStatus) { sesSkip++; continue; }

    var foundGood = false;

    // ── 1st: Verify firstname@ (col E = PAT_FN) ──────────────
    var patFN  = row[PCOL.PAT_FN]  ? String(row[PCOL.PAT_FN]).trim()  : '';
    var statFN = row[PCOL.STAT_FN] ? String(row[PCOL.STAT_FN]).trim() : '';

    var resFN = verifyPattern(patFN, PCOL.STAT_FN, statFN);
    if (resFN.stopped) return;
    statFN = resFN.status;

    if (statFN && CONFIG.VALID_STATUSES.indexOf(statFN) !== -1) {
      permSheet.getRange(i + 1, PCOL.VER_EMAIL  + 1).setValue(patFN);
      permSheet.getRange(i + 1, PCOL.VER_STATUS + 1).setValue(statFN);
      foundGood = true;
    }

    // ── 2nd: Verify firstname.lastname@ (col G = PAT_FNLN) ───
    if (!foundGood) {
      var patFNLN  = row[PCOL.PAT_FNLN]  ? String(row[PCOL.PAT_FNLN]).trim()  : '';
      var statFNLN = row[PCOL.STAT_FNLN] ? String(row[PCOL.STAT_FNLN]).trim() : '';

      var resFNLN = verifyPattern(patFNLN, PCOL.STAT_FNLN, statFNLN);
      if (resFNLN.stopped) return;
      statFNLN = resFNLN.status;

      if (statFNLN && CONFIG.VALID_STATUSES.indexOf(statFNLN) !== -1) {
        permSheet.getRange(i + 1, PCOL.VER_EMAIL  + 1).setValue(patFNLN);
        permSheet.getRange(i + 1, PCOL.VER_STATUS + 1).setValue(statFNLN);
        foundGood = true;
      }
    }

    // ── 3rd: Verify firstinitiallast@ (col I = PAT_FILN) ────
    if (!foundGood) {
      var patFILN  = row[PCOL.PAT_FILN]  ? String(row[PCOL.PAT_FILN]).trim()  : '';
      var statFILN = row[PCOL.STAT_FILN] ? String(row[PCOL.STAT_FILN]).trim() : '';

      var resFILN = verifyPattern(patFILN, PCOL.STAT_FILN, statFILN);
      if (resFILN.stopped) return;
      statFILN = resFILN.status;

      if (statFILN && CONFIG.VALID_STATUSES.indexOf(statFILN) !== -1) {
        permSheet.getRange(i + 1, PCOL.VER_EMAIL  + 1).setValue(patFILN);
        permSheet.getRange(i + 1, PCOL.VER_STATUS + 1).setValue(statFILN);
        foundGood = true;
      }
    }

    // ── Fallback: no safe/catch_all found — still record 1st pattern (firstname@) ──
    if (!foundGood) {
      if (patFN) {
        permSheet.getRange(i + 1, PCOL.VER_EMAIL  + 1).setValue(patFN);
        permSheet.getRange(i + 1, PCOL.VER_STATUS + 1).setValue(statFN || 'unknown');
      }
    }

    // Flush every 10 rows to save incremental progress
    if ((i - startRow) % 10 === 0) SpreadsheetApp.flush();
  }

  // ── All rows processed ────────────────────────────────────
  props.setProperty(PROP_RUNNING, 'false');
  props.setProperty(PROP_ROW,    String(data.length));
  props.setProperty(PROP_VTOTAL, String(sesVerify));
  props.setProperty(PROP_SSKIP,  String(sesSkip));
  SpreadsheetApp.flush();
  _deleteTriggers();

  ss.toast(
    '✅ Verified: ' + sesVerify + '  |  Skipped: ' + sesSkip + '\n' +
    'Now run Step 4.',
    '🎉 Verification complete!',
    20
  );
}

// ════════════════════════════════════════════════════════════
//  Reoon API caller
// ════════════════════════════════════════════════════════════

function _callReoon(email, apiKey) {
  var url = 'https://emailverifier.reoon.com/api/v1/verify' +
            '?email=' + encodeURIComponent(email) +
            '&key='   + encodeURIComponent(apiKey) +
            '&mode=power';
  var maxRetries = 2;
  for (var attempt = 0; attempt < maxRetries; attempt++) {
    try {
      var resp = UrlFetchApp.fetch(url, {
        method            : 'GET',
        muteHttpExceptions: true,
        headers           : { 'Accept': 'application/json' }
      });
      var code = resp.getResponseCode();
      var body = resp.getContentText();
      Logger.log('[Reoon] %s → HTTP %s (attempt %s)', email, code, attempt + 1);

      if (code === 200) {
        var json = JSON.parse(body);
        return { status: json.status || 'unknown', keyExhausted: false };
      }
      if (code === 401 || code === 403) {
        return { status: 'key_error', keyExhausted: true };
      }
      if (code === 429) {
        // Rate-limited: wait 3 sec then retry (iterative, not recursive)
        Utilities.sleep(3000);
        continue;
      }
      // Try to detect credit exhaustion in body
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
  // All retries exhausted (only happens on repeated 429)
  return { status: 'error', keyExhausted: false };
}

/**
 * Checks remaining daily credits for a Reoon API key.
 * Returns the number of daily credits, or 0 if invalid/exhausted.
 */
function _getDailyCredits(apiKey) {
  var url = 'https://emailverifier.reoon.com/api/v1/check-account-balance/' +
            '?key=' + encodeURIComponent(apiKey);
  try {
    var resp = UrlFetchApp.fetch(url, {
      method            : 'GET',
      muteHttpExceptions: true,
      headers           : { 'Accept': 'application/json' }
    });
    var code = resp.getResponseCode();
    if (code === 200) {
      var json = JSON.parse(resp.getContentText());
      if (json.status === 'success' && json.api_status === 'active') {
        return parseInt(json.remaining_daily_credits || '0');
      }
    }
  } catch (e) {
    Logger.log('[Reoon Balance Check] Exception: ' + e.message);
  }
  return 0;
}

// ════════════════════════════════════════════════════════════
//  Trigger helpers
// ════════════════════════════════════════════════════════════

function _scheduleNext() {
  _deleteTriggers();
  ScriptApp.newTrigger(TRIGGER_FN)
    .timeBased()
    .after(60 * 1000) // 1 minute from now
    .create();
}

function _scheduleNextDelayed(minutes) {
  _deleteTriggers();
  ScriptApp.newTrigger(TRIGGER_FN)
    .timeBased()
    .after(minutes * 60 * 1000)
    .create();
}

function _deleteTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === TRIGGER_FN) {
      ScriptApp.deleteTrigger(t);
    }
  });
}

// ════════════════════════════════════════════════════════════
//  Progress checker (menu action)
// ════════════════════════════════════════════════════════════

function checkProgress() {
  var props     = PropertiesService.getScriptProperties();
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var permSheet = ss.getSheetByName(CONFIG.PERMUTATOR_SHEET_NAME);
  if (!permSheet) {
    SpreadsheetApp.getUi().alert('"Email Permutator" sheet not found.');
    return;
  }

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
    '📋 Total founders : ' + total  + '\n' +
    '✅ Verified done  : ' + done   + '  (' + pct + '%)\n' +
    '⏳ Remaining      : ' + (total - done) + '\n\n' +
    '🟢 Safe           : ' + safe    + '\n' +
    '🟡 Catch-all      : ' + catchAll + '\n' +
    '🔴 Invalid        : ' + inv
  );
}
