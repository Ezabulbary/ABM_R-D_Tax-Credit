// ============================================================
// step4_finalFormat.gs — Build "Final Format" sheet
//
// Reads from:
//   - "ABM R&D Tax Credit - Automation" (this SS) → Email Permutator tab
//   - "ABM_R&D_Tax Credit"              → Main_Crunchbase tab
//   - "Apollo Leads - Cleaner"          → Apollo Leads tab
//
// Writes to:
//   - "ABM_R&D_Tax Credit"              → Final Format tab (output)
//   - Also updates Main_Crunchbase summary columns:
//       "How Many Leads" and "Verification Status" per domain
//
// Auto-trigger: runs in batches of 500 rows, resumes every 30 sec.
// ============================================================

var FINAL_COLUMNS = [
  'Organization Name',
  'Domain',
  'First Name',
  'Last Name',
  'Email',
  'Verification Status',
  'Title',
  'Seniority',
  'Departments',
  'Person Linkedin Url',
  'City',
  'State',
  'Country',
  'headline',
  'Founded Date',
  'Number of Employees',
  'Full Description',
  'Industry',
  'Patents Granted',
  'Total Funding Amount',
  'IT Spend (Aberdeen)',
  'Most Recent Valuation Range',
  'Estimated Revenue Range'
];

// Mapping: Final Format column name → Apollo Leads column name (where they differ)
var FINAL_TO_APOLLO = (function() {
  var m = Object.create(null);
  m['Organization Name']       = 'Company Name';
  m['Founded Date']            = 'Company Founded Year';
  m['Number of Employees']     = '# Employees';
  m['Total Funding Amount']    = 'Total Funding';
  m['Estimated Revenue Range'] = 'Annual Revenue';
  return m;
}());

// These columns come from Email Permutator — never overwrite with Apollo
var PERM_COL_NAMES = {
  'Domain': true, 'First Name': true, 'Last Name': true,
  'Email': true, 'Verification Status': true, 'Organization Name': true
};

var PROP_F_PERM_IDX    = 'FINAL_PERM_IDX';
var PROP_F_APOLLO_IDX  = 'FINAL_APOLLO_IDX';
var PROP_F_RUNNING     = 'FINAL_RUNNING';
var PROP_F_ADDED_COUNT = 'FINAL_ADDED_COUNT';
var TRIGGER_FN_FINAL   = 'autoFinalFormatContinue';

// ════════════════════════════════════════════════════════════
//  PUBLIC — Menu actions
// ════════════════════════════════════════════════════════════

function startFinalFormat() {
  var automationSS = SpreadsheetApp.getActiveSpreadsheet();
  var permSheet    = automationSS.getSheetByName(CONFIG.PERMUTATOR_SHEET_NAME);
  var crunchbaseSS = SpreadsheetApp.openById(CONFIG.CRUNCHBASE_SS_ID);
  var mainSheet    = crunchbaseSS.getSheetByName(CONFIG.MAIN_SHEET_NAME);

  if (!mainSheet) {
    SpreadsheetApp.getUi().alert('❌ "' + CONFIG.MAIN_SHEET_NAME + '" not found in ABM_R&D_Tax Credit.');
    return;
  }
  if (!permSheet) {
    SpreadsheetApp.getUi().alert('❌ "Email Permutator" not found. Run Steps 2 & 3 first.');
    return;
  }

  var props = PropertiesService.getScriptProperties();
  props.setProperty(PROP_F_PERM_IDX,    '1');
  props.setProperty(PROP_F_APOLLO_IDX,  '1');
  props.setProperty(PROP_F_RUNNING,     'true');
  props.setProperty(PROP_F_ADDED_COUNT, '0');

  _deleteTriggersForFunction(TRIGGER_FN_FINAL);

  SpreadsheetApp.getUi().alert(
    '▶️ Step 4 Starting!\n\n' +
    'Final Format sheet is being built in the background.\n' +
    'Auto-resumes every 30 sec until complete.\n\n' +
    'Output: "Final Format" tab in ABM_R&D_Tax Credit spreadsheet.'
  );

  _runFinalFormatBatch();
}

function autoFinalFormatContinue() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty(PROP_F_RUNNING) !== 'true') {
    _deleteTriggersForFunction(TRIGGER_FN_FINAL);
    return;
  }
  _runFinalFormatBatch();
}

function stopFinalFormat() {
  var props = PropertiesService.getScriptProperties();
  props.setProperty(PROP_F_RUNNING, 'false');
  _deleteTriggersForFunction(TRIGGER_FN_FINAL);

  var added = parseInt(props.getProperty(PROP_F_ADDED_COUNT) || '0');
  SpreadsheetApp.getUi().alert(
    '⛔ Final Format build stopped.\n\nLeads added so far: ' + added +
    '\n\nUse "⏩ Resume" to continue.'
  );
}

function resumeFinalFormat() {
  var props   = PropertiesService.getScriptProperties();
  var permIdx = parseInt(props.getProperty(PROP_F_PERM_IDX) || '1');

  if (permIdx <= 1 && props.getProperty(PROP_F_RUNNING) !== 'true') {
    SpreadsheetApp.getUi().alert('ℹ️ No saved progress. Use "▶️ Start" to begin.');
    return;
  }

  props.setProperty(PROP_F_RUNNING, 'true');
  _deleteTriggersForFunction(TRIGGER_FN_FINAL);
  SpreadsheetApp.getUi().alert('▶️ Resuming Final Format build…');
  _runFinalFormatBatch();
}

// ════════════════════════════════════════════════════════════
//  CORE — Single batch processor (≤ 4.5 min, 500 rows)
// ════════════════════════════════════════════════════════════

function _runFinalFormatBatch() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty(PROP_F_RUNNING) !== 'true') {
    _deleteTriggersForFunction(TRIGGER_FN_FINAL);
    return;
  }

  var automationSS = SpreadsheetApp.getActiveSpreadsheet();
  var crunchbaseSS = SpreadsheetApp.openById(CONFIG.CRUNCHBASE_SS_ID);
  var apolloSS     = SpreadsheetApp.openById(CONFIG.APOLLO_SS_ID);

  var mainSheet   = crunchbaseSS.getSheetByName(CONFIG.MAIN_SHEET_NAME);
  var permSheet   = automationSS.getSheetByName(CONFIG.PERMUTATOR_SHEET_NAME);
  var apolloSheet = apolloSS.getSheetByName(CONFIG.APOLLO_SHEET_NAME);

  if (!mainSheet || !permSheet) {
    props.setProperty(PROP_F_RUNNING, 'false');
    return;
  }

  // Get or create Final Format tab in Crunchbase SS
  var finalSheet = crunchbaseSS.getSheetByName(CONFIG.FINAL_FORMAT_SHEET_NAME);
  var isNew      = false;
  if (!finalSheet) {
    finalSheet = crunchbaseSS.insertSheet(CONFIG.FINAL_FORMAT_SHEET_NAME);
    isNew      = true;
  }

  var lastRow = finalSheet.getLastRow();
  var numCols = FINAL_COLUMNS.length;

  var PI = {
    org   : FINAL_COLUMNS.indexOf('Organization Name'),
    domain: FINAL_COLUMNS.indexOf('Domain'),
    first : FINAL_COLUMNS.indexOf('First Name'),
    last  : FINAL_COLUMNS.indexOf('Last Name'),
    email : FINAL_COLUMNS.indexOf('Email'),
    status: FINAL_COLUMNS.indexOf('Verification Status')
  };

  // ── Load existing emails (dedup) ──────────────────────────
  var seenKeys = Object.create(null);
  if (lastRow > 1) {
    var existingEmails = finalSheet.getRange(1, PI.email + 1, lastRow, 1).getValues();
    for (var j = 1; j < existingEmails.length; j++) {
      var eEmail = String(existingEmails[j][0]).trim().toLowerCase();
      if (eEmail) seenKeys[eEmail] = true;
    }
  }

  // ── Headers ───────────────────────────────────────────────
  if (isNew || lastRow === 0) {
    finalSheet.getRange(1, 1, 1, numCols).setValues([FINAL_COLUMNS]);
    finalSheet.getRange(1, 1, 1, numCols)
      .setBackground('#0d1b2a').setFontColor('#e2e8f0')
      .setFontWeight('bold').setFontSize(10);
    finalSheet.setFrozenRows(1);
    lastRow = 1;
  }

  // ── Read source data ──────────────────────────────────────
  var mainData    = mainSheet.getDataRange().getValues();
  var mainHeaders = mainData[0].map(String);
  var permData    = permSheet.getDataRange().getValues();

  // Main sheet column map
  var mainColMap = Object.create(null);
  for (var h = 0; h < mainHeaders.length; h++) {
    var cn = mainHeaders[h].trim();
    if (cn && !(cn in mainColMap)) mainColMap[cn] = h;
  }

  var orgColIdx = ('Organization Name' in mainColMap) ? mainColMap['Organization Name'] : -1;
  var mainLookup = Object.create(null);
  for (var mi = 1; mi < mainData.length; mi++) {
    var mr  = mainData[mi];
    var org = orgColIdx >= 0 && mr[orgColIdx] ? String(mr[orgColIdx]).trim() : '';
    if (!org) continue;
    var mk = org.toLowerCase();
    if (!(mk in mainLookup)) mainLookup[mk] = [];
    mainLookup[mk].push(mr);
  }

  var mainColForFinal = new Array(numCols);
  for (var fc = 0; fc < numCols; fc++) {
    mainColForFinal[fc] = (FINAL_COLUMNS[fc] in mainColMap) ? mainColMap[FINAL_COLUMNS[fc]] : -1;
  }

  // Apollo Leads lookup
  var apolloByDomain = Object.create(null);
  var apolloByPerson = Object.create(null);
  var apolloFillMap  = [];
  var hasApollo      = false;

  if (apolloSheet) {
    var apolloData    = apolloSheet.getDataRange().getValues();
    var apolloHeaders = apolloData.length > 0 ? apolloData[0].map(String) : [];

    if (apolloHeaders.length > 0) {
      var apolloColMap = Object.create(null);
      apolloHeaders.forEach(function(h, i) {
        var ht = h.trim();
        if (ht && !(ht in apolloColMap)) apolloColMap[ht] = i;
      });

      for (var fc2 = 0; fc2 < numCols; fc2++) {
        var fName = FINAL_COLUMNS[fc2];
        if (fName in PERM_COL_NAMES) continue;
        var aName = (fName in FINAL_TO_APOLLO) ? FINAL_TO_APOLLO[fName] : fName;
        if (aName in apolloColMap) {
          apolloFillMap.push({ fi: fc2, ai: apolloColMap[aName] });
        }
      }

      var aDomainCol  = ('Domain'     in apolloColMap) ? apolloColMap['Domain']     : -1;
      var aWebsiteCol = ('Website'    in apolloColMap) ? apolloColMap['Website']    : -1;
      var aFirstCol   = ('First Name' in apolloColMap) ? apolloColMap['First Name'] : -1;

      for (var ar = 1; ar < apolloData.length; ar++) {
        var aRow    = apolloData[ar];
        var aDomain = '';
        if (aDomainCol >= 0 && aRow[aDomainCol]) {
          aDomain = String(aRow[aDomainCol]).trim().toLowerCase();
        } else if (aWebsiteCol >= 0 && aRow[aWebsiteCol]) {
          aDomain = extractDomain(String(aRow[aWebsiteCol]).trim()).toLowerCase();
        }
        if (!aDomain) continue;

        if (!(aDomain in apolloByDomain)) apolloByDomain[aDomain] = [];
        apolloByDomain[aDomain].push(aRow);

        if (aFirstCol >= 0 && aRow[aFirstCol]) {
          var aFirst = String(aRow[aFirstCol]).trim().toLowerCase();
          var pKey   = aDomain + '|||' + aFirst;
          if (!(pKey in apolloByPerson)) apolloByPerson[pKey] = aRow;
        }
      }
      hasApollo = apolloData.length > 1;
    }
  }

  // ── Load progress ─────────────────────────────────────────
  var startPermIdx = parseInt(props.getProperty(PROP_F_PERM_IDX)   || '1');
  var startApolIdx = parseInt(props.getProperty(PROP_F_APOLLO_IDX) || '1');
  var sesAdded     = parseInt(props.getProperty(PROP_F_ADDED_COUNT) || '0');

  var startTime  = Date.now();
  var MAX_ROWS   = 500;
  var finalRows  = [];
  var hitTimeout = false;

  // ── Phase 1: Email Permutator leads ───────────────────────
  var p = startPermIdx;
  for (; p < permData.length; p++) {
    if (Date.now() - startTime > CONFIG.MAX_RUN_MS || finalRows.length >= MAX_ROWS) {
      hitTimeout = true;
      break;
    }

    var pRow      = permData[p];
    var verEmail  = pRow[10] ? String(pRow[10]).trim() : '';
    var verStatus = pRow[11] ? String(pRow[11]).trim() : '';
    if (!verEmail) continue;

    var emailKey = verEmail.toLowerCase();
    if (emailKey in seenKeys) continue;
    seenKeys[emailKey] = true;

    var pOrg    = pRow[0] ? String(pRow[0]).trim() : '';
    var pFirst  = pRow[1] ? String(pRow[1]).trim() : '';
    var pLast   = pRow[2] ? String(pRow[2]).trim() : '';
    var pDomain = pRow[3] ? String(pRow[3]).trim() : '';
    if (!pOrg) continue;

    var mRows   = (pOrg.toLowerCase() in mainLookup) ? mainLookup[pOrg.toLowerCase()] : null;
    var mainRow = (mRows && mRows.length > 0) ? mRows[0] : null;

    var outRow = new Array(numCols);
    outRow[PI.org]    = pOrg;
    outRow[PI.domain] = pDomain;
    outRow[PI.first]  = pFirst;
    outRow[PI.last]   = pLast;
    outRow[PI.email]  = verEmail;
    outRow[PI.status] = verStatus;

    for (var fi = 0; fi < numCols; fi++) {
      if (outRow[fi] !== undefined) continue;
      var midx = mainColForFinal[fi];
      var mv   = (mainRow && midx >= 0) ? mainRow[midx] : '';
      outRow[fi] = (mv !== null && mv !== undefined) ? mv : '';
    }

    if (hasApollo && pDomain) {
      var dk   = pDomain.toLowerCase();
      var pk   = dk + '|||' + pFirst.toLowerCase();
      var aRow = (pk in apolloByPerson) ? apolloByPerson[pk]
               : ((dk in apolloByDomain) ? apolloByDomain[dk][0] : null);

      if (aRow) {
        for (var af = 0; af < apolloFillMap.length; af++) {
          var entry = apolloFillMap[af];
          var cv    = outRow[entry.fi];
          if (cv !== '' && cv !== null && cv !== undefined) continue;
          var aVal  = aRow[entry.ai];
          if (aVal !== undefined && aVal !== null && String(aVal).trim() !== '') {
            outRow[entry.fi] = aVal;
          }
        }
      }
    }

    finalRows.push(outRow);
  }

  // ── Phase 2: Apollo-only leads (not in permutator) ────────
  var ap = startApolIdx;
  if (!hitTimeout && hasApollo) {
    var apolloDataRef = apolloSheet.getDataRange().getValues();
    var apolloHdrs    = apolloDataRef[0].map(String);
    var aEmailColIdx  = apolloHdrs.indexOf('Email');

    if (aEmailColIdx !== -1) {
      var apolloColForFinal = FINAL_COLUMNS.map(function(fName) {
        var aName = (fName in FINAL_TO_APOLLO) ? FINAL_TO_APOLLO[fName] : fName;
        return apolloHdrs.indexOf(aName);
      });

      for (; ap < apolloDataRef.length; ap++) {
        if (Date.now() - startTime > CONFIG.MAX_RUN_MS || finalRows.length >= MAX_ROWS) {
          hitTimeout = true;
          break;
        }

        var aRow    = apolloDataRef[ap];
        var aEmail  = aRow[aEmailColIdx] ? String(aRow[aEmailColIdx]).trim() : '';
        if (!aEmail) continue;

        var aEmailKey = aEmail.toLowerCase();
        if (aEmailKey in seenKeys) continue;
        seenKeys[aEmailKey] = true;

        var outRow = apolloColForFinal.map(function(aColIdx) {
          var v = (aColIdx !== -1) ? aRow[aColIdx] : '';
          return (v !== null && v !== undefined) ? v : '';
        });
        finalRows.push(outRow);
      }
    }
  }

  // ── Write batch ───────────────────────────────────────────
  if (finalRows.length > 0) {
    finalSheet.getRange(lastRow + 1, 1, finalRows.length, numCols).setValues(finalRows);
    sesAdded += finalRows.length;
  }

  finalSheet.setColumnWidths(1, numCols, 160);
  SpreadsheetApp.flush();

  if (hitTimeout) {
    props.setProperty(PROP_F_PERM_IDX,    String(p));
    props.setProperty(PROP_F_APOLLO_IDX,  String(ap));
    props.setProperty(PROP_F_ADDED_COUNT, String(sesAdded));
    _scheduleTrigger(TRIGGER_FN_FINAL);

    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Perm row ' + p + ' / Apollo row ' + ap + ' | Added: ' + sesAdded,
      '🔄 Auto-resuming Step 4 in 30 sec…',
      15
    );
    return;
  }

  // ── Complete: update Main_Crunchbase summary ──────────────
  props.setProperty(PROP_F_RUNNING, 'false');
  _deleteTriggersForFunction(TRIGGER_FN_FINAL);

  var localValues  = finalSheet.getDataRange().getValues();
  var fDomainIdx   = FINAL_COLUMNS.indexOf('Domain');
  var fStatusIdx   = FINAL_COLUMNS.indexOf('Verification Status');
  var domainStats  = Object.create(null);

  if (fDomainIdx !== -1 && fStatusIdx !== -1) {
    for (var f = 1; f < localValues.length; f++) {
      var fDom  = String(localValues[f][fDomainIdx]).trim().toLowerCase();
      var fStat = String(localValues[f][fStatusIdx]).trim();
      if (!fDom) continue;
      if (!(fDom in domainStats)) domainStats[fDom] = { count: 0, statuses: [] };
      domainStats[fDom].count += 1;
      if (fStat && domainStats[fDom].statuses.indexOf(fStat) === -1) {
        domainStats[fDom].statuses.push(fStat);
      }
    }
  }

  var mainRange   = mainSheet.getDataRange();
  var mainValues  = mainRange.getValues();
  var mHeaders    = mainValues[0].map(function(h) { return String(h).trim(); });
  var howManyIdx  = mHeaders.indexOf('How Many Leads');
  var mVerStatIdx = mHeaders.indexOf('Verification Status');
  var mDomainIdx  = mHeaders.indexOf('Domain');

  if (mainValues.length > 1 && mDomainIdx !== -1) {
    var countOutput  = [];
    var statusOutput = [];
    for (var row = 1; row < mainValues.length; row++) {
      var dom = mainValues[row][mDomainIdx] ? String(mainValues[row][mDomainIdx]).trim().toLowerCase() : '';
      if (dom && dom in domainStats) {
        countOutput.push([domainStats[dom].count]);
        statusOutput.push([domainStats[dom].statuses.join(', ')]);
      } else {
        countOutput.push(['']);
        statusOutput.push(['']);
      }
    }
    if (howManyIdx !== -1) {
      mainSheet.getRange(2, howManyIdx + 1, countOutput.length, 1).setValues(countOutput);
    }
    if (mVerStatIdx !== -1) {
      mainSheet.getRange(2, mVerStatIdx + 1, statusOutput.length, 1).setValues(statusOutput);
    }
  }

  SpreadsheetApp.flush();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Total leads added: ' + sesAdded,
    '🎉 Step 4 Complete! Final Format is ready.',
    20
  );

  SpreadsheetApp.getUi().showModelessDialog(
    HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;color:#e2e8f0;background:#0f172a;' +
      'padding:15px;border-radius:8px;font-size:13px;">' +
      '<h3 style="color:#22c55e;margin-top:0;">✅ Step 4 Complete!</h3>' +
      '<p>1. "Final Format" tab updated in <b>ABM_R&D_Tax Credit</b> spreadsheet.</p>' +
      '<p>2. "Main_Crunchbase" summary columns updated (How Many Leads, Verification Status).</p>' +
      '<p style="font-weight:bold;margin-top:15px;">Total new unique leads added: ' + sesAdded + '</p>' +
      '<button onclick="google.script.host.close()" style="background:#2563eb;color:#fff;' +
      'border:none;padding:6px 16px;border-radius:4px;cursor:pointer;margin-top:10px;">OK</button>' +
      '</div>'
    ).setWidth(420).setHeight(220),
    '✅ Step 4 Complete'
  );
}
