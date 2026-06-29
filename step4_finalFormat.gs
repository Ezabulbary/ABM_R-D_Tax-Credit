// ============================================================
// step4_finalFormat.gs — Build "Final Format" sheet in batches
//
// ✅ Fixed column order — always the same 23 columns in sequence.
// ✅ Batch-based Execution with Time-based Triggers to prevent
//    "Exceeded maximum execution time" timeouts.
//
// Data source priority (per cell):
//   1. Email Permutator  → Domain, First Name, Last Name, Email, Verification Status
//   2. Main sheet        → remaining columns matched by header name
//   3. Apollo Leads tab  → fills any still-empty cells (matched by domain,
//                          then refined by first name if multiple rows share domain)
// ============================================================

// ── Fixed output column order (23 columns) ────────────────────
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

// Final Format column name → Apollo Leads column name
// (only needed where header names differ between the two sheets)
var FINAL_TO_APOLLO = (function() {
  var m = Object.create(null);
  m['Organization Name']       = 'Company Name';  // Apollo "Company Name" → Final "Organization Name"
  m['Founded Date']            = 'Company Founded Year';
  m['Number of Employees']     = '# Employees';
  m['Total Funding Amount']    = 'Total Funding';
  m['Estimated Revenue Range'] = 'Annual Revenue';
  m['Headquarters Location']   = 'Company Address';
  return m;
}());

// Permutator-supplied column names (these positions are set directly — never from Apollo)
var PERM_COL_NAMES = {
  'Domain': true, 'First Name': true, 'Last Name': true,
  'Email': true, 'Verification Status': true, 'Organization Name': true
};

// Script Properties keys for background execution
var PROP_F_PERM_IDX    = 'FINAL_PERM_IDX';
var PROP_F_APOLLO_IDX  = 'FINAL_APOLLO_IDX';
var PROP_F_RUNNING     = 'FINAL_RUNNING';
var PROP_F_ADDED_COUNT = 'FINAL_ADDED_COUNT';
var FINAL_TRIGGER_FN   = 'autoFinalFormatContinue';

// ════════════════════════════════════════════════════════════
//  PUBLIC — Menu actions
// ════════════════════════════════════════════════════════════

/** Step 4 START — resets progress and starts fresh */
function startFinalFormat() {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName(CONFIG.MAIN_SHEET_NAME);
  var permSheet = ss.getSheetByName(CONFIG.PERMUTATOR_SHEET_NAME);
  if (!mainSheet) {
    SpreadsheetApp.getUi().alert('❌ "' + CONFIG.MAIN_SHEET_NAME + '" sheet not found.');
    return;
  }
  if (!permSheet) {
    SpreadsheetApp.getUi().alert('❌ "' + CONFIG.PERMUTATOR_SHEET_NAME + '" sheet not found. Run Step 2 and Step 3 first.');
    return;
  }

  // Reset properties
  var props = PropertiesService.getScriptProperties();
  props.setProperty(PROP_F_PERM_IDX,    '1');
  props.setProperty(PROP_F_APOLLO_IDX,  '1');
  props.setProperty(PROP_F_RUNNING,     'true');
  props.setProperty(PROP_F_ADDED_COUNT,  '0');
  props.deleteProperty('TEMP_EXT_SS_ID');

  _deleteFinalTriggers();

  SpreadsheetApp.getUi().alert(
    '▶️ Step 4 is starting!\n\n' +
    'The Final Format sheet will be built in the background in batches to prevent execution timeouts.\n\n' +
    'To check progress or stop, use the Step 4 sub-menu.'
  );

  _runFinalFormatBatch();
}

/** Auto-trigger handler — called by Google Trigger */
function autoFinalFormatContinue() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty(PROP_F_RUNNING) !== 'true') {
    _deleteFinalTriggers();
    return;
  }
  _runFinalFormatBatch();
}

/** Step 4 STOP — deletes the trigger */
function stopFinalFormat() {
  var props = PropertiesService.getScriptProperties();
  props.setProperty(PROP_F_RUNNING, 'false');
  _deleteFinalTriggers();

  var added = parseInt(props.getProperty(PROP_F_ADDED_COUNT) || '0');
  SpreadsheetApp.getUi().alert(
    '⛔ Building Final Format has been stopped.\n\n' +
    'Leads added this session: ' + added + '\n\n' +
    'To continue, use Step 4 → Resume.'
  );
}

/** Step 4 RESUME — continues from where it left off */
function resumeFinalFormat() {
  var props = PropertiesService.getScriptProperties();
  var permIdx = parseInt(props.getProperty(PROP_F_PERM_IDX) || '1');
  
  if (permIdx <= 1 && props.getProperty(PROP_F_RUNNING) !== 'true') {
    SpreadsheetApp.getUi().alert('ℹ️ No saved progress found. Use "Start" to begin.');
    return;
  }

  props.setProperty(PROP_F_RUNNING, 'true');
  _deleteFinalTriggers();

  SpreadsheetApp.getUi().alert('▶️ Resuming Final Format sheet build!');
  _runFinalFormatBatch();
}

// ════════════════════════════════════════════════════════════
//  CORE — Single batch processor (max 3.5 min per call)
// ════════════════════════════════════════════════════════════

function _runFinalFormatBatch() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty(PROP_F_RUNNING) !== 'true') {
    _deleteFinalTriggers();
    return;
  }

  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet   = ss.getSheetByName(CONFIG.MAIN_SHEET_NAME);
  var permSheet   = ss.getSheetByName(CONFIG.PERMUTATOR_SHEET_NAME);
  var apolloSheet = ss.getSheetByName(CONFIG.APOLLO_SHEET_NAME); // optional

  if (!mainSheet || !permSheet) {
    props.setProperty(PROP_F_RUNNING, 'false');
    return;
  }

  // ── Read or Create Final Format sheet ─────────────────────
  var finalSheet = ss.getSheetByName(CONFIG.FINAL_FORMAT_SHEET_NAME);
  var isNew = false;
  if (!finalSheet) {
    finalSheet = ss.insertSheet(CONFIG.FINAL_FORMAT_SHEET_NAME);
    isNew = true;
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

  // ── Open or Create external spreadsheet ───────────────────
  var extSS;
  var extSheet;
  var isExtNew = false;
  
  if (CONFIG.EXTERNAL_SPREADSHEET_ID && CONFIG.EXTERNAL_SPREADSHEET_ID.trim() !== '') {
    try {
      extSS = SpreadsheetApp.openById(CONFIG.EXTERNAL_SPREADSHEET_ID.trim());
      extSheet = extSS.getSheetByName(CONFIG.FINAL_FORMAT_SHEET_NAME);
      if (!extSheet) {
        extSheet = extSS.insertSheet(CONFIG.FINAL_FORMAT_SHEET_NAME);
        isExtNew = true;
      }
    } catch (e) {
      extSS = null;
    }
  }

  if (!extSS) {
    var storedExtId = props.getProperty('TEMP_EXT_SS_ID');
    if (storedExtId) {
      try {
        extSS = SpreadsheetApp.openById(storedExtId);
        extSheet = extSS.getSheetByName(CONFIG.FINAL_FORMAT_SHEET_NAME);
      } catch (ignore) {}
    }
    if (!extSS) {
      extSS = SpreadsheetApp.create('Final Format Export - ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'));
      extSheet = extSS.getSheets()[0];
      extSheet.setName(CONFIG.FINAL_FORMAT_SHEET_NAME);
      isExtNew = true;
      props.setProperty('TEMP_EXT_SS_ID', extSS.getId());
    }
  }

  var extLastRow = extSheet.getLastRow();

  // ── Load existing keys (to prevent duplicates) ───────────
  var seenKeys = Object.create(null);
  if (lastRow > 1) {
    var existingValues = finalSheet.getRange(1, PI.email + 1, lastRow, 1).getValues();
    for (var j = 1; j < existingValues.length; j++) {
      var eEmail = String(existingValues[j][0]).trim().toLowerCase();
      if (eEmail) seenKeys[eEmail] = true;
    }
  }
  if (extLastRow > 1) {
    var extValues = extSheet.getRange(1, PI.email + 1, extLastRow, 1).getValues();
    for (var k = 1; k < extValues.length; k++) {
      var extEmail = String(extValues[k][0]).trim().toLowerCase();
      if (extEmail) seenKeys[extEmail] = true;
    }
  }

  // ── Headers Setup ─────────────────────────────────────────
  if (isNew || lastRow === 0) {
    finalSheet.getRange(1, 1, 1, numCols).setValues([FINAL_COLUMNS]);
    finalSheet.getRange(1, 1, 1, numCols)
      .setBackground('#0d1b2a').setFontColor('#e2e8f0')
      .setFontWeight('bold').setFontSize(10);
    finalSheet.setFrozenRows(1);
    lastRow = 1;
  }
  if (isExtNew || extLastRow === 0) {
    extSheet.getRange(1, 1, 1, numCols).setValues([FINAL_COLUMNS]);
    extSheet.getRange(1, 1, 1, numCols)
      .setBackground('#0d1b2a').setFontColor('#e2e8f0')
      .setFontWeight('bold').setFontSize(10);
    extSheet.setFrozenRows(1);
    extLastRow = 1;
  }

  // ── Read source data ──────────────────────────────────────
  var mainData    = mainSheet.getDataRange().getValues();
  var mainHeaders = mainData[0].map(String);
  var permData    = permSheet.getDataRange().getValues();

  // ── Build Main sheet column map & lookup ──────────────────
  var mainColMap = Object.create(null);
  for (var h = 0; h < mainHeaders.length; h++) {
    var cn = mainHeaders[h].trim();
    if (cn && !(cn in mainColMap)) mainColMap[cn] = h;
  }
  var orgColIdx  = ('Organization Name' in mainColMap) ? mainColMap['Organization Name'] : -1;
  var mainLookup = Object.create(null);
  for (var mi = 1; mi < mainData.length; mi++) {
    var mr   = mainData[mi];
    var org  = orgColIdx >= 0 && mr[orgColIdx] ? String(mr[orgColIdx]).trim() : '';
    if (!org) continue;
    var mk = org.toLowerCase();
    if (!(mk in mainLookup)) mainLookup[mk] = [];
    mainLookup[mk].push(mr);
  }
  var mainColForFinal = new Array(numCols);
  for (var fc = 0; fc < numCols; fc++) {
    mainColForFinal[fc] = (FINAL_COLUMNS[fc] in mainColMap) ? mainColMap[FINAL_COLUMNS[fc]] : -1;
  }

  // ── Read Apollo Leads ─────────────────────────────────────
  var apolloByDomain = Object.create(null);
  var apolloByPerson = Object.create(null);
  var hasApollo      = false;
  var apolloFillMap  = [];
  var apolloHeaders  = [];
  var apolloData     = [];

  if (apolloSheet) {
    apolloData = apolloSheet.getDataRange().getValues();
    if (apolloData.length > 0) {
      apolloHeaders = apolloData[0].map(String);
      var apolloColMap = Object.create(null);
      for (var ah = 0; ah < apolloHeaders.length; ah++) {
        var aHdr = apolloHeaders[ah].trim();
        if (aHdr && !(aHdr in apolloColMap)) apolloColMap[aHdr] = ah;
      }
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
      hasApollo = true;
    }
  }

  // ── Load progress ─────────────────────────────────────────
  var startPermIdx = parseInt(props.getProperty(PROP_F_PERM_IDX)   || '1');
  var startApolIdx = parseInt(props.getProperty(PROP_F_APOLLO_IDX) || '1');
  var sesAdded     = parseInt(props.getProperty(PROP_F_ADDED_COUNT) || '0');

  var startTime = Date.now();
  var MAX_MS    = 2 * 60 * 1000; // 2 min safe execution window
  var MAX_ROWS  = 500;           // limit batch rows to prevent huge writes
  var finalRows = [];
  var hitTimeout = false;

  // 1. Process Email Permutator leads
  var p = startPermIdx;
  for (; p < permData.length; p++) {
    if (Date.now() - startTime > MAX_MS || finalRows.length >= MAX_ROWS) {
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

    var lKey    = pOrg.toLowerCase();
    var mRows   = (lKey in mainLookup) ? mainLookup[lKey] : null;
    var mainRow = (mRows && mRows.length > 0) ? mRows[0] : null;

    var outRow = new Array(numCols);
    outRow[PI.org]    = pOrg;
    outRow[PI.domain] = pDomain;
    outRow[PI.first]  = pFirst;
    outRow[PI.last]   = pLast;
    outRow[PI.email]  = verEmail;
    outRow[PI.status] = verStatus;

    for (var mi = 0; mi < numCols; mi++) {
      if (outRow[mi] !== undefined) continue;
      var midx = mainColForFinal[mi];
      var mv   = (mainRow && midx >= 0) ? mainRow[midx] : '';
      outRow[mi] = (mv !== null && mv !== undefined) ? mv : '';
    }

    if (hasApollo && pDomain) {
      var dk  = pDomain.toLowerCase();
      var pk  = dk + '|||' + pFirst.toLowerCase();
      var aRow = (pk in apolloByPerson)
                 ? apolloByPerson[pk]
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

  // 2. Process independent Apollo Leads
  var ap = startApolIdx;
  if (!hitTimeout && hasApollo && apolloData.length > 1) {
    var aEmailColIdx = apolloHeaders.indexOf('Email');
    if (aEmailColIdx !== -1) {
      // Precompute column indices for Apollo to speed up loop
      var apolloColForFinal = new Array(numCols);
      for (var fc = 0; fc < numCols; fc++) {
        var fName = FINAL_COLUMNS[fc];
        var aName = (fName in FINAL_TO_APOLLO) ? FINAL_TO_APOLLO[fName] : fName;
        apolloColForFinal[fc] = apolloHeaders.indexOf(aName);
      }

      for (; ap < apolloData.length; ap++) {
        if (Date.now() - startTime > MAX_MS || finalRows.length >= MAX_ROWS) {
          hitTimeout = true;
          break;
        }

        var aRow = apolloData[ap];
        var aEmail = aRow[aEmailColIdx] ? String(aRow[aEmailColIdx]).trim() : '';
        if (!aEmail) continue;

        var emailKey = aEmail.toLowerCase();
        if (emailKey in seenKeys) continue;
        seenKeys[emailKey] = true;

        var outRow = new Array(numCols);
        for (var fc = 0; fc < numCols; fc++) {
          var aColIdx = apolloColForFinal[fc];
          var aVal = (aColIdx !== -1 && aColIdx !== undefined) ? aRow[aColIdx] : '';
          outRow[fc] = (aVal !== null && aVal !== undefined) ? aVal : '';
        }
        finalRows.push(outRow);
      }
    }
  }

  // ── Write batch data ──────────────────────────────────────
  if (finalRows.length > 0) {
    finalSheet.getRange(lastRow + 1, 1, finalRows.length, numCols).setValues(finalRows);
    extSheet.getRange(extLastRow + 1, 1, finalRows.length, numCols).setValues(finalRows);
    sesAdded += finalRows.length;
  }

  finalSheet.setColumnWidths(1, numCols, 160);
  extSheet.setColumnWidths(1, numCols, 160);
  SpreadsheetApp.flush();

  if (hitTimeout) {
    // Save progress and set next trigger
    props.setProperty(PROP_F_PERM_IDX,    String(p));
    props.setProperty(PROP_F_APOLLO_IDX,  String(ap));
    props.setProperty(PROP_F_ADDED_COUNT, String(sesAdded));
    _scheduleFinalNext();

    ss.toast(
      'Processed through Permutator row ' + p + ' / Apollo row ' + ap + '. Added ' + sesAdded + ' leads.',
      '🔄 Auto-resuming Step 4 in 1 min…',
      15
    );
  } else {
    // ── Complete! Update summary in Main_Crunchbase ─────────
    props.setProperty(PROP_F_RUNNING, 'false');
    _deleteFinalTriggers();
    props.deleteProperty('TEMP_EXT_SS_ID');

    var localValues = finalSheet.getDataRange().getValues();
    var domainStats = Object.create(null);
    var fDomainIdx = FINAL_COLUMNS.indexOf('Domain');
    var fStatusIdx = FINAL_COLUMNS.indexOf('Verification Status');
    
    if (fDomainIdx !== -1 && fStatusIdx !== -1) {
      for (var f = 1; f < localValues.length; f++) {
        var fDom = String(localValues[f][fDomainIdx]).trim().toLowerCase();
        var fStat = String(localValues[f][fStatusIdx]).trim();
        if (!fDom) continue;
        
        if (!(fDom in domainStats)) {
          domainStats[fDom] = { count: 0, statuses: [] };
        }
        domainStats[fDom].count += 1;
        if (fStat && domainStats[fDom].statuses.indexOf(fStat) === -1) {
          domainStats[fDom].statuses.push(fStat);
        }
      }
    }

    var mainRange = mainSheet.getDataRange();
    var mainValues = mainRange.getValues();
    var mainHeaders = mainValues[0].map(String).map(function(h) { return h.trim(); });
    
    var howManyLeadsIdx = mainHeaders.indexOf('How Many Leads');
    var mVerStatusIdx   = mainHeaders.indexOf('Verification Status');
    var mDomainIdx      = mainHeaders.indexOf('Domain');

    if (mainValues.length > 1 && mDomainIdx !== -1) {
      var countOutput = [];
      var statusOutput = [];
      for (var r = 1; r < mainValues.length; r++) {
        var row = mainValues[r];
        var dom = row[mDomainIdx] ? String(row[mDomainIdx]).trim().toLowerCase() : '';
        if (dom && dom in domainStats) {
          countOutput.push([domainStats[dom].count]);
          statusOutput.push([domainStats[dom].statuses.join(', ')]);
        } else {
          countOutput.push(['']);
          statusOutput.push(['']);
        }
      }
      
      if (howManyLeadsIdx !== -1) {
        mainSheet.getRange(2, howManyLeadsIdx + 1, countOutput.length, 1).setValues(countOutput);
      }
      if (mVerStatusIdx !== -1) {
        mainSheet.getRange(2, mVerStatusIdx + 1, statusOutput.length, 1).setValues(statusOutput);
      }
    }

    SpreadsheetApp.flush();

    var extNote = (CONFIG.EXTERNAL_SPREADSHEET_ID && CONFIG.EXTERNAL_SPREADSHEET_ID.trim() !== '')
      ? '🔗 External sheet ID updated:\n' + extSS.getUrl()
      : '🔗 New external sheet created:\n' + extSS.getUrl();

    ss.toast('Leads added: ' + sesAdded, '🎉 Step 4 Complete!', 20);

    // Show a modal completion alert since it finished running in background
    var html = HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;color:#e2e8f0;background:#0f172a;padding:15px;border-radius:8px;font-size:13px;">' +
      '<h3 style="color:#22c55e;margin-top:0;">✅ Step 4 Complete!</h3>' +
      '<p>1. Local "Final Format" sheet has been populated.</p>' +
      '<p>2. "Main_Crunchbase" summary columns populated.</p>' +
      '<p>' + extNote + '</p>' +
      '<p style="font-weight:bold;margin-top:15px;">Total new unique leads added: ' + sesAdded + '</p>' +
      '<button onclick="google.script.host.close()" style="background:#2563eb;color:#fff;border:none;padding:6px 16px;border-radius:4px;cursor:pointer;margin-top:10px;">OK</button>' +
      '</div>'
    ).setWidth(400).setHeight(230);
    SpreadsheetApp.getUi().showModelessDialog(html, '✅ Step 4 Complete');
  }
}

// ── Background trigger helpers ──────────────────────────────

function _scheduleFinalNext() {
  _deleteFinalTriggers();
  ScriptApp.newTrigger(FINAL_TRIGGER_FN)
    .timeBased()
    .after(60 * 1000) // 1 minute from now
    .create();
}

function _deleteFinalTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === FINAL_TRIGGER_FN) {
      ScriptApp.deleteTrigger(t);
    }
  });
}

/**
 * Helper to extract domain
 */
function extractDomain(url) {
  if (!url) return '';
  var d = url
    .replace(/^https?:\/\//i, '')
    .replace(/^www\./i, '')
    .split('/')[0]
    .split('?')[0]
    .split('#')[0]
    .split(':')[0]
    .toLowerCase()
    .trim();
  return d;
}
