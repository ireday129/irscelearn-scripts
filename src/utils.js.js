function recheckMaster() {
  try {
    clearMasterIssuesFromFixedIssues_(true); // safe stub

    backfillMasterPtinFromRoster_(true);

    const ss = SpreadsheetApp.getActive();
    const master = mustGet_(ss, CFG.SHEET_MASTER);
    const mVals = master.getDataRange().getValues();
    if (mVals.length<=1) { toast_('Master is empty.', true); return; }

    const mHdr = normalizeHeaderRow_(mVals[0]);
    const mMap = mapHeaders_(mHdr);
    const roster = getRosterMap_(ss);

    // Valid *format* only — does NOT bless P00000000
    const ptinRe = /^P0\d{7}$/i;

    const out = [];
    for (let i=1;i<mVals.length;i++){
      const row   = mVals[i];
      const first = String(row[mMap.firstName]||'').trim();
      const last  = String(row[mMap.lastName]||'').trim();
      const ptin  = formatPtinP0_(row[mMap.ptin]||'');
      const mIssue= String(row[mMap.masterIssueCol]||'').trim();

      // If the row is already Reported? -> do not compute issues
      if (parseBool_(row[mMap.reportedCol])) {
        out.push(['']); // leave blank
        continue;
      }

      // Preserve existing "PTIN does not exist" exactly as requested
      if (mIssue === 'PTIN does not exist') {
        out.push(['PTIN does not exist']);
        continue;
      }

      // Compute fresh status
      let status = 'Good';

      if (!ptin) {
        status = 'Missing PTIN';
      } else if (ptin.toUpperCase() === 'P00000000') {
        // Explicitly treat all-zero PTIN as invalid
        status = 'PTIN does not exist';
      } else if (!ptinRe.test(ptin)) {
        status = 'PTIN does not exist';
      } else if (roster) {
        const ro = roster.get(ptin);
        if (ro && !namesMatchFull_(first, last, ro.first, ro.last)) {
          status = 'PTIN & name do not match';
        }
      }

      // Respect "Fixed" marker
      if (/^fixed$/i.test(mIssue)) status = 'Good';

      out.push([status==='Good' ? '' : mapFreeTextToStandardIssue_(status)]);
    }

    if (mMap.masterIssueCol!=null) {
      master.getRange(2, mMap.masterIssueCol+1, out.length, 1).setValues(out);
    }

    toast_('Master rechecked (sticky "PTIN does not exist" preserved).', false);
  } catch(e) {
    toast_('Recheck failed: ' + e.message, true);
    Logger.log(e.stack||e.message);
  }
}

/***** ====================== ROSTER ↔ WEBHOOK & HIGHLIGHT ====================== *****/

/** Define webhook URL only if not already defined elsewhere */
if (typeof WEBHOOK_URL === 'undefined') {
  var WEBHOOK_URL = 'https://irscelearn.com/wp-json/uap/v2/uap-5213-5214';
}

/**
 * Post webhook for a given Roster row (expects mapRosterHeaders_ to be available).
 * Sends: { email, first_name, last_name }
 */
function postRosterWebhookForRow_(row, rMap) {
  try {
    const email = String(row[rMap.email] || '').trim().toLowerCase();
    const first = String(row[rMap.first] || '').trim();
    const last  = String(row[rMap.last]  || '').trim();
    if (!email || !first || !last) {
      Logger.log('Webhook skipped: missing email/first/last on row.');
      return;
    }

    const payload = {
      email: email,
      first_name: first,
      last_name: last
    };

    const res = UrlFetchApp.fetch(WEBHOOK_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const code = res.getResponseCode();
    if (code >= 200 && code < 300) {
      toast_('Webhook sent for ' + email);
    } else {
      Logger.log('Webhook non-2xx (' + code + '): ' + res.getContentText());
      toast_('Webhook failed (' + code + ') for ' + email, true);
    }
  } catch (e) {
    Logger.log('Webhook error: ' + (e.stack || e.message));
    toast_('Webhook error: ' + e.message, true);
  }
}

/**
 * Highlight all Roster rows (yellow) whose PTIN appears in Reported Hours.
 * Removes highlight for rows not in Reported Hours.
 * - Uses PTIN only (program-agnostic).
 */
function refreshRosterHighlightsFromReportedHours(quiet) {
  const ss = SpreadsheetApp.getActive();
  const roster = ss.getSheetByName(CFG.SHEET_ROSTER);
  const rh     = ss.getSheetByName('Reported Hours');
  if (!roster || !rh) { if(!quiet) toast_('Roster or Reported Hours not found.', true); return; }

  // Build set of PTINs from Reported Hours
  const rhVals = rh.getDataRange().getValues();
  if (rhVals.length <= 1) { if(!quiet) toast_('Reported Hours is empty; nothing to highlight.'); return; }
  const rhHdr  = rhVals[0].map(s=>String(s||'').trim());
  const iRhPT  = rhHdr.findIndex(h => /^ptin$/i.test(String(h)));
  const iRhPT2 = rhHdr.findIndex(h => /attendee ptin/i.test(String(h)));
  const idxPT  = iRhPT >= 0 ? iRhPT : iRhPT2;
  if (idxPT < 0) { if(!quiet) toast_('Reported Hours missing PTIN column.', true); return; }

  const ptinSet = new Set();
  for (let r=1; r<rhVals.length; r++) {
    const pt = formatPtinP0_(rhVals[r][idxPT] || '');
    if (pt) ptinSet.add(pt);
  }

  // Map Roster headers
  const rMap = mapRosterHeaders_(roster);
  if (!rMap) { if(!quiet) toast_('Roster headers not recognized.', true); return; }

  const rVals = roster.getDataRange().getValues();
  if (rVals.length <= 1) return;

  // Prepare background updates
  const startRow = 2;
  const numRows  = rVals.length - 1;
  const numCols  = roster.getLastColumn();

  if (numRows <= 0) return;

  // Current backgrounds
  const bgRange = roster.getRange(startRow, 1, numRows, numCols);
  const bgs = bgRange.getBackgrounds();

  const YELLOW  = '#fff59d';
  const CLEAR   = '#ffffff';

  let changed = 0;
  for (let i=0; i<numRows; i++) {
    const row = rVals[i+1];
    const pt  = formatPtinP0_(row[rMap.ptin] || '');
    const shouldHighlight = !!(pt && ptinSet.has(pt));

    // Only set the first column’s bg; then copy across row for speed
    const currentIsYellow = (bgs[i][0] || '').toLowerCase() === YELLOW;
    if (shouldHighlight && !currentIsYellow) {
      for (let c=0; c<numCols; c++) bgs[i][c] = YELLOW;
      changed++;
    } else if (!shouldHighlight && currentIsYellow) {
      for (let c=0; c<numCols; c++) bgs[i][c] = CLEAR;
      changed++;
    }
  }

  if (changed) {
    bgRange.setBackgrounds(bgs);
  }

  if (!quiet) toast_('Roster highlight sync done (' + changed + ' row style update' + (changed===1?'':'s') + ').');
}

/** Safe function caller to avoid onEdit hard-crashes (only define if missing). */
if (typeof safeCall_ !== 'function') {
  function safeCall_(fn, arg) {
    try { if (typeof fn === 'function') fn(arg); } catch (e) { Logger.log(e.stack || e.message); }
  }
}

/**
 * Unified onEdit:
 *  - Calls existing processMasterEdits(e) safely (if present)
 *  - Detects Roster → Valid? changes; when set TRUE => posts webhook and clears highlight for that row
 * If an onEdit already exists, we wrap it to preserve prior behavior.
 */
(function(){
  var __hadOnEdit = (typeof onEdit === 'function');
  var __prevOnEdit = __hadOnEdit ? onEdit : null;

  function rosterOnEditBlock_(e) {
    try {
      if (!e || !e.range) return;
      const sh = e.range.getSheet();
      if (!sh) return;

      const rosterName = String(CFG.SHEET_ROSTER || 'Roster').trim();
      if (sh.getName() !== rosterName) return;

      const rMap = mapRosterHeaders_(sh);
      if (!rMap || rMap.valid < 0) return;

      const editedCol = e.range.getColumn() - 1; // zero-based
      if (editedCol !== rMap.valid) return;

      // React only when set to TRUE
      const nowTrue = parseBool_(e.value);
      if (!nowTrue) return;

      const rowIdx = e.range.getRow(); // 1-based
      if (rowIdx <= 1) return; // skip header

      const rowVals = sh.getRange(rowIdx, 1, 1, sh.getLastColumn()).getValues()[0];

      // Fire webhook
      postRosterWebhookForRow_(rowVals, rMap);

      // Clear highlight for that row
      const numCols = sh.getLastColumn();
      const clearBg = new Array(numCols).fill('#ffffff');
      sh.getRange(rowIdx, 1, 1, numCols).setBackgrounds([clearBg]);

    } catch (err) {
      Logger.log('onEdit roster block error: ' + (err.stack || err.message));
    }
  }

  onEdit = function(e) {
    // Run any previously defined onEdit
    if (__prevOnEdit) {
      try { __prevOnEdit(e); } catch (err) { Logger.log('prev onEdit error: ' + (err.stack || err.message)); }
    }

    // Master-side edit handling if you have it
    safeCall_(processMasterEdits, e);

    // Roster handler for Valid? TRUE → webhook + unhighlight
    rosterOnEditBlock_(e);
  };
})();
(function() {
  // Add or replace writeCleanDataOnly_ at the end of the file.
  /**
   * Clears only the data body of the Clean sheet (keeps header row).
   * Safe to call even if the sheet is empty. The second parameter is accepted
   * for backward compatibility with older callers but is ignored.
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss Active spreadsheet handle
   * @param {*} _cleanRows (ignored)
   */
  function writeCleanDataOnly_(ss, _cleanRows) {
    var sh = mustGet_(ss, CFG.SHEET_CLEAN);
    var last = sh.getLastRow();
    if (last > 1) {
      sh.getRange(2, 1, last - 1, sh.getLastColumn()).clearContent();
    }
  }
  // Expose to global scope if not already present or replace if needed
  this.writeCleanDataOnly_ = writeCleanDataOnly_;
})();

/**
 * Build an index of unresolved issues from Master.
 * Returns sets keyed by PTIN|Program, email|Program, and "first last"|Program (all normalized).
 * Only includes rows where Master."Reporting Issue?" is non-blank and not "fixed".
 */
function buildUnresolvedIssueIndex_() {
  var ss = SpreadsheetApp.getActive();
  var out = {
    byPtinProg: new Set(),
    byEmailProg: new Set(),
    byNameProg: new Set()
  };

  var master = ss.getSheetByName(CFG.SHEET_MASTER);
  if (!master || master.getLastRow() <= 1) return out;

  var mVals = master.getDataRange().getValues();
  var mHdr  = normalizeHeaderRow_(mVals[0]);
  var mm    = mapHeaders_(mHdr);

  // Required indexes we rely on; guard for missing columns
  var iIssue = mm.masterIssueCol;
  var iProg  = mm.program;
  var iPtin  = mm.ptin;
  var iEmail = mm.email;
  var iFname = mm.firstName;
  var iLname = mm.lastName;

  if (iIssue == null || iProg == null) return out;

  var body = mVals.slice(1);
  for (var r = 0; r < body.length; r++) {
    var row   = body[r];
    var issue = String(row[iIssue] || '').trim();

    // Skip blank or explicitly 'fixed'
    if (!issue || /^fixed$/i.test(issue)) continue;

    var prog = normalizeProgram_(row[iProg] || '');
    if (!prog) continue;

    var ptin  = (iPtin  != null) ? formatPtinP0_(row[iPtin]  || '') : '';
    var email = (iEmail != null) ? String(row[iEmail] || '').toLowerCase().trim() : '';
    var fn    = (iFname != null) ? String(row[iFname] || '').trim().toLowerCase() : '';
    var ln    = (iLname != null) ? String(row[iLname] || '').trim().toLowerCase() : '';

    if (ptin) out.byPtinProg.add(ptin + '|' + prog);
    if (email) out.byEmailProg.add(email + '|' + prog);
    if (fn && ln) out.byNameProg.add(fn + ' ' + ln + '|' + prog);
  }

return out;
}

// --- name comparison helper (define once) ---
if (typeof namesMatchFull_ !== 'function') {
  function namesMatchFull_(f1, l1, f2, l2) {
    const norm = s => String(s || '')
      .trim()
      .replace(/\s+/g, ' ')
      .toLowerCase();
    return norm(f1) === norm(f2) && norm(l1) === norm(l2);
  }
}