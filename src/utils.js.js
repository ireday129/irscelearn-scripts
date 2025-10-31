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

// --- normalizeCompletionForUpload_ (define once) ---
if (typeof normalizeCompletionForUpload_ !== 'function') {
  /**
   * Normalize a completion date for IRS upload rules:
   * - Parses many date shapes.
   * - If the date is in the future OR more than 4 days in the past,
   *   coerce it to "yesterday" (local tz, no time).
   * - Otherwise return the parsed Date.
   *
   * @param {*} v  A Date, serial number, or string
   * @return {Date|*} normalized Date or original value if unparsable
   */
  function normalizeCompletionForUpload_(v) {
    // Prefer project-wide parser if present
    var d = (typeof parseDate_ === 'function') ? parseDate_(v) : null;

    // Simple fallback parse if parseDate_ is not available
    if (!d) {
      if (v instanceof Date && !isNaN(v)) {
        d = new Date(v.getFullYear(), v.getMonth(), v.getDate());
      } else if (typeof v === 'number' && v > 20000) {
        // Excel serial (1899-12-30 base)
        var base = new Date(1899, 11, 30);
        d = new Date(base.getTime() + v * 24 * 60 * 60 * 1000);
        d = new Date(d.getFullYear(), d.getMonth(), d.getDate());
      } else if (v) {
        var t = new Date(String(v));
        if (!isNaN(t)) d = new Date(t.getFullYear(), t.getMonth(), t.getDate());
      }
    }

    if (!d || isNaN(d)) return v; // leave as-is if we can’t parse

    var today = new Date();
    var todayMid = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    var diffDays = Math.floor((todayMid - d) / 86400000); // 24*60*60*1000

    if (diffDays > 4 || diffDays < 0) {
      var yesterday = new Date(todayMid.getFullYear(), todayMid.getMonth(), todayMid.getDate() - 1);
      return yesterday;
    }
    return d;
  }
}

/**
 * Clears only the data body of the Clean sheet and writes new rows.
 * @param {SpreadsheetApp.Spreadsheet} ss
 * @param {Array<Object>} cleanRows - objects like:
 *   { first, last, ptin, email, prog, hours, completion, issue }
 */
function writeCleanDataOnly_(ss, cleanRows) {
  const sh = mustGet_(ss, CFG.SHEET_CLEAN);

  // 1) Clear existing body (keep headers)
  const lastRow = sh.getLastRow();
  if (lastRow > 1) {
    sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).clearContent();
  }

  // Nothing to write? we're done.
  if (!cleanRows || !cleanRows.length) return;

  // 2) Map headers to column indexes
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(x => String(x || '').trim());
  const cMap = mapCleanHeaders_(hdr);
  const iF  = cMap.firstName;
  const iL  = cMap.lastName;
  const iP  = cMap.ptin;
  const iE  = cMap.email;
  const iG  = cMap.program;
  const iH  = cMap.hours;
  const iC  = cMap.completion;
  const iRI = cMap.issue;

  if ([iF, iL, iP, iE, iG, iH, iC].some(v => v < 0)) {
    toast_('Clean headers missing/renamed; cannot write rows.', true);
    return;
  }

  // 3) Build rows in header order
  const data = cleanRows.map(r => {
    const arr = new Array(hdr.length).fill('');
    arr[iF]  = r.first || '';
    arr[iL]  = r.last || '';
    arr[iP]  = formatPtinP0_(r.ptin || '');
    arr[iE]  = String(r.email || '').toLowerCase().trim();
    arr[iG]  = normalizeProgram_(r.prog || '');
    arr[iH]  = r.hours || '';
    arr[iC]  = formatToMDY_(r.completion || '');
    if (iRI >= 0) arr[iRI] = r.issue ? String(r.issue).trim() : '';
    return arr;
  });

  // 4) Write values and set date format on completion column
  sh.getRange(2, 1, data.length, hdr.length).setValues(data);
  if (iC >= 0) {
    sh.getRange(2, iC + 1, data.length, 1).setNumberFormat('mm/dd/yyyy');
  }
}

/**
 * Append staged rows to the Clean sheet (no duplicate removal here).
 * Intended for resumable batch writes.
 * cleanRows = [{ first, last, ptin, email, prog, hours, completion, issue }]
 */
function appendToClean_(ss, cleanRows) {
  if (!cleanRows || !cleanRows.length) return;

  const sh = mustGet_(ss, CFG.SHEET_CLEAN);
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
    .map(x => String(x || '').trim());
  const cMap = mapCleanHeaders_(hdr);

  const iF  = cMap.firstName;
  const iL  = cMap.lastName;
  const iP  = cMap.ptin;
  const iE  = cMap.email;
  const iG  = cMap.program;
  const iH  = cMap.hours;
  const iC  = cMap.completion;
  const iRI = cMap.issue;

  if ([iF,iL,iP,iE,iG,iH,iC].some(v => v < 0)) {
    toast_('Clean headers missing/renamed — append aborted', true);
    return;
  }

  const start = sh.getLastRow() + 1;
  const data = cleanRows.map(r => {
    const arr = new Array(hdr.length).fill('');
    arr[iF]  = r.first || '';
    arr[iL]  = r.last || '';
    arr[iP]  = formatPtinP0_(r.ptin || '');
    arr[iE]  = String(r.email || '').toLowerCase().trim();
    arr[iG]  = normalizeProgram_(r.prog || '');
    arr[iH]  = r.hours || '';
    arr[iC]  = formatToMDY_(r.completion || '');
    if (iRI >= 0) arr[iRI] = r.issue ? String(r.issue).trim() : '';
    return arr;
  });

  sh.getRange(start, 1, data.length, hdr.length).setValues(data);

  if (iC >= 0) {
    sh.getRange(start, iC + 1, data.length, 1).setNumberFormat('mm/dd/yyyy');
  }
}

// --- unresolved issue index (define once) ---
// Scans Master and returns keys for rows still carrying a Reporting Issue?
// Output: { byPtinProg:Set, byEmailProg:Set, byNameProg:Set }
if (typeof buildUnresolvedIssueIndex_ !== 'function') {
  function buildUnresolvedIssueIndex_() {
    const ss = SpreadsheetApp.getActive();
    const out = {
      byPtinProg: new Set(),
      byEmailProg: new Set(),
      byNameProg: new Set()
    };

    const master = ss.getSheetByName(CFG.SHEET_MASTER);
    if (!master || master.getLastRow() <= 1) return out;

    const mVals = master.getDataRange().getValues();
    const mHdr  = normalizeHeaderRow_(mVals[0]);
    const mm    = mapHeaders_(mHdr);

    const iF = mm.firstName;
    const iL = mm.lastName;
    const iP = mm.ptin;
    const iE = mm.email;
    const iG = mm.program;
    const iIssue = mm.masterIssueCol;

    if (iIssue == null || iG == null) return out;

    const body = mVals.slice(1);
    for (let r = 0; r < body.length; r++) {
      const row = body[r];

      // consider any non-blank, non-'fixed' issue as unresolved
      const issue = String(row[iIssue] || '').trim().toLowerCase();
      if (!issue || issue === 'fixed') continue;

      const prog  = normalizeProgram_(row[iG] || '');
      if (!prog) continue;

      const ptin  = iP != null ? formatPtinP0_(row[iP] || '') : '';
      const email = iE != null ? String(row[iE] || '').toLowerCase().trim() : '';
      const fn    = iF != null ? String(row[iF] || '').trim().toLowerCase() : '';
      const ln    = iL != null ? String(row[iL] || '').trim().toLowerCase() : '';

      if (ptin) out.byPtinProg.add(ptin + '|' + prog);
      if (email) out.byEmailProg.add(email + '|' + prog);
      if (fn && ln) out.byNameProg.add(fn + ' ' + ln + '|' + prog);
    }

    return out;
  }
}
// Close the if wrapper for buildUnresolvedIssueIndex_ here

if (typeof mapHeaders_ !== 'function') {
  /**
   * Master header mapper (case-insensitive, with aliases).
   * Prefers CFG.COL_HEADERS keys when present; otherwise falls back to common aliases.
   */
  function mapHeaders_(hdr) {
    const lower = hdr.map(h => String(h || '').trim().toLowerCase().replace(/\s+/g, ' '));

    const findAny = (candidates) => {
      for (const c of candidates) {
        const i = lower.indexOf(String(c || '').toLowerCase().trim().replace(/\s+/g, ' '));
        if (i >= 0) return i;
      }
      return null;
    };

    // Pull desired canonical labels from CFG if available
    const H = (typeof CFG !== 'undefined' && CFG.COL_HEADERS) ? CFG.COL_HEADERS : {};

    // Build candidate lists (CFG value first, then aliases)
    const C = {
      firstName:  [H.firstName,  'attendee first name', 'first name', 'first'],
      lastName:   [H.lastName,   'attendee last name',  'last name',  'last', 'surname'],
      ptin:       [H.ptin,       'ptin', 'attendee ptin'],
      email:      [H.email,      'email', 'e-mail', 'mail'],
      program:    [H.program,    'program number', 'program', 'program #', 'course number'],
      hours:      [H.hours,      'ce hours', 'ce hours awarded', 'hours'],
      completion: [H.completion, 'program completion date', 'completion date', 'date completed'],
      group:      [H.group,      'group'],
      masterIssueCol: [H.masterIssueCol, 'reporting issue?', 'reporting issue'],
      reportedCol:    [H.reportedCol,    'reported?', 'reported'],
      updatedCol:     [H.updatedCol,     'last updated'],
      reportedAtCol:  [H.reportedAtCol,  'date reported', 'reported at', 'reported date'],
      source:         [H.source,         'source']
    };

    const out = {};
    for (const [key, candidates] of Object.entries(C)) {
      const candList = candidates.filter(Boolean);
      out[key] = candList.length ? findAny(candList) : null;
    }
    return out;
  }
}

/** Minimal recheckMaster so the menu isn’t broken. */
function recheckMaster() {
  try {
    const ss = SpreadsheetApp.getActive();
    const master = mustGet_(ss, CFG.SHEET_MASTER);
    const mVals = master.getDataRange().getValues();
    if (mVals.length <= 1) { toast_('Master is empty.', true); return; }

    const hdr = normalizeHeaderRow_(mVals[0]);
    const mm  = mapHeaders_(hdr);
    const rosterMap = getRosterMap_(ss); // ok if null
    const ptinRe = /^P0\d{7}$/i;

    const out = [];
    for (let r = 1; r < mVals.length; r++) {
      const row = mVals[r];
      const first = String(row[mm.firstName] || '').trim();
      const last  = String(row[mm.lastName]  || '').trim();
      const ptin  = formatPtinP0_(row[mm.ptin] || '');
      let status  = 'Good';

      if (!ptin) status = 'Missing PTIN';
      else if (!ptinRe.test(ptin) || ptin === 'P00000000') status = 'PTIN does not exist';

      if (status === 'Good' && rosterMap && ptin) {
        const ro = rosterMap.get(ptin);
        if (ro && !namesMatchFull_(first, last, ro.first, ro.last)) status = 'PTIN & name do not match';
      }
      out.push([status === 'Good' ? '' : status]);
    }

    if (mm.masterIssueCol != null) {
      master.getRange(2, mm.masterIssueCol + 1, out.length, 1).setValues(out);
    }
    toast_('Master rechecked.');
  } catch (e) {
    toast_('Recheck failed: ' + e.message, true);
    Logger.log(e.stack || e);
  }
}
/** ——— SANITY HELPERS (global, lightweight) ——— **/

/** Clean header mapper (case-insensitive) */
function mapCleanHeaders_(hdr) {
  const lower = hdr.map(h => String(h || '').toLowerCase().trim().replace(/\s+/g, ' '));
  const findAny = (names) => {
    for (const n of names) {
      const i = lower.indexOf(String(n || '').toLowerCase().trim().replace(/\s+/g, ' '));
      if (i >= 0) return i;
    }
    return -1;
  };
  return {
    firstName:  findAny(['attendee first name','first name','first']),
    lastName:   findAny(['attendee last name','last name','last','surname']),
    ptin:       findAny(['attendee ptin','ptin']),
    email:      findAny(['email','e-mail','mail']),
    program:    findAny(['program number','program','# program','program #','course number']),
    hours:      findAny(['ce hours awarded','ce hours','hours']),
    completion: findAny(['program completion date','completion date','date completed']),
    issue:      findAny(['reporting issue?','reporting issue'])
  };
}

/** Parse many date shapes into a Date (date-only) or null */
function parseDate_(v) {
  if (v == null) return null;
  if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  const s = String(v).trim();
  if (!s) return null;
  const n = Number(s);
  if (!isNaN(n) && n > 20000) { // Excel serial
    const base = new Date(1899, 11, 30);
    const d = new Date(base.getTime() + n * 86400000);
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }
  const d2 = new Date(s);
  return isNaN(d2.getTime()) ? null : new Date(d2.getFullYear(), d2.getMonth(), d2.getDate());
}

/** Format a date-ish thing to MM/dd/yyyy (string) */
function formatToMDY_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'MM/dd/yyyy');
  }
  const d = parseDate_(v);
  return d ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'MM/dd/yyyy') : '';
}