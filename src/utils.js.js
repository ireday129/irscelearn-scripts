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