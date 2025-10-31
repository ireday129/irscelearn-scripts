/**
 * Highlight Roster rows (entire row) for anyone who has ANY row marked Reported?=TRUE in "Master".
 * Matching precedence: EMAIL → PTIN → "first last".
 * - Uses a soft yellow highlight for the whole row.
 * - If Roster.Valid? is TRUE for that row, clear the highlight (white).
 * - Optionally backfills missing EMAIL/PTIN on the Roster from Master where columns exist.
 * - Designed to be robust even if PTIN is missing on Roster.
 *
 * Menu entry helper: highlightRosterFromReportedHoursMenu()
 */

// Local truthiness helper: uses global truthy_ if present, else a safe fallback.
function _asTrue(val) {
  try {
    if (typeof truthy_ === 'function') return truthy_(val);
  } catch (e) {}
  if (typeof val === 'boolean') return val === true;
  const s = String(val == null ? '' : val).trim().toLowerCase();
  return s === 'true' || s === 'yes' || s === 'y' || s === '1' || s === '✓' || s === '✔';
}

// Build a normalized "first last" key
function _nameKey(first, last) {
  const f = String(first || '').trim().toLowerCase();
  const l = String(last  || '').trim().toLowerCase();
  return f && l ? (f + ' ' + l) : '';
}

function highlightRosterFromReportedHours(quiet) {
  const ss = SpreadsheetApp.getActive();
  const roster = mustGet_(ss, CFG.SHEET_ROSTER);
  const master = mustGet_(ss, CFG.SHEET_MASTER);

  // --- Load Master and build indices ---
  const mVals = master.getDataRange().getValues();
  if (mVals.length <= 1) { if (!quiet) toast_('Master is empty; nothing to highlight.'); return; }

  const mHdr = normalizeHeaderRow_(mVals[0]);
  const mm   = mapHeaders_(mHdr);

  if (mm.reportedCol == null) {
    if (!quiet) toast_('Master missing "Reported?" column mapping in CFG.COL_HEADERS.reportedCol.', true);
    return;
  }
  if (mm.email == null && mm.ptin == null && (mm.firstName == null || mm.lastName == null)) {
    if (!quiet) toast_('Master missing Email/PTIN/Name columns needed for matching.', true);
    return;
  }

  const reportedEmails = new Set();    // emails that have ANY reported row
  const reportedPtins  = new Set();    // ptins that have ANY reported row
  const reportedNames  = new Set();    // "first last" that have ANY reported row

  const emailToSnapshot = new Map();   // email -> { ptin, first, last }
  const ptinToEmail     = new Map();   // ptin  -> email (best-guess backfill)
  const nameToEmail     = new Map();   // "first last" -> email (best-guess backfill)

  const mBody = mVals.slice(1);
  for (let i = 0; i < mBody.length; i++) {
    const row = mBody[i];
    const email = mm.email != null ? String(row[mm.email] || '').toLowerCase().trim() : '';
    const ptin  = mm.ptin  != null ? formatPtinP0_(row[mm.ptin] || '') : '';
    const first = mm.firstName != null ? String(row[mm.firstName] || '').trim() : '';
    const last  = mm.lastName  != null ? String(row[mm.lastName]  || '').trim() : '';
    const reported = _asTrue(row[mm.reportedCol]);
    const nkey = _nameKey(first, last);

    if (ptin && email) ptinToEmail.set(ptin, email);
    if (nkey && email) nameToEmail.set(nkey, email);
    if (email && !emailToSnapshot.has(email)) emailToSnapshot.set(email, { ptin, first, last });

    if (reported) {
      if (email) reportedEmails.add(email);
      if (ptin)  reportedPtins.add(ptin);
      if (nkey)  reportedNames.add(nkey);
    }
  }

  // --- Map Roster headers (PTIN is OPTIONAL) ---
  const rMap = mapRosterHeaders_(roster);
  if (!rMap) { if (!quiet) toast_('Roster header mapping failed.', true); return; }

  const lastRow = roster.getLastRow();
  const lastCol = roster.getLastColumn();
  if (lastRow < 2) { if (!quiet) toast_('Roster has no data to highlight.'); return; }

  const dataRange = roster.getRange(2, 1, lastRow - 1, lastCol);
  const vals   = dataRange.getValues();
  const colors = dataRange.getBackgrounds();

  const YELLOW  = '#fff59d';
  const WHITE   = '#ffffff';

  let wroteBack = false;
  let matches = 0;

  for (let r = 0; r < vals.length; r++) {
    const row = vals[r];

    // Pull current values (guard missing columns)
    const rFirst = (typeof rMap.first === 'number' && rMap.first >= 0) ? row[rMap.first] : '';
    const rLast  = (typeof rMap.last  === 'number' && rMap.last  >= 0) ? row[rMap.last]  : '';
    let   rEmail = (typeof rMap.email === 'number' && rMap.email >= 0) ? String(row[rMap.email] || '').toLowerCase().trim() : '';
    let   rPtin  = (typeof rMap.ptin  === 'number' && rMap.ptin  >= 0) ? formatPtinP0_(row[rMap.ptin] || '') : '';
    const rValid = (typeof rMap.valid === 'number' && rMap.valid >= 0) ? _asTrue(row[rMap.valid]) : false;

    // --- Backfill EMAIL/PTIN if missing using Master lookups (optional) ---
    if (!rEmail) {
      if (rPtin && ptinToEmail.has(rPtin)) {
        rEmail = ptinToEmail.get(rPtin);
      } else {
        const nk = _nameKey(rFirst, rLast);
        if (nk && nameToEmail.has(nk)) rEmail = nameToEmail.get(nk);
      }
      if (rEmail && typeof rMap.email === 'number' && rMap.email >= 0) {
        row[rMap.email] = rEmail;
        wroteBack = true;
      }
    }
    if (!rPtin && rEmail && emailToSnapshot.has(rEmail)) {
      const snap = emailToSnapshot.get(rEmail);
      if (snap.ptin && typeof rMap.ptin === 'number' && rMap.ptin >= 0) {
        rPtin = snap.ptin;
        row[rMap.ptin] = rPtin;
        wroteBack = true;
      }
    }

    // Match precedence: EMAIL -> PTIN -> NAME
    const nkeyR = _nameKey(rFirst, rLast);
    const hasReportedMatch =
      (!!rEmail && reportedEmails.has(rEmail)) ||
      (!!rPtin  && reportedPtins.has(rPtin))   ||
      (!!nkeyR  && reportedNames.has(nkeyR));

    // Decide highlight color
    const rowColor = (!rValid && hasReportedMatch) ? YELLOW : WHITE;
    if (!rValid && hasReportedMatch) matches++;

    // Paint entire row
    colors[r] = new Array(lastCol).fill(rowColor);
  }

  if (wroteBack) dataRange.setValues(vals);
  dataRange.setBackgrounds(colors);

  if (!quiet) {
    toast_(`Roster highlighting updated from Master. Matches: ${matches}`);
  }
}

function highlightRosterFromReportedHoursMenu() {
  try {
    highlightRosterFromReportedHours(true);
    toast_('Roster highlighting updated from Master.Reported?');
  } catch (e) {
    toast_('Failed to update roster highlighting: ' + e.message, true);
    Logger.log(e.stack || e);
  }
}