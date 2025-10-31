/**
 * Highlight Roster rows (entire row) for anyone who has ANY row marked Reported?=TRUE in "Master".
 * - Matches by PTIN (normalized to P0#######).
 * - Uses a soft yellow highlight for the whole row.
 * - If Roster.Valid? is TRUE for that row, clear the highlight (no background).
 */
// Local truthiness helper: uses global truthy_ if present, else a safe fallback
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
  {
    const ss = SpreadsheetApp.getActive();
    const roster = mustGet_(ss, CFG.SHEET_ROSTER);
    const master = mustGet_(ss, CFG.SHEET_MASTER);

    // --- Load Master and build indices ---
    const mVals = master.getDataRange().getValues();
    if (mVals.length <= 1) { if (!quiet) toast_('Master is empty; nothing to highlight.'); return; }

    const mHdr = normalizeHeaderRow_(mVals[0]);
    const mm   = mapHeaders_(mHdr);

    if (mm.ptin == null || mm.reportedCol == null) {
      if (!quiet) toast_('Master missing PTIN and/or Reported? column.', true);
      return;
    }

    const reportedPtins = new Set();
    const emailToMaster = new Map(); // email -> { ptin, email, first, last, reported }
    const nameToMaster  = new Map(); // "first last" -> same
    const ptinToEmail   = new Map(); // ptin -> email (for backfilling email)
    const reportedEmails = new Set();   // email with ANY reported row
    const reportedNames  = new Set();   // "first last" with ANY reported row

    const mBody = mVals.slice(1);
    for (let i = 0; i < mBody.length; i++) {
      const row = mBody[i];
      const ptin = formatPtinP0_(row[mm.ptin] || '');
      const reported = _asTrue(row[mm.reportedCol]);

      const email = mm.email != null ? String(row[mm.email] || '').toLowerCase().trim() : '';
      const first = mm.firstName != null ? String(row[mm.firstName] || '').trim() : '';
      const last  = mm.lastName  != null ? String(row[mm.lastName]  || '').trim() : '';
      const nkey  = _nameKey(first, last);

      if (ptin) ptinToEmail.set(ptin, email);
      if (reported) {
        if (ptin)  reportedPtins.add(ptin);
        if (email) reportedEmails.add(email);
        if (nkey)  reportedNames.add(nkey);
      }
      if (email) emailToMaster.set(email, { ptin, email, first, last, reported });
      if (nkey)  nameToMaster.set(nkey, { ptin, email, first, last, reported });
    }

    // --- Map Roster headers ---
    const rMap = mapRosterHeaders_(roster);
    if (!rMap) { if (!quiet) toast_('Roster header mapping failed.', true); return; }
    if (rMap.ptin < 0) { if (!quiet) toast_('Roster missing PTIN column.', true); return; }
    const hasValidCol = typeof rMap.valid === 'number' && rMap.valid >= 0;

    const lastRow = roster.getLastRow();
    const lastCol = roster.getLastColumn();
    if (lastRow < 2) { if (!quiet) toast_('Roster has no data to highlight.'); return; }

    const dataRange = roster.getRange(2, 1, lastRow - 1, lastCol);
    const vals   = dataRange.getValues();
    const colors = dataRange.getBackgrounds();

    const YELLOW  = '#fff59d';
    const DEFAULT = '#ffffff'; // always clear to white when not highlighted or when Valid? is TRUE

    let anyWriteBack = false;

    for (let r = 0; r < vals.length; r++) {
      const row = vals[r];

      // Pull current values
      const rFirst = rMap.first >= 0 ? row[rMap.first] : '';
      const rLast  = rMap.last  >= 0 ? row[rMap.last]  : '';
      let   rEmail = rMap.email >= 0 ? String(row[rMap.email] || '').toLowerCase().trim() : '';
      let   rPtin  = formatPtinP0_(row[rMap.ptin] || '');
      const rValid = hasValidCol ? _asTrue(row[rMap.valid]) : false;

      // --- Backfill PTIN or Email if missing, using Master ---
      // Prefer email -> PTIN, else name -> PTIN/email, else PTIN -> email
      if (!rPtin) {
        if (rEmail && emailToMaster.has(rEmail)) {
          const hit = emailToMaster.get(rEmail);
          if (hit.ptin) {
            rPtin = hit.ptin;
            row[rMap.ptin] = rPtin;
            anyWriteBack = true;
          }
        } else {
          const key = _nameKey(rFirst, rLast);
          if (key && nameToMaster.has(key)) {
            const hit = nameToMaster.get(key);
            if (hit.ptin && !rPtin) {
              rPtin = hit.ptin;
              row[rMap.ptin] = rPtin;
              anyWriteBack = true;
            }
            if (!rEmail && hit.email) {
              rEmail = hit.email;
              if (rMap.email >= 0) {
                row[rMap.email] = rEmail;
                anyWriteBack = true;
              }
            }
          }
        }
      }

      if (!rEmail && rPtin && ptinToEmail.has(rPtin)) {
        rEmail = ptinToEmail.get(rPtin);
        if (rMap.email >= 0 && rEmail) {
          row[rMap.email] = rEmail;
          anyWriteBack = true;
        }
      }

      const nkeyR = _nameKey(rFirst, rLast);
      const hasReportedMatch =
        (rPtin && reportedPtins.has(rPtin)) ||
        (rEmail && reportedEmails.has(rEmail)) ||
        (nkeyR && reportedNames.has(nkeyR));

      // --- Decide highlight color
      let rowColor = DEFAULT;
      if (!rValid && hasReportedMatch) {
        rowColor = YELLOW;
      }

      const newRowColors = new Array(lastCol).fill(rowColor);
      colors[r] = newRowColors;
    }

    // Write back any roster PTIN/email backfills
    if (anyWriteBack) {
      dataRange.setValues(vals);
    }

    // Apply highlights
    dataRange.setBackgrounds(colors);
    if (!quiet) toast_('Roster highlighting & backfill updated from Master.');
  }
}

function highlightRosterFromReportedHoursMenu() {
  try {
    highlightRosterFromReportedHours(true);
    toast_('Roster highlighting updated from Master.Reported?.');
  } catch (e) {
    toast_('Failed to update roster highlighting: ' + e.message, true);
    Logger.log(e.stack || e);
  }
}

/**
 * Highlight Roster rows (entire row) for anyone who has ANY row marked Reported?=TRUE in "Master".
 * Primary match: EMAIL. Secondary fallbacks: PTIN, then "first last".
 * - Uses a soft yellow highlight for the whole row.
 * - If Roster.Valid? is TRUE for that row, clear the highlight (white).
 * - PTIN/Email backfill is OPTIONAL: if the target column is missing, we skip writing it.
 */

// Local truthiness helper: uses global truthy_ if present, else a safe fallback
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
    if (!quiet) toast_('Master missing Reported? column.', true);
    return;
  }
  if (mm.email == null && mm.ptin == null && (mm.firstName == null || mm.lastName == null)) {
    if (!quiet) toast_('Master missing Email/PTIN/Name columns needed for matching.', true);
    return;
  }

  const reportedEmails = new Set();    // emails that have ANY reported row
  const reportedPtins  = new Set();    // ptins that have ANY reported row
  const reportedNames  = new Set();    // "first last" that have ANY reported row

  const emailToMaster  = new Map();    // email -> { ptin, first, last }
  const ptinToEmail    = new Map();    // ptin  -> email (best-guess backfill)
  const nameToEmail    = new Map();    // "first last" -> email (best-guess backfill)

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
    if (email && !emailToMaster.has(email)) emailToMaster.set(email, { ptin, first, last });

    if (reported) {
      if (email) reportedEmails.add(email);
      if (ptin)  reportedPtins.add(ptin);
      if (nkey)  reportedNames.add(nkey);
    }
  }

  // --- Map Roster headers (PTIN is now OPTIONAL) ---
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
    if (!rPtin && rEmail && emailToMaster.has(rEmail)) {
      const snap = emailToMaster.get(rEmail);
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

    // Paint entire row
    colors[r] = new Array(lastCol).fill(rowColor);
  }

  if (wroteBack) dataRange.setValues(vals);
  dataRange.setBackgrounds(colors);
  if (!quiet) toast_('Roster highlighting updated from Master (email-first matching).');
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