/* global SpreadsheetApp */
/* global mustGet_, normalizeHeaderRow_, mapHeaders_, formatPtinP0_, mapRosterHeaders_, parseBool_ */
/* global CFG, toast_ */

/**
 * Backfill Roster from Master for rows where Master.Reported? == TRUE.
 * - Columns synced: First Name, Last Name, Email, PTIN
 * - Updates existing roster rows (match by PTIN preferred, fallback Email)
 * - Appends new roster rows if no match
 * - Highlights affected roster rows yellow
 * - Does NOT change Roster "Valid?" column
 */
function backfillRosterFromMasterReported_() {
  const ss = SpreadsheetApp.getActive();

  // Sheets
  const shMaster = mustGet_(ss, CFG.SHEET_MASTER);
  const shRoster = mustGet_(ss, CFG.SHEET_ROSTER);

  // ----- MASTER: read & map
  const mVals = shMaster.getDataRange().getValues();
  if (mVals.length <= 1) { toast_('Master is empty; nothing to backfill.', true); return; }
  const mHdr = normalizeHeaderRow_(mVals[0]);
  const mm   = mapHeaders_(mHdr);

  // sanity: required cols
  const reqMaster = ['firstName','lastName','email','ptin','reportedCol'];
  const missM = reqMaster.filter(k => mm[k] == null);
  if (missM.length) { toast_('Master missing columns: ' + missM.join(', '), true); return; }

  // Build list of Master rows to use (Reported? == TRUE)
  const masterBody = mVals.slice(1);
  const masterRows = [];
  for (let i = 0; i < masterBody.length; i++) {
    const row = masterBody[i];
    if (!parseBool_(row[mm.reportedCol])) continue; // only reported
    const first = String(row[mm.firstName] || '').trim();
    const last  = String(row[mm.lastName]  || '').trim();
    const email = String(row[mm.email]     || '').trim().toLowerCase();
    const ptin  = formatPtinP0_(row[mm.ptin] || '');
    // Skip if we have neither PTIN nor Email
    if (!ptin && !email) continue;

    masterRows.push({ first, last, email, ptin });
  }
  if (!masterRows.length) { toast_('No Reported rows found on Master.', true); return; }

  // ----- ROSTER: read & map
  const rVals = shRoster.getDataRange().getValues();
  if (rVals.length < 1) { toast_('Roster has no header; cannot backfill.', true); return; }

  const rMap = mapRosterHeaders_(shRoster);
  if (!rMap) { toast_('Roster headers not recognized; ensure First/Last/Email/PTIN exist.', true); return; }

  const rHdr   = rMap.hdr;
  const rBody  = rVals.length > 1 ? rVals.slice(1) : [];
  const lastCol= shRoster.getLastColumn();

  // Build indexes for quick match
  const idxByPtin  = new Map();
  const idxByEmail = new Map();

  for (let i = 0; i < rBody.length; i++) {
    const row = rBody[i];
    const pt  = formatPtinP0_(row[rMap.ptin] || '');
    const em  = String(row[rMap.email] || '').trim().toLowerCase();
    if (pt)  idxByPtin.set(pt, i);
    if (em)  idxByEmail.set(em, i);
  }

  // ----- Apply updates + collect appends
  const YELLOW = '#fff59d';
  const rowsToHighlight = new Set(); // 0-based within body

  const toAppend = []; // arrays matching roster width
  let updates = 0, appends = 0;

  masterRows.forEach(entry => {
    let hitIndex = -1;

    if (entry.ptin && idxByPtin.has(entry.ptin)) {
      hitIndex = idxByPtin.get(entry.ptin);
    } else if (entry.email && idxByEmail.has(entry.email)) {
      hitIndex = idxByEmail.get(entry.email);
    }

    if (hitIndex >= 0) {
      // Update existing row (fill blanks only to avoid clobbering)
      const row = rBody[hitIndex];
      if (rMap.first >= 0 && !String(row[rMap.first]||'').trim() && entry.first) row[rMap.first] = entry.first;
      if (rMap.last  >= 0 && !String(row[rMap.last] ||'').trim() && entry.last)  row[rMap.last]  = entry.last;
      if (rMap.email >= 0 && !String(row[rMap.email]||'').trim() && entry.email) row[rMap.email] = entry.email;
      if (rMap.ptin  >= 0 && !String(row[rMap.ptin] ||'').trim() && entry.ptin)  row[rMap.ptin]  = entry.ptin;

      updates++;
      rowsToHighlight.add(hitIndex);
    } else {
      // Append new roster row
      const arr = new Array(lastCol).fill('');
      if (rMap.first >= 0) arr[rMap.first] = entry.first;
      if (rMap.last  >= 0) arr[rMap.last]  = entry.last;
      if (rMap.email >= 0) arr[rMap.email] = entry.email;
      if (rMap.ptin  >= 0) arr[rMap.ptin]  = entry.ptin;

      toAppend.push(arr);
      appends++;
    }
  });

  // Write updates back (if any)
  if (rBody.length && (updates > 0)) {
    shRoster.getRange(2, 1, rBody.length, lastCol).setValues(rBody);
  }

  // Append new rows (if any)
  let appendStartRow = null;
  if (toAppend.length) {
    appendStartRow = shRoster.getLastRow() + 1;
    shRoster.getRange(appendStartRow, 1, toAppend.length, lastCol).setValues(toAppend);
    // Mark appended rows for highlight as well
    for (let i = 0; i < toAppend.length; i++) {
      // convert appended rows to 0-based index within "body" space:
      rowsToHighlight.add(rBody.length + i);
    }
  }

  // Highlight all affected rows yellow
  // We’ll compute row numbers on the sheet: header is row 1, body starts row 2.
  if (rowsToHighlight.size) {
    const maxRow = shRoster.getLastRow();
    const ranges = [];
    rowsToHighlight.forEach(i => {
      const rowNumber = 2 + i; // offset for header
      if (rowNumber >= 2 && rowNumber <= maxRow) {
        ranges.push(shRoster.getRange(rowNumber, 1, 1, lastCol));
      }
    });
    // Apply backgrounds (looping is OK for a handful; for hundreds, we could merge ranges)
    ranges.forEach(r => r.setBackground(YELLOW));
  }

  toast_(`Backfill Roster from Master complete: ${updates} updated, ${appends} appended. Highlighted ${rowsToHighlight.size} row(s).`);
}