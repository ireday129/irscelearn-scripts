/**
 * Highlight Roster rows (entire row) for anyone who has ANY row marked Reported?=TRUE in "Master".
 * - Matches by PTIN (normalized to P0#######).
 * - Uses a soft yellow highlight for the whole row.
 * - If Roster.Valid? is TRUE for that row, clear the highlight (no background).
 */
function highlightRosterFromReportedHours(quiet) {
  const ss = SpreadsheetApp.getActive();
  const roster = mustGet_(ss, CFG.SHEET_ROSTER);
  const master = mustGet_(ss, CFG.SHEET_MASTER);

  // --- Load Master and build set of reported PTINs ---
  const mVals = master.getDataRange().getValues();
  if (mVals.length <= 1) { if (!quiet) toast_('Master is empty; nothing to highlight.'); return; }

  const mHdr = normalizeHeaderRow_(mVals[0]);
  const mm   = mapHeaders_(mHdr);

  if (mm.ptin == null || mm.reportedCol == null) {
    if (!quiet) toast_('Master missing PTIN and/or Reported? column.', true);
    return;
  }

  const reportedPtins = new Set();
  const mBody = mVals.slice(1);
  for (let i = 0; i < mBody.length; i++) {
    const row = mBody[i];
    const ptin = formatPtinP0_(row[mm.ptin] || '');
    const isReported = truthy_(row[mm.reportedCol]);
    if (ptin && isReported) reportedPtins.add(ptin);
  }

  if (!reportedPtins.size) { if (!quiet) toast_('No Reported?=TRUE rows in Master; nothing to highlight.'); return; }

  // --- Map Roster headers ---
  const rMap = mapRosterHeaders_(roster);
  if (!rMap) { if (!quiet) toast_('Roster header mapping failed.', true); return; }
  if (rMap.ptin < 0) { if (!quiet) toast_('Roster missing PTIN column.', true); return; }
  // rMap.valid may be -1 if the column is missing, handle gracefully
  const hasValidCol = typeof rMap.valid === 'number' && rMap.valid >= 0;

  const lastRow = roster.getLastRow();
  const lastCol = roster.getLastColumn();
  if (lastRow < 2) { if (!quiet) toast_('Roster has no data to highlight.'); return; }

  const dataRange = roster.getRange(2, 1, lastRow - 1, lastCol);
  const vals = dataRange.getValues();
  const colors = dataRange.getBackgrounds();

  const YELLOW = '#fff59d';

  for (let r = 0; r < vals.length; r++) {
    const row = vals[r];
    const ptin  = formatPtinP0_(row[rMap.ptin] || '');
    const valid = hasValidCol ? truthy_(row[rMap.valid]) : false;

    let rowColor = null; // null clears to default background
    if (!valid && ptin && reportedPtins.has(ptin)) {
      rowColor = YELLOW;
    }

    const newRowColors = new Array(lastCol).fill(rowColor);
    colors[r] = newRowColors;
  }

  dataRange.setBackgrounds(colors);
  if (!quiet) toast_('Roster highlighting updated from Master.Reported?.');
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
