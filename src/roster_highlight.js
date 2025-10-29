/**
 * Highlight Roster rows (entire row) for anyone found in "Reported Hours".
 * - Matches by PTIN (normalized to P0#######).
 * - Uses a soft yellow highlight for the whole row.
 * - Does NOT change "Valid?"—it’s purely visual so you can manually review.
 */
function highlightRosterFromReportedHours(quiet) {
  const ss = SpreadsheetApp.getActive();
  const roster = mustGet_(ss, CFG.SHEET_ROSTER);
  const rh     = mustGet_(ss, 'Reported Hours');

  // --- Load Reported Hours ---
  const rhVals = rh.getDataRange().getValues();
  if (rhVals.length <= 1) { if (!quiet) toast_('Reported Hours is empty; nothing to highlight.'); return; }
  const rhHdr  = rhVals[0].map(s => String(s || '').trim().toLowerCase());
  const iRhPT  = Math.max(rhHdr.indexOf('ptin'), rhHdr.indexOf('attendee ptin'));
  if (iRhPT < 0) { if (!quiet) toast_('Reported Hours missing PTIN column.', true); return; }

  // Build a set of PTINs that have any reported row
  const reportedPtins = new Set();
  for (let r = 1; r < rhVals.length; r++) {
    const ptin = formatPtinP0_(rhVals[r][iRhPT] || '');
    if (ptin) reportedPtins.add(ptin);
  }
  if (!reportedPtins.size) { if (!quiet) toast_('No PTINs found in Reported Hours.', true); return; }

  // --- Map Roster headers ---
  const rMap = mapRosterHeaders_(roster);
  if (!rMap) { if (!quiet) toast_('Roster header mapping failed.', true); return; }

  const lastRow = roster.getLastRow();
  const lastCol = roster.getLastColumn();
  if (lastRow < 2) { if (!quiet) toast_('Roster has no data to highlight.'); return; }

  const dataRange = roster.getRange(2, 1, lastRow - 1, lastCol);
  const vals = dataRange.getValues();
  const colors = dataRange.getBackgrounds();

  // Soft yellow
  const YELLOW = '#fff59d';

  for (let i = 0; i < vals.length; i++) {
    const row = vals[i];
    const ptin = formatPtinP0_(row[rMap.ptin] || '');
    const shouldHighlight = ptin && reportedPtins.has(ptin);

    // Paint or clear the entire row
    const newRowColors = new Array(lastCol).fill(shouldHighlight ? YELLOW : null);
    colors[i] = newRowColors;
  }

  dataRange.setBackgrounds(colors);
  if (!quiet) toast_('Roster highlighting updated from Reported Hours.');
}

/** Menu-safe wrapper */
function highlightRosterFromReportedHoursMenu() {
  try {
    highlightRosterFromReportedHours(true);
    toast_('Roster highlighting updated from Reported Hours.');
  } catch (e) {
    toast_('Failed to update roster highlighting: ' + e.message, true);
    Logger.log(e.stack || e);
  }
}
