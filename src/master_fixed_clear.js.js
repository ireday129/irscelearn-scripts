/**
 * Global function to process Master sheet edit side-effects.
 * This function MUST be run by an Installable Trigger set to 'On edit'.
 */
function processMasterEdits(e){
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    const col = e.range.getColumn();
    
    // --- Setup: Defensive Sheet Name Check ---
    const masterSheetName = String(CFG.SHEET_MASTER || 'Master').trim().toLowerCase();
    if (sh.getName().trim().toLowerCase() !== masterSheetName) return;

    // --- Core Logic ---
    const newVal = e.value;
    
    // NOTE: Relying on mapHeaders_ being globally available from utils.js.gs
    const mHdr = normalizeHeaderRow_(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);
    const mMap = mapHeaders_(mHdr);
    
    // Find the column indexes: Reporting Issue? is mMap.masterIssueCol.
    const ISSUE_COL_INDEX = mMap.masterIssueCol != null ? mMap.masterIssueCol + 1 : 10; // Fallback to J=10
    const REPORTED_AT_COL_INDEX = mMap.reportedAtCol != null ? mMap.reportedAtCol + 1 : null;

    // --- 3. Roster Valid? Check (Placeholder for other onEdit logic) ---
    // If you have any other onEdit logic (like for the Roster sheet) 
    // it must be merged into this file here.
    if (sh.getName() === CFG.SHEET_ROSTER) {
      // Logic for Roster onEdit would go here...
    }

  } catch (err) {
    Logger.log('processMasterEdits error: ' + err.message + ' Stack: ' + err.stack);
  }
}
/** Normalize Master PTINs and flag obvious invalids (P00000000). */
function normalizeAndFlagMasterPtins_(quiet) {
  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName(CFG.SHEET_MASTER);
  if (!master) { if (!quiet) toast_(`Sheet "${CFG.SHEET_MASTER}" not found.`, true); return; }

  const mVals = master.getDataRange().getValues();
  if (mVals.length <= 1) return;

  const hdr = normalizeHeaderRow_(mVals[0]);
  const mm = mapHeaders_(hdr);
  if (mm.ptin == null || mm.masterIssueCol == null) {
    if (!quiet) toast_('Master is missing PTIN and/or Reporting Issue? columns.', true);
    return;
  }

  const body = mVals.slice(1);
  let changed = 0;

  for (let i = 0; i < body.length; i++) {
    const row = body[i];

    // --- normalize PTIN ---
    let raw = String(row[mm.ptin] || '').trim();
    if (!raw) continue;

    // Fix "PO" (letter O) to "P0" (zero) at the start, case-insensitive
    if (/^pO/i.test(raw)) raw = 'P0' + raw.slice(2);

    // Use existing normalizer to force P0####### format when possible
    const normalized = formatPtinP0_(raw);

    // Write back normalized value if changed
    if (normalized !== row[mm.ptin]) {
      row[mm.ptin] = normalized;
      changed++;
    }

    // --- flag invalid all-zeros PTIN ---
    if (normalized === 'P00000000') {
      const curIssue = String(row[mm.masterIssueCol] || '').trim();
      // Don’t stomp 'Fixed' if that’s your workflow; otherwise set the issue.
      if (!/^fixed$/i.test(curIssue)) {
        if (curIssue !== 'PTIN does not exist') {
          row[mm.masterIssueCol] = 'PTIN does not exist';
          changed++;
        }
      }
    }
  }

  if (changed) {
    master.getRange(2, 1, body.length, hdr.length).setValues(body);
    if (!quiet) toast_(`Master PTINs normalized & invalids flagged (${changed} cell change${changed===1?'':'s'}).`);
  }
}