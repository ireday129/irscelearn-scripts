/**
 * Global function to process Master sheet fixed status clear.
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
    
    // Find the column indexes: Fixed? is Col 12. Reported? is Col 9. Reporting Issue? is mMap.masterIssueCol.
    const FIXED_COL_INDEX = 12;      // Column L
    const REPORTED_COL_INDEX = 9;    // Column I (Based on standard Master layout)
    const ISSUE_COL_INDEX = mMap.masterIssueCol != null ? mMap.masterIssueCol + 1 : 10; // Fallback to J=10

    
    // 1. FIXED? Check (Column L = 12): If FIXED is checked, clear Reporting Issue? (Column J)
    if (col === FIXED_COL_INDEX && parseBool_(newVal)) {
        // Clear the Reporting Issue column in the same row
        sh.getRange(e.range.getRow(), ISSUE_COL_INDEX).setValue('');
        toast_('Master issue cleared by checking Fixed?.', false);
    }
    
    // 2. REPORTED? Check (Column I = 9): If REPORTED is checked, set Fixed? (Column L) to FALSE
    if (col === REPORTED_COL_INDEX && parseBool_(newVal)) {
        // Set the Fixed? column in the same row to FALSE
        sh.getRange(e.range.getRow(), FIXED_COL_INDEX).setValue(false);
        toast_('Fixed? status cleared by marking row Reported.', false);
    }

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