/**
 * Ensure every row in "Reported Hours" is reflected in "Master"
 * for date-related fields.
 *
 * Matching key: Attendee PTIN + Program Number
 *
 * HARD-CODED COLUMN MAPPING (0-based indexes):
 *
 *  Reported Hours sheet:
 *    - Attendee PTIN      = column C (index 2)
 *    - Program Number     = column D (index 3)
 *    - Program Completion = column F (index 5)
 *    - Date Reported      = column G (index 6)
 *
 *  Master sheet:
 *    - Attendee PTIN      = column C (index 2)
 *    - Program Number     = column D (index 3)
 *    - Program Completion = column G (index 6)
 *    - Reported At        = column K (index 10)
 *
 * Behavior:
 *  - If a PTIN+Program combo exists in Reported Hours,
 *    FORCE-OVERWRITE on Master:
 *      * Program Completion Date
 *      * Reported At
 *  - Does NOT touch any other columns (Reported?, issues, last updated, names, etc.)
 */
function syncMasterWithReportedHours(quiet) {
  const ss     = SpreadsheetApp.getActive();
  const master = mustGet_(ss, CFG.SHEET_MASTER);
  const rh     = mustGet_(ss, 'Reported Hours');

  // Optional: sweep duplicates in RH first (if this helper exists)
  try {
    if (typeof enforceReportedHoursUniqueness_ === 'function') {
      enforceReportedHoursUniqueness_();
    }
  } catch (e) {
    Logger.log(e);
  }

  const rhVals = rh.getDataRange().getValues();
  if (rhVals.length <= 1) {
    if (!quiet) toast_('Reported Hours is empty; nothing to sync.');
    return;
  }

  // --- Hard-coded column indexes for Reported Hours (0-based) ---
  const RH_COL_PTIN          = 2; // C - Attendee PTIN
  const RH_COL_PROGRAM       = 3; // D - Program Number
  const RH_COL_COMPLETION    = 5; // F - Program Completion Date
  const RH_COL_DATE_REPORTED = 6; // G - Date Reported

  // Build a map from Reported Hours: key = prog|ptin -> { comp, dRep }
  const rhMap = new Map();
  for (let r = 1; r < rhVals.length; r++) {
    const row = rhVals[r];
    const ptinRaw = row[RH_COL_PTIN];
    const progRaw = row[RH_COL_PROGRAM];

    // Use raw, trimmed strings for matching instead of any normalizers
    const ptin = String(ptinRaw || '').trim();
    const prog = String(progRaw || '').trim();
    if (!prog || !ptin) continue;

    const comp = row[RH_COL_COMPLETION]    || '';
    const dRep = row[RH_COL_DATE_REPORTED] || '';

    rhMap.set(prog + '|' + ptin, { comp, dRep });
  }

  const mVals = master.getDataRange().getValues();
  if (mVals.length <= 1) {
    if (!quiet) toast_('Master is empty; cannot sync.', true);
    return;
  }

  // --- Hard-coded column indexes for Master (0-based) ---
  const MASTER_COL_PTIN          = 2;  // C - Attendee PTIN
  const MASTER_COL_PROGRAM       = 3;  // D - Program Number
  const MASTER_COL_COMPLETION    = 6;  // G - Program Completion Date
  const MASTER_COL_REPORTED_AT   = 10; // K - Reported At

  const mBody   = mVals.slice(1);
  const numCols = mVals[0].length;
  let updates   = 0;

  // Walk Master rows and, when a PTIN+Program combo exists in RH, override only:
  // - Program Completion Date (col G)
  // - Reported At (col K)
  for (let r = 0; r < mBody.length; r++) {
    const row = mBody[r];

    const ptinRaw = row[MASTER_COL_PTIN];
    const progRaw = row[MASTER_COL_PROGRAM];

    // Use raw, trimmed strings for matching
    const ptin = String(ptinRaw || '').trim();
    const prog = String(progRaw || '').trim();
    if (!prog || !ptin) continue;

    const key  = prog + '|' + ptin;
    const from = rhMap.get(key);
    if (!from) continue;

    const comp = from.comp;
    const dRep = from.dRep;

    let changed = false;

    // FORCE Program Completion Date from Reported Hours when RH has any non-empty value
    if (comp !== null && comp !== '') {
      row[MASTER_COL_COMPLETION] = parseDate_(comp) || comp;
      changed = true;
    }

    // FORCE Reported At from Reported Hours (or now) regardless of current Master value
    const repVal = dRep || new Date();
    row[MASTER_COL_REPORTED_AT] =
      (repVal instanceof Date ? repVal : parseDate_(repVal) || new Date());
    changed = true;

    if (changed) updates++;
  }

  // Write back all Master rows (only the 2 targeted columns are effectively changed)
  if (updates > 0) {
    master.getRange(2, 1, mBody.length, numCols).setValues(mBody);
  }

  if (!quiet) {
    toast_('Reported Hours → Master date repair: ' + updates + ' Master rows updated (Program Completion Date + Reported At).');
  }
}