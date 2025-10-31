/** roster_highlight.js
 * Highlight Roster rows yellow if that attendee has ANY Reported?=TRUE on Master.
 * - Match by Email (case-insensitive) OR PTIN (normalized P0#######).
 * - If Roster.Valid? is checked/TRUE, force background WHITE (no highlight).
 * - Non-destructive backfill: if Roster email/PTIN is blank and Master has it, fill it.
 * - Requires utils helpers if available; includes light fallbacks.
 */

function highlightRosterFromReportedHours() {
  const ss = SpreadsheetApp.getActive();
  const master = mustGet_(ss, CFG.SHEET_MASTER);
  const roster = mustGet_(ss, CFG.SHEET_ROSTER);

  // ---- Helpers / fallbacks ----
  const truthy = (typeof truthy_ === 'function')
    ? truthy_
    : (v => {
        if (typeof v === 'boolean') return v;
        if (v == null) return false;
        const s = String(v).trim().toLowerCase();
        return s === 'true' || s === 'yes' || s === 'y' || s === '1' || s === '✓';
      });

  const normEmail = e => String(e || '').trim().toLowerCase();
  const normPtin  = v => (typeof formatPtinP0_ === 'function')
    ? formatPtinP0_(v || '')
    : String(v || '').toUpperCase().replace(/[^0-9P]/g, '').replace(/^P(?!0)/, 'P0');

  const normHeaders = arr => (typeof normalizeHeaderRow_ === 'function')
    ? normalizeHeaderRow_(arr)
    : arr.map(v => String(v || '').trim().replace(/\s+/g, ' '));

  const mapMaster = hdr => (typeof mapHeaders_ === 'function')
    ? mapHeaders_(hdr)
    : (() => {
        const lower = hdr.map(h => String(h || '').toLowerCase().trim());
        const find = label => lower.indexOf(String(label || '').toLowerCase().trim());
        return {
          firstName: find('attendee first name'),
          lastName:  find('attendee last name'),
          ptin:      find('ptin'),
          email:     find('email'),
          program:   find('program number'),
          hours:     find('ce hours'),
          completion:find('program completion date'),
          masterIssueCol: find('reporting issue?'),
          reportedCol:    find('reported?'),
          reportedAtCol:  find('date reported')
        };
      })();

  const mapRoster = sh => {
    if (typeof mapRosterHeaders_ === 'function') return mapRosterHeaders_(sh);
    const hdr = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(s => String(s || '').trim());
    const lower = hdr.map(h => h.toLowerCase());
    const find = (names) => {
      for (const n of (Array.isArray(names) ? names : [names])) {
        const i = lower.indexOf(String(n).toLowerCase());
        if (i >= 0) return i;
      }
      return -1;
    };
    return {
      first: find(['attendee first name','first name']),
      last:  find(['attendee last name','last name']),
      ptin:  find(['attendee ptin','ptin']),
      email: find(['email','e-mail']),
      valid: find(['valid?','valid']),
      group: find(['group']),
      hdr
    };
  };

  // ---- Build reported sets & backfill maps from MASTER ----
  const mVals = master.getDataRange().getValues();
  if (mVals.length <= 1) { toast_('Master is empty; nothing to highlight.', true); return; }

  const mHdr = normHeaders(mVals[0]);
  const mm   = mapMaster(mHdr);
  if (mm.reportedCol == null || (mm.email == null && mm.ptin == null)) {
    toast_('Master is missing Reported? and/or key columns (Email/PTIN).', true);
    return;
  }

  const reportedEmails = new Set();  // lowercased
  const reportedPtins  = new Set();  // normalized
  const emailToPtin    = new Map();  // for Roster backfill
  const ptinToEmail    = new Map();  // for Roster backfill

  for (let r = 1; r < mVals.length; r++) {
    const row = mVals[r];
    const isReported = truthy(row[mm.reportedCol]);
    const em = (mm.email != null) ? normEmail(row[mm.email]) : '';
    const pt = (mm.ptin  != null) ? normPtin(row[mm.ptin])   : '';

    // Build cross maps for backfill even if not reported
    if (em && pt && !ptinToEmail.has(pt)) ptinToEmail.set(pt, em);
    if (em && pt && !emailToPtin.has(em)) emailToPtin.set(em, pt);

    if (!isReported) continue;
    if (em) reportedEmails.add(em);
    if (pt) reportedPtins.add(pt);
  }

  // ---- Process ROSTER rows ----
  const rVals = roster.getDataRange().getValues();
  if (rVals.length <= 1) { toast_('Roster is empty.', true); return; }

  const rMap  = mapRoster(roster);
  const iF = rMap.first, iL = rMap.last, iP = rMap.ptin, iE = rMap.email, iV = rMap.valid;

  if ([iF, iL].some(i => i < 0) || (iE < 0 && iP < 0)) {
    toast_('Roster missing First/Last and Email/PTIN columns.', true);
    return;
  }

  const body      = rVals.slice(1);
  const height    = body.length;
  const width     = rVals[0].length;
  const startRow  = 2;
  const startCol  = 1;

  // Prepare backgrounds array (rows x cols)
  const bgRange  = roster.getRange(startRow, startCol, height, width);
  const bgs      = bgRange.getBackgrounds();
  let valueChanges = 0, bgChanges = 0;

  for (let r = 0; r < body.length; r++) {
    const row = body[r];

    // Current keys
    let em = (iE >= 0) ? normEmail(row[iE]) : '';
    let pt = (iP >= 0) ? normPtin(row[iP])  : '';

    // Gentle backfill from Master maps when blank
    if (!em && pt && ptinToEmail.has(pt)) {
      em = ptinToEmail.get(pt);
      if (iE >= 0) { row[iE] = em; valueChanges++; }
    }
    if (!pt && em && emailToPtin.has(em)) {
      pt = emailToPtin.get(em);
      if (iP >= 0) { row[iP] = pt; valueChanges++; }
    }

    // Evaluate state
    const isValid = (iV >= 0) ? truthy(row[iV]) : false;
    const hasReported = (!!em && reportedEmails.has(em)) || (!!pt && reportedPtins.has(pt));

    // Decide background color
    const want = isValid ? '#ffffff' : (hasReported ? '#fff8b1' : '#ffffff');

    // Apply background across the entire row (visual clarity)
    for (let c = 0; c < width; c++) {
      if (bgs[r][c] !== want) {
        bgs[r][c] = want;
        bgChanges++;
      }
    }
  }

  // Write updates (values first, then backgrounds)
  if (valueChanges) {
    roster.getRange(startRow, startCol, height, width).setValues(body);
  }
  if (bgChanges) {
    bgRange.setBackgrounds(bgs);
  }

  toast_(
    `Roster highlight complete: ${bgChanges ? 'backgrounds updated' : 'no highlight changes'}; ` +
    `${valueChanges ? 'filled some blanks from Master.' : 'no backfill needed.'}`
  );
}