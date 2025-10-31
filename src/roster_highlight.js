/** roster_highlight.js
 * Highlights Roster rows yellow if that attendee has ANY reported hours.
 * Sources of truth:
 *   1) MASTER: any row where Reported? == TRUE (match by Email and/or Attendee PTIN)
 *   2) Reported Hours sheet: any ledger row for that Email/PTIN (fallback)
 *
 * If Roster.Valid? is TRUE, force white (no highlight).
 * Backfills missing PTIN/Email on Roster from Master cross-maps when possible.
 *
 * This version is resilient to header variants BUT explicitly prefers exact
 *   "Attendee First Name", "Attendee Last Name", "Attendee PTIN"
 * across MASTER, ROSTER, and REPORTED HOURS when present.
 */

function highlightRosterFromReportedHours() {
  const ss = SpreadsheetApp.getActive();
  const master = mustGet_(ss, CFG.SHEET_MASTER);
  const roster = mustGet_(ss, CFG.SHEET_ROSTER);

  // -------- helpers --------
  const truthy = (typeof truthy_ === 'function')
    ? truthy_
    : (v) => {
        if (typeof v === 'boolean') return v;
        if (v == null) return false;
        const s = String(v).trim().toLowerCase();
        return s === 'true' || s === 'yes' || s === 'y' || s === '1' || s === '✓';
      };

  const normEmail = (e) => String(e || '').trim().toLowerCase();

  const normPtin = (v) => {
    if (typeof formatPtinP0_ === 'function') return formatPtinP0_(v);
    // Fallback: P0 + last 7 digits
    const digits = (String(v||'').match(/\d+/g) || []).join('');
    if (digits.length >= 7) return 'P0' + digits.slice(-7);
    const s = String(v||'').toUpperCase().replace(/[^0-9P]/g,'');
    if (/^P0\d{7}$/.test(s)) return s;
    if (/^P\d{7}$/.test(s)) return 'P0' + s.slice(1);
    return '';
  };

  const normHeaders = (arr) =>
    (typeof normalizeHeaderRow_ === 'function')
      ? normalizeHeaderRow_(arr)
      : arr.map((v)=>String(v||'').trim().replace(/\s+/g,' '));

  const idxExact = (hdr, label) => {
    const h = hdr.map(x => String(x||'').trim());
    const i = h.indexOf(label);
    return i;
  };

  const idxFlex = (hdr, candidates) => {
    const low = hdr.map(h => String(h||'').toLowerCase().trim());
    const strip = s => s.replace(/[^a-z0-9]/g,'');
    const low2 = low.map(strip);
    const list = Array.isArray(candidates) ? candidates : [candidates];
    for (const c of list) {
      const cl = String(c).toLowerCase().trim();
      const cs = strip(cl);
      let i = low.indexOf(cl); if (i>=0) return i;
      i = low2.indexOf(cs);    if (i>=0) return i;
    }
    return -1;
  };

  // Prefer the EXACT canonical column names if present; otherwise fall back.
  const preferCanonical = (hdr, canonical, fallbackList) => {
    const ix = idxExact(hdr, canonical);
    if (ix >= 0) return ix;
    return idxFlex(hdr, fallbackList);
  };

  // -------- MASTER: collect reported sets + cross maps --------
  const reportedEmails = new Set();
  const reportedPtins  = new Set();
  const emailToPtin    = new Map(); // backfill maps
  const ptinToEmail    = new Map();

  (function buildFromMaster(){
    const mVals = master.getDataRange().getValues();
    if (mVals.length <= 1) return;
    const mHdr = normHeaders(mVals[0]);

    const iFirst = preferCanonical(mHdr, 'Attendee First Name', ['first name','attendee first name']);
    const iLast  = preferCanonical(mHdr, 'Attendee Last Name',  ['last name','attendee last name']);
    const iPtin  = preferCanonical(mHdr, 'Attendee PTIN',       ['ptin','attendee ptin']);
    const iEmail = preferCanonical(mHdr, 'Email',                ['email','e-mail','attendee email','email address']);
    const iRep   = preferCanonical(mHdr, 'Reported?',            ['reported?','reported','is reported']);

    if (iPtin < 0 && iEmail < 0) {
      toast_('MASTER missing Email/PTIN columns; cannot highlight.', true);
      return;
    }

    for (let r=1; r<mVals.length; r++){
      const row = mVals[r];
      const em = (iEmail>=0) ? normEmail(row[iEmail]) : '';
      const pt = (iPtin >=0) ? normPtin(row[iPtin])  : '';

      if (em && pt) {
        if (!emailToPtin.has(em)) emailToPtin.set(em, pt);
        if (!ptinToEmail.has(pt)) ptinToEmail.set(pt, em);
      }

      const rep = (iRep>=0) ? truthy(row[iRep]) : false;
      if (!rep) continue;
      if (em) reportedEmails.add(em);
      if (pt) reportedPtins.add(pt);
    }
  })();

  // -------- Reported Hours: merge sets (fallback) --------
  (function mergeFromReportedHours(){
    const sh = ss.getSheetByName('Reported Hours');
    if (!sh || sh.getLastRow() < 2) return;
    const vals = sh.getDataRange().getValues();
    const hdr  = normHeaders(vals[0]);

    const iEmail = preferCanonical(hdr, 'Email', ['email','attendee email','email address']);
    const iPtin  = preferCanonical(hdr, 'PTIN',  ['ptin','attendee ptin']);

    if (iEmail < 0 && iPtin < 0) return;

    for (let r=1; r<vals.length; r++){
      const row = vals[r];
      const em = (iEmail>=0) ? normEmail(row[iEmail]) : '';
      const pt = (iPtin >=0) ? normPtin(row[iPtin])  : '';
      if (em) reportedEmails.add(em);
      if (pt) reportedPtins.add(pt);
      if (em && pt) {
        if (!emailToPtin.has(em)) emailToPtin.set(em, pt);
        if (!ptinToEmail.has(pt)) ptinToEmail.set(pt, em);
      }
    }
  })();

  // -------- ROSTER: backfill + color --------
  const rVals = roster.getDataRange().getValues();
  if (rVals.length <= 1) { toast_('Roster is empty.', true); return; }

  const rHdr = normHeaders(rVals[0]);
  const iFirst = preferCanonical(rHdr, 'Attendee First Name', ['first name','attendee first name']);
  const iLast  = preferCanonical(rHdr, 'Attendee Last Name',  ['last name','attendee last name']);
  const iEmail = preferCanonical(rHdr, 'Email',               ['email','e-mail','email address']);
  const iPtin  = preferCanonical(rHdr, 'Attendee PTIN',       ['ptin','attendee ptin']);
  const iValid = preferCanonical(rHdr, 'Valid?',              ['valid?','valid','is valid']);

  if ((iEmail < 0 && iPtin < 0) || iFirst < 0 || iLast < 0) {
    toast_('Roster missing First/Last and Email/PTIN columns.', true);
    return;
  }

  const body     = rVals.slice(1);
  const height   = body.length;
  const width    = rVals[0].length;
  const startRow = 2;
  const startCol = 1;

  const bgRange = roster.getRange(startRow, startCol, height, width);
  const bgs     = bgRange.getBackgrounds();

  let filled = 0, painted = 0, cleared = 0;

  for (let r=0; r<body.length; r++){
    const row = body[r];

    // Get current keys & backfill if missing (use canonical preference)
    let em = (iEmail>=0) ? normEmail(row[iEmail]) : '';
    let pt = (iPtin >=0) ? normPtin(row[iPtin])  : '';

    if (!em && pt && ptinToEmail.has(pt)) { em = ptinToEmail.get(pt); if (iEmail>=0) { row[iEmail] = em; filled++; } }
    if (!pt && em && emailToPtin.has(em)) { pt = emailToPtin.get(em); if (iPtin >=0) { row[iPtin]  = pt; filled++; } }

    const isValid = (iValid>=0) ? truthy(row[iValid]) : false;

    // Decide highlight from combined sets (works even if only PTIN is shared)
    const hasReported = (!isValid) && ((em && reportedEmails.has(em)) || (pt && reportedPtins.has(pt)));
    const want = hasReported ? '#fff8b1' : '#ffffff';

    for (let c=0; c<width; c++){
      if (bgs[r][c] !== want) {
        bgs[r][c] = want;
        if (hasReported) painted++; else cleared++;
      }
    }
  }

  if (filled) roster.getRange(startRow, startCol, height, width).setValues(body);
  if (painted || cleared) bgRange.setBackgrounds(bgs);

  toast_(`Roster highlight: backfilled=${filled}, yellowed=${painted>0?painted:0}, cleared=${cleared>0?cleared:0}.`);
}

/** Menu shim */
function highlightRosterFromReportedHoursMenu() {
  highlightRosterFromReportedHours();
}

/** Debug a single roster row by email or PTIN (paste value). */
function debugRosterMatchOne_(value) {
  const ss = SpreadsheetApp.getActive();
  const roster = mustGet_(ss, CFG.SHEET_ROSTER);
  const v = String(value||'').trim();
  const isPt = /^P0?\d{7}$/i.test(v) || /\d{7}$/.test(v);
  const keyPt = isPt ? (typeof formatPtinP0_==='function'?formatPtinP0_(v):('P0'+(v.replace(/\D/g,'').slice(-7)))) : '';
  const keyEm = isPt ? '' : v.toLowerCase();

  const rows = roster.getDataRange().getValues();
  const hdr  = rows[0].map(s=>String(s||'').trim());
  const iE = preferCanonical(hdr, 'Email', ['email','attendee email','email address']);
  const iP = preferCanonical(hdr, 'Attendee PTIN', ['ptin','attendee ptin']);
  for (let r=1; r<rows.length; r++){
    const em = iE>=0 ? String(rows[r][iE]||'').toLowerCase().trim() : '';
    const pt = iP>=0 ? String(rows[r][iP]||'').toUpperCase().trim() : '';
    if ((keyEm && em===keyEm) || (keyPt && pt===keyPt)) {
      Logger.log('Roster row %s matches value=%s (em=%s, pt=%s)', r+1, value, em, pt);
      return;
    }
  }
  Logger.log('No roster row matched value=%s (normalized PTIN=%s).', value, keyPt);
}