/** ROSTER HELPERS + onEdit (override-style dedupe, supports Group) **/

function mapRosterHeaders_(sh){
  if (!sh || sh.getLastRow() === 0) return null;
  const hdr = sh.getRange(1,1,1, sh.getLastColumn()).getValues()[0].map(s=>String(s||'').trim());
  const lower = hdr.map(h => h.toLowerCase());

  function find(names){
    for (const n of (Array.isArray(names)?names:[names])) {
      const i = lower.indexOf(String(n).toLowerCase());
      if (i >= 0) return i;
    }
    return -1;
  }

  // Be flexible about capitalization; include Group
  const first = find(['Attendee First Name','Attendee first name']);
  const last  = find(['Attendee Last Name','Attendee last name']);
  const ptin  = find(['Attendee PTIN','attendee ptin','PTIN']);
  const email = find(['Email','email']);
  const valid = find(['Valid?','valid?']);
  const group = find(['Group','group']);

  if ([first,last,ptin,email,valid].some(i=>i<0)) return null; // group is optional but supported
  return {first, last, ptin, email, valid, group, hdr};
}

/**
 * Helper to map required columns in the Master sheet.
 */
function mapMasterHeaders_(sh) {
  if (!sh || sh.getLastRow() === 0) return null;
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(s => String(s || '').trim().toLowerCase());
  
  const map = {
    // Attempt to match common Master fields to Roster fields
    firstName: hdr.indexOf('attendee first name') >= 0 ? hdr.indexOf('attendee first name') : hdr.indexOf('first name'),
    lastName: hdr.indexOf('attendee last name') >= 0 ? hdr.indexOf('attendee last name') : hdr.indexOf('last name'),
    ptin: hdr.indexOf('attendee ptin') >= 0 ? hdr.indexOf('attendee ptin') : hdr.indexOf('ptin'),
    email: hdr.indexOf('email'),
    group: hdr.indexOf('group')
  };
  
  // Must have First Name, Last Name, and Email for Roster generation
  if (map.firstName < 0 || map.lastName < 0 || map.email < 0) return null;
  
  return map;
}

function getRosterMap_(ss){
  const sh = ss.getSheetByName(CFG.SHEET_ROSTER);
  if (!sh) return null;
  const map = mapRosterHeaders_(sh);
  if (!map) return null;
  const vals = sh.getDataRange().getValues();
  if (vals.length <= 1) return new Map();
  const m = new Map();
  for (let i=1;i<vals.length;i++){
    const row = vals[i];
    const ptin = formatPtinP0_(row[map.ptin]||'');
    const first = String(row[map.first]||'').trim();
    const last  = String(row[map.last]||'').trim();
    if (ptin) m.set(ptin, {first, last});
  }
  return m;
}

/**
 * Generates the Roster sheet by pulling unique entries from the Master sheet.
 * Deduplicates by Email (with PTIN as fallback key), keeping the latest entry.
 * Filters columns to match Roster's expected schema and ensures PTIN is formatted as P0....
 * The final list is sorted by Attendee First Name.
 */
function generateRosterFromMaster(quiet) {
  const ss = SpreadsheetApp.getActive();
  const masterSh = ss.getSheetByName(CFG.SHEET_MASTER);
  const rosterSh = ss.getSheetByName(CFG.SHEET_ROSTER);
  
  if (!masterSh || !rosterSh) {
    if (!quiet) toast_(`Sheets "${CFG.SHEET_MASTER}" and/or "${CFG.SHEET_ROSTER}" not found.`, true);
    return;
  }
  
  const masterMap = mapMasterHeaders_(masterSh);
  if (!masterMap) {
    if (!quiet) toast_('Master sheet is missing required columns (First Name, Last Name, Email).', true);
    return;
  }
  
  const rosterMap = mapRosterHeaders_(rosterSh);
  if (!rosterMap) {
     if (!quiet) toast_('Roster headers missing or renamed.', true); 
     return;
  }

  const masterVals = masterSh.getDataRange().getValues();
  if (masterVals.length <= 1) {
    // Clear Roster body if master is empty
    if (rosterSh.getLastRow() > 1) rosterSh.getRange(2, 1, rosterSh.getLastRow() - 1, rosterMap.hdr.length).clearContent();
    return;
  }
  
  const seen = new Map(); // key (email/ptin) -> [row data for Roster]

  for (let r = 1; r < masterVals.length; r++) {
    const row = masterVals[r];
    
    // Extract & Format required fields from Master row
    const email = String(row[masterMap.email] || '').toLowerCase().trim();
    // PTIN is formatted P0... here, satisfying the user requirement.
    let ptin = masterMap.ptin >= 0 ? formatPtinP0_(row[masterMap.ptin] || '') : ''; 
    const first = String(row[masterMap.firstName] || '').trim();
    const last  = String(row[masterMap.lastName] || '').trim();
    const group = masterMap.group >= 0 ? String(row[masterMap.group] || '').trim() : '';
    
    // Deduplication Key: Email primary, PTIN fallback (ensures unique users)
    const key = email || ptin;
    if (!key) continue;

    // Create an array that matches the Roster column order expected by mapRosterHeaders_
    const rosterRow = Array(rosterMap.hdr.length).fill('');
    
    // Map values to Roster sheet's column order: First, Last, PTIN, Email, Valid?, Group
    rosterRow[rosterMap.first] = first;
    rosterRow[rosterMap.last] = last;
    rosterRow[rosterMap.ptin] = ptin; // Formatted PTIN
    rosterRow[rosterMap.email] = email; 
    rosterRow[rosterMap.group] = group; 
    rosterRow[rosterMap.valid] = false; // Initial Roster entries are usually set to false/unreviewed
    
    // Latest row data overrides previous data
    seen.set(key, rosterRow);
  }

  let newRosterBody = Array.from(seen.values());
  
  // Sort the final roster alphabetically by Attendee First Name
  const firstNameIndex = rosterMap.first;
  newRosterBody.sort((a, b) => {
    const nameA = String(a[firstNameIndex] || '').toUpperCase();
    const nameB = String(b[firstNameIndex] || '').toUpperCase();
    if (nameA < nameB) return -1;
    if (nameA > nameB) return 1;
    return 0;
  });

  // Clear existing Roster body and write the new data
  const numCols = rosterMap.hdr.length;
  const rowsToClear = rosterSh.getLastRow() > 1 ? rosterSh.getLastRow() - 1 : 0;
  if (rowsToClear > 0) rosterSh.getRange(2, 1, rowsToClear, numCols).clearContent();
  
  if (newRosterBody.length) rosterSh.getRange(2, 1, newRosterBody.length, numCols).setValues(newRosterBody);

  if (!quiet) toast_(`Roster generated from Master: ${newRosterBody.length} unique entr${newRosterBody.length===1?'y':'ies'} by Email/PTIN.`);
}


/**
 * Deduplicate Roster by key (Email primary, PTIN fallback).
 * Latest row for a key **overrides** earlier rows.
 * The final output is sorted alphabetically by Attendee First Name.
 * Keeps headers as-is; rewrites the body with unique, sorted rows.
 */
function dedupeRosterByEmail(quiet){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG.SHEET_ROSTER);
  if (!sh) { if(!quiet) toast_(`Sheet "${CFG.SHEET_ROSTER}" not found.`, true); return; }
  const map = mapRosterHeaders_(sh);
  if (!map) { if(!quiet) toast_('Roster headers missing or renamed.', true); return; }

  const vals = sh.getDataRange().getValues();
  if (vals.length <= 1) return;

  const headers = map.hdr;
  const seen = new Map();               // key -> row array (latest wins)
  // Removed `order` array since we will sort, not maintain insertion order

  for (let r=1; r<vals.length; r++){
    const row = vals[r].slice();        // copy
    // Normalize PTIN in place for consistency
    if (map.ptin >= 0) row[map.ptin] = formatPtinP0_(row[map.ptin] || '');

    const email = String(row[map.email]||'').toLowerCase().trim();
    const ptin  = String(row[map.ptin]||'').trim();
    const key   = email || ptin;
    if (!key) continue;

    // Latest row overrides the stored one (Required behavior)
    seen.set(key, row);
  }

  // 1. Get the unique rows (latest version for each key)
  let deduped = Array.from(seen.values());
  
  // 2. Sort alphabetically by Attendee First Name (New requirement)
  const firstNameIndex = map.first;

  deduped.sort((a, b) => {
    // Ensure case-insensitive comparison
    const nameA = String(a[firstNameIndex] || '').toUpperCase();
    const nameB = String(b[firstNameIndex] || '').toUpperCase();
    
    if (nameA < nameB) return -1;
    if (nameA > nameB) return 1;
    return 0; // names are equal
  });

  // Clear body then write back with exact header width
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 1, lastRow-1, headers.length).clearContent();
  if (deduped.length) sh.getRange(2, 1, deduped.length, headers.length).setValues(deduped);

  if (!quiet) toast_(`Roster deduplicated and sorted by First Name: ${deduped.length} unique entr${deduped.length===1?'y':'ies'}.`);
}

function setRosterValidAndPtinForEmails_(emailToPtin){
  if (!emailToPtin || emailToPtin.size === 0) return;
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG.SHEET_ROSTER);
  if (!sh) return;
  const map = mapRosterHeaders_(sh);
  if (!map) return;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const rng = sh.getRange(2,1,lastRow-1, sh.getLastColumn());
  const vals = rng.getValues();
  let updated = 0;

  for (let i=0;i<vals.length;i++){
    const row = vals[i];
    const email = String(row[map.email]||'').toLowerCase().trim();
    if (!email) continue;

    if (emailToPtin.has(email)) {
      const ptin = formatPtinP0_(emailToPtin.get(email) || '');
      if (ptin && row[map.ptin] !== ptin) row[map.ptin] = ptin;
      if (!parseBool_(row[map.valid])) row[map.valid] = true;
      updated++;
    }
  }

  if (updated) rng.setValues(vals);
}

function updateRosterValidityFromIssues_() {
  const ss = SpreadsheetApp.getActive();
  const roster = ss.getSheetByName(CFG.SHEET_ROSTER);
  if (!roster) return;
  const rMap = mapRosterHeaders_(roster);
  if (!rMap) return;

  const issues = ss.getSheetByName(CFG.SHEET_ISSUES);
  const unresolvedPtins = new Set();
  const unresolvedEmails = new Set();
  if (issues) {
    const ivals = issues.getDataRange().getValues();
    if (ivals.length > 1) {
      const ih = ivals[0].map(s=>String(s||'').trim());
      const iP = ih.indexOf('Attendee PTIN');
      const iE = ih.indexOf('Email');
      const iFx = ih.indexOf('Fixed?');
      for (let r=1;r<ivals.length;r++){
        const row = ivals[r];
        if (!parseBool_(row[iFx])) {
          const p = formatPtinP0_(row[iP]||'');
          const e = String(row[iE]||'').toLowerCase().trim();
          if (p) unresolvedPtins.add(p);
          if (e) unresolvedEmails.add(e);
        }
      }
    }
  }

  const master = ss.getSheetByName(CFG.SHEET_MASTER);
  if (master) {
    const mVals = master.getDataRange().getValues();
    if (mVals.length > 1) {
      const mh = normalizeHeaderRow_(mVals[0]);
      const mMap = mapHeaders_(mh);
      for (let r=1;r<mVals.length;r++){
        const row = mVals[r];
        const iss = String(row[mMap.masterIssueCol]||'').trim();
        if (iss && iss.toLowerCase()!=='fixed') {
          const p = formatPtinP0_(row[mMap.ptin]||'');
          const e = String(row[mMap.email]||'').toLowerCase().trim();
          if (p) unresolvedPtins.add(p);
          if (e) unresolvedEmails.add(e);
        }
      }
    }
  }

  if (unresolvedPtins.size===0 && unresolvedEmails.size===0) return;
  const rng = roster.getRange(2,1, Math.max(roster.getLastRow()-1,0), roster.getLastColumn());
  if (rng.getNumRows() === 0) return;
  const vals = rng.getValues();
  let changed = 0;
  for (let i=0;i<vals.length;i++){
    const row = vals[i];
    const p = formatPtinP0_(row[rMap.ptin]||'');
    const e = String(row[rMap.email]||'').toLowerCase().trim();
    const hit = (p && unresolvedPtins.has(p)) || (e && unresolvedEmails.has(e));
    if (hit && parseBool_(row[rMap.valid])) { row[rMap.valid] = false; changed++; }
  }
  if (changed) rng.setValues(vals);
}

/** Backfill Master PTIN from Roster by Email **/
function backfillMasterPtinFromRoster_(quiet){
  const ss = SpreadsheetApp.getActive();
  const roster = ss.getSheetByName(CFG.SHEET_ROSTER);
  if (!roster) return;
  const rMap = mapRosterHeaders_(roster);
  if (!rMap) return;

  const master = ss.getSheetByName(CFG.SHEET_MASTER);
  if (!master) return;
  const mVals = master.getDataRange().getValues();
  if (mVals.length <= 1) return;

  const mHdr = normalizeHeaderRow_(mVals[0]);
  const mMap = mapHeaders_(mHdr);
  if (mMap.ptin==null || mMap.email==null) return;

  const rVals = roster.getDataRange().getValues();
  const emailToPtin = new Map();
  for (let i=1;i<rVals.length;i++){
    const row = rVals[i];
    const email = String(row[rMap.email]||'').toLowerCase().trim();
    const ptin  = formatPtinP0_(row[rMap.ptin]||'');
    if (email && ptin) emailToPtin.set(email, ptin);
  }

  const body = mVals.slice(1);
  let changes = 0;
  for (let i=0;i<body.length;i++){
    const row = body[i];
    const email = String(row[mMap.email]||'').toLowerCase().trim();
    const ptin  = formatPtinP0_(row[mMap.ptin]||'');
    if (email && !ptin && emailToPtin.has(email)) {
      row[mMap.ptin] = emailToPtin.get(email);
      changes++;
    }
  }
  if (changes) master.getRange(2,1,body.length,mVals[0].length).setValues(body);
}

/** onEdit: mark Roster Valid? TRUE pushes fixes to Master **/
function onEdit(e){
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (sh.getName() !== CFG.SHEET_ROSTER) return;
    const map = mapRosterHeaders_(sh);
    if (!map) return;

    if (e.range.getRow() >= 2 && e.range.getColumn() === (map.valid + 1)) {
      const newVal = e.value;
      if (parseBool_(newVal)) {
        const r = e.range.getRow();
        const vals = sh.getRange(r, 1, 1, sh.getLastColumn()).getValues()[0];

        const first = String(vals[map.first]||'').trim();
        const last  = String(vals[map.last]||'').trim();
        const ptin  = formatPtinP0_(vals[map.ptin]||'');
        const email = String(vals[map.email]||'').toLowerCase().trim();

        if (!email) { toast_('Roster row has no Email; cannot sync to Master.', true); return; }

        const master = mustGet_(SpreadsheetApp.getActive(), CFG.SHEET_MASTER);
        const mVals = master.getDataRange().getValues();
        if (mVals.length <= 1) return;
        const mHdr = normalizeHeaderRow_(mVals[0]);
        const mMap = mapHeaders_(mHdr);
        const body = mVals.slice(1);
        let updated = 0;
        for (let i=0;i<body.length;i++){
          const row = body[i];
          const em  = String(row[mMap.email]||'').toLowerCase().trim();
          if (em === email) {
            if (first) row[mMap.firstName] = first;
            if (last)  row[mMap.lastName]  = last;
            if (ptin)  row[mMap.ptin]      = ptin;
            row[mMap.masterIssueCol] = ''; // clear issue
            updated++;
          }
        }
        if (updated) {
          master.getRange(2,1,body.length,mVals[0].length).setValues(body);
          toast_(`Roster → Master: cleared Reporting Issue & synced ${updated} row(s) for ${email}.`);
        }
      }
    }
  } catch (err) {
    toast_('onEdit error: ' + err.message, true);
  }
}

/** Utility: Normalize PTIN format to P0####### uppercase (idempotent). */
function formatPtinP0_(ptinRaw) {
  let v = String(ptinRaw || '').trim().toUpperCase();
  if (!v) return '';
  // strip non-digits after the leading P if user pasted oddly
  v = v.replace(/^P0?(\d{0,7}).*$/, (_, d) => 'P0' + (d || '').padStart(7,'0')).replace(/[^P0\d]/g,'');
  if (!/^P0\d{7}$/.test(v)) {
    // fallback: try to build P0 + last 7 digits we can find
    const digits = (String(ptinRaw).match(/\d+/g) || []).join('');
    if (digits) v = 'P0' + digits.slice(-7).padStart(7,'0');
  }
  return v;
}
function onEdit(e){
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    const sheetName = sh.getName();

    // ---------------- ROSTER onEdit (unchanged) ----------------
    if (sheetName === CFG.SHEET_ROSTER) {
      const map = mapRosterHeaders_(sh);
      if (!map) return;

      if (map.valid >= 0 && e.range.getRow() >= 2 && e.range.getColumn() === (map.valid + 1)) {
        const newVal = e.value;
        if (parseBool_(newVal)) {
          const r = e.range.getRow();
          const vals = sh.getRange(r, 1, 1, sh.getLastColumn()).getValues()[0];
          const first = String(vals[map.first]||'').trim();
          const last  = String(vals[map.last]||'').trim();
          const ptin  = formatPtinP0_(vals[map.ptin]||'');
          const email = String(vals[map.email]||'').toLowerCase().trim();
          if (!email) { toast_('Roster row has no Email; cannot sync to Master.', true); return; }

          const master = mustGet_(SpreadsheetApp.getActive(), CFG.SHEET_MASTER);
          const mVals = master.getDataRange().getValues();
          if (mVals.length <= 1) return;
          const mHdr = normalizeHeaderRow_(mVals[0]);
          const mMap = mapHeaders_(mHdr);
          const body = mVals.slice(1);
          let updated = 0;
          for (let i=0;i<body.length;i++){
            const row = body[i];
            const em  = String(row[mMap.email]||'').toLowerCase().trim();
            if (em === email) {
              if (first) row[mMap.firstName] = first;
              if (last)  row[mMap.lastName]  = last;
              if (ptin)  row[mMap.ptin]      = ptin;
              row[mMap.masterIssueCol] = ''; // clear issue
              updated++;
            }
          }
          if (updated) {
            master.getRange(2,1,body.length,mVals[0].length).setValues(body);
            toast_(`Roster → Master: cleared Reporting Issue & synced ${updated} row(s) for ${email}.`);
          }
        }
      }
      return; // end Roster branch
    }

    // ---------------- MASTER onEdit: PTIN normalize + issue flag ----------------
    if (sheetName === CFG.SHEET_MASTER) {
      const mVals = sh.getRange(1,1,1, sh.getLastColumn()).getValues();
      if (mVals.length === 0) return;
      const hdr = normalizeHeaderRow_(mVals[0]);
      const mm = mapHeaders_(hdr);
      if (mm.ptin == null) return; // nothing to do

      // If a PTIN cell was edited, normalize and flag immediately for that row
      if (e.range.getRow() >= 2 && e.range.getColumn() === (mm.ptin + 1)) {
        const r = e.range.getRow();
        const rowVals = sh.getRange(r, 1, 1, sh.getLastColumn()).getValues()[0];

        let raw = String(rowVals[mm.ptin] || '').trim();
        if (!raw) return;

        if (/^pO/i.test(raw)) raw = 'P0' + raw.slice(2);      // PO -> P0
        const normalized = formatPtinP0_(raw);

        let dirty = false;
        if (normalized !== rowVals[mm.ptin]) {
          rowVals[mm.ptin] = normalized;
          dirty = true;
        }

        if (mm.masterIssueCol != null) {
          if (normalized === 'P00000000') {
            const curIssue = String(rowVals[mm.masterIssueCol] || '').trim();
            if (!/^fixed$/i.test(curIssue) && curIssue !== 'PTIN does not exist') {
              rowVals[mm.masterIssueCol] = 'PTIN does not exist';
              dirty = true;
            }
          }
        }

        if (dirty) sh.getRange(r, 1, 1, sh.getLastColumn()).setValues([rowVals]);
      }
    }
  } catch (err) {
    toast_('onEdit error: ' + err.message, true);
  }
}