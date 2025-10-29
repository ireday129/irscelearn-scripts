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

  // Only First Name, Last Name, PTIN, and Email are mandatory for Roster headers. 
  // 'Valid?' and 'Group' are now treated as optional.
  if ([first,last,ptin,email].some(i=>i<0)) return null; 
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
    group: hdr.indexOf('group'),
    // ADDED: Find Source index (checking 'source' and 'program source')
    source: hdr.indexOf('source') >= 0 ? hdr.indexOf('source') : hdr.indexOf('program source')
  };
  
  // Must have First Name, Last Name, and Email for Roster generation
  if (map.firstName < 0 || map.lastName < 0 || map.email < 0) return null;
  
  return map;
}

/**
 * Helper to map required columns in the Reported Hours sheet.
 */
function mapReportedHoursHeaders_(sh) {
  if (!sh || sh.getLastRow() === 0) return null;
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(s => String(s || '').trim().toLowerCase());

  function find(names) {
    for (const n of (Array.isArray(names) ? names : [names])) {
      const i = hdr.indexOf(String(n).toLowerCase().trim());
      if (i >= 0) return i;
    }
    return -1;
  }
  
  const map = {
    firstName: find(['attendee first name']),
    lastName: find(['attendee last name']),
    ptin: find(['ptin', 'attendee ptin']),
    program: find(['program number']),
    hours: find(['ce hours']),
    completion: find(['program completion date']),
    dateReported: find(['date reported'])
  };
  
  // All columns are mandatory for reporting
  if (Object.values(map).some(i => i < 0)) return null;
  
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
 * **NEW LOGIC: Merges Master data into existing Roster data, preserving unique Roster rows.**
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

  // --- 1. Load existing Roster data to preserve manual entries ---
  const existingRosterVals = rosterSh.getDataRange().getValues().slice(1);
  const rosterDataMap = new Map(); // key (email/ptin) -> [row data for Roster]
  
  // Populate map with existing Roster rows
  for (let r = 0; r < existingRosterVals.length; r++) {
    const row = existingRosterVals[r];
    const email = String(row[rosterMap.email] || '').toLowerCase().trim();
    const ptin = formatPtinP0_(row[rosterMap.ptin] || ''); 
    const key = email || ptin;
    if (key) rosterDataMap.set(key, row.slice()); // Store a copy
  }

  // --- 2. Iterate Master and merge/update Roster data ---
  const masterVals = masterSh.getDataRange().getValues();
  if (masterVals.length > 1) {
      for (let r = 1; r < masterVals.length; r++) {
      const row = masterVals[r];
      
      // Extract & Format required fields from Master row
      const email = String(row[masterMap.email] || '').toLowerCase().trim();
      let ptin = masterMap.ptin >= 0 ? formatPtinP0_(row[masterMap.ptin] || '') : ''; 
      const first = String(row[masterMap.firstName] || '').trim();
      const last  = String(row[masterMap.lastName] || '').trim();
      const rosterGroupValue = masterMap.source >= 0 ? String(row[masterMap.source] || '').trim() : '';

      const key = email || ptin;
      if (!key) continue;

      // Create a new Roster row array from Master data
      const masterRosterRow = Array(rosterMap.hdr.length).fill('');
      masterRosterRow[rosterMap.first] = first;
      masterRosterRow[rosterMap.last] = last;
      masterRosterRow[rosterMap.ptin] = ptin; // Formatted PTIN
      masterRosterRow[rosterMap.email] = email; 
      masterRosterRow[rosterMap.group] = rosterGroupValue;

      // Update or add: latest Master row always overrides the current data for that key
      rosterDataMap.set(key, masterRosterRow);
    }
  }
  
  // --- 3. Finalize and Sort ---
  let newRosterBody = Array.from(rosterDataMap.values());
  
  // Sort the final roster alphabetically by Attendee First Name
  const firstNameIndex = rosterMap.first;
  newRosterBody.sort((a, b) => {
    const nameA = String(a[firstNameIndex] || '').toUpperCase();
    const nameB = String(b[firstNameIndex] || '').toUpperCase();
    if (nameA < nameB) return -1;
    if (nameA > nameB) return 1;
    return 0;
  });

  // --- 4. Write back merged data ---
  const numCols = rosterMap.hdr.length;
  const rowsToClear = rosterSh.getLastRow() > 1 ? rosterSh.getLastRow() - 1 : 0;
  if (rowsToClear > 0) rosterSh.getRange(2, 1, rowsToClear, numCols).clearContent();
  
  if (newRosterBody.length) rosterSh.getRange(2, 1, newRosterBody.length, numCols).setValues(newRosterBody);

  if (!quiet) toast_(`Roster generated from Master: ${newRosterBody.length} unique entr${newRosterBody.length===1?'y':'ies'} (manual rows preserved).`);
}

/**
 * Updates the Master sheet with CE reporting data from the Reported Hours sheet.
 * Matches rows using PTIN + Program Number. Updates CE hours, completion date,
 * and sets 'Reported?' to TRUE.
 */
function updateMasterFromReportedHours(quiet) {
  const ss = SpreadsheetApp.getActive();
  const masterSh = ss.getSheetByName(CFG.SHEET_MASTER);
  const reportedSh = ss.getSheetByName('Reported Hours'); // Use literal name based on request

  if (!masterSh || !reportedSh) {
    if (!quiet) toast_('Master or Reported Hours sheet not found.', true);
    return;
  }

  const reportedMap = mapReportedHoursHeaders_(reportedSh);
  if (!reportedMap) {
    if (!quiet) toast_('Reported Hours sheet is missing required columns.', true);
    return;
  }
  
  const masterVals = masterSh.getDataRange().getValues();
  if (masterVals.length <= 1) return;

  const mHdr = normalizeHeaderRow_(masterVals[0]);
  const mMap = mapHeaders_(mHdr); // Assumes CFG.COL_HEADERS for reported/hours/completion is available

  // Ensure Master has required columns for matching and updating
  if (mMap.ptin == null || mMap.program == null || mMap.hours == null || mMap.completion == null || mMap.reportedCol == null || mMap.reportedAtCol == null) {
    if (!quiet) toast_('Master sheet is missing required columns (PTIN, Program Number, CE Hours, Completion, Reported?, Date Reported).', true);
    return;
  }
  
  const reportedVals = reportedSh.getDataRange().getValues();
  const reportedHoursMap = new Map(); // Key: PTIN|ProgramNumber -> {hours, completion, dateReported}

  // 1. Build map of reported data for fast lookup
  for (let r = 1; r < reportedVals.length; r++) {
    const row = reportedVals[r];
    const ptin = formatPtinP0_(row[reportedMap.ptin] || '');
    const program = normalizeProgram_(row[reportedMap.program] || '');
    
    if (ptin && program) {
      const key = `${ptin}|${program}`;
      reportedHoursMap.set(key, {
        hours: row[reportedMap.hours],
        completion: row[reportedMap.completion],
        dateReported: row[reportedMap.dateReported]
      });
    }
  }

  if (reportedHoursMap.size === 0) {
    if (!quiet) toast_('Reported Hours sheet contains no valid PTIN/Program Number pairs.', false);
    return;
  }

  // 2. Iterate Master sheet and apply updates
  const body = masterVals.slice(1);
  let changes = 0;

  for (let i = 0; i < body.length; i++) {
    const row = body[i];
    const ptin = formatPtinP0_(row[mMap.ptin] || '');
    const program = normalizeProgram_(row[mMap.program] || '');
    
    if (ptin && program) {
      const key = `${ptin}|${program}`;
      if (reportedHoursMap.has(key)) {
        const reportedData = reportedHoursMap.get(key);
        
        // Update data fields
        row[mMap.hours] = reportedData.hours;
        row[mMap.completion] = reportedData.completion;
        row[mMap.reportedAtCol] = reportedData.dateReported;

        // Set Reported? column to TRUE
        row[mMap.reportedCol] = true;
        changes++;
      }
    }
  }

  // 3. Write back updated Master data
  if (changes) {
    masterSh.getRange(2, 1, body.length, masterVals[0].length).setValues(body);
    if (!quiet) toast_(`Master sheet updated with reported hours: ${changes} row(s) marked as reported.`);
  } else {
    if (!quiet) toast_('No Master rows matched reported hours data.', false);
  }
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
      // Guard: Only update 'Valid?' if the column exists (map.valid >= 0)
      if (map.valid >= 0 && !parseBool_(row[map.valid])) row[map.valid] = true;
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

  // The logic below relies on global CFG values and helper functions defined outside this block.
  // We'll proceed assuming they are available in the full script environment.
  const issues = ss.getSheetByName(CFG.SHEET_ISSUES) || ss.getSheetByName('Reporting Issue') || ss.getSheetByName('Reporting Issues');
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
    // Guard: Only set 'Valid?' to false if the column exists
    if (hit && rMap.valid >= 0 && parseBool_(row[rMap.valid])) { row[rMap.valid] = false; changed++; }
  }
  if (changed) rng.setValues(vals);
}

/**
 * Clears the "Reporting Issue?" column on the Master sheet for any rows
 * corresponding to issues marked as "Fixed?" on the Issues sheet.
 */
function clearMasterIssuesFromFixedIssues_(quiet) {
  const ss = SpreadsheetApp.getActive();
  const masterSh = ss.getSheetByName(CFG.SHEET_MASTER);
  // Assuming CFG.SHEET_ISSUES is defined and using fallbacks for safety
  const issuesSh = ss.getSheetByName(CFG.SHEET_ISSUES) || ss.getSheetByName('Reporting Issue') || ss.getSheetByName('Reporting Issues');

  if (!masterSh || !issuesSh) {
    if (!quiet) toast_('Master or Issues sheet not found.', true);
    return;
  }
  
  const issuesVals = issuesSh.getDataRange().getValues();
  if (issuesVals.length <= 1) return;

  // Header mapping for Issues Sheet (must assume standard headers)
  const ih = issuesVals[0].map(s=>String(s||'').trim());
  const iP = ih.indexOf('Attendee PTIN');
  const iE = ih.indexOf('Email');
  const iFx = ih.indexOf('Fixed?');

  if (iP < 0 || iE < 0 || iFx < 0) {
    if (!quiet) toast_('Issues sheet is missing PTIN, Email, or Fixed? columns.', true);
    return;
  }
  
  const fixedKeys = new Set(); // Stores key (Email or PTIN) of fixed issues

  for (let r = 1; r < issuesVals.length; r++) {
    const row = issuesVals[r];
    if (parseBool_(row[iFx])) {
      // Issue is marked as fixed
      const p = formatPtinP0_(row[iP] || '');
      const e = String(row[iE] || '').toLowerCase().trim();
      
      // Store both formatted PTIN and lowercased email as potential keys
      if (e) fixedKeys.add(e);
      if (p) fixedKeys.add(p);
    }
  }

  if (fixedKeys.size === 0) return;
  
  const masterVals = masterSh.getDataRange().getValues();
  if (masterVals.length <= 1) return;

  const mHdr = normalizeHeaderRow_(masterVals[0]); // normalizeHeaderRow_ is assumed to be in UTILITIES
  const mMap = mapHeaders_(mHdr); // mapHeaders_ is assumed to be in UTILITIES

  if (mMap.ptin == null || mMap.email == null || mMap.masterIssueCol == null) {
    if (!quiet) toast_('Master sheet is missing PTIN, Email, or Reporting Issue? columns.', true);
    return;
  }
  
  const body = masterVals.slice(1);
  let changes = 0;
  
  for (let i = 0; i < body.length; i++) {
    const row = body[i];
    const email = String(row[mMap.email] || '').toLowerCase().trim();
    const ptin = formatPtinP0_(row[mMap.ptin] || '');
    
    // Check if this Master row corresponds to a fixed issue key (by email or PTIN)
    const isFixed = (email && fixedKeys.has(email)) || (ptin && fixedKeys.has(ptin));
    
    // Check if the issue column currently has content (is not blank)
    const hasIssue = String(row[mMap.masterIssueCol] || '').trim() !== '';

    if (isFixed && hasIssue) {
      // Clear the Reporting Issue column
      row[mMap.masterIssueCol] = '';
      changes++;
    }
  }
  
  if (changes) {
    masterSh.getRange(2, 1, body.length, masterVals[0].length).setValues(body);
    if (!quiet) toast_(`Master issues cleared based on Fixed? status: ${changes} row(s) updated.`);
  } else {
     if (!quiet) toast_('No Master issues needed clearing based on Fixed? status.');
  }
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

/**
 * Backfill Roster from Master, but ONLY from Master rows where Reported? === TRUE.
 * - Match by Email first, then by PTIN as fallback.
 * - Fill blanks only on the Roster row (do not overwrite non-empty cells).
 * - Does NOT append new rows to Roster; it only updates existing ones.
 * - Fields backfilled: First Name, Last Name, PTIN, Group (if present on Roster), Email (if blank).
 */
function backfillRosterFromMasterReported_(quiet) {
  const ss = SpreadsheetApp.getActive();
  const rosterSh = ss.getSheetByName(CFG.SHEET_ROSTER);
  const masterSh = ss.getSheetByName(CFG.SHEET_MASTER);

  if (!rosterSh || !masterSh) { if(!quiet) toast_('Roster or Master sheet not found.', true); return; }

  // Map headers
  const rMap = mapRosterHeaders_(rosterSh);
  if (!rMap) { if(!quiet) toast_('Roster headers missing or renamed.', true); return; }

  const mVals = masterSh.getDataRange().getValues();
  if (mVals.length <= 1) { if(!quiet) toast_('Master is empty; nothing to backfill.', false); return; }

  const mHdr = normalizeHeaderRow_(mVals[0]);
  const mMap = mapHeaders_(mHdr);

  // Ensure Master has the columns we need
  if (mMap.email == null || mMap.program == null || mMap.reportedCol == null) {
    if (!quiet) toast_('Master is missing Email/Program/Reported? columns.', true);
    return;
  }

  // Build Roster index by Email and PTIN
  const rLastRow = rosterSh.getLastRow();
  if (rLastRow < 2) { if(!quiet) toast_('Roster has no data rows to backfill.', false); return; }

  const rRange = rosterSh.getRange(2, 1, rLastRow - 1, rosterSh.getLastColumn());
  const rVals = rRange.getValues();

  const byEmail = new Map(); // email -> rowIndex in rVals
  const byPtin  = new Map(); // ptin  -> rowIndex in rVals
  for (let i = 0; i < rVals.length; i++) {
    const row = rVals[i];
    const email = rMap.email >= 0 ? String(row[rMap.email] || '').toLowerCase().trim() : '';
    const ptin  = rMap.ptin  >= 0 ? formatPtinP0_(row[rMap.ptin]  || '') : '';
    if (email) byEmail.set(email, i);
    if (ptin)  byPtin.set(ptin, i);
  }

  let updates = 0;

  // Walk Master (skip header)
  for (let r = 1; r < mVals.length; r++) {
    const mrow = mVals[r];

    // Only use rows that are already Reported? === TRUE
    const isReported = parseBool_(mrow[mMap.reportedCol]);
    if (!isReported) continue;

    const mEmail = mMap.email != null ? String(mrow[mMap.email] || '').toLowerCase().trim() : '';
    const mPtin  = mMap.ptin  != null ? formatPtinP0_(mrow[mMap.ptin] || '') : '';
    if (!mEmail && !mPtin) continue; // need at least one key

    // Locate roster row by email, then PTIN
    let rIdx = -1;
    if (mEmail && byEmail.has(mEmail)) rIdx = byEmail.get(mEmail);
    else if (mPtin && byPtin.has(mPtin)) rIdx = byPtin.get(mPtin);

    if (rIdx < 0) continue; // do not append new rows; update only

    const rRow = rVals[rIdx];

    // Fill blanks only
    if (rMap.first >= 0 && isBlankCell_(rRow[rMap.first]) && mMap.firstName != null) {
      const v = String(mrow[mMap.firstName] || '').trim();
      if (v) { rRow[rMap.first] = v; updates++; }
    }
    if (rMap.last >= 0 && isBlankCell_(rRow[rMap.last]) && mMap.lastName != null) {
      const v = String(mrow[mMap.lastName] || '').trim();
      if (v) { rRow[rMap.last] = v; updates++; }
    }
    if (rMap.ptin >= 0 && isBlankCell_(rRow[rMap.ptin]) && mMap.ptin != null) {
      const v = formatPtinP0_(mrow[mMap.ptin] || '');
      if (v) { rRow[rMap.ptin] = v; updates++; }
    }
    if (rMap.email >= 0 && isBlankCell_(rRow[rMap.email]) && mMap.email != null) {
      const v = String(mrow[mMap.email] || '').toLowerCase().trim();
      if (v) { rRow[rMap.email] = v; updates++; }
    }

    // If Roster has a Group column and it's blank, try to fill from Master 'group' or 'source'
    if (rMap.group >= 0 && isBlankCell_(rRow[rMap.group])) {
      let g = '';
      if (mMap.group != null) g = String(mrow[mMap.group] || '').trim();
      else if (typeof mMap.source !== 'undefined' && mMap.source != null) g = String(mrow[mMap.source] || '').trim();
      if (g) { rRow[rMap.group] = g; updates++; }
    }
  }

  if (updates) {
    rRange.setValues(rVals);
    if (!quiet) toast_(`Roster backfilled from Master (Reported?=TRUE only): ${updates} cell update${updates===1?'':'s'}.`);
  } else if (!quiet) {
    toast_('No Roster cells needed backfilling from Reported Master rows.');
  }
}

/** onEdit: mark Roster Valid? TRUE pushes fixes to Master **/
function onEdit(e){
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (sh.getName() !== CFG.SHEET_ROSTER) return;
    const map = mapRosterHeaders_(sh);
    if (!map) return;

    // Guard: Only check for an edit in the 'Valid?' column if it exists (map.valid >= 0)
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
