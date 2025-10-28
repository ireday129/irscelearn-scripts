/** CONFIG / CONSTANTS (Placeholder for external configuration CFG object) **/
// NOTE: This file assumes an external 'config.js.gs' file defines the CFG object
// which includes sheet names and column headers. The logic below relies on the
// mapping provided in the user's latest 'config.js.gs'.

// ===========================================================================
// UTILITIES (Necessary helper functions for the rest of the script)
// ===========================================================================
function toast_(msg, isErr){ SpreadsheetApp.getActive().toast(msg, isErr?'Error':'IRS CE Tools', 5); }

function mustGet_(ss, name){
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`Sheet "${name}" not found.`);
  return sh;
}

function normalizeHeaderRow_(arr){ return arr.map(v=>String(v||'').replace(/\uFEFF/g,'').trim().replace(/\s+/g,' ')); }

function mapHeaders_(hdr){
  const lower = hdr.map(h=>h.toLowerCase());
  const find = (label)=> lower.indexOf(String(label||'').toLowerCase().trim().replace(/\s+/g,' '));
  return {
    firstName: find(CFG.COL_HEADERS.firstName),
    lastName:  find(CFG.COL_HEADERS.lastName),
    ptin:      find(CFG.COL_HEADERS.ptin),
    email:     find(CFG.COL_HEADERS.email),
    program:   find(CFG.COL_HEADERS.program),
    hours:     find(CFG.COL_HEADERS.hours),
    completion:find(CFG.COL_HEADERS.completion),
    group:     find(CFG.COL_HEADERS.group), // Now correctly maps to 'Group'
    masterIssueCol: find(CFG.COL_HEADERS.masterIssueCol),
    reportedCol:    find(CFG.COL_HEADERS.reportedCol),
    updatedCol:     find(CFG.COL_HEADERS.updatedCol),
    reportedAtCol:  find(CFG.COL_HEADERS.reportedAtCol)
  };
}

function parseBool_(val){
  if (typeof val==='boolean') return val;
  if (val==null) return false;
  const s = String(val).trim().toLowerCase();
  return s==='true'||s==='yes'||s==='y'||s==='1';
}

function formatPtinP0_(ptinRaw) {
  let v = String(ptinRaw || '').trim().toUpperCase();
  if (!v) return '';
  v = v.replace(/^P0?(\d{0,7}).*$/, (_, d) => 'P0' + (d || '').padStart(7,'0')).replace(/[^P0\d]/g,'');
  if (!/^P0\d{7}$/.test(v)) {
    const digits = (String(ptinRaw).match(/\d+/g) || []).join('');
    if (digits) v = 'P0' + digits.slice(-7).padStart(7,'0');
  }
  return v;
}

function normalizeProgram_(v) {
  return String(v || '').toUpperCase().replace(/\s+/g, '').trim();
}

/** Merge two rows, filling blank cells in 'kept' with non-blank values from 'incoming'. */
function mergeRowsFillBlanks_(kept, incoming) {
  const out = kept.slice();
  const n = Math.max(kept.length, incoming.length);
  for (let c = 0; c < n; c++) {
    const a = out[c];
    const b = incoming[c];
    if (isBlankCell_(a) && hasValue_(b)) out[c] = b;
  }
  return out;
}
function isBlankCell_(v) {
  if (v === null || v === undefined) return true;
  if (v instanceof Date) return false;
  if (typeof v === 'string') return v.trim() === '';
  return false;
}
function hasValue_(v) { return !isBlankCell_(v); }

// ===========================================================================
// MASTER DEDUPLICATION
// ===========================================================================
/** Deduplicate MASTER by Program Number + Email (older row wins, blanks filled). */
function dedupeMasterByEmailProgram(quiet) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG.SHEET_MASTER);
  if (!sh) { if (!quiet) toast_(`Sheet "${CFG.SHEET_MASTER}" not found.`, true); return; }

  const vals = sh.getDataRange().getValues();
  if (vals.length <= 1) return;

  const headersNorm = normalizeHeaderRow_(vals[0]);
  const mm = mapHeaders_(headersNorm);
  if (mm.program == null || mm.email == null) {
    if (!quiet) toast_('Master is missing Program Number and/or Email columns.', true);
    return;
  }

  const body = vals.slice(1);
  if (!body.length) return;

  const hasPTIN = mm.ptin != null;
  const keepMap = new Map();  // key -> keptRow (older)
  const order   = [];         // order of first appearances

  for (let i = 0; i < body.length; i++) {
    const row = body[i].slice();
    if (hasPTIN) row[mm.ptin] = formatPtinP0_(row[mm.ptin] || '');

    const email = String(row[mm.email] || '').toLowerCase().trim();
    const prog  = normalizeProgram_(row[mm.program] || '');
    const key = prog + '|' + email;

    if (!prog) continue;
    if (!keepMap.has(key)) { keepMap.set(key, row); order.push(key); }
    else keepMap.set(key, mergeRowsFillBlanks_(keepMap.get(key), row));
  }

  const deduped = order.map(k => keepMap.get(k));
  const cols = vals[0].length;
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow()-1, cols).clearContent();
  if (deduped.length) sh.getRange(2, 1, deduped.length, cols).setValues(deduped);

  if (!quiet) toast_(`Master deduped (older wins; blanks filled): ${deduped.length} row(s) remain.`);
}

// ===========================================================================
// ROSTER HELPERS
// ===========================================================================
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

  const first = find(['Attendee First Name','Attendee first name']);
  const last  = find(['Attendee Last Name','Attendee last name']);
  const ptin  = find(['Attendee PTIN','attendee ptin','PTIN']);
  const email = find(['Email','email']);
  const valid = find(['Valid?','valid?']);
  const group = find(['Group','group']);

  if ([first,last,ptin,email].some(i=>i<0)) return null; 
  return {first, last, ptin, email, valid, group, hdr};
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

/** Map required columns in the Master sheet. */
function mapMasterHeaders_(sh) {
  if (!sh || sh.getLastRow() === 0) return null;
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(s => String(s || '').trim().toLowerCase());
  const map = {
    firstName: hdr.indexOf('attendee first name') >= 0 ? hdr.indexOf('attendee first name') : hdr.indexOf('first name'),
    lastName:  hdr.indexOf('attendee last name')  >= 0 ? hdr.indexOf('attendee last name')  : hdr.indexOf('last name'),
    ptin:      hdr.indexOf('attendee ptin')       >= 0 ? hdr.indexOf('attendee ptin')       : hdr.indexOf('ptin'),
    email:     hdr.indexOf('email'),
    group:     hdr.indexOf('group'),
  };
  if (map.firstName < 0 || map.lastName < 0 || map.email < 0) return null;
  return map;
}

/** Map required columns in the Reported Hours sheet. */
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
    lastName:  find(['attendee last name']),
    ptin:      find(['ptin', 'attendee ptin']),
    program:   find(['program number']),
    hours:     find(['ce hours']),
    completion:find(['program completion date']),
    dateReported: find(['date reported'])
  };
  if (Object.values(map).some(i => i < 0)) return null;
  return map;
}

// ===========================================================================
// MAIN FUNCTIONALITY
// ===========================================================================
/** Build/refresh Roster from Master (latest overrides by key email/PTIN). */
function generateRosterFromMaster(quiet) {
  const ss = SpreadsheetApp.getActive();
  const masterSh = ss.getSheetByName(CFG.SHEET_MASTER);
  const rosterSh = ss.getSheetByName(CFG.SHEET_ROSTER);
  if (!masterSh || !rosterSh) { if (!quiet) toast_(`Sheets "${CFG.SHEET_MASTER}" and/or "${CFG.SHEET_ROSTER}" not found.`, true); return; }
  
  const masterMap = mapMasterHeaders_(masterSh);
  if (!masterMap) { if (!quiet) toast_('Master sheet is missing required columns (First Name, Last Name, Email).', true); return; }
  const rosterMap = mapRosterHeaders_(rosterSh);
  if (!rosterMap) { if (!quiet) toast_('Roster headers missing or renamed.', true); return; }

  const existingRosterVals = rosterSh.getDataRange().getValues().slice(1);
  const rosterDataMap = new Map(); // key (email/ptin) -> row array

  for (let r = 0; r < existingRosterVals.length; r++) {
    const row = existingRosterVals[r];
    const email = String(row[rosterMap.email] || '').toLowerCase().trim();
    const ptin = formatPtinP0_(row[rosterMap.ptin] || ''); 
    const key = email || ptin;
    if (key) rosterDataMap.set(key, row.slice());
  }

  const masterVals = masterSh.getDataRange().getValues();
  if (masterVals.length > 1) {
    for (let r = 1; r < masterVals.length; r++) {
      const row = masterVals[r];
      const email = String(row[masterMap.email] || '').toLowerCase().trim();
      const ptin  = masterMap.ptin >= 0 ? formatPtinP0_(row[masterMap.ptin] || '') : ''; 
      const first = String(row[masterMap.firstName] || '').trim();
      const last  = String(row[masterMap.lastName]  || '').trim();
      const rosterGroupValue = masterMap.group >= 0 ? String(row[masterMap.group] || '').trim() : ''; 
      const key = email || ptin;
      if (!key) continue;

      const masterRosterRow = Array(rosterMap.hdr.length).fill('');
      masterRosterRow[rosterMap.first] = first;
      masterRosterRow[rosterMap.last]  = last;
      masterRosterRow[rosterMap.ptin]  = ptin;
      masterRosterRow[rosterMap.email] = email; 
      masterRosterRow[rosterMap.group] = rosterGroupValue;

      rosterDataMap.set(key, masterRosterRow);
    }
  }
  
  let newRosterBody = Array.from(rosterDataMap.values());
  const firstNameIndex = rosterMap.first;
  newRosterBody.sort((a, b) => {
    const nameA = String(a[firstNameIndex] || '').toUpperCase();
    const nameB = String(b[firstNameIndex] || '').toUpperCase();
    if (nameA < nameB) return -1;
    if (nameA > nameB) return 1;
    return 0;
  });

  const numCols = rosterMap.hdr.length;
  const rowsToClear = rosterSh.getLastRow() > 1 ? rosterSh.getLastRow() - 1 : 0;
  if (rowsToClear > 0) rosterSh.getRange(2, 1, rowsToClear, numCols).clearContent();
  if (newRosterBody.length) rosterSh.getRange(2, 1, newRosterBody.length, numCols).setValues(newRosterBody);

  if (!quiet) toast_(`Roster generated from Master: ${newRosterBody.length} unique entr${newRosterBody.length===1?'y':'ies'} (manual rows preserved).`);
}

/** Update Master from Reported Hours (match PTIN+Program). */
function updateMasterFromReportedHours(quiet) {
  const ss = SpreadsheetApp.getActive();
  const masterSh = ss.getSheetByName(CFG.SHEET_MASTER);
  const reportedSh = ss.getSheetByName('Reported Hours');
  if (!masterSh || !reportedSh) { if (!quiet) toast_('Master or Reported Hours sheet not found.', true); return; }

  const reportedMap = mapReportedHoursHeaders_(reportedSh);
  if (!reportedMap) { if (!quiet) toast_('Reported Hours sheet is missing required columns.', true); return; }
  
  const masterVals = masterSh.getDataRange().getValues();
  if (masterVals.length <= 1) return;

  const mHdr = normalizeHeaderRow_(masterVals[0]);
  const mMap = mapHeaders_(mHdr); 
  if (mMap.ptin == null || mMap.program == null || mMap.hours == null || mMap.completion == null || mMap.reportedCol == null || mMap.reportedAtCol == null) {
    if (!quiet) toast_('Master sheet is missing required columns (PTIN, Program Number, CE Hours, Completion, Reported?, Date Reported).', true);
    return;
  }
  
  const reportedVals = reportedSh.getDataRange().getValues();
  const reportedHoursMap = new Map(); // PTIN|Program -> {hours, completion, dateReported}
  for (let r = 1; r < reportedVals.length; r++) {
    const row = reportedVals[r];
    const ptin = formatPtinP0_(row[reportedMap.ptin] || '');
    const program = normalizeProgram_(row[reportedMap.program] || '');
    if (ptin && program) {
      reportedHoursMap.set(`${ptin}|${program}`, {
        hours: row[reportedMap.hours],
        completion: row[reportedMap.completion],
        dateReported: row[reportedMap.dateReported]
      });
    }
  }

  if (reportedHoursMap.size === 0) { if (!quiet) toast_('Reported Hours sheet contains no valid PTIN/Program Number pairs.', false); return; }

  const body = masterVals.slice(1);
  let changes = 0;
  for (let i = 0; i < body.length; i++) {
    const row = body[i];
    const ptin = formatPtinP0_(row[mMap.ptin] || '');
    const program = normalizeProgram_(row[mMap.program] || '');
    if (!ptin || !program) continue;

    const key = `${ptin}|${program}`;
    if (reportedHoursMap.has(key)) {
      const reportedData = reportedHoursMap.get(key);
      row[mMap.hours]         = reportedData.hours;
      row[mMap.completion]    = reportedData.completion;
      row[mMap.reportedAtCol] = reportedData.dateReported;
      row[mMap.reportedCol]   = true;
      changes++;
    }
  }

  if (changes) {
    masterSh.getRange(2, 1, body.length, masterVals[0].length).setValues(body);
    if (!quiet) toast_(`Master sheet updated with reported hours: ${changes} row(s) marked as reported.`);
  } else {
    if (!quiet) toast_('No Master rows matched reported hours data.', false);
  }
}

/** Dedupe Roster by key (email primary, PTIN fallback). Latest wins. */
function dedupeRosterByEmail(quiet){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG.SHEET_ROSTER);
  if (!sh) { if(!quiet) toast_(`Sheet "${CFG.SHEET_ROSTER}" not found.`, true); return; }
  const map = mapRosterHeaders_(sh);
  if (!map) { if(!quiet) toast_('Roster headers missing or renamed.', true); return; }

  const vals = sh.getDataRange().getValues();
  if (vals.length <= 1) return;

  const headers = map.hdr;
  const seen = new Map();

  for (let r=1; r<vals.length; r++){
    const row = vals[r].slice();
    if (map.ptin >= 0) row[map.ptin] = formatPtinP0_(row[map.ptin] || '');

    const email = String(row[map.email]||'').toLowerCase().trim();
    const ptin  = String(row[map.ptin]||'').trim();
    const key   = email || ptin;
    if (!key) continue;

    seen.set(key, row);
  }

  let deduped = Array.from(seen.values());
  const firstNameIndex = map.first;
  deduped.sort((a, b) => {
    const nameA = String(a[firstNameIndex] || '').toUpperCase();
    const nameB = String(b[firstNameIndex] || '').toUpperCase();
    if (nameA < nameB) return -1;
    if (nameA > nameB) return 1;
    return 0;
  });

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
    if (hit && rMap.valid >= 0 && parseBool_(row[rMap.valid])) { row[rMap.valid] = false; changed++; }
  }
  if (changed) rng.setValues(vals);
}

/** Clear Master "Reporting Issue?" where Issues tab has Fixed?=TRUE. */
function clearMasterIssuesFromFixedIssues_(quiet) {
  const ss = SpreadsheetApp.getActive();
  const masterSh = ss.getSheetByName(CFG.SHEET_MASTER);
  const issuesSh = ss.getSheetByName(CFG.SHEET_ISSUES) || ss.getSheetByName('Reporting Issue') || ss.getSheetByName('Reporting Issues');

  if (!masterSh || !issuesSh) { if (!quiet) toast_('Master or Issues sheet not found.', true); return; }
  
  const issuesVals = issuesSh.getDataRange().getValues();
  if (issuesVals.length <= 1) return;

  const ih = issuesVals[0].map(s=>String(s||'').trim());
  const iP = ih.indexOf('Attendee PTIN');
  const iE = ih.indexOf('Email');
  const iFx = ih.indexOf('Fixed?');
  if (iP < 0 || iE < 0 || iFx < 0) { if (!quiet) toast_('Issues sheet is missing PTIN, Email, or Fixed? columns.', true); return; }
  
  const fixedKeys = new Set(); 
  for (let r = 1; r < issuesVals.length; r++) {
    const row = issuesVals[r];
    if (parseBool_(row[iFx])) {
      const p = formatPtinP0_(row[iP] || '');
      const e = String(row[iE] || '').toLowerCase().trim();
      if (e) fixedKeys.add(e);
      if (p) fixedKeys.add(p);
    }
  }
  if (fixedKeys.size === 0) return;
  
  const masterVals = masterSh.getDataRange().getValues();
  if (masterVals.length <= 1) return;

  const mHdr = normalizeHeaderRow_(masterVals[0]); 
  const mMap = mapHeaders_(mHdr); 
  if (mMap.ptin == null || mMap.email == null || mMap.masterIssueCol == null) { if (!quiet) toast_('Master sheet is missing PTIN, Email, or Reporting Issue? columns.', true); return; }
  
  const body = masterVals.slice(1);
  let changes = 0;
  for (let i = 0; i < body.length; i++) {
    const row = body[i];
    const email = String(row[mMap.email] || '').toLowerCase().trim();
    const ptin = formatPtinP0_(row[mMap.ptin] || '');
    const isFixed = (email && fixedKeys.has(email)) || (ptin && fixedKeys.has(ptin));
    const hasIssue = String(row[mMap.masterIssueCol] || '').trim() !== '';
    if (isFixed && hasIssue) { row[mMap.masterIssueCol] = ''; changes++; }
  }
  
  if (changes) {
    masterSh.getRange(2, 1, body.length, masterVals[0].length).setValues(body);
    if (!quiet) toast_(`Master issues cleared based on Fixed? status: ${changes} row(s) updated.`);
  } else {
     if (!quiet) toast_('No Master issues needed clearing based on Fixed? status.');
  }
}

// ===========================================================================
// BACKFILL FROM ROSTER
// ===========================================================================
/** Master PTIN from Roster (by Email). */
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

/** Master Group from Roster Group (by Email/PTIN). */
function backfillMasterGroupFromRoster_(quiet){
  const ss = SpreadsheetApp.getActive();
  const roster = ss.getSheetByName(CFG.SHEET_ROSTER);
  if (!roster) return;
  const rMap = mapRosterHeaders_(roster);
  if (!rMap || rMap.group == null) return; 

  const master = ss.getSheetByName(CFG.SHEET_MASTER);
  if (!master) return;
  const mVals = master.getDataRange().getValues();
  if (mVals.length <= 1) return;

  const mHdr = normalizeHeaderRow_(mVals[0]);
  let mMap = mapHeaders_(mHdr);
  if (mMap.group == null) {
    const groupIndex = mHdr.map(h => h.toLowerCase()).indexOf('group');
    if (groupIndex !== -1) mMap = { ...mMap, group: groupIndex };
  }
  if (mMap.ptin == null || mMap.email == null || mMap.group == null) {
     if (!quiet) toast_('Master sheet is missing PTIN, Email, or Group columns. Cannot backfill Group.', true);
     return;
  }

  const rVals = roster.getDataRange().getValues();
  const emailOrPtinToGroup = new Map();
  for (let i=1;i<rVals.length;i++){
    const row = rVals[i];
    const email = String(row[rMap.email]||'').toLowerCase().trim();
    const ptin  = formatPtinP0_(row[rMap.ptin]||'');
    const group = String(row[rMap.group]||'').trim();
    if (!group) continue; 
    const key = email || ptin;
    if (key) emailOrPtinToGroup.set(key, group);
  }

  const body = mVals.slice(1);
  let changes = 0;
  for (let i=0;i<body.length;i++){
    const row = body[i];
    const email = String(row[mMap.email]||'').toLowerCase().trim();
    const ptin  = formatPtinP0_(row[mMap.ptin]||'');
    const masterGroupValue  = String(row[mMap.group]||'').trim();
    const key = email || ptin;
    if (key && !masterGroupValue && emailOrPtinToGroup.has(key)) {
      row[mMap.group] = emailOrPtinToGroup.get(key);
      changes++;
    }
  }
  if (changes) {
    master.getRange(2,1,body.length,mVals[0].length).setValues(body);
    if (!quiet) toast_(`Master Group backfilled from Roster Group: ${changes} row(s) updated.`);
  }
}

/** Backfill Reported Hours (names fill blanks, EMAIL ALWAYS refreshed) from Roster by PTIN */
function backfillReportedHoursFromRoster_(quiet){
  const ss = SpreadsheetApp.getActive();
  const roster = ss.getSheetByName(CFG.SHEET_ROSTER);
  if (!roster) { if(!quiet) toast_('Roster sheet not found.', true); return; }

  const rMap = mapRosterHeaders_(roster);
  if (!rMap) { if(!quiet) toast_('Roster headers missing/renamed.', true); return; }

  const rVals = roster.getDataRange().getValues();
  if (rVals.length <= 1) return;

  const ptinToProfile = new Map();
  for (let i=1;i<rVals.length;i++){
    const row = rVals[i];
    const ptin  = formatPtinP0_(row[rMap.ptin]||'');
    if (!ptin) continue;
    const first = String(row[rMap.first]||'').trim();
    const last  = String(row[rMap.last]||'').trim();
    const email = String(row[rMap.email]||'').toLowerCase().trim();
    ptinToProfile.set(ptin, {first, last, email});
  }

  const rh = ss.getSheetByName('Reported Hours');
  if (!rh) { if(!quiet) toast_('Reported Hours sheet not found (skipping backfill).'); return; }

  const hVals = rh.getDataRange().getValues();
  if (hVals.length <= 1) return;

  const hdr = hVals[0].map(s=>String(s||'').trim());
  const lower = hdr.map(h=>h.toLowerCase());
  const idx = (label) => lower.indexOf(String(label||'').toLowerCase());
  const iPTIN  = idx('ptin');
  const iFName = idx('attendee first name');
  const iLName = idx('attendee last name');
  const iEmail = idx('email');
  if (iPTIN < 0) { if(!quiet) toast_('Reported Hours is missing "PTIN" column.', true); return; }

  const body = hVals.slice(1);
  let emailsUpdated = 0, nameFieldsFilled = 0;

  for (let r=0; r<body.length; r++){
    const row = body[r];
    const ptin = formatPtinP0_(row[iPTIN]||'');
    if (!ptin) continue;

    const prof = ptinToProfile.get(ptin);
    if (!prof) continue;

    // names: fill blanks only
    if (iFName >= 0 && isBlankCell_(row[iFName]) && hasValue_(prof.first)) { row[iFName] = prof.first; nameFieldsFilled++; }
    if (iLName >= 0 && isBlankCell_(row[iLName]) && hasValue_(prof.last))  { row[iLName] = prof.last;  nameFieldsFilled++; }

    // email: ALWAYS refresh from roster (if roster has a value)
    if (iEmail >= 0 && hasValue_(prof.email)) {
      const current = String(row[iEmail]||'').toLowerCase().trim();
      if (current !== prof.email) { row[iEmail] = prof.email; emailsUpdated++; }
    }
  }

  if (emailsUpdated || nameFieldsFilled) {
    rh.getRange(2, 1, body.length, hdr.length).setValues(body);
    if(!quiet) toast_(`Reported Hours backfill: ${emailsUpdated} email(s) refreshed, ${nameFieldsFilled} name field(s) filled.`);
  } else if(!quiet) {
    toast_('Reported Hours backfill: nothing to update.');
  }
}

/** Combined backfill from Roster to Master + Reported Hours. */
function backfillMasterFromRosterCombined(quiet) {
  try {
    backfillMasterPtinFromRoster_(quiet);
    backfillMasterGroupFromRoster_(quiet);
    backfillReportedHoursFromRoster_(quiet);
    if (!quiet) toast_('Backfill complete (Master PTIN, Master Group, Reported Hours).');
  } catch (e) {
    toast_('Error during Backfill: ' + e.message, true);
  }
}

/** onEdit: Roster Valid? TRUE pushes fixes to Master */
function onEdit(e){
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (sh.getName() !== CFG.SHEET_ROSTER) return;
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
  } catch (err) {
    toast_('onEdit error: ' + err.message, true);
  }
}